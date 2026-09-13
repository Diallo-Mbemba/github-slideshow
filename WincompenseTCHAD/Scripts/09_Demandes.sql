/*
    =========================================================================
    Double regard sur le référentiel des points de vente
    =========================================================================

    Toute écriture sur les sous-agents, les agences propres et les groupes statistiques passe
    désormais par deux personnes : un INPUTER la saisit, un AUTHORIZER l'autorise. Tant qu'elle
    n'est pas autorisée, elle n'existe pas pour l'application.

    POURQUOI UNE FILE D'ATTENTE SÉPARÉE, ET NON UNE COLONNE « Statut » SUR LES VRAIES TABLES

    Les colonnes de T_Pdv_SA sont lues par la comptabilisation quotidienne ET par d'autres
    applications. Une ligne non autorisée qui séjournerait dans T_Pdv_SA serait vue par ces
    applications, qui n'ont aucune raison de connaître le nouveau statut : le contrôle serait
    contourné sans que personne n'y touche.

    Avec une file séparée, T_Pdv_SA, T_Pdv_EC et T_GroupeStatistique ne contiennent QUE de la
    donnée autorisée. Aucun lecteur, interne ou externe, n'a à être modifié.

    UNE SEULE TABLE POUR LES TROIS OBJETS

    Les trois objets partagent presque tous leurs champs : une clé, une désignation, un groupe,
    un taux, deux comptes et un code de rattachement. Trois tables miroir tripleraient le code
    pour la même garantie. Les vues V_Demande_* rendent les noms de colonnes de chaque table
    cible, pour qu'un auditeur lise « CompteCompense » et non un nom générique.

    À exécuter APRÈS 01 à 08, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. Fonction de l'utilisateur dans le dispositif
--
--    Le rôle dit le domaine (COMMERCIAL, COMPENSE, ADMIN), la fonction dit le pouvoir.
--    NULL = aucune : l'utilisateur ne peut ni saisir ni autoriser les points de vente.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns
               WHERE object_id = OBJECT_ID(N'T_UtilisateurWU') AND name = N'Fonction')
BEGIN
    ALTER TABLE dbo.T_UtilisateurWU ADD Fonction NVARCHAR(20) NULL;
    PRINT 'Colonne Fonction ajoutée à T_UtilisateurWU.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_UtilisateurWU_Fonction')
BEGIN
    ALTER TABLE dbo.T_UtilisateurWU WITH CHECK
        ADD CONSTRAINT CK_T_UtilisateurWU_Fonction
        CHECK (Fonction IS NULL OR Fonction IN ('INPUTER', 'AUTHORIZER'));
    PRINT 'Contrainte CK_T_UtilisateurWU_Fonction posée.';
END
GO

-- =========================================================================
-- 2. La file des demandes
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_DemandeWU')
BEGIN
    CREATE TABLE dbo.T_DemandeWU
    (
        IdDemande       BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,

        -- SOUS_AGENT (T_Pdv_SA), AGENCE (T_Pdv_EC), GROUPE (T_GroupeStatistique)
        TypeObjet       NVARCHAR(20)    NOT NULL,

        -- CREATION, MODIFICATION, SUPPRESSION, SYNCHRONISATION
        Operation       NVARCHAR(20)    NOT NULL,

        -- EN_ATTENTE, AUTORISE, REJETE
        Statut          NVARCHAR(20)    NOT NULL DEFAULT ('EN_ATTENTE'),

        /*
            Valeurs proposées. Correspondance avec les tables cibles :

              Colonne             Sous-agent            Agence propre           Groupe
              ------------------  --------------------  ----------------------  -----------------
              Cle                 Code_Pdv              Codesite                Groupe
              Designation         Designationagence     Designationagence       (inutilisée)
              GroupeStatistique   GroupeStatistique     (inutilisée)            (inutilisée)
              CompteActivite      CompteCompense        (inutilisée)            CompteActivite
              CompteCommission    CompteCommission      (inutilisée)            CompteCommission
              Taux                Taux                  (inutilisée)            Taux
              CodeRattachement    codeagence            [CodeAgenc-Voyager]     (inutilisée)
        */
        Cle                 NVARCHAR(255)   NOT NULL,
        Designation         NVARCHAR(255)   NULL,
        GroupeStatistique   NVARCHAR(255)   NULL,
        CompteActivite      NVARCHAR(255)   NULL,
        CompteCommission    NVARCHAR(255)   NULL,
        Taux                DECIMAL(4,2)    NULL,
        CodeRattachement    NVARCHAR(255)   NULL,

        -- Qui a saisi, qui a décidé.
        SaisiPar        NVARCHAR(50)    NOT NULL,
        DateSaisie      DATETIME        NOT NULL DEFAULT (GETDATE()),
        DecidePar       NVARCHAR(50)    NULL,
        DateDecision    DATETIME        NULL,
        MotifRejet      NVARCHAR(500)   NULL,

        CONSTRAINT CK_T_DemandeWU_TypeObjet
            CHECK (TypeObjet IN ('SOUS_AGENT', 'AGENCE', 'GROUPE')),

        CONSTRAINT CK_T_DemandeWU_Operation
            CHECK (Operation IN ('CREATION', 'MODIFICATION', 'SUPPRESSION', 'SYNCHRONISATION')),

        CONSTRAINT CK_T_DemandeWU_Statut
            CHECK (Statut IN ('EN_ATTENTE', 'AUTORISE', 'REJETE')),

        /*
            LE CŒUR DU DISPOSITIF, POSÉ DANS LA BASE ELLE-MÊME.

            Personne ne décide de sa propre saisie. La règle est déjà appliquée par
            l'application ; elle est redoublée ici pour qu'une écriture faite hors de
            l'application — un UPDATE à la main dans Management Studio — ne puisse pas
            la contourner.
        */
        CONSTRAINT CK_T_DemandeWU_PasSoiMeme
            CHECK (DecidePar IS NULL OR DecidePar <> SaisiPar)
    );

    -- Une seule demande en attente à la fois par objet : deux demandes contradictoires sur le
    -- même sous-agent s'appliqueraient sinon dans l'ordre où l'authorizer les traite.
    CREATE UNIQUE INDEX UX_T_DemandeWU_EnAttente
        ON dbo.T_DemandeWU (TypeObjet, Cle)
        WHERE Statut = 'EN_ATTENTE';

    CREATE INDEX IX_T_DemandeWU_Statut ON dbo.T_DemandeWU (Statut, DateSaisie);

    PRINT 'Table T_DemandeWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_DemandeWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 3. Vues d'audit : les mêmes demandes, aux noms de colonnes de chaque table cible
-- =========================================================================
IF OBJECT_ID(N'dbo.V_Demande_SousAgent', N'V') IS NOT NULL DROP VIEW dbo.V_Demande_SousAgent;
GO
CREATE VIEW dbo.V_Demande_SousAgent
AS
    SELECT  IdDemande, Operation, Statut,
            Code_Pdv          = Cle,
            Designationagence = Designation,
            GroupeStatistique = GroupeStatistique,
            CompteCompense    = CompteActivite,
            CompteCommission  = CompteCommission,
            Taux              = Taux,
            codeagence        = CodeRattachement,
            SaisiPar, DateSaisie, DecidePar, DateDecision, MotifRejet
    FROM    dbo.T_DemandeWU
    WHERE   TypeObjet = 'SOUS_AGENT';
GO

IF OBJECT_ID(N'dbo.V_Demande_Agence', N'V') IS NOT NULL DROP VIEW dbo.V_Demande_Agence;
GO
CREATE VIEW dbo.V_Demande_Agence
AS
    SELECT  IdDemande, Operation, Statut,
            Codesite               = Cle,
            Designationagence      = Designation,
            [CodeAgenc-Voyager]    = CodeRattachement,
            SaisiPar, DateSaisie, DecidePar, DateDecision, MotifRejet
    FROM    dbo.T_DemandeWU
    WHERE   TypeObjet = 'AGENCE';
GO

IF OBJECT_ID(N'dbo.V_Demande_Groupe', N'V') IS NOT NULL DROP VIEW dbo.V_Demande_Groupe;
GO
CREATE VIEW dbo.V_Demande_Groupe
AS
    SELECT  IdDemande, Operation, Statut,
            Groupe            = Cle,
            CompteActivite    = CompteActivite,
            CompteCommission  = CompteCommission,
            Taux              = Taux,
            SaisiPar, DateSaisie, DecidePar, DateDecision, MotifRejet
    FROM    dbo.T_DemandeWU
    WHERE   TypeObjet = 'GROUPE';
GO

/*
    =========================================================================
    4. APRÈS L'EXÉCUTION
    =========================================================================

    Les données déjà présentes dans T_Pdv_SA, T_Pdv_EC et T_GroupeStatistique sont réputées
    autorisées : elles précèdent le dispositif, il n'y a pas de validation rétroactive.

    Reste à attribuer les fonctions. Aucun utilisateur n'en a au départ — pas même
    l'administrateur — donc PERSONNE ne peut créer de sous-agent tant que ce n'est pas fait.
    C'est voulu : la première décision à prendre est qui saisit et qui autorise.

    Cela se fait dans l'application, écran « Utilisateurs et connexions ». Il faut AU MOINS un
    inputer et AU MOINS un authorizer, et ce ne peut pas être la même personne.

        UPDATE dbo.T_UtilisateurWU SET Fonction = 'INPUTER'    WHERE Identifiant = N'...';
        UPDATE dbo.T_UtilisateurWU SET Fonction = 'AUTHORIZER' WHERE Identifiant = N'...';

    Prévoyez PLUSIEURS authorizers : avec un seul, une semaine d'absence bloque toute création
    de point de vente.
*/

-- Contrôle final.
SELECT Identifiant, NomComplet, Role, Fonction, Actif
FROM   dbo.T_UtilisateurWU
ORDER BY Role, Identifiant;
GO
