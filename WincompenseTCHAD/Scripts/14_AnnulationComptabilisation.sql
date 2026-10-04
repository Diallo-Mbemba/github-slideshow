/*
    =========================================================================
    Annulation d'une comptabilisation — archives et file de décision
    =========================================================================

    POURQUOI ANNULER N'EST PAS SUPPRIMER

    Une journée peut avoir été comptabilisée à tort : rapport Western Union vide, rapport
    d'une autre journée, double chargement. Il faut pouvoir la retirer des rapports.

    Un DELETE le ferait — et effacerait du même coup la preuve. Or c'est précisément quand
    une journée est annulée que l'auditeur veut savoir qui, quand, et pourquoi. Les lignes
    sont donc DÉPLACÉES vers les tables ci-dessous, avec le motif et les deux signatures.
    Rien n'est perdu ; rien ne gêne.

    POURQUOI DES TABLES SÉPARÉES, ET NON UNE COLONNE « Annulee » SUR LES TABLES VIVANTES

    T_HistoriqueWU a pour clé primaire (DateActivite, Account), et T_PieceWU
    (DateActivite, Ligne). Une journée annulée qui resterait sur place entrerait en
    collision avec la journée recomptabilisée qui la remplace — et il faudrait alors
    ajouter l'annulation à chaque clé primaire, donc à chaque requête de lecture, donc
    aux rapports d'activité, qui n'ont aucune raison de connaître ce nouveau statut.

    Avec des tables d'archive, T_HistoriqueWU, T_HistoriqueMTCN et T_PieceWU ne contiennent
    QUE des journées vivantes. Aucun lecteur existant n'a à être modifié, et les rapports
    d'activité cessent de compter la journée annulée sans qu'on y touche.

    UNE JOURNÉE PEUT ÊTRE ANNULÉE PLUSIEURS FOIS

    Comptabilisée, annulée, recomptabilisée, annulée de nouveau : la même date d'activité
    peut donc figurer plusieurs fois dans les archives. C'est IdAnnulation, et non la date,
    qui identifie une archive.

    LE DOUBLE REGARD PASSE PAR LA FILE EXISTANTE

    L'annulation n'a pas sa propre file : elle emprunte T_DemandeWU, avec
    TypeObjet = 'COMPTABILISATION' et Operation = 'ANNULATION'. L'authorizer garde un seul
    écran et une seule habitude, la contrainte « personne ne décide de sa propre saisie »
    s'applique telle quelle, et l'index unique sur (TypeObjet, Cle) empêche deux demandes
    d'annulation simultanées sur la même journée.

    À exécuter APRÈS 01 à 13, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. L'en-tête de l'annulation : une ligne par annulation
-- =========================================================================
IF OBJECT_ID(N'dbo.T_AnnulationWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_AnnulationWU
    (
        IdAnnulation        BIGINT IDENTITY(1,1) NOT NULL,

        DateActivite        DATE            NOT NULL,   -- Journée retirée

        -- RAPPORT_VIDE, RAPPORT_ERRONE, MAUVAISE_JOURNEE, DOUBLON, AUTRE
        Motif               NVARCHAR(30)    NOT NULL,
        Commentaire         NVARCHAR(500)   NULL,

        /*
            Le fichier destiné au core banking a-t-il déjà été injecté ?

            L'application ne peut pas le savoir seule : elle sait qu'elle a PRODUIT le
            fichier (T_FichierCoreBankingWU), pas qu'il a été chargé dans les livres de la
            banque. La réponse est donc celle de l'agent, et elle est conservée : si elle
            est « oui », l'écriture doit être extournée en comptabilité, et Wincompense ne
            le fait pas.
        */
        CoreBankingInjecte  BIT             NOT NULL DEFAULT (0),

        -- Ce qui a été déplacé : de quoi afficher l'archive sans la rouvrir.
        NombrePdv           INT             NOT NULL DEFAULT (0),
        NombreTransactions  INT             NOT NULL DEFAULT (0),
        NombreEcritures     INT             NOT NULL DEFAULT (0),
        TotalDebit          BIGINT          NOT NULL DEFAULT (0),
        TotalCredit         BIGINT          NOT NULL DEFAULT (0),

        -- Les deux signatures. DemandeePar ne peut pas valoir AutoriseePar : c'est tout
        -- l'objet du double regard, et la contrainte le redit ici, dans la base.
        IdDemande           BIGINT          NULL,
        DemandeePar         NVARCHAR(50)    NOT NULL,
        DateDemande         DATETIME        NOT NULL DEFAULT (GETDATE()),
        AutoriseePar        NVARCHAR(50)    NOT NULL,
        DateAutorisation    DATETIME        NOT NULL DEFAULT (GETDATE()),

        CONSTRAINT PK_T_AnnulationWU PRIMARY KEY (IdAnnulation),

        CONSTRAINT CK_T_AnnulationWU_Motif
            CHECK (Motif IN ('RAPPORT_VIDE', 'RAPPORT_ERRONE', 'MAUVAISE_JOURNEE',
                             'DOUBLON', 'AUTRE')),

        CONSTRAINT CK_T_AnnulationWU_PasSoiMeme
            CHECK (AutoriseePar <> DemandeePar)
    );

    CREATE INDEX IX_T_AnnulationWU_DateActivite ON dbo.T_AnnulationWU (DateActivite);

    PRINT 'Table T_AnnulationWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_AnnulationWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. L'historique annulé — mêmes colonnes que T_HistoriqueWU, plus IdAnnulation
--
--    Pas de clé primaire sur (IdAnnulation, DateActivite, Account) : l'archive ne se
--    consulte que par annulation, et une contrainte de plus ne protégerait rien qui ne
--    le soit déjà par la table d'origine.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_HistoriqueAnnuleWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_HistoriqueAnnuleWU
    (
        IdAnnulation        BIGINT          NOT NULL,

        DateActivite        DATE            NOT NULL,
        Account             NVARCHAR(255)   NOT NULL,
        Designation         NVARCHAR(255)   NULL,
        GroupeStatistique   NVARCHAR(255)   NULL,
        TypePdv             NVARCHAR(20)    NULL,

        NombreEnvois        INT             NOT NULL DEFAULT (0),
        NombrePaiements     INT             NOT NULL DEFAULT (0),
        NombreAnnulations   INT             NOT NULL DEFAULT (0),

        PrincipalEnvoi      DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        ChargeEnvoi         DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        Taxes               DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        PrincipalPaye       DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        CommissionEnvoi     DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        CommissionPaiement  DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        CommissionTransfert DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        TVA                 DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TTAEnvoi            DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TTAReception        DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TaxeEnvoi           DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        -- Part revenant a la banque, recopiee telle qu'elle etait dans l'historique.
        -- Une archive qui perd une partie de ce qu'elle archive ne prouve plus grand-chose.
        CommissionEnvoiBanque     DECIMAL(18, 2) NULL,
        CommissionPaiementBanque  DECIMAL(18, 2) NULL,
        CommissionTransfertBanque DECIMAL(18, 2) NULL,
        TauxSA                    DECIMAL(4, 2)  NULL,

        DateEnregistrement      DATETIME    NULL,
        ComptabilisePar         NVARCHAR(50) NULL,
        DateComptabilisation    DATETIME    NULL
    );

    CREATE INDEX IX_T_HistoriqueAnnuleWU_Annulation
        ON dbo.T_HistoriqueAnnuleWU (IdAnnulation);

    PRINT 'Table T_HistoriqueAnnuleWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueAnnuleWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 3. Le détail des MTCN annulés
-- =========================================================================
IF OBJECT_ID(N'dbo.T_HistoriqueMTCNAnnuleWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_HistoriqueMTCNAnnuleWU
    (
        IdAnnulation        BIGINT          NOT NULL,

        DateActivite        DATE            NOT NULL,
        Account             NVARCHAR(255)   NOT NULL,
        MTCN                NVARCHAR(50)    NOT NULL,

        Sens                NVARCHAR(10)    NOT NULL,
        Statut              NVARCHAR(10)    NULL,
        Montant             DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        Designation         NVARCHAR(255)   NULL,
        GroupeStatistique   NVARCHAR(255)   NULL,
        TypePdv             NVARCHAR(20)    NULL,

        DateEnregistrement      DATETIME    NULL,
        ComptabilisePar         NVARCHAR(50) NULL,
        DateComptabilisation    DATETIME    NULL
    );

    CREATE INDEX IX_T_HistoriqueMTCNAnnuleWU_Annulation
        ON dbo.T_HistoriqueMTCNAnnuleWU (IdAnnulation);

    -- La recherche d'une transaction part toujours de son MTCN, y compris dans l'archive :
    -- « ce MTCN a-t-il été comptabilisé ? » doit trouver la réponse même si la journée a
    -- été annulée depuis.
    CREATE INDEX IX_T_HistoriqueMTCNAnnuleWU_MTCN ON dbo.T_HistoriqueMTCNAnnuleWU (MTCN);

    PRINT 'Table T_HistoriqueMTCNAnnuleWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueMTCNAnnuleWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 4. La pièce annulée
--
--    La contrainte CK_T_PieceWU_UnSeulSens n'est PAS reprise. Une archive conserve ce qui
--    a été écrit ; elle ne rejuge pas. Refuser une ligne au moment du déplacement ferait
--    échouer l'annulation et laisserait la journée en place, ce qui est exactement le
--    contraire du but.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_PieceAnnuleeWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_PieceAnnuleeWU
    (
        IdAnnulation        BIGINT          NOT NULL,

        DateActivite        DATE            NOT NULL,
        Ligne               INT             NOT NULL,

        Compte              NVARCHAR(50)    NOT NULL,
        Libelle             NVARCHAR(255)   NULL,
        Debit               BIGINT          NOT NULL DEFAULT (0),
        Credit              BIGINT          NOT NULL DEFAULT (0),
        CodeAgence          NVARCHAR(50)    NULL,

        DateEnregistrement  DATETIME        NULL,
        EnregistrePar       NVARCHAR(100)   NULL,

        CONSTRAINT PK_T_PieceAnnuleeWU PRIMARY KEY (IdAnnulation, Ligne)
    );

    CREATE INDEX IX_T_PieceAnnuleeWU_DateActivite ON dbo.T_PieceAnnuleeWU (DateActivite);

    PRINT 'Table T_PieceAnnuleeWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_PieceAnnuleeWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 5. La file des demandes accueille un quatrième objet
--
--    Deux colonnes manquent à T_DemandeWU pour porter une demande d'annulation : le
--    commentaire libre et la réponse sur le core banking. Les autres colonnes servent
--    telles quelles :
--
--      Cle          = la journée, au format AAAA-MM-JJ
--      Designation  = le motif (RAPPORT_VIDE, RAPPORT_ERRONE, …)
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns
               WHERE object_id = OBJECT_ID(N'T_DemandeWU') AND name = N'Commentaire')
BEGIN
    ALTER TABLE dbo.T_DemandeWU ADD Commentaire NVARCHAR(500) NULL;
    PRINT 'Colonne Commentaire ajoutée à T_DemandeWU.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns
               WHERE object_id = OBJECT_ID(N'T_DemandeWU') AND name = N'CoreBankingInjecte')
BEGIN
    ALTER TABLE dbo.T_DemandeWU ADD CoreBankingInjecte BIT NULL;
    PRINT 'Colonne CoreBankingInjecte ajoutée à T_DemandeWU.';
END
GO

-- Les deux contraintes de valeurs doivent accepter le nouvel objet et la nouvelle
-- opération. Elles sont remplacées, et non assouplies : une contrainte supprimée sans
-- être reposée laisserait passer n'importe quel libellé.
IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_TypeObjet')
BEGIN
    ALTER TABLE dbo.T_DemandeWU DROP CONSTRAINT CK_T_DemandeWU_TypeObjet;
END
GO

ALTER TABLE dbo.T_DemandeWU WITH CHECK
    ADD CONSTRAINT CK_T_DemandeWU_TypeObjet
    CHECK (TypeObjet IN ('SOUS_AGENT', 'AGENCE', 'GROUPE', 'COMPTABILISATION'));
GO
PRINT 'Contrainte CK_T_DemandeWU_TypeObjet reposée avec COMPTABILISATION.';
GO

IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_Operation')
BEGIN
    ALTER TABLE dbo.T_DemandeWU DROP CONSTRAINT CK_T_DemandeWU_Operation;
END
GO

ALTER TABLE dbo.T_DemandeWU WITH CHECK
    ADD CONSTRAINT CK_T_DemandeWU_Operation
    CHECK (Operation IN ('CREATION', 'MODIFICATION', 'SUPPRESSION', 'SYNCHRONISATION',
                         'ANNULATION'));
GO
PRINT 'Contrainte CK_T_DemandeWU_Operation reposée avec ANNULATION.';
GO

-- =========================================================================
-- 6. Vue d'audit : les journées annulées, lisibles d'un coup d'œil
-- =========================================================================
IF OBJECT_ID(N'dbo.V_JourneesAnnulees', N'V') IS NOT NULL DROP VIEW dbo.V_JourneesAnnulees;
GO
CREATE VIEW dbo.V_JourneesAnnulees
AS
    SELECT  a.IdAnnulation,
            a.DateActivite,
            Motif             = a.Motif,
            a.Commentaire,
            CoreBanking       = CASE WHEN a.CoreBankingInjecte = 1
                                     THEN N'INJECTÉ — extourne à demander'
                                     ELSE N'non injecté' END,
            a.NombrePdv,
            a.NombreTransactions,
            a.NombreEcritures,
            a.TotalDebit,
            a.TotalCredit,
            Equilibre         = CASE WHEN a.TotalDebit = a.TotalCredit
                                     THEN N'oui' ELSE N'NON' END,
            a.DemandeePar, a.DateDemande,
            a.AutoriseePar, a.DateAutorisation
    FROM    dbo.T_AnnulationWU AS a;
GO

-- =========================================================================
-- 7. Droits
--
--    Le rôle de compense LIT les archives et n'y écrit pas : c'est l'autorisation qui
--    déplace les lignes, et l'autorisation appartient à l'authorizer. Comme la fonction
--    d'authorizer se porte indifféremment sur un compte de compense, de commercial ou
--    d'administration, les trois rôles reçoivent l'écriture — le contrôle du double
--    regard, lui, est posé par la contrainte CK_T_AnnulationWU_PasSoiMeme et par
--    l'application, non par les droits SQL.
--
--    AUCUN rôle ne reçoit DELETE sur les archives : une archive que ses propres
--    utilisateurs peuvent effacer ne prouve rien.
-- =========================================================================
DECLARE @tables TABLE (nom SYSNAME);
INSERT INTO @tables (nom)
VALUES (N'T_AnnulationWU'), (N'T_HistoriqueAnnuleWU'),
       (N'T_HistoriqueMTCNAnnuleWU'), (N'T_PieceAnnuleeWU');

DECLARE @role SYSNAME, @table SYSNAME, @sql NVARCHAR(400);

DECLARE curRoles CURSOR LOCAL FAST_FORWARD FOR
    SELECT name FROM sys.database_principals
    WHERE type = 'R' AND name IN (N'wu_compense', N'wu_commercial', N'wu_admin');

OPEN curRoles;
FETCH NEXT FROM curRoles INTO @role;

WHILE @@FETCH_STATUS = 0
BEGIN
    DECLARE curTablesAnnul CURSOR LOCAL FAST_FORWARD FOR SELECT nom FROM @tables;
    OPEN curTablesAnnul;
    FETCH NEXT FROM curTablesAnnul INTO @table;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @sql = N'GRANT SELECT, INSERT ON dbo.' + QUOTENAME(@table) + N' TO ' + QUOTENAME(@role);
        EXEC sp_executesql @sql;
        FETCH NEXT FROM curTablesAnnul INTO @table;
    END

    CLOSE curTablesAnnul;
    DEALLOCATE curTablesAnnul;

    PRINT 'Droits accordés à ' + @role + ' sur les tables d''annulation.';
    FETCH NEXT FROM curRoles INTO @role;
END

CLOSE curRoles;
DEALLOCATE curRoles;
GO

-- La vue d'audit se lit par tout le monde.
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.V_JourneesAnnulees TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.V_JourneesAnnulees TO wu_commercial');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT ON dbo.V_JourneesAnnulees TO wu_admin');
GO

-- =========================================================================
-- 8. Contrôle
-- =========================================================================
SELECT  annulations = COUNT(*),
        journees    = COUNT(DISTINCT DateActivite),
        premiere    = MIN(DateActivite),
        derniere    = MAX(DateActivite)
FROM    dbo.T_AnnulationWU;
GO

PRINT N'';
PRINT N'Annulation : une journee retiree sort des rapports d''activite, pas des livres';
PRINT N'de la banque. Si le fichier core banking a ete injecte, l''extourne se demande';
PRINT N'en comptabilite — Wincompense ne la fait pas.';
GO
