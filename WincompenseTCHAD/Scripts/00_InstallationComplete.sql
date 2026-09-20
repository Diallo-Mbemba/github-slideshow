/*
================================================================================================
    WINCOMPENSE TCHAD — INSTALLATION COMPLÈTE DE LA BASE DE DONNÉES
================================================================================================

    UN SEUL SCRIPT. Il crée la base, ses tables, ses rôles et ses droits, et rend compte de
    ce qu'il a fait. Il remplace l'exécution des dix scripts 01 à 10 un par un.

    ------------------------------------------------------------------------------------------
    AVANT DE L'EXÉCUTER
    ------------------------------------------------------------------------------------------

      1. Ouvrir SQL Server Management Studio, connecté au serveur de PRODUCTION.
      2. Vérifier en haut de l'écran que c'est bien le bon serveur.
      3. Exécuter le script entier (F5).

    Aucune base n'est écrasée : chaque objet est protégé par un contrôle d'existence. Le script
    est REJOUABLE — le relancer sur une base déjà installée ne détruit rien et ne crée que ce
    qui manque. Il peut donc servir aussi à rattraper une installation incomplète.

    ------------------------------------------------------------------------------------------
    CE QU'IL FAIT, DANS L'ORDRE
    ------------------------------------------------------------------------------------------

      Partie 1   La base GWC_WINCOMPENSE_ETD
      Partie 2   Les douze tables et leurs index
      Partie 3   Les données de paramétrage de départ (comptes comptables, jours fériés)
      Partie 4   Les trois rôles et leurs droits
      Partie 5   L'accès du compte applicatif à la base
      Partie 6   Le compte rendu

    L'ORDRE DIFFÈRE VOLONTAIREMENT de la numérotation des scripts d'origine : les droits sont
    accordés EN DERNIER, une fois toutes les tables créées. Exécutés dans l'ordre numérique,
    les scripts laissaient T_DemandeWU et T_JourFerieWU sans aucun droit, et l'application
    échouait là où on ne l'attendait pas — calendrier illisible au moment de dater le fichier
    core banking, file du double regard vide alors qu'elle contenait des demandes.

    ------------------------------------------------------------------------------------------
    CE QU'IL NE FAIT PAS
    ------------------------------------------------------------------------------------------

      - Il n'insère AUCUN sous-agent ni agence : le référentiel se charge depuis l'application,
        ou par vos propres scripts. Les données d'exemple du script 02 sont volontairement
        exclues : elles n'ont leur place que sur un environnement de test.

      - Il ne crée AUCUN compte utilisateur : le premier administrateur se crée au premier
        démarrage de l'application, qui demande confirmation de la base visée.

      - Il ne crée AUCUN login SQL Server : il faudrait un mot de passe, qui n'a pas sa place
        dans un fichier qui circule. La PARTIE 5 rattache à la base un login DÉJÀ CRÉÉ par la
        banque, et lui accorde son rôle.

      - Il ne règle NI le mode de récupération NI la sauvegarde. À vérifier séparément : une
        base en mode FULL sans sauvegarde du journal finit par remplir le disque.

    ------------------------------------------------------------------------------------------
    APRÈS L'EXÉCUTION
    ------------------------------------------------------------------------------------------

      1. Lire l'onglet « Messages » : chaque objet créé y est annoncé.
      2. Lire le compte rendu de la partie 6 : il dit ce qui manque, s'il manque quelque chose.
      3. Compléter la PARTIE 5 avec vos groupes Active Directory, puis réexécuter cette partie.
      4. Renseigner les neuf comptes comptables depuis l'application (écran Comptes systèmes),
         ou par le script 03 si vous préférez les poser directement.
      5. Compléter T_JourFerieWU des fêtes musulmanes de l'année : elles ne se calculent pas
         d'avance et le script ne sème que les fêtes à date fixe et les lundis de Pâques.

================================================================================================
*/

SET NOCOUNT ON;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.1 — Points de vente : sous-agents et agences propres
    (source : Scripts\01_CreateTables_GWC_WINCOMPENSE_ETD.sql)
----------------------------------------------------------------------------------------------*/

/*
    Script de création des tables de paramétrage des points de vente Western Union
    pour la base GWC_WINCOMPENSE_ETD (SQL Server Express, instance .\SQLEXPRESS).

    IMPORTANT : l'Account (Code_Pdv / Codesite) est l'identifiant métier unique.
    Ne jamais utiliser codeagence comme clé d'identification.

    Types alignés sur le schéma réellement en place à Ecobank Tchad (NVARCHAR(255) partout,
    Taux en DECIMAL(4,2)). Le script étant protégé par IF NOT EXISTS, il ne modifie jamais
    une table déjà créée : il ne sert qu'à monter un environnement neuf.

    Ces deux tables se gèrent depuis l'application : boutons « Sous-agents... » et
    « Agences propres... » de l'écran principal.
*/

USE master;
GO

IF NOT EXISTS (SELECT 1 FROM sys.databases WHERE name = N'GWC_WINCOMPENSE_ETD')
BEGIN
    CREATE DATABASE GWC_WINCOMPENSE_ETD;
END
GO

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- Table des SOUS-AGENTS Western Union
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_Pdv_SA')
BEGIN
    CREATE TABLE dbo.T_Pdv_SA
    (
        Code_Pdv          NVARCHAR(255)   NOT NULL PRIMARY KEY, -- Account (identifiant métier unique)
        Designationagence NVARCHAR(255)   NOT NULL,
        GroupeStatistique NVARCHAR(255)   NOT NULL,
        Taux              DECIMAL(4, 2)   NOT NULL DEFAULT (0),  -- Quote-part du sous-agent (0.70 = 70 %), DEUX décimales
        CompteCompense    NVARCHAR(255)   NOT NULL,              -- Compte de compensation du sous-agent
        CompteCommission  NVARCHAR(255)   NOT NULL,              -- Compte de commission du sous-agent
        codeagence        NVARCHAR(255)   NOT NULL               -- Informatif uniquement, jamais utilisé comme clé
    );
END
GO

-- =========================================================================
-- Table des AGENCES PROPRES (Ecobank)
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_Pdv_EC')
BEGIN
    CREATE TABLE dbo.T_Pdv_EC
    (
        Codesite            NVARCHAR(255)   NOT NULL PRIMARY KEY, -- Account (identifiant métier unique)
        Designationagence   NVARCHAR(255)   NULL,
        [CodeAgenc-Voyager] NVARCHAR(255)   NOT NULL
    );
END
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.2 — Comptes comptables paramétrés (SystemeWU)
    (source : Scripts\03_SystemeWU.sql)
----------------------------------------------------------------------------------------------*/

/*
    Table SystemeWU : paramétrage des comptes comptables utilisés par la pièce comptable
    Western Union. Ces comptes sont lus au démarrage de l'application et modifiables depuis
    le formulaire « Comptes Systèmes WU » (bouton « Paramètres des comptes... »).

    Ce script est IDEMPOTENT : il ne crée la table et la ligne de paramétrage que si elles
    n'existent pas déjà, et ne modifie JAMAIS une ligne existante — le paramétrage en place,
    qui fait foi, ne doit pas être écrasé par un script.

    À n'exécuter que si la table SystemeWU n'existe pas encore dans l'environnement visé.
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'SystemeWU')
BEGIN
    CREATE TABLE dbo.SystemeWU
    (
        Passif                      NVARCHAR(255) NULL,
        Actif                       NVARCHAR(255) NULL,
        Tob                         NVARCHAR(255) NULL,   -- TVA collectée Western Union
        Tthu                        NVARCHAR(255) NULL,   -- Impôts et taxe sur envoi
        Cpte_PositionNette          NVARCHAR(255) NULL,   -- Compte courant Western Union ETD
        Cpte_Produit                NVARCHAR(255) NULL,   -- Commission sur Transfert_Ecobank
        Cpte_Charge_Publicitaire    NVARCHAR(255) NULL,
        Cpte_Gainde_Change          NVARCHAR(255) NULL,
        Cpte_Envoi                  NVARCHAR(255) NULL,   -- TTA sur envoi WU
        Cpte_Paiement               NVARCHAR(255) NULL,   -- TTA sur paiement WU
        Cpte_attenteDEBIT           NVARCHAR(255) NULL,   -- Compte inter bancaire (débit)
        Cpte_attenteCREDIT          NVARCHAR(255) NULL,   -- Compte inter bancaire (crédit)
        Cpte_Produit_Envoi          NVARCHAR(255) NULL,   -- Commission sur Envoi_Ecobank
        Cpte_Produit_Paiement       NVARCHAR(255) NULL,   -- Commission sur Paiement_Ecobank
        Cpte_Envoi_agence           NVARCHAR(255) NULL,
        Cpte_Paiement_agence        NVARCHAR(255) NULL,
        code                        NCHAR(10)     NULL    -- Identifie la ligne de paramétrage
    );
END
GO

/*
    Ligne de paramétrage Western Union.

    Les valeurs ci-dessous sont celles du paramétrage en service au moment de la rédaction,
    reprises du formulaire « Comptes Systèmes WU » de la Direction Comptable. Les colonnes
    non utilisées par l'application (Passif, Actif, Cpte_Charge_Publicitaire,
    Cpte_Gainde_Change, Cpte_Envoi_agence, Cpte_Paiement_agence) sont laissées à NULL :
    l'application ne les lit ni ne les écrit jamais.

    La colonne « code » doit être renseignée : sans elle, l'application refuse d'enregistrer
    dès que la table contient plusieurs lignes, ne pouvant pas déterminer laquelle mettre à
    jour sans risquer d'altérer le paramétrage d'une autre application.
*/
IF NOT EXISTS (SELECT 1 FROM dbo.SystemeWU)
BEGIN
    INSERT INTO dbo.SystemeWU
    (
        Tob, Tthu,
        Cpte_PositionNette,
        Cpte_Produit, Cpte_Produit_Envoi, Cpte_Produit_Paiement,
        Cpte_Envoi, Cpte_Paiement,
        Cpte_attenteDEBIT, Cpte_attenteCREDIT,
        code
    )
    VALUES
    (
        '434000104',    -- TVA collectée Western Union
        '434000147',    -- Impôts et taxe sur envoi
        '32100003292',  -- Compte courant Western Union ETD
        '728300148',    -- Commission sur Transfert_Ecobank
        '728300148',    -- Commission sur Envoi_Ecobank
        '728300149',    -- Commission sur Paiement_Ecobank
        '434000145',    -- TTA sur envoi WU
        '434000159',    -- TTA sur paiement WU
        '381000101',    -- Compte inter bancaire (débit)
        '381000101',    -- Compte inter bancaire (crédit)
        '123'
    );
END
GO

-- Contrôle : comptes effectivement lus par l'application.
SELECT  code                    AS [Code],
        Cpte_PositionNette      AS [Compte courant WU ETD],
        Cpte_attenteDEBIT       AS [Compte inter bancaire],
        Cpte_Produit            AS [Commission Transfert],
        Cpte_Produit_Envoi      AS [Commission Envoi],
        Cpte_Produit_Paiement   AS [Commission Paiement],
        Tthu                    AS [Impots et taxe envoi],
        Tob                     AS [TVA],
        Cpte_Envoi              AS [TTA envoi],
        Cpte_Paiement           AS [TTA paiement]
FROM    dbo.SystemeWU;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.3 — Groupes statistiques
    (source : Scripts\04_GroupeStatistique.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Table T_GroupeStatistique + MIGRATION depuis T_Pdv_SA
    =========================================================================

    Règle métier mise en oeuvre :
      - un groupe statistique porte UN compte d'activité, UN compte de commission, UN taux ;
      - tout Account (sous-agent) appartient à un seul groupe et en hérite ces trois valeurs ;
      - un compte d'activité ou de commission n'appartient qu'à UN SEUL groupe.

    Jusqu'ici, ces trois valeurs étaient portées par chaque sous-agent, et la règle ne pouvait
    être que vérifiée par l'application. Cette table en fait la source de vérité, et les
    contraintes ci-dessous rendent les violations IMPOSSIBLES, y compris en écriture directe.

    Le script est PARTIEL et RÉPÉTABLE : il reprend tous les groupes sains et laisse de côté
    ceux qui demandent un arbitrage, en disant lesquels et pourquoi. Il se relance autant de
    fois que nécessaire, sans jamais retoucher un groupe déjà enregistré.

    Un groupe non repris n'est PAS un groupe perdu : l'application continue de le proposer,
    avec les valeurs lues dans T_Pdv_SA, et permet d'y rattacher des sous-agents comme de
    l'enregistrer d'un clic après arbitrage. Rien n'est bloqué en attendant.

    Il se déroule en quatre temps :
        1. création de la table (si absente) ;
        2. DIAGNOSTIC des données existantes ;
        3. MIGRATION des groupes sains, puis liste des groupes restant à arbitrer ;
        4. contrôle final.

    Les colonnes CompteCompense / CompteCommission / Taux de T_Pdv_SA sont CONSERVÉES et
    tenues synchronisées avec le groupe : d'autres applications peuvent les lire. L'application
    Wincompense, elle, lit désormais les valeurs du groupe.

    À exécuter une seule fois, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. Création de la table
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    CREATE TABLE dbo.T_GroupeStatistique
    (
        Groupe           NVARCHAR(255) NOT NULL PRIMARY KEY,  -- Libellé du groupe (clé métier)
        CompteActivite   NVARCHAR(255) NOT NULL,              -- Porte la ligne « CCS_... ACTIVITE WU »
        CompteCommission NVARCHAR(255) NOT NULL,              -- Rétrocession des commissions
        Taux             DECIMAL(4, 2) NOT NULL DEFAULT (0)   -- Quote-part du groupe (0.70 = 70 %)
    );

    -- Un compte ne peut appartenir qu'à un seul groupe : la contrainte le garantit.
    CREATE UNIQUE INDEX UQ_T_GroupeStatistique_CompteActivite
        ON dbo.T_GroupeStatistique (CompteActivite);

    CREATE UNIQUE INDEX UQ_T_GroupeStatistique_CompteCommission
        ON dbo.T_GroupeStatistique (CompteCommission);

    PRINT 'Table T_GroupeStatistique créée.';
END
ELSE
BEGIN
    PRINT 'Table T_GroupeStatistique déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. DIAGNOSTIC des données existantes
-- =========================================================================
PRINT '';
PRINT '--- Anomalie 1 : groupes dont les sous-agents ne portent pas tous les mêmes valeurs ---';
PRINT '    (à corriger avant migration : le groupe ne peut avoir qu''un seul jeu de valeurs)';

SELECT  LTRIM(RTRIM(GroupeStatistique))     AS Groupe,
        COUNT(*)                            AS NombreSousAgents,
        MIN(CompteCompense)                 AS ActiviteMin,
        MAX(CompteCompense)                 AS ActiviteMax,
        MIN(CompteCommission)               AS CommissionMin,
        MAX(CompteCommission)               AS CommissionMax,
        MIN(Taux)                           AS TauxMin,
        MAX(Taux)                           AS TauxMax
FROM    dbo.T_Pdv_SA
WHERE   GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> ''
GROUP BY LTRIM(RTRIM(GroupeStatistique))
HAVING  MIN(CompteCompense)   <> MAX(CompteCompense)
     OR MIN(CompteCommission) <> MAX(CompteCommission)
     OR MIN(Taux)             <> MAX(Taux)
ORDER BY 1;
GO

PRINT '';
PRINT '--- Anomalie 2 : un même compte utilisé par plusieurs groupes ---';
PRINT '    (à corriger avant migration : un compte n''appartient qu''à un seul groupe)';

WITH Groupes AS
(
    SELECT  LTRIM(RTRIM(GroupeStatistique)) AS Groupe,
            MIN(CompteCompense)             AS CompteActivite,
            MIN(CompteCommission)           AS CompteCommission
    FROM    dbo.T_Pdv_SA
    WHERE   GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> ''
    GROUP BY LTRIM(RTRIM(GroupeStatistique))
)
SELECT 'Compte d''activité' AS TypeDeCompte, CompteActivite AS Compte, COUNT(*) AS NombreGroupes
FROM   Groupes GROUP BY CompteActivite HAVING COUNT(*) > 1
UNION ALL
SELECT 'Compte de commission', CompteCommission, COUNT(*)
FROM   Groupes GROUP BY CompteCommission HAVING COUNT(*) > 1;
GO

PRINT '';
PRINT '--- Information : sous-agents sans groupe statistique ---';
PRINT '    (non bloquant pour la migration, mais ils n''hériteront d''aucune valeur)';

SELECT  Code_Pdv, Designationagence, CompteCompense, CompteCommission, Taux
FROM    dbo.T_Pdv_SA
WHERE   GroupeStatistique IS NULL OR LTRIM(RTRIM(GroupeStatistique)) = ''
ORDER BY Code_Pdv;
GO

-- =========================================================================
-- 3. MIGRATION PARTIELLE ET RÉPÉTABLE
--
--    Sont repris tous les groupes SAINS, c'est-à-dire ceux qui satisfont les deux
--    conditions : leurs sous-agents portent tous les mêmes valeurs, et aucun de leurs
--    deux comptes n'est utilisé par un autre groupe.
--
--    Les groupes écartés sont listés en fin de script. Ils restent parfaitement
--    utilisables dans l'application, qui continue de les proposer et permet de les
--    enregistrer un à un après arbitrage : rien n'est bloqué en attendant.
--
--    Ce script se relance autant de fois que nécessaire : il ne reprend chaque fois
--    que les groupes devenus sains et jamais ceux déjà enregistrés.
-- =========================================================================

WITH Groupes AS
(
    SELECT  LTRIM(RTRIM(GroupeStatistique)) AS Groupe,
            MIN(CompteCompense)             AS CompteActivite,
            MAX(CompteCompense)             AS ActiviteMax,
            MIN(CompteCommission)           AS CompteCommission,
            MAX(CompteCommission)           AS CommissionMax,
            MIN(Taux)                       AS Taux,
            MAX(Taux)                       AS TauxMax
    FROM    dbo.T_Pdv_SA
    WHERE   GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> ''
    GROUP BY LTRIM(RTRIM(GroupeStatistique))
),
Coherents AS   -- groupes dont tous les sous-agents portent les mêmes valeurs
(
    SELECT * FROM Groupes
    WHERE  CompteActivite = ActiviteMax
      AND  CompteCommission = CommissionMax
      AND  Taux = TauxMax
),
Sains AS       -- ... et dont aucun compte n'est revendiqué par un autre groupe
(
    SELECT c.*
    FROM   Coherents c
    WHERE  NOT EXISTS (SELECT 1 FROM Groupes a
                       WHERE a.Groupe <> c.Groupe AND a.CompteActivite = c.CompteActivite)
      AND  NOT EXISTS (SELECT 1 FROM Groupes a
                       WHERE a.Groupe <> c.Groupe AND a.CompteCommission = c.CompteCommission)
      -- ni par un groupe déjà enregistré lors d'une exécution précédente
      AND  NOT EXISTS (SELECT 1 FROM dbo.T_GroupeStatistique g
                       WHERE g.Groupe <> c.Groupe
                         AND (g.CompteActivite = c.CompteActivite
                           OR g.CompteCommission = c.CompteCommission))
)
INSERT INTO dbo.T_GroupeStatistique (Groupe, CompteActivite, CompteCommission, Taux)
SELECT  s.Groupe, s.CompteActivite, s.CompteCommission, s.Taux
FROM    Sains s
WHERE   NOT EXISTS (SELECT 1 FROM dbo.T_GroupeStatistique g WHERE g.Groupe = s.Groupe);

PRINT '';
PRINT 'MIGRATION : ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' groupe(s) repris lors de cette exécution.';
GO

-- =========================================================================
-- 3 bis. Groupes NON repris, et pourquoi
-- =========================================================================
PRINT '';
PRINT '--- Groupes restant à arbitrer (non repris dans T_GroupeStatistique) ---';

WITH Groupes AS
(
    SELECT  LTRIM(RTRIM(GroupeStatistique)) AS Groupe,
            MIN(CompteCompense)             AS ActiviteMin,
            MAX(CompteCompense)             AS ActiviteMax,
            MIN(CompteCommission)           AS CommissionMin,
            MAX(CompteCommission)           AS CommissionMax,
            MIN(Taux)                       AS TauxMin,
            MAX(Taux)                       AS TauxMax,
            COUNT(*)                        AS NombreSousAgents
    FROM    dbo.T_Pdv_SA
    WHERE   GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> ''
    GROUP BY LTRIM(RTRIM(GroupeStatistique))
)
SELECT  g.Groupe,
        g.NombreSousAgents,
        CASE WHEN g.ActiviteMin <> g.ActiviteMax
               OR g.CommissionMin <> g.CommissionMax
               OR g.TauxMin <> g.TauxMax
             THEN 'Sous-agents aux valeurs divergentes : choisir lesquelles retenir'
             ELSE 'Un de ses comptes est deja utilise par un autre groupe'
        END                             AS RaisonDuRejet,
        g.ActiviteMin, g.ActiviteMax,
        g.CommissionMin, g.CommissionMax,
        g.TauxMin, g.TauxMax
FROM    Groupes g
WHERE   NOT EXISTS (SELECT 1 FROM dbo.T_GroupeStatistique t WHERE t.Groupe = g.Groupe)
ORDER BY g.Groupe;
GO

-- =========================================================================
-- 4. Contrôle final
-- =========================================================================
PRINT '';
PRINT '--- Groupes statistiques et nombre de sous-agents rattachés ---';

SELECT  g.Groupe,
        g.CompteActivite,
        g.CompteCommission,
        g.Taux,
        COUNT(p.Code_Pdv) AS NombreSousAgents
FROM    dbo.T_GroupeStatistique g
        LEFT JOIN dbo.T_Pdv_SA p
               ON LTRIM(RTRIM(p.GroupeStatistique)) = g.Groupe
GROUP BY g.Groupe, g.CompteActivite, g.CompteCommission, g.Taux
ORDER BY g.Groupe;
GO

/*
    ÉTAPE FACULTATIVE — clé étrangère
    ---------------------------------
    Elle interdirait de rattacher un sous-agent à un groupe inexistant. Elle n'est PAS posée
    automatiquement : elle échouerait tant qu'il reste des sous-agents sans groupe (chaîne
    vide), et le choix de leur en attribuer un relève de la Direction Comptable.

    À exécuter une fois tous les sous-agents rattachés à un groupe existant :

    ALTER TABLE dbo.T_Pdv_SA
        ADD CONSTRAINT FK_T_Pdv_SA_GroupeStatistique
        FOREIGN KEY (GroupeStatistique) REFERENCES dbo.T_GroupeStatistique (Groupe);
*/



/*----------------------------------------------------------------------------------------------
    PARTIE 2.4 — Historique des journées comptabilisées
    (source : Scripts\05_HistoriqueWU.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Table T_HistoriqueWU — historique des journées comptabilisées
    =========================================================================

    L'application traite une journée à la fois, à partir des deux rapports Western Union, et
    ne conservait jusqu'ici aucune trace de ce qu'elle avait calculé. Cette table enregistre
    le résultat de chaque journée COMPTABILISÉE — une ligne par point de vente et par jour —
    et sert de source aux rapports d'activité sur une période.

    Elle restitue donc ce qui a réellement été comptabilisé, et non un recalcul a posteriori :
    c'est tout l'intérêt d'un historique. Elle n'est alimentée qu'au moment où la pièce
    comptable est générée, jamais lors d'un simple affichage.

    Une journée regénérée remplace intégralement ses lignes : l'écriture est donc répétable
    sans jamais produire de doublon.

    À exécuter une fois, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueWU')
BEGIN
    CREATE TABLE dbo.T_HistoriqueWU
    (
        DateActivite        DATE            NOT NULL,   -- Journée d'opérations (txnDateLOC)
        Account             NVARCHAR(255)   NOT NULL,   -- Point de vente

        -- Identification du point de vente TELLE QU'ELLE ÉTAIT le jour de la comptabilisation.
        -- Recopiée et non rattachée par clé étrangère : un point de vente peut changer de
        -- groupe ou de désignation, ce qui ne doit pas réécrire le passé.
        Designation         NVARCHAR(255)   NULL,
        GroupeStatistique   NVARCHAR(255)   NULL,
        TypePdv             NVARCHAR(20)    NULL,       -- SA, EC ou INCONNU

        -- Volumes
        NombreEnvois        INT             NOT NULL DEFAULT (0),
        NombrePaiements     INT             NOT NULL DEFAULT (0),
        NombreAnnulations   INT             NOT NULL DEFAULT (0),

        -- Montants issus des rapports
        PrincipalEnvoi      DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        ChargeEnvoi         DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        Taxes               DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        PrincipalPaye       DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        -- Commissions calculées
        CommissionEnvoi     DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        CommissionPaiement  DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        CommissionTransfert DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        -- Taxes calculées
        TVA                 DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TTAEnvoi            DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TTAReception        DECIMAL(18, 2)  NOT NULL DEFAULT (0),
        TaxeEnvoi           DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        DateEnregistrement  DATETIME        NOT NULL DEFAULT (GETDATE()),

        CONSTRAINT PK_T_HistoriqueWU PRIMARY KEY (DateActivite, Account)
    );

    -- Les rapports interrogent toujours par intervalle de dates.
    CREATE INDEX IX_T_HistoriqueWU_DateActivite ON dbo.T_HistoriqueWU (DateActivite);

    PRINT 'Table T_HistoriqueWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueWU déjà présente : création ignorée.';
END
GO

-- Contrôle : journées déjà historisées.
SELECT  DateActivite,
        COUNT(*)                AS PointsDeVente,
        SUM(NombreEnvois)       AS Envois,
        SUM(NombrePaiements)    AS Paiements,
        SUM(PrincipalEnvoi)     AS PrincipalEnvoye,
        SUM(PrincipalPaye)      AS PrincipalPaye,
        MIN(DateEnregistrement) AS PremiereEcriture
FROM    dbo.T_HistoriqueWU
GROUP BY DateActivite
ORDER BY DateActivite DESC;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.5 — Historique détaillé, MTCN par MTCN
    (source : Scripts\06_HistoriqueMTCN.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Table T_HistoriqueMTCN — transactions détaillées des journées comptabilisées
    =========================================================================

    T_HistoriqueWU agrège la journée par point de vente : elle ne peut pas porter le MTCN, qui
    identifie UNE transaction. Cette table conserve donc le détail, une ligne par transaction
    d'envoi ou de paiement, afin de pouvoir retrouver et justifier une opération précise.

    Alimentée en même temps que T_HistoriqueWU, dans la MÊME transaction : les deux tables ne
    peuvent pas diverger.

    Volume : de l'ordre de 400 transactions par jour, soit environ 100 000 lignes par an.

    Pas de clé primaire sur (DateActivite, Account, MTCN, Sens) : rien ne garantit qu'un MTCN
    ne puisse pas apparaître deux fois le même jour pour le même point de vente — un ajustement
    en produirait un — et une contrainte trop stricte ferait échouer l'historisation d'une
    journée par ailleurs valable. L'unicité technique est assurée par une colonne d'identité.

    À exécuter une fois, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueMTCN')
BEGIN
    CREATE TABLE dbo.T_HistoriqueMTCN
    (
        IdTransaction       BIGINT IDENTITY(1,1) NOT NULL,

        DateActivite        DATE            NOT NULL,   -- Journée d'opérations (txnDateLOC)
        Account             NVARCHAR(255)   NOT NULL,   -- Point de vente
        MTCN                NVARCHAR(50)    NOT NULL,   -- Référence Western Union de la transaction

        Sens                NVARCHAR(10)    NOT NULL,   -- ENVOI ou PAIEMENT
        Statut              NVARCHAR(10)    NULL,       -- S (réglée), W (en attente), C (annulée)
        Montant             DECIMAL(18, 2)  NOT NULL DEFAULT (0),

        -- Identification du point de vente telle qu'elle était ce jour-là, comme dans
        -- T_HistoriqueWU : changer un point de vente de groupe ne doit pas réécrire le passé.
        Designation         NVARCHAR(255)   NULL,
        GroupeStatistique   NVARCHAR(255)   NULL,
        TypePdv             NVARCHAR(20)    NULL,

        DateEnregistrement  DATETIME        NOT NULL DEFAULT (GETDATE()),

        CONSTRAINT PK_T_HistoriqueMTCN PRIMARY KEY (IdTransaction)
    );

    -- Les rapports interrogent par intervalle de dates, puis filtrent par groupe.
    CREATE INDEX IX_T_HistoriqueMTCN_DateActivite ON dbo.T_HistoriqueMTCN (DateActivite);

    -- La recherche d'une transaction précise part toujours de son MTCN.
    CREATE INDEX IX_T_HistoriqueMTCN_MTCN ON dbo.T_HistoriqueMTCN (MTCN);

    PRINT 'Table T_HistoriqueMTCN créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueMTCN déjà présente : création ignorée.';
END
GO

-- Contrôle : transactions historisées, par journée et par sens.
SELECT  DateActivite,
        Sens,
        COUNT(*)        AS NombreTransactions,
        SUM(Montant)    AS Montant,
        SUM(CASE WHEN Statut = 'C' THEN 1 ELSE 0 END) AS Annulations
FROM    dbo.T_HistoriqueMTCN
GROUP BY DateActivite, Sens
ORDER BY DateActivite DESC, Sens;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.6 — Utilisateurs et journal des connexions
    (source : Scripts\07_Utilisateurs.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Gestion des utilisateurs et traçabilité
    =========================================================================

    Trois rôles :
      COMPENSE    — charge les rapports, calcule et génère les pièces comptables ;
      COMMERCIAL  — crée et modifie les sous-agents, les agences et les groupes ;
      ADMIN       — les deux, plus les comptes systèmes et les utilisateurs.

    Les trois voient les rapports d'activité.

    Les mots de passe ne sont JAMAIS stockés : seule une empreinte PBKDF2 l'est, avec son sel
    et son nombre d'itérations. Conserver le nombre d'itérations en base permet de le relever
    plus tard sans invalider les comptes existants : chaque empreinte se vérifie avec le sien.

    À exécuter une fois, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. Utilisateurs
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_UtilisateurWU')
BEGIN
    CREATE TABLE dbo.T_UtilisateurWU
    (
        Identifiant         NVARCHAR(50)    NOT NULL PRIMARY KEY,   -- Nom de connexion
        NomComplet          NVARCHAR(150)   NOT NULL,
        Role                NVARCHAR(20)    NOT NULL,               -- COMPENSE, COMMERCIAL, ADMIN

        -- Empreinte du mot de passe : PBKDF2, en Base64. Le mot de passe n'existe nulle part.
        Empreinte           NVARCHAR(256)   NOT NULL,
        Sel                 NVARCHAR(128)   NOT NULL,
        Iterations          INT             NOT NULL DEFAULT (100000),

        Actif               BIT             NOT NULL DEFAULT (1),
        DoitChangerMotDePasse BIT           NOT NULL DEFAULT (1),   -- Vrai après création ou réinitialisation

        -- Verrouillage après échecs répétés : le compteur est remis à zéro par une connexion réussie.
        EchecsConsecutifs   INT             NOT NULL DEFAULT (0),
        DateVerrouillage    DATETIME        NULL,

        DerniereConnexion   DATETIME        NULL,

        DateCreation        DATETIME        NOT NULL DEFAULT (GETDATE()),
        CreePar             NVARCHAR(50)    NULL,
        DateModification    DATETIME        NULL,
        ModifiePar          NVARCHAR(50)    NULL,

        CONSTRAINT CK_T_UtilisateurWU_Role CHECK (Role IN ('COMPENSE', 'COMMERCIAL', 'ADMIN'))
    );

    PRINT 'Table T_UtilisateurWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_UtilisateurWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. Journal des connexions
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_ConnexionWU')
BEGIN
    CREATE TABLE dbo.T_ConnexionWU
    (
        IdConnexion     BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Identifiant     NVARCHAR(50)    NOT NULL,
        DateConnexion   DATETIME        NOT NULL DEFAULT (GETDATE()),
        Poste           NVARCHAR(100)   NULL,       -- Nom de la machine
        CompteWindows   NVARCHAR(100)   NULL,       -- Session Windows sous laquelle l'application tournait
        Succes          BIT             NOT NULL,
        Motif           NVARCHAR(200)   NULL        -- Raison du refus, le cas échéant
    );

    CREATE INDEX IX_T_ConnexionWU_DateConnexion ON dbo.T_ConnexionWU (DateConnexion);
    CREATE INDEX IX_T_ConnexionWU_Identifiant   ON dbo.T_ConnexionWU (Identifiant);

    PRINT 'Table T_ConnexionWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_ConnexionWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 3. Traçabilité sur les tables de paramétrage
--
--    Ajout non destructif : chaque colonne n'est créée que si elle manque, et les lignes
--    existantes restent à NULL — on ne réécrit pas un passé que l'on ne connaît pas.
-- =========================================================================
DECLARE @table SYSNAME, @sql NVARCHAR(MAX);
DECLARE curTables CURSOR FOR
    SELECT name FROM sys.tables
    WHERE name IN (N'T_Pdv_SA', N'T_Pdv_EC', N'T_GroupeStatistique', N'SystemeWU');

OPEN curTables;
FETCH NEXT FROM curTables INTO @table;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(@table) AND name = N'CreePar')
    BEGIN
        SET @sql = N'ALTER TABLE dbo.' + QUOTENAME(@table) + N'
                     ADD CreePar NVARCHAR(50) NULL,
                         DateCreation DATETIME NULL,
                         ModifiePar NVARCHAR(50) NULL,
                         DateModification DATETIME NULL';
        EXEC sp_executesql @sql;
        PRINT 'Colonnes de traçabilité ajoutées à ' + @table + '.';
    END
    ELSE
    BEGIN
        PRINT 'Colonnes de traçabilité déjà présentes sur ' + @table + '.';
    END

    FETCH NEXT FROM curTables INTO @table;
END

CLOSE curTables;
DEALLOCATE curTables;
GO

-- =========================================================================
-- 4. Traçabilité sur l'historique : qui a comptabilisé la journée, et quand
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueWU')
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'T_HistoriqueWU') AND name = N'ComptabilisePar')
BEGIN
    ALTER TABLE dbo.T_HistoriqueWU
        ADD ComptabilisePar NVARCHAR(50) NULL,
            DateComptabilisation DATETIME NULL;
    PRINT 'Colonnes de comptabilisation ajoutées à T_HistoriqueWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueMTCN')
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'T_HistoriqueMTCN') AND name = N'ComptabilisePar')
BEGIN
    ALTER TABLE dbo.T_HistoriqueMTCN
        ADD ComptabilisePar NVARCHAR(50) NULL,
            DateComptabilisation DATETIME NULL;
    PRINT 'Colonnes de comptabilisation ajoutées à T_HistoriqueMTCN.';
END
GO

/*
    =========================================================================
    5. PREMIER ADMINISTRATEUR
    =========================================================================

    Ce script ne crée AUCUN compte, et surtout pas un compte à mot de passe connu : une
    empreinte figée dans un fichier versionné est un mot de passe publié, que personne ne
    pense ensuite à changer.

    C'est l'application qui s'en charge. Au tout premier démarrage, si la table ne contient
    aucun administrateur ACTIF dont l'empreinte est exploitable, l'écran de connexion propose
    de créer ce premier compte et demande son mot de passe, qui est haché comme tous les
    autres. C'est le seul moment où l'application accepte de créer un compte sans que
    personne ne soit connecté, et uniquement dans ce cas précis.

    Il n'y a donc rien à faire ici : exécutez ce script, puis lancez l'application.
*/

-- Contrôle final : la liste est vide tant que le premier administrateur n'a pas été créé.
SELECT Identifiant, NomComplet, Role, Actif, DoitChangerMotDePasse, DerniereConnexion
FROM   dbo.T_UtilisateurWU
ORDER BY Role, Identifiant;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 2.7 — Double regard : file des demandes
    (source : Scripts\09_Demandes.sql)
----------------------------------------------------------------------------------------------*/

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



/*----------------------------------------------------------------------------------------------
    PARTIE 3.1 — Jours fériés de la banque
    (source : Scripts\10_JoursFeries.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Jours fériés — date de valeur des écritures chargées dans le core banking
    =========================================================================

    La compensation porte l'activité du jour J en valeur au PREMIER JOUR OUVRÉ SUIVANT : ni un
    samedi, ni un dimanche, ni un jour férié. Une écriture datée d'un jour chômé est rejetée par
    le core banking, ou passée d'office au jour suivant sans que personne ne le sache.

    Les samedis et dimanches sont déduits du calendrier par l'application. Les jours fériés, non :
    ils changent chaque année, et une partie d'entre eux suit le calendrier lunaire. Ils sont donc
    tenus dans cette table, que la banque complète — sans recompiler l'application.

    ATTENTION — CETTE LISTE EST À VALIDER PAR LA BANQUE.

    Les dates ci-dessous sont les fêtes à date fixe et les lundis de Pâques, calculés jusqu'en
    2030. Les fêtes musulmanes — Aïd el-Fitr, Aïd el-Kébir, Mawlid — suivent le calendrier lunaire
    et ne se calculent pas d'avance : elles sont annoncées chaque année. Elles ne figurent donc PAS
    ici et doivent être ajoutées à mesure, faute de quoi une écriture sera datée d'un jour chômé.

    À exécuter APRÈS 01 à 09, sur la base GWC_WINCOMPENSE_ETD.
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_JourFerieWU')
BEGIN
    CREATE TABLE dbo.T_JourFerieWU
    (
        DateFerie   DATE            NOT NULL PRIMARY KEY,
        Libelle     NVARCHAR(100)   NOT NULL,
        SaisiPar    NVARCHAR(50)    NULL,
        DateSaisie  DATETIME        NOT NULL DEFAULT (GETDATE())
    );

    PRINT 'Table T_JourFerieWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_JourFerieWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- T_PieceWU — les pièces comptables produites, conservées ligne à ligne
--
-- L'historique garde les montants ; il ne garde pas le taux du sous-agent ce jour-là, ni
-- ses comptes, ni la ligne d'écart posée sur le compte inter bancaire. Reconstituer une
-- pièce ancienne avec le paramétrage d'aujourd'hui réécrirait le passé : un sous-agent
-- passé de 70 % à 60 % ferait apparaître une pièce qui n'a jamais été visée ni signée.
--
-- Une pièce comptable est un justificatif. Elle est donc conservée telle quelle.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_PieceWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_PieceWU
    (
        DateActivite        DATE            NOT NULL,
        Ligne               INT             NOT NULL,

        Compte              NVARCHAR(50)    NOT NULL,
        Libelle             NVARCHAR(255)   NULL,

        -- En FCFA, donc entier : la monnaie n'a pas de décimale, et la pièce est déjà
        -- arrondie à l'unité.
        Debit               BIGINT          NOT NULL DEFAULT (0),
        Credit              BIGINT          NOT NULL DEFAULT (0),

        -- Alimente la colonne ACBRN du fichier destiné au core banking.
        CodeAgence          NVARCHAR(50)    NULL,

        DateEnregistrement  DATETIME        NOT NULL DEFAULT (GETDATE()),
        EnregistrePar       NVARCHAR(100)   NULL,

        CONSTRAINT PK_T_PieceWU PRIMARY KEY (DateActivite, Ligne),

        CONSTRAINT CK_T_PieceWU_UnSeulSens CHECK
            ((Debit <> 0 AND Credit = 0) OR (Credit <> 0 AND Debit = 0))
    );

    CREATE INDEX IX_T_PieceWU_DateActivite ON dbo.T_PieceWU (DateActivite);

    PRINT 'Table T_PieceWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_PieceWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- Annulation d'une comptabilisation — archives et traces
--
-- ANNULER N'EST PAS SUPPRIMER. Une journée retirée voit ses lignes DÉPLACÉES vers les
-- tables d'archive ci-dessous, sous un identifiant qui porte le motif et les deux
-- signatures. Un DELETE effacerait la preuve au moment précis où l'on en a besoin.
--
-- Les tables vivantes (T_HistoriqueWU, T_HistoriqueMTCN, T_PieceWU) ne contiennent ainsi
-- que des journées en vigueur : les rapports d'activité cessent de compter une journée
-- annulée sans qu'on touche à une seule de leurs requêtes.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_AnnulationWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_AnnulationWU
    (
        IdAnnulation        BIGINT IDENTITY(1,1) NOT NULL,

        DateActivite        DATE            NOT NULL,
        Motif               NVARCHAR(30)    NOT NULL,
        Commentaire         NVARCHAR(500)   NULL,

        -- L'application sait qu'elle a PRODUIT le fichier, pas qu'il a été chargé dans les
        -- livres de la banque. Cette colonne garde la réponse de l'agent.
        CoreBankingInjecte  BIT             NOT NULL DEFAULT (0),

        NombrePdv           INT             NOT NULL DEFAULT (0),
        NombreTransactions  INT             NOT NULL DEFAULT (0),
        NombreEcritures     INT             NOT NULL DEFAULT (0),
        TotalDebit          BIGINT          NOT NULL DEFAULT (0),
        TotalCredit         BIGINT          NOT NULL DEFAULT (0),

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

        DateEnregistrement      DATETIME        NULL,
        ComptabilisePar         NVARCHAR(50)    NULL,
        DateComptabilisation    DATETIME        NULL
    );

    CREATE INDEX IX_T_HistoriqueAnnuleWU_Annulation ON dbo.T_HistoriqueAnnuleWU (IdAnnulation);

    PRINT 'Table T_HistoriqueAnnuleWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueAnnuleWU déjà présente : création ignorée.';
END
GO

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

        DateEnregistrement      DATETIME        NULL,
        ComptabilisePar         NVARCHAR(50)    NULL,
        DateComptabilisation    DATETIME        NULL
    );

    CREATE INDEX IX_T_HistoriqueMTCNAnnuleWU_Annulation ON dbo.T_HistoriqueMTCNAnnuleWU (IdAnnulation);
    CREATE INDEX IX_T_HistoriqueMTCNAnnuleWU_MTCN ON dbo.T_HistoriqueMTCNAnnuleWU (MTCN);

    PRINT 'Table T_HistoriqueMTCNAnnuleWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_HistoriqueMTCNAnnuleWU déjà présente : création ignorée.';
END
GO

-- La contrainte CK_T_PieceWU_UnSeulSens n'est pas reprise ici : une archive conserve ce
-- qui a été écrit, elle ne rejuge pas. La refuser ferait échouer l'annulation et
-- laisserait la journée en place, soit l'inverse du but.
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
-- T_FichierCoreBankingWU — trace des fichiers produits pour le core banking
--
-- Le fichier lui-même n'est pas conservé : il dérive entièrement de la pièce. Mais le FAIT
-- de l'avoir produit ne se déduit de rien, et c'est le seul moment où la journée quitte
-- Wincompense pour entrer dans les livres de la banque.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_FichierCoreBankingWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_FichierCoreBankingWU
    (
        IdFichier           BIGINT IDENTITY(1,1) NOT NULL,

        DateActivite        DATE            NOT NULL,
        DateValeur          DATE            NOT NULL,
        NumeroLot           NVARCHAR(10)    NULL,

        NomFichier          NVARCHAR(255)   NULL,
        CheminFichier       NVARCHAR(500)   NULL,

        NombreLignes        INT             NOT NULL DEFAULT (0),
        TotalDebit          BIGINT          NOT NULL DEFAULT (0),
        TotalCredit         BIGINT          NOT NULL DEFAULT (0),

        DateProduction      DATETIME        NOT NULL DEFAULT (GETDATE()),
        ProduitPar          NVARCHAR(50)    NULL,

        CONSTRAINT PK_T_FichierCoreBankingWU PRIMARY KEY (IdFichier)
    );

    CREATE INDEX IX_T_FichierCoreBankingWU_DateActivite
        ON dbo.T_FichierCoreBankingWU (DateActivite, DateProduction DESC);

    PRINT 'Table T_FichierCoreBankingWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_FichierCoreBankingWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- La file des demandes accueille l'annulation
--
--   Cle         = la journée, au format AAAA-MM-JJ
--   Designation = le motif (RAPPORT_VIDE, RAPPORT_ERRONE, …)
--
-- Deux colonnes manquaient : le commentaire libre et la réponse sur le core banking.
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

IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_TypeObjet')
    ALTER TABLE dbo.T_DemandeWU DROP CONSTRAINT CK_T_DemandeWU_TypeObjet;
GO

ALTER TABLE dbo.T_DemandeWU WITH CHECK
    ADD CONSTRAINT CK_T_DemandeWU_TypeObjet
    CHECK (TypeObjet IN ('SOUS_AGENT', 'AGENCE', 'GROUPE', 'COMPTABILISATION'));
GO

IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_Operation')
    ALTER TABLE dbo.T_DemandeWU DROP CONSTRAINT CK_T_DemandeWU_Operation;
GO

ALTER TABLE dbo.T_DemandeWU WITH CHECK
    ADD CONSTRAINT CK_T_DemandeWU_Operation
    CHECK (Operation IN ('CREATION', 'MODIFICATION', 'SUPPRESSION', 'SYNCHRONISATION',
                         'ANNULATION'));
GO

PRINT 'File des demandes ouverte aux annulations de comptabilisation.';
GO

-- =========================================================================
-- Fêtes à date fixe et lundis de Pâques, 2026 à 2030
--
-- Rejouable : seules les dates absentes sont ajoutées, celles que la banque aurait corrigées
-- ne sont pas écrasées.
-- =========================================================================
;WITH Feries(DateFerie, Libelle) AS
(
    SELECT * FROM (VALUES
        ('2026-01-01', N'Jour de l''An'),
        ('2026-04-06', N'Lundi de Pâques'),
        ('2026-05-01', N'Fête du Travail'),
        ('2026-08-11', N'Fête de l''Indépendance'),
        ('2026-11-28', N'Fête de la République'),
        ('2026-12-01', N'Journée de la Liberté et de la Démocratie'),
        ('2026-12-25', N'Noël'),
        ('2027-01-01', N'Jour de l''An'),
        ('2027-03-29', N'Lundi de Pâques'),
        ('2027-05-01', N'Fête du Travail'),
        ('2027-08-11', N'Fête de l''Indépendance'),
        ('2027-11-28', N'Fête de la République'),
        ('2027-12-01', N'Journée de la Liberté et de la Démocratie'),
        ('2027-12-25', N'Noël'),
        ('2028-01-01', N'Jour de l''An'),
        ('2028-04-17', N'Lundi de Pâques'),
        ('2028-05-01', N'Fête du Travail'),
        ('2028-08-11', N'Fête de l''Indépendance'),
        ('2028-11-28', N'Fête de la République'),
        ('2028-12-01', N'Journée de la Liberté et de la Démocratie'),
        ('2028-12-25', N'Noël'),
        ('2029-01-01', N'Jour de l''An'),
        ('2029-04-02', N'Lundi de Pâques'),
        ('2029-05-01', N'Fête du Travail'),
        ('2029-08-11', N'Fête de l''Indépendance'),
        ('2029-11-28', N'Fête de la République'),
        ('2029-12-01', N'Journée de la Liberté et de la Démocratie'),
        ('2029-12-25', N'Noël'),
        ('2030-01-01', N'Jour de l''An'),
        ('2030-04-22', N'Lundi de Pâques'),
        ('2030-05-01', N'Fête du Travail'),
        ('2030-08-11', N'Fête de l''Indépendance'),
        ('2030-11-28', N'Fête de la République'),
        ('2030-12-01', N'Journée de la Liberté et de la Démocratie'),
        ('2030-12-25', N'Noël')
    ) AS Source(DateFerie, Libelle)
)
INSERT INTO dbo.T_JourFerieWU (DateFerie, Libelle, SaisiPar)
SELECT  f.DateFerie, f.Libelle, N'script'
FROM    Feries AS f
WHERE   NOT EXISTS (SELECT 1 FROM dbo.T_JourFerieWU AS t WHERE t.DateFerie = f.DateFerie);

PRINT CAST(@@ROWCOUNT AS NVARCHAR(10)) + ' jour(s) férié(s) ajouté(s).';
GO

/*
    =========================================================================
    À COMPLÉTER CHAQUE ANNÉE : les fêtes musulmanes
    =========================================================================

    Leurs dates sont annoncées par les autorités et ne se calculent pas d'avance. Ajoutez-les
    dès qu'elles sont connues, sur ce modèle :

        INSERT INTO dbo.T_JourFerieWU (DateFerie, Libelle, SaisiPar)
        VALUES ('2026-03-20', N'Aïd el-Fitr',   N'saisie manuelle'),
               ('2026-05-27', N'Aïd el-Kébir',  N'saisie manuelle'),
               ('2026-08-25', N'Mawlid',        N'saisie manuelle');

    Les dates ci-dessus sont des EXEMPLES, à ne pas reprendre telles quelles.

    Tant qu'une année n'a aucun jour férié enregistré, l'application le signale au moment de
    produire le fichier : elle ne se tait pas sur une liste vide.
*/

-- Contrôle : ce qui est enregistré, année par année.
SELECT  Annee = YEAR(DateFerie), Nombre = COUNT(*)
FROM    dbo.T_JourFerieWU
GROUP BY YEAR(DateFerie)
ORDER BY Annee;
GO



/*----------------------------------------------------------------------------------------------
    PARTIE 4 — Rôles de base de données et droits
    (source : Scripts\08_RolesSQLServer.sql)
----------------------------------------------------------------------------------------------*/

/*
    =========================================================================
    Rôles SQL Server de la base GWC_WINCOMPENSE_ETD
    =========================================================================

    L'application contrôle déjà les droits par son propre mot de passe applicatif. Ces rôles
    de base de données en sont le second verrou : ils s'appliquent aux connexions SQL Server
    elles-mêmes, y compris à quelqu'un qui contournerait l'application — un poste relié à la
    base avec SQL Server Management Studio, par exemple.

    Trois rôles, calqués sur ceux de l'application :

      wu_compense    — lit le paramétrage, écrit l'historique et le journal des connexions.
                       Il ne peut PAS modifier les sous-agents, les agences, les groupes
                       ni les comptes comptables : c'est ce paramétrage qui détermine les
                       écritures, le laisser modifiable par celui qui les génère reviendrait
                       à supprimer le contrôle croisé voulu par la banque.

      wu_commercial  — écrit le paramétrage des points de vente, lit l'historique pour
                       consulter les rapports d'activité. Il ne peut pas écrire l'historique :
                       il ne comptabilise rien.

      wu_admin       — les deux, plus les comptes systèmes et la table des utilisateurs.

    AUCUN rôle ne reçoit DELETE sur le journal des connexions : un journal que ses propres
    utilisateurs peuvent effacer ne prouve rien. Seul le propriétaire de la base pourra le
    purger, et cela se verra.

    L'historique, lui, accepte DELETE pour le rôle de compense : rejouer une journée suppose
    d'effacer la précédente version, ce que fait EnregistrerJournee dans une transaction.

    Ce script est rejouable : il ne crée que ce qui manque.

    À exécuter APRÈS 01 à 07, sur la base GWC_WINCOMPENSE_ETD.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. Création des rôles
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_compense' AND type = 'R')
BEGIN
    CREATE ROLE wu_compense;
    PRINT 'Rôle wu_compense créé.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_commercial' AND type = 'R')
BEGIN
    CREATE ROLE wu_commercial;
    PRINT 'Rôle wu_commercial créé.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_admin' AND type = 'R')
BEGIN
    CREATE ROLE wu_admin;
    PRINT 'Rôle wu_admin créé.';
END
GO

-- =========================================================================
-- 2. Droits du rôle wu_compense
--
--    Lecture du paramétrage, écriture de l'historique.
-- =========================================================================
GRANT SELECT ON dbo.T_Pdv_SA            TO wu_compense;
GRANT SELECT ON dbo.T_Pdv_EC            TO wu_compense;
GRANT SELECT ON dbo.SystemeWU           TO wu_compense;

GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueWU   TO wu_compense;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueMTCN TO wu_compense;

-- Se connecter suppose de lire sa propre fiche et d'écrire au journal ; le compteur d'échecs
-- et la date de dernière connexion sont mis à jour par l'application elle-même.
GRANT SELECT, UPDATE ON dbo.T_UtilisateurWU TO wu_compense;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU   TO wu_compense;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_GroupeStatistique TO wu_compense');
END
GO

-- =========================================================================
-- 3. Droits du rôle wu_commercial
--
--    Écriture du paramétrage des points de vente, lecture de l'historique.
-- =========================================================================
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_SA TO wu_commercial;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_EC TO wu_commercial;

GRANT SELECT ON dbo.T_HistoriqueWU   TO wu_commercial;
GRANT SELECT ON dbo.T_HistoriqueMTCN TO wu_commercial;

-- Les comptes comptables se consultent — un commercial doit pouvoir vérifier à quel compte
-- un groupe renvoie — mais ne se modifient pas.
GRANT SELECT ON dbo.SystemeWU TO wu_commercial;

GRANT SELECT, UPDATE ON dbo.T_UtilisateurWU TO wu_commercial;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU   TO wu_commercial;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_GroupeStatistique TO wu_commercial');
END
GO

-- =========================================================================
-- 4. Droits du rôle wu_admin
--
--    Tout ce que font les deux autres, plus les comptes systèmes et les utilisateurs.
-- =========================================================================
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_SA         TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_EC         TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueWU   TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueMTCN TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_UtilisateurWU  TO wu_admin;

GRANT SELECT, UPDATE ON dbo.SystemeWU     TO wu_admin;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU TO wu_admin;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_GroupeStatistique TO wu_admin');
END
GO

-- =========================================================================
-- 4 bis. Tables créées par les scripts 09 et 10
--
--    Elles n'existent pas encore lorsque ce script est exécuté dans l'ordre numérique :
--    les GRANT sont donc protégés par IF EXISTS et ne s'appliquent qu'au second passage,
--    ou lorsque ce script est exécuté en dernier (voir 00_InstallationComplete.sql).
--
--    SANS CES DROITS, l'application échoue là où on ne l'attend pas : le calendrier des
--    jours fériés devient illisible au moment de dater le fichier core banking, et la file
--    du double regard reste vide alors qu'elle contient des demandes.
--
--    RELANCEZ CE SCRIPT APRÈS 09 ET 10 si vous exécutez les scripts un par un.
-- =========================================================================

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_JourFerieWU')
BEGIN
    -- Lecture pour tous : dater une écriture suppose de connaître les jours chômés,
    -- quel que soit le rôle. L'écriture reste réservée à l'administrateur.
    EXEC('GRANT SELECT ON dbo.T_JourFerieWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.T_JourFerieWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_JourFerieWU TO wu_admin');
    PRINT 'Droits accordés sur T_JourFerieWU.';
END
ELSE
BEGIN
    PRINT 'T_JourFerieWU absente : relancez ce script après 10_JoursFeries.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PieceWU')
BEGIN
    -- DELETE pour le rôle de compense, comme sur l'historique : rejouer une journée suppose
    -- d'effacer la version précédente, ce que fait EnregistrerJournee dans une transaction.
    -- Le commercial lit, il ne comptabilise rien.
    EXEC('GRANT SELECT, INSERT, DELETE ON dbo.T_PieceWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.T_PieceWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_PieceWU TO wu_admin');
    PRINT 'Droits accordés sur T_PieceWU.';
END
ELSE
BEGIN
    PRINT 'T_PieceWU absente : relancez ce script après 13_PiecesComptables.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_AnnulationWU')
BEGIN
    -- Les archives d'annulation : SELECT et INSERT pour les trois rôles, JAMAIS de DELETE.
    -- L'INSERT est accordé largement parce que la fonction d'authorizer se porte
    -- indifféremment sur un compte de compense, de commercial ou d'administration ; le
    -- double regard, lui, est posé par la contrainte CK_T_AnnulationWU_PasSoiMeme et par
    -- l'application, non par les droits SQL.
    --
    -- Une archive que ses propres utilisateurs peuvent effacer ne prouve rien : c'est
    -- exactement quand une journée est annulée que l'auditeur veut la retrouver.
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_admin');

    PRINT 'Droits accordés sur les tables d''annulation.';
END
ELSE
BEGIN
    PRINT 'T_AnnulationWU absente : relancez ce script après 14_AnnulationComptabilisation.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_FichierCoreBankingWU')
BEGIN
    -- Le rôle de compense produit les fichiers : il écrit. Le commercial lit. Aucun DELETE :
    -- une trace que l'on peut effacer ne prouve rien.
    EXEC('GRANT SELECT, INSERT ON dbo.T_FichierCoreBankingWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.T_FichierCoreBankingWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_FichierCoreBankingWU TO wu_admin');
    PRINT 'Droits accordés sur T_FichierCoreBankingWU.';
END
ELSE
BEGIN
    PRINT 'T_FichierCoreBankingWU absente : relancez ce script après 15_FichierCoreBanking.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_DemandeWU')
BEGIN
    -- La file du double regard portait d'abord le seul paramétrage, et l'agent de compense
    -- n'y avait aucun accès. Elle porte maintenant AUSSI les demandes d'annulation d'une
    -- journée comptabilisée, qui appartiennent, elles, à la compense : le rôle doit donc
    -- pouvoir y déposer, y lire, et y décider s'il porte la fonction d'authorizer.
    --
    -- Ce qui protège le référentiel n'est pas l'absence de ce droit SQL, mais la fonction
    -- INPUTER / AUTHORIZER portée par l'utilisateur et la contrainte
    -- CK_T_DemandeWU_PasSoiMeme, qu'un UPDATE fait à la main ne contourne pas davantage.
    --
    -- Aucun rôle ne reçoit DELETE : une demande rejetée se conserve, c'est elle qui prouve
    -- qu'un contrôle a eu lieu.
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_admin');
    PRINT 'Droits accordés sur T_DemandeWU.';
END
ELSE
BEGIN
    PRINT 'T_DemandeWU absente : relancez ce script après 09_Demandes.sql.';
END
GO

-- =========================================================================
-- 5. Accès du compte applicatif
--
--    C'est ici que se règle l'erreur « Cannot open database ... requested by the login »
--    (4060), celle que l'application affiche quand le login existe sur le serveur mais que
--    la base ne le connaît pas encore.
--
--    Trois niveaux, qu'on confond facilement :
--      le LOGIN ouvre la porte du bâtiment  — créé par la banque, avec son mot de passe ;
--      l'UTILISATEUR ouvre celle du bureau  — créé ci-dessous ;
--      le RÔLE dit ce qu'on a le droit d'y faire — accordé ci-dessous.
--
--    Ce script ne crée PAS le login : il faudrait son mot de passe, qui n'a pas sa place
--    dans un fichier qui circule. Si le login manque, le script le dit et s'arrête là.
-- =========================================================================

-- ---- LA SEULE LIGNE À ADAPTER -------------------------------------------
DECLARE @compteApplicatif SYSNAME = N'etdwincompense';

-- Rôle accordé. wu_admin parce qu'un compte unique, employé par tous les postes, doit porter
-- la réunion des droits de tous les postes. Si la banque fournit un compte par agent,
-- remplacer par wu_compense (agent de compense) ou wu_commercial (saisie des points de vente)
-- et exécuter cette partie une fois par compte : les trois rôles retrouvent alors leur
-- utilité, et un accès direct à la base par SSMS reste borné au métier réel de chacun.
DECLARE @roleApplicatif SYSNAME = N'wu_admin';

DECLARE @ordreAcces NVARCHAR(400);

IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @compteApplicatif)
BEGIN
    PRINT N'ARRET PARTIE 5 : le login ' + @compteApplicatif + N' n''existe pas sur ce serveur.';
    PRINT N'                Creez-le d''abord avec son mot de passe, puis reexecutez ce script :';
    PRINT N'                    CREATE LOGIN ' + QUOTENAME(@compteApplicatif) + N' WITH PASSWORD = N''...'';';
END
ELSE
BEGIN

    IF DATABASE_PRINCIPAL_ID(@compteApplicatif) IS NULL
    BEGIN
        SET @ordreAcces = N'CREATE USER ' + QUOTENAME(@compteApplicatif) +
                          N' FOR LOGIN ' + QUOTENAME(@compteApplicatif);
        EXEC sp_executesql @ordreAcces;
        PRINT N'Utilisateur cree dans la base : ' + @compteApplicatif;
    END
    ELSE
    BEGIN
        PRINT N'Deja en place : utilisateur ' + @compteApplicatif;
    END

    -- ISNULL : si la creation de l'utilisateur a echoue, IS_ROLEMEMBER rend NULL, et un
    -- NULL = 0 vaut « inconnu », donc faux : on passerait dans la branche « deja en place »
    -- sans que rien ne le soit.
    IF ISNULL(IS_ROLEMEMBER(@roleApplicatif, @compteApplicatif), 0) = 0
    BEGIN
        SET @ordreAcces = N'ALTER ROLE ' + QUOTENAME(@roleApplicatif) +
                          N' ADD MEMBER ' + QUOTENAME(@compteApplicatif);
        EXEC sp_executesql @ordreAcces;
        PRINT N'Role ' + @roleApplicatif + N' accorde a ' + @compteApplicatif;
    END
    ELSE
    BEGIN
        PRINT N'Deja en place : ' + @compteApplicatif + N' dans ' + @roleApplicatif;
    END
END
GO

-- =========================================================================
-- 6. Contrôle : droits effectivement accordés, par rôle et par table
-- =========================================================================
SELECT  principal    = dp.name,
        objet        = OBJECT_NAME(pe.major_id),
        droit        = pe.permission_name,
        etat         = pe.state_desc
FROM    sys.database_permissions AS pe
        INNER JOIN sys.database_principals AS dp ON dp.principal_id = pe.grantee_principal_id
WHERE   dp.name IN (N'wu_compense', N'wu_commercial', N'wu_admin')
ORDER BY dp.name, OBJECT_NAME(pe.major_id), pe.permission_name;
GO


/*----------------------------------------------------------------------------------------------
    PARTIE 6 — COMPTE RENDU
    Ce que le script a réellement produit, et ce qu'il reste à faire.
----------------------------------------------------------------------------------------------*/

USE GWC_WINCOMPENSE_ETD;
GO

PRINT '';
PRINT '================================================================================';
PRINT '  WINCOMPENSE TCHAD - COMPTE RENDU D''INSTALLATION';
PRINT '================================================================================';
GO

-- ---------------------------------------------------------------------------------------------
-- 6.1  Les onze tables attendues
-- ---------------------------------------------------------------------------------------------
PRINT '';
PRINT '--- Tables ---';
GO

;WITH Attendues AS (
    SELECT nom = N'T_Pdv_SA',            libelle = N'Sous-agents'                    UNION ALL
    SELECT N'T_Pdv_EC',                  N'Agences propres'                       UNION ALL
    SELECT N'SystemeWU',                 N'Comptes comptables parametres'         UNION ALL
    SELECT N'T_GroupeStatistique',       N'Groupes statistiques'                  UNION ALL
    SELECT N'T_HistoriqueWU',            N'Journees comptabilisees'               UNION ALL
    SELECT N'T_HistoriqueMTCN',          N'Transactions, MTCN par MTCN'           UNION ALL
    SELECT N'T_UtilisateurWU',           N'Comptes utilisateurs'                  UNION ALL
    SELECT N'T_ConnexionWU',             N'Journal des connexions'                UNION ALL
    SELECT N'T_DemandeWU',               N'Double regard : file des demandes'     UNION ALL
    SELECT N'T_JourFerieWU',             N'Jours feries'                          UNION ALL
    SELECT N'T_PieceWU',                 N'Pieces comptables conservees'
)
SELECT  [Table]   = a.nom,
        [Role]    = a.libelle,
        [Etat]    = CASE WHEN t.name IS NULL THEN '*** ABSENTE ***' ELSE 'creee' END,
        [Lignes]  = ISNULL((SELECT SUM(p.rows) FROM sys.partitions p
                            WHERE p.object_id = t.object_id AND p.index_id IN (0,1)), 0)
FROM    Attendues AS a
        LEFT JOIN sys.tables AS t ON t.name = a.nom
ORDER BY CASE WHEN t.name IS NULL THEN 0 ELSE 1 END, a.nom;
GO

-- ---------------------------------------------------------------------------------------------
-- 6.2  Les trois roles et le nombre de droits accordes
-- ---------------------------------------------------------------------------------------------
PRINT '';
PRINT '--- Roles ---';
GO

;WITH Roles AS (
    SELECT nom = N'wu_compense'   UNION ALL
    SELECT N'wu_commercial'       UNION ALL
    SELECT N'wu_admin'
)
SELECT  [Role]     = r.nom,
        [Etat]     = CASE WHEN dp.name IS NULL THEN '*** ABSENT ***' ELSE 'cree' END,
        [Droits]   = ISNULL((SELECT COUNT(*) FROM sys.database_permissions pe
                             WHERE pe.grantee_principal_id = dp.principal_id), 0),
        [Membres]  = ISNULL((SELECT COUNT(*) FROM sys.database_role_members m
                             WHERE m.role_principal_id = dp.principal_id), 0)
FROM    Roles AS r
        LEFT JOIN sys.database_principals AS dp ON dp.name = r.nom AND dp.type = 'R'
ORDER BY r.nom;
GO

-- ---------------------------------------------------------------------------------------------
-- 6.3  Ce qu'il reste a faire
--
--      Passe par EXEC et sous condition d'existence : un lot ad hoc resout les noms de tables
--      a la compilation, et une seule table absente ferait echouer tout le lot - donc
--      disparaitre le compte rendu au moment precis ou il sert.
-- ---------------------------------------------------------------------------------------------
PRINT '';
PRINT '--- Reste a faire ---';
GO

IF OBJECT_ID(N'dbo.SystemeWU')       IS NOT NULL
   AND OBJECT_ID(N'dbo.T_JourFerieWU')   IS NOT NULL
   AND OBJECT_ID(N'dbo.T_Pdv_SA')        IS NOT NULL
   AND OBJECT_ID(N'dbo.T_Pdv_EC')        IS NOT NULL
   AND OBJECT_ID(N'dbo.T_UtilisateurWU') IS NOT NULL
BEGIN
    EXEC(N'SELECT  [#]       = ordre,
        [Sujet]   = sujet,
        [Etat]    = CASE WHEN fait = 1 THEN ''fait'' ELSE ''A FAIRE'' END,
        [Comment] = quoi
FROM (
    SELECT  ordre = 1,
            sujet = N''Compte applicatif rattache a la base'',
            fait  = CASE WHEN EXISTS (SELECT 1 FROM sys.database_role_members m
                                      INNER JOIN sys.database_principals r
                                              ON r.principal_id = m.role_principal_id
                                      WHERE r.name IN (N''wu_compense'', N''wu_commercial'', N''wu_admin''))
                         THEN 1 ELSE 0 END,
            quoi  = N''PARTIE 5 de ce script : renseigner @compteApplicatif, puis reexecuter''
    UNION ALL
    SELECT  2, N''Comptes comptables renseignes'',
            CASE WHEN EXISTS (SELECT 1 FROM dbo.SystemeWU) THEN 1 ELSE 0 END,
            N''Application, ecran Comptes systemes WU - ils determinent toute la piece comptable''
    UNION ALL
    SELECT  3, N''Jours feries enregistres pour l''''annee en cours'',
            CASE WHEN EXISTS (SELECT 1 FROM dbo.T_JourFerieWU WHERE YEAR(DateFerie) = YEAR(GETDATE()))
                 THEN 1 ELSE 0 END,
            N''Ajouter les fetes musulmanes : elles ne se calculent pas d''''avance''
    UNION ALL
    SELECT  4, N''Sous-agents charges'',
            CASE WHEN EXISTS (SELECT 1 FROM dbo.T_Pdv_SA) THEN 1 ELSE 0 END,
            N''Application, ecran Sous-agents - avec compte de compensation ET compte de commission''
    UNION ALL
    SELECT  5, N''Agences propres chargees'',
            CASE WHEN EXISTS (SELECT 1 FROM dbo.T_Pdv_EC) THEN 1 ELSE 0 END,
            N''Application, ecran Agences propres''
    UNION ALL
    SELECT  6, N''Premier administrateur cree'',
            CASE WHEN EXISTS (SELECT 1 FROM dbo.T_UtilisateurWU WHERE Role = ''''ADMIN'''' AND Actif = 1)
                 THEN 1 ELSE 0 END,
            N''Se cree au premier demarrage de l''''application''
) AS Controles
ORDER BY fait, ordre;');
END
ELSE
BEGIN
    PRINT 'Des tables manquent : voir le tableau des tables ci-dessus.';
    PRINT 'Le detail du reste a faire ne peut pas etre etabli tant qu''elles ne sont pas creees.';
END
GO

-- ---------------------------------------------------------------------------------------------
-- 6.4  Ou l'application devra pointer
-- ---------------------------------------------------------------------------------------------
PRINT '';
PRINT '--- Chaine de connexion a inscrire dans App.config ---';
GO

SELECT  [Serveur]  = @@SERVERNAME,
        [Base]     = DB_NAME(),
        [Chaine]   = 'Server=' + @@SERVERNAME + ';Database=' + DB_NAME()
                     + ';Integrated Security=True;Connect Timeout=10;';
GO

PRINT '';
PRINT '================================================================================';
PRINT '  Fin. Lisez les trois tableaux ci-dessus avant de declarer l''installation faite.';
PRINT '================================================================================';
GO
