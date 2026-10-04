/*
    =========================================================================
    Les écarts de change — gains et pertes, et leur pièce comptable
    =========================================================================

    POURQUOI CES TABLES

    Western Union règle la banque en EUROS ; le guichet encaisse en FRANCS CFA. Converti à
    la parité fixe, le montant réglé ne retombe pas exactement sur le montant encaissé :
    d'un ou deux francs sur une transaction sans conversion, de plusieurs milliers sur un
    gros envoi. Cette différence est un gain ou une perte de change, et elle doit être
    comptabilisée pour elle-même, sur une pièce SÉPARÉE de celle du principal et des
    commissions.

    Sur le rapport de référence — trois journées, 2 395 lignes — cela représente 555 849 F
    de gains et 1 544 F de pertes, soit 554 305 F nets. Ce n'est pas un détail d'arrondi.

    CE QUE CE SCRIPT AJOUTE

      1. Deux colonnes sur SystemeWU : les comptes de gain et de perte de change.
      2. T_PariteChangeWU   : la parité, historisée avec sa date d'effet.
      3. T_EcartChangeWU    : le détail, transaction par transaction.
      4. T_ExclusionChangeWU : les lignes écartées du calcul, avec leur motif.
      5. T_PieceChangeWU    : les pièces de change produites, conservées.

    LA COLONNE Cpte_Gainde_Change EXISTAIT DÉJÀ, ET N'ÉTAIT LUE PAR PERSONNE

    La table SystemeWU la porte depuis l'origine, à NULL. La banque avait réservé la place.
    Ce script ne la crée donc que si elle manque, et n'y écrit RIEN : c'est à la Direction
    Comptable de la renseigner depuis l'écran « Comptes systèmes ». Il n'existait en revanche
    aucune colonne pour la PERTE de change : celle-là est ajoutée.

    AUCUN COMPTE N'EST AMORCÉ, ET C'EST VOULU

    Un numéro de compte inventé par un script d'installation est un numéro de compte qui
    finira par être comptabilisé. Tant que les deux comptes ne sont pas saisis, l'application
    REFUSE de produire la pièce de change et dit lequel manque. Elle continue en revanche de
    produire la pièce principale comme avant : les deux paramétrages sont indépendants.

    À exécuter APRÈS 00 (ou 01 à 20), sur la base GWC_WINCOMPENSE_ETD.
    Ce script est REJOUABLE : il crée ce qui manque et ne réécrit rien de ce qui est saisi.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. Les deux comptes, sur la table de paramétrage existante
--
--    Pas de nouvelle table : l'écran « Comptes systèmes » que la Direction Comptable
--    connaît déjà reçoit deux champs de plus. Un paramétrage éclaté sur deux tables
--    aurait fini par diverger.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'SystemeWU' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    RAISERROR(N'Table SystemeWU absente : exécutez d''abord 00_InstallationComplete.sql (ou 03_SystemeWU.sql).', 16, 1);
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'SystemeWU' AND schema_id = SCHEMA_ID(N'dbo'))
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'dbo.SystemeWU') AND name = N'Cpte_Gainde_Change')
BEGIN
    ALTER TABLE dbo.SystemeWU ADD Cpte_Gainde_Change NVARCHAR(255) NULL;
    PRINT 'Colonne Cpte_Gainde_Change ajoutée à SystemeWU.';
END
ELSE
    PRINT 'Colonne Cpte_Gainde_Change déjà présente sur SystemeWU (elle existe depuis l''origine).';
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'SystemeWU' AND schema_id = SCHEMA_ID(N'dbo'))
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'dbo.SystemeWU') AND name = N'Cpte_Pertede_Change')
BEGIN
    ALTER TABLE dbo.SystemeWU ADD Cpte_Pertede_Change NVARCHAR(255) NULL;
    PRINT 'Colonne Cpte_Pertede_Change ajoutée à SystemeWU.';
END
ELSE
    PRINT 'Colonne Cpte_Pertede_Change déjà présente sur SystemeWU.';
GO

-- =========================================================================
-- 2. La parité, historisée
--
--    POURQUOI UNE HISTORIQUE ET NON UNE VALEUR. La parité EUR/XAF est fixe depuis 1999 et
--    ne bougera probablement jamais. Mais une pièce comptable conservée doit pouvoir être
--    relue dans dix ans AVEC LA PARITÉ DE SON JOUR, et non avec celle du jour où on la
--    relit. Le jour où la parité changerait, les pièces anciennes resteraient justes.
--
--    La parité employée est en outre recopiée sur CHAQUE ligne de détail : l'historique dit
--    laquelle s'appliquait, le détail dit laquelle a servi. Les deux doivent concorder, et
--    c'est vérifiable.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PariteChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE dbo.T_PariteChangeWU
    (
        Id                  INT             IDENTITY(1,1) PRIMARY KEY,

        -- Premier jour où cette parité s'applique. La parité en vigueur pour une journée est
        -- celle dont la date d'effet est la plus récente parmi celles qui la précèdent.
        DateEffet           DATE            NOT NULL,

        DeviseSource        NVARCHAR(3)     NOT NULL DEFAULT (N'EUR'),
        DeviseCible         NVARCHAR(3)     NOT NULL DEFAULT (N'XAF'),

        -- Six décimales : 655,957 en a trois, et une parité non fixe en demanderait plus.
        Parite              DECIMAL(18,6)   NOT NULL,

        Commentaire         NVARCHAR(255)   NULL,

        CreePar             NVARCHAR(50)    NULL,
        DateCreation        DATETIME        NULL,

        -- Une parité nulle ou négative rendrait toutes les contre-valeurs nulles, et donc le
        -- gain de change égal au chiffre d'affaires. La base refuse, elle n'avertit pas.
        CONSTRAINT CK_T_PariteChangeWU_Parite CHECK (Parite > 0),

        CONSTRAINT UQ_T_PariteChangeWU_Effet
            UNIQUE (DeviseSource, DeviseCible, DateEffet)
    );

    PRINT 'Table T_PariteChangeWU créée.';
END
ELSE
    PRINT 'Table T_PariteChangeWU déjà présente : création ignorée.';
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PariteChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
    AND NOT EXISTS (SELECT 1 FROM dbo.T_PariteChangeWU WHERE DeviseSource = N'EUR' AND DeviseCible = N'XAF')
BEGIN
    INSERT INTO dbo.T_PariteChangeWU
        (DateEffet, DeviseSource, DeviseCible, Parite, Commentaire, CreePar, DateCreation)
    VALUES
        ('1999-01-01', N'EUR', N'XAF', 655.957,
         N'Parité fixe franc CFA / euro, inchangée depuis l''introduction de l''euro.',
         N'installation', GETDATE());

    PRINT 'Parité EUR/XAF 655,957 amorcée au 01/01/1999.';
END
ELSE
    PRINT 'Parité EUR/XAF déjà présente : amorçage ignoré.';
GO

-- =========================================================================
-- 3. Le détail, transaction par transaction
--
--    POURQUOI CONSERVER LE DÉTAIL ET PAS SEULEMENT LES TOTAUX. La banque conteste des
--    LIGNES, pas des totaux : « et le transfert 0050908785, vous trouvez quoi ? ». Sans le
--    détail conservé, répondre exige de retrouver le rapport du jour, de le recharger et de
--    refaire le calcul — en espérant que le paramétrage n'ait pas changé entre-temps.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_EcartChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE dbo.T_EcartChangeWU
    (
        Id                  BIGINT          IDENTITY(1,1) PRIMARY KEY,

        Mtcn                NVARCHAR(30)    NOT NULL,
        Account             NVARCHAR(20)    NULL,       -- point de vente, pour la narrative
        DateReglement       DATE            NOT NULL,
        Sens                NCHAR(1)        NOT NULL,   -- S envoi, P paiement
        CodeProduit         NVARCHAR(10)    NULL,       -- IMTR, FTSS, AVSS...
        Statut              NVARCHAR(10)    NULL,       -- TxnStatus : S réglée, W en attente
        DeviseLocale        NVARCHAR(10)    NULL,       -- LOCCurrencyCode

        -- Les montants. Quatre décimales : c'est la précision du rapport Western Union.
        MontantLocal        DECIMAL(19,4)   NOT NULL,
        ClearPrincipalLoc   DECIMAL(19,4)   NOT NULL,
        ClearFxLoc          DECIMAL(19,4)   NOT NULL,
        MontantEnDevise     DECIMAL(19,4)   NOT NULL,
        Parite              DECIMAL(18,6)   NOT NULL,

        -- Contre-valeur et écart sont RECOPIÉS et non calculés par la base. Ils sont produits
        -- par le code, qui applique l'arrondi au franc à la transaction ; les recalculer ici
        -- avec une autre règle d'arrondi donnerait deux vérités pour une même ligne.
        ContreValeur        DECIMAL(19,4)   NOT NULL,
        Ecart               DECIMAL(19,4)   NOT NULL,
        Nature              NVARCHAR(10)    NOT NULL,   -- Gain, Perte, Neutre

        FichierSource       NVARCHAR(260)   NULL,

        DateEnregistrement  DATETIME        NULL,
        EnregistrePar       NVARCHAR(50)    NULL,

        CONSTRAINT CK_T_EcartChangeWU_Sens CHECK (Sens IN (N'S', N'P')),
        CONSTRAINT CK_T_EcartChangeWU_Parite CHECK (Parite > 0),

        -- LE GARDE-FOU DU DOUBLE CHARGEMENT. Un rapport rechargé par erreur doublerait le
        -- gain de change de la journée, et rien ne le signalerait : les deux pièces seraient
        -- équilibrées. La base refuse donc deux fois la même transaction, et c'est la base
        -- qui refuse — pas un contrôle applicatif qu'un correctif pressé pourrait contourner.
        CONSTRAINT UQ_T_EcartChangeWU_Transaction
            UNIQUE (Mtcn, DateReglement, Sens)
    );

    CREATE INDEX IX_T_EcartChangeWU_Journee
        ON dbo.T_EcartChangeWU (DateReglement, Sens, CodeProduit);

    PRINT 'Table T_EcartChangeWU créée.';
END
ELSE
    PRINT 'Table T_EcartChangeWU déjà présente : création ignorée.';
GO

-- =========================================================================
-- 4. Les lignes écartées
--
--    Une exclusion muette est une exclusion indéfendable. Vingt-huit lignes du rapport de
--    référence n'entrent pas dans le calcul : la banque doit pouvoir lire lesquelles et
--    pourquoi, et reconnaître ses propres remboursements.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_ExclusionChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE dbo.T_ExclusionChangeWU
    (
        Id                  BIGINT          IDENTITY(1,1) PRIMARY KEY,

        FichierSource       NVARCHAR(260)   NULL,
        NumeroDeLigne       INT             NOT NULL,
        Mtcn                NVARCHAR(30)    NULL,
        Motif               NVARCHAR(255)   NOT NULL,
        Detail              NVARCHAR(255)   NULL,

        DateEnregistrement  DATETIME        NULL,
        EnregistrePar       NVARCHAR(50)    NULL
    );

    CREATE INDEX IX_T_ExclusionChangeWU_Fichier
        ON dbo.T_ExclusionChangeWU (FichierSource, NumeroDeLigne);

    PRINT 'Table T_ExclusionChangeWU créée.';
END
ELSE
    PRINT 'Table T_ExclusionChangeWU déjà présente : création ignorée.';
GO

-- =========================================================================
-- 5. Les pièces de change conservées
--
--    MÊME FORME QUE LA PIÈCE PRINCIPALE — compte, libellé, débit, crédit, code agence. Ce
--    n'est pas une coïncidence : c'est ce qui permet à l'export Excel, au fichier core
--    banking et au contrôle d'équilibre déjà écrits de la traiter sans une ligne de plus.
--
--    LA CLÉ DE GROUPE, ET POURQUOI ELLE EST UNE CHAÎNE. Le découpage des pièces est un
--    réglage : une pièce par journée, par journée et sens, ou par journée, sens et produit.
--    Trois colonnes de clé auraient été vides selon le réglage, et l'unicité n'aurait plus
--    rien protégé. Une clé composée une fois par le code — « 2026-03-27 » ou
--    « 2026-03-27|S|IMTR » — reste unique dans les trois cas.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PieceChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE dbo.T_PieceChangeWU
    (
        Id                  BIGINT          IDENTITY(1,1) PRIMARY KEY,

        CleGroupe           NVARCHAR(60)    NOT NULL,
        DateReglement       DATE            NOT NULL,
        Sens                NVARCHAR(10)    NULL,       -- vide si le découpage ne distingue pas le sens
        CodeProduit         NVARCHAR(10)    NULL,       -- vide si le découpage ne distingue pas le produit
        Account             NVARCHAR(20)    NULL,       -- renseigné si la pièce ne porte qu'un point de vente
        Mtcn                NVARCHAR(30)    NULL,       -- renseigné si la pièce ne porte qu'une transaction

        Ligne               INT             NOT NULL,
        Compte              NVARCHAR(50)    NOT NULL,
        Libelle             NVARCHAR(255)   NOT NULL,
        Debit               BIGINT          NOT NULL DEFAULT (0),
        Credit              BIGINT          NOT NULL DEFAULT (0),
        CodeAgence          NVARCHAR(10)    NULL,

        NombreTransactions  INT             NOT NULL DEFAULT (0),
        Parite              DECIMAL(18,6)   NOT NULL,
        FichierSource       NVARCHAR(260)   NULL,

        DateEnregistrement  DATETIME        NULL,
        EnregistrePar       NVARCHAR(50)    NULL,

        -- Une écriture sans montant n'a rien à faire dans une pièce ; une écriture qui serait
        -- à la fois au débit et au crédit n'a aucun sens comptable.
        CONSTRAINT CK_T_PieceChangeWU_Montant
            CHECK ((Debit > 0 AND Credit = 0) OR (Credit > 0 AND Debit = 0)),

        CONSTRAINT UQ_T_PieceChangeWU_Ligne UNIQUE (CleGroupe, Ligne)
    );

    CREATE INDEX IX_T_PieceChangeWU_Journee
        ON dbo.T_PieceChangeWU (DateReglement);

    PRINT 'Table T_PieceChangeWU créée.';
END
ELSE
    PRINT 'Table T_PieceChangeWU déjà présente : création ignorée.';
GO

-- =========================================================================
-- 5 bis. Les colonnes de la narrative, sur une base où le script est rejoué
--
--    LA BANQUE A DEMANDÉ QUE LE MTCN ET LE ACCOUNT RESSORTENT DANS LA NARRATIVE. Les tables
--    créées plus haut les portent ; celles d'une base où ce script a déjà été exécuté avant
--    cette demande ne les portent pas. Ces trois ALTER les ajoutent sans toucher aux données.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_EcartChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'dbo.T_EcartChangeWU') AND name = N'Account')
BEGIN
    ALTER TABLE dbo.T_EcartChangeWU ADD Account NVARCHAR(20) NULL;
    PRINT 'Colonne Account ajoutée à T_EcartChangeWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PieceChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'dbo.T_PieceChangeWU') AND name = N'Account')
BEGIN
    ALTER TABLE dbo.T_PieceChangeWU ADD Account NVARCHAR(20) NULL;
    PRINT 'Colonne Account ajoutée à T_PieceChangeWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_PieceChangeWU' AND schema_id = SCHEMA_ID(N'dbo'))
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'dbo.T_PieceChangeWU') AND name = N'Mtcn')
BEGIN
    ALTER TABLE dbo.T_PieceChangeWU ADD Mtcn NVARCHAR(30) NULL;
    PRINT 'Colonne Mtcn ajoutée à T_PieceChangeWU.';
END
GO

-- =========================================================================
-- 6. Droits
--
--    POSÉS ICI, JUSTE APRÈS LES CRÉATIONS, ET C'EST UNE LEÇON PAYÉE EN PRODUCTION. Le
--    script 08 pose les droits de tout le référentiel ; exécuté avant les scripts qui
--    créent les tables, il les saute en silence — un PRINT que personne ne lit. La banque
--    l'a découvert sur un « INSERT permission was denied » en pleine journée de travail.
--    Chaque table nouvelle porte donc désormais ses droits avec elle.
--
--    QUI PEUT QUOI :
--      wu_compense   lit et écrit le détail, les exclusions et les pièces — c'est lui qui
--                    traite la journée. Il ne SUPPRIME rien : une correction passe par une
--                    annulation tracée, jamais par un DELETE.
--      wu_commercial lit, et rien de plus.
--      wu_admin      lit tout, écrit le paramétrage de la parité, et peut supprimer un
--                    détail ou une pièce — pour défaire un double chargement, et pour cela
--                    seulement.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_PariteChangeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_EcartChangeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_ExclusionChangeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceChangeWU TO wu_compense');
    PRINT 'Droits accordés à wu_compense.';
END
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_PariteChangeWU TO wu_commercial');
    EXEC('GRANT SELECT ON dbo.T_EcartChangeWU TO wu_commercial');
    EXEC('GRANT SELECT ON dbo.T_ExclusionChangeWU TO wu_commercial');
    EXEC('GRANT SELECT ON dbo.T_PieceChangeWU TO wu_commercial');
    PRINT 'Droits accordés à wu_commercial.';
END
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_PariteChangeWU TO wu_admin');
    EXEC('GRANT SELECT, INSERT, DELETE ON dbo.T_EcartChangeWU TO wu_admin');
    EXEC('GRANT SELECT, INSERT, DELETE ON dbo.T_ExclusionChangeWU TO wu_admin');
    EXEC('GRANT SELECT, INSERT, DELETE ON dbo.T_PieceChangeWU TO wu_admin');
    PRINT 'Droits accordés à wu_admin.';
END
GO

-- =========================================================================
-- 7. Compte rendu
-- =========================================================================
SELECT  N'Comptes de change' AS Objet,
        CASE WHEN EXISTS (SELECT 1 FROM sys.columns
                          WHERE object_id = OBJECT_ID(N'dbo.SystemeWU') AND name = N'Cpte_Gainde_Change')
             THEN N'colonne présente' ELSE N'COLONNE ABSENTE' END AS Gain,
        CASE WHEN EXISTS (SELECT 1 FROM sys.columns
                          WHERE object_id = OBJECT_ID(N'dbo.SystemeWU') AND name = N'Cpte_Pertede_Change')
             THEN N'colonne présente' ELSE N'COLONNE ABSENTE' END AS Perte;
GO

SELECT  DateEffet, DeviseSource, DeviseCible, Parite, CreePar, DateCreation
FROM    dbo.T_PariteChangeWU
ORDER BY DeviseSource, DeviseCible, DateEffet DESC;
GO

PRINT '';
PRINT '=========================================================================';
PRINT 'Écarts de change : tables posées, parité amorcée, droits accordés.';
PRINT '';
PRINT 'IL RESTE UNE CHOSE À FAIRE, ET ELLE N''EST PAS DANS CE SCRIPT :';
PRINT 'saisir les deux comptes dans Paramétrage > Comptes systèmes —';
PRINT '    « Gain de change »  et  « Perte de change ».';
PRINT '';
PRINT 'Tant qu''ils sont vides, l''application refuse de produire la pièce de';
PRINT 'change et dit lequel manque. La pièce principale, elle, n''est pas';
PRINT 'affectée : les deux paramétrages sont indépendants.';
PRINT '=========================================================================';
GO
