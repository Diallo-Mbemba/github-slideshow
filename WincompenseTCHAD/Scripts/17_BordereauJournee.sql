/*
    =========================================================================
    Bordereau de fin de journée — en-tête de traitement et visa
    =========================================================================

    CE QUE LA PIÈCE COMPTABLE NE PROUVE PAS

    La pièce porte déjà les quatre cartouches de la banque : Initié, Contrôlé, Approuvé,
    puis Écriture passée et autorisée. Les ÉCRITURES sont donc couvertes par une signature.

    Mais qui signe la pièce signe ce qui y figure. Cinq choses n'y figurent pas :

      — ce qui n'a PAS été comptabilisé : les Accounts non paramétrés sont écartés, et
        n'apparaissent donc nulle part sur la pièce. C'est pourtant l'information qu'un
        contrôleur cherche en premier ;
      — quels rapports Western Union ont servi. Si une journée est contestée dans six mois,
        rien ne prouve que c'est bien le rapport de ce jour-là qui a été traité ;
      — l'écart d'arrondi, noyé au milieu des lignes de la pièce ;
      — le fichier destiné au core banking, qui est pourtant le point de non-retour ;
      — que la journée est complète : volumes, points de vente, concordance.

    Cette table porte l'en-tête du TRAITEMENT, là où la pièce porte les écritures. C'est
    d'elle que se tire le bordereau de fin de journée, qui se signe et se classe.

    UNE LIGNE PAR JOURNÉE, ÉCRITE AVEC L'HISTORIQUE

    Dans la même transaction que T_HistoriqueWU et T_PieceWU : trois documents d'une même
    journée ne doivent jamais diverger. Et RÉPÉTABLE comme eux — recomptabiliser une journée
    remplace son en-tête.

    RECOMPTABILISER EFFACE LE VISA, ET C'EST VOULU

    Le visa atteste d'un traitement précis. Refaire la journée produit un autre traitement :
    laisser le visa en place ferait croire qu'un supérieur a vu des chiffres qu'il n'a jamais
    vus. Il faut viser de nouveau.

    PERSONNE NE VISE SON PROPRE TRAITEMENT

    La contrainte est posée dans la base, et non seulement dans l'application : un UPDATE fait
    à la main dans Management Studio ne la contourne pas davantage.

    À exécuter APRÈS 01 à 16, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. L'en-tête de traitement d'une journée
-- =========================================================================
IF OBJECT_ID(N'dbo.T_TraitementWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_TraitementWU
    (
        DateActivite        DATE            NOT NULL,   -- Journée comptabilisée
        DateValeur          DATE            NULL,       -- Jour où les écritures sont passées
        NumeroLot           NVARCHAR(10)    NULL,

        /*
            Les deux rapports Western Union qui ont servi, et leur empreinte SHA-256.

            L'empreinte ne protège de rien — qui remplace un fichier peut recalculer la
            sienne. Elle répond à une question : « le fichier que vous me montrez est-il
            celui qui a été traité ce jour-là ? ». Sans elle, la question n'a pas de réponse.
        */
        FichierActivite     NVARCHAR(255)   NULL,
        EmpreinteActivite   NVARCHAR(64)    NULL,
        FichierReglement    NVARCHAR(255)   NULL,
        EmpreinteReglement  NVARCHAR(64)    NULL,

        -- Ce qui a été traité, et ce qui a été écarté.
        NombrePdv           INT             NOT NULL DEFAULT (0),
        NombreSousAgents    INT             NOT NULL DEFAULT (0),
        NombreAgences       INT             NOT NULL DEFAULT (0),
        NombreEcartes       INT             NOT NULL DEFAULT (0),

        NombreEnvois        INT             NOT NULL DEFAULT (0),
        NombrePaiements     INT             NOT NULL DEFAULT (0),
        NombreAnnulations   INT             NOT NULL DEFAULT (0),

        -- La pièce, en trois chiffres : l'équilibre se constate, il ne se suppose pas.
        TotalDebit          BIGINT          NOT NULL DEFAULT (0),
        TotalCredit         BIGINT          NOT NULL DEFAULT (0),
        EcartArrondi        BIGINT          NOT NULL DEFAULT (0),
        CompteEcart         NVARCHAR(50)    NULL,

        ComptabilisePar     NVARCHAR(50)    NULL,
        DateComptabilisation DATETIME       NOT NULL DEFAULT (GETDATE()),

        -- Le visa du supérieur. NULL tant que la journée n'a pas été relue.
        VisePar             NVARCHAR(50)    NULL,
        DateVisa            DATETIME        NULL,
        CommentaireVisa     NVARCHAR(500)   NULL,

        CONSTRAINT PK_T_TraitementWU PRIMARY KEY (DateActivite),

        CONSTRAINT CK_T_TraitementWU_PasSoiMeme
            CHECK (VisePar IS NULL OR VisePar <> ComptabilisePar),

        -- Un visa sans date, ou une date sans visa, ne veut rien dire.
        CONSTRAINT CK_T_TraitementWU_VisaComplet
            CHECK ((VisePar IS NULL AND DateVisa IS NULL)
                OR (VisePar IS NOT NULL AND DateVisa IS NOT NULL))
    );

    PRINT 'Table T_TraitementWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_TraitementWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. L'archive des en-têtes, pour les journées annulées
--
--    Une journée annulée emporte son historique, ses MTCN et sa pièce. Elle doit emporter
--    aussi la façon dont elle avait été traitée : c'est ce que l'on relira pour comprendre
--    POURQUOI elle a dû être retirée.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_TraitementAnnuleWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_TraitementAnnuleWU
    (
        IdAnnulation        BIGINT          NOT NULL,

        DateActivite        DATE            NOT NULL,
        DateValeur          DATE            NULL,
        NumeroLot           NVARCHAR(10)    NULL,

        FichierActivite     NVARCHAR(255)   NULL,
        EmpreinteActivite   NVARCHAR(64)    NULL,
        FichierReglement    NVARCHAR(255)   NULL,
        EmpreinteReglement  NVARCHAR(64)    NULL,

        NombrePdv           INT             NULL,
        NombreSousAgents    INT             NULL,
        NombreAgences       INT             NULL,
        NombreEcartes       INT             NULL,

        NombreEnvois        INT             NULL,
        NombrePaiements     INT             NULL,
        NombreAnnulations   INT             NULL,

        TotalDebit          BIGINT          NULL,
        TotalCredit         BIGINT          NULL,
        EcartArrondi        BIGINT          NULL,
        CompteEcart         NVARCHAR(50)    NULL,

        ComptabilisePar     NVARCHAR(50)    NULL,
        DateComptabilisation DATETIME       NULL,

        VisePar             NVARCHAR(50)    NULL,
        DateVisa            DATETIME        NULL,
        CommentaireVisa     NVARCHAR(500)   NULL,

        CONSTRAINT PK_T_TraitementAnnuleWU PRIMARY KEY (IdAnnulation)
    );

    CREATE INDEX IX_T_TraitementAnnuleWU_DateActivite
        ON dbo.T_TraitementAnnuleWU (DateActivite);

    PRINT 'Table T_TraitementAnnuleWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_TraitementAnnuleWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 3. Vue d'audit : les journées en attente de visa
--
--    C'est la liste que le chef de service ouvre le matin.
-- =========================================================================
IF OBJECT_ID(N'dbo.V_JourneesAViser', N'V') IS NOT NULL DROP VIEW dbo.V_JourneesAViser;
GO
CREATE VIEW dbo.V_JourneesAViser
AS
    SELECT  t.DateActivite,
            t.NumeroLot,
            t.ComptabilisePar,
            t.DateComptabilisation,
            attente     = DATEDIFF(DAY, t.DateComptabilisation, GETDATE()),
            t.NombrePdv,
            t.NombreEcartes,
            equilibre   = CASE WHEN t.TotalDebit = t.TotalCredit THEN N'oui' ELSE N'NON' END,
            t.TotalDebit,
            t.EcartArrondi
    FROM    dbo.T_TraitementWU AS t
    WHERE   t.VisePar IS NULL;
GO

-- =========================================================================
-- 4. Droits
--
--    Le rôle de compense écrit l'en-tête : il le produit en comptabilisant.
--    Le visa est un UPDATE, accordé aux trois rôles pour la même raison que l'annulation —
--    la fonction d'authorizer se porte indifféremment sur un compte de compense, de
--    commercial ou d'administration. Le contrôle du double regard est posé par la contrainte
--    CK_T_TraitementWU_PasSoiMeme et par l'application, non par les droits SQL.
--
--    DELETE est accordé sur T_TraitementWU seulement : recomptabiliser une journée remplace
--    son en-tête, comme pour l'historique et la pièce. Sur l'ARCHIVE, aucun DELETE.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_TraitementWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_TraitementAnnuleWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.V_JourneesAViser TO wu_compense');
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
BEGIN
    EXEC('GRANT SELECT, UPDATE ON dbo.T_TraitementWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_TraitementAnnuleWU TO wu_commercial');
    EXEC('GRANT SELECT ON dbo.V_JourneesAViser TO wu_commercial');
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_TraitementWU TO wu_admin');
    EXEC('GRANT SELECT, INSERT ON dbo.T_TraitementAnnuleWU TO wu_admin');
    EXEC('GRANT SELECT ON dbo.V_JourneesAViser TO wu_admin');
END
GO

-- =========================================================================
-- 5. Contrôle
-- =========================================================================
SELECT  journees    = COUNT(*),
        visees      = SUM(CASE WHEN VisePar IS NOT NULL THEN 1 ELSE 0 END),
        enAttente   = SUM(CASE WHEN VisePar IS NULL THEN 1 ELSE 0 END),
        premiere    = MIN(DateActivite),
        derniere    = MAX(DateActivite)
FROM    dbo.T_TraitementWU;
GO

PRINT N'';
PRINT N'Les journees comptabilisees AVANT ce script n''ont pas d''en-tete de traitement :';
PRINT N'leur bordereau se reconstitue depuis l''historique et la piece, mais sans le nom des';
PRINT N'rapports Western Union, qui n''etait conserve nulle part. L''ecran le dit.';
GO
