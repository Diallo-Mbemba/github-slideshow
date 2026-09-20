/*
    =========================================================================
    Trace des fichiers produits pour le core banking — T_FichierCoreBankingWU
    =========================================================================

    POURQUOI GARDER TRACE D'UN FICHIER QU'ON NE STOCKE PAS

    Le fichier destiné au core banking n'est pas conservé, et n'a pas à l'être : il dérive
    entièrement de la pièce, par les mêmes règles. Le garder serait garder deux fois la
    même chose, avec le risque que les deux copies divergent.

    Mais le FAIT de l'avoir produit, lui, ne se déduit de rien. Or c'est le seul moment où
    la journée quitte Wincompense pour entrer dans les livres de la banque. Sans cette
    trace, l'application ne peut ni avertir avant de recomptabiliser une journée déjà
    partie, ni prévenir avant de l'annuler.

    CE QUE LA TRACE NE DIT PAS

    Qu'il a été INJECTÉ. Wincompense écrit un classeur ; elle ne voit pas ce que le core
    banking en fait. La trace dit « ce fichier est sorti, tel jour, par telle personne,
    sous tel numéro de lot ». La suite se demande à l'agent, et c'est la colonne
    CoreBankingInjecte de T_AnnulationWU qui garde sa réponse.

    PLUSIEURS PRODUCTIONS POUR UNE MÊME JOURNÉE

    Rien ne l'interdit : un fichier mal enregistré se refait, et l'écran des pièces
    archivées sait reconstruire celui d'une journée ancienne. Toutes les productions sont
    conservées ; c'est la plus récente qui fait foi.

    À exécuter APRÈS 01 à 13, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF OBJECT_ID(N'dbo.T_FichierCoreBankingWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_FichierCoreBankingWU
    (
        IdFichier           BIGINT IDENTITY(1,1) NOT NULL,

        DateActivite        DATE            NOT NULL,   -- Journée comptabilisée
        DateValeur          DATE            NOT NULL,   -- Jour où les écritures sont passées
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

    -- La question posée est toujours « cette journée est-elle déjà sortie ? ».
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
-- Droits
--
-- Le rôle de compense produit les fichiers : il écrit. Le commercial lit.
-- AUCUN rôle ne reçoit DELETE : une trace que l'on peut effacer ne prouve rien.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_compense' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT, INSERT ON dbo.T_FichierCoreBankingWU TO wu_compense');
    PRINT 'Droits accordés à wu_compense sur T_FichierCoreBankingWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_commercial' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_FichierCoreBankingWU TO wu_commercial');
    PRINT 'Droits accordés à wu_commercial sur T_FichierCoreBankingWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_admin' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_FichierCoreBankingWU TO wu_admin');
    PRINT 'Droits accordés à wu_admin sur T_FichierCoreBankingWU.';
END
GO

-- =========================================================================
-- Contrôle
-- =========================================================================
SELECT  fichiers = COUNT(*),
        journees = COUNT(DISTINCT DateActivite),
        premiere = MIN(DateActivite),
        derniere = MAX(DateActivite)
FROM    dbo.T_FichierCoreBankingWU;
GO

PRINT N'';
PRINT N'Les fichiers produits AVANT cette table ne sont pas traces : l''application dira';
PRINT N'« aucune production connue » pour ces journees-la, ce qui n''est pas la meme chose';
PRINT N'que « jamais produit ».';
GO
