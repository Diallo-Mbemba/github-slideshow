/*
    =========================================================================
    Pièces comptables conservées — table T_PieceWU
    =========================================================================

    POURQUOI CONSERVER LA PIÈCE, ET NON LA RECALCULER

    L'historique (T_HistoriqueWU) garde les volumes, les montants et les totaux de
    commissions. Il ne garde PAS le taux du sous-agent ce jour-là, ni ses comptes de
    compensation et de commission, ni la ligne d'écart posée sur le compte inter bancaire.

    Reconstituer une pièce ancienne avec le paramétrage d'aujourd'hui réécrirait donc le
    passé : un sous-agent passé de 70 % à 60 % ferait apparaître une pièce qui n'a jamais
    été visée ni signée. Une pièce comptable est un justificatif ; « à peu près la même »
    n'a pas de sens devant un inspecteur.

    Cette table conserve donc la pièce TELLE QU'ELLE A ÉTÉ PRODUITE, ligne à ligne. La
    consulter, c'est relire ce qui a été écrit — rien n'est recalculé.

    Elle porte aussi de quoi reconstruire le fichier destiné au core banking : le compte,
    le libellé, le sens, le montant et le code agence suffisent, et ce sont exactement les
    colonnes ci-dessous.

    ÉCRITE DANS LA MÊME TRANSACTION QUE L'HISTORIQUE
    Deux documents d'une même journée ne doivent jamais diverger. Si l'un échoue, aucun des
    deux n'est écrit.

    À exécuter sur la base GWC_WINCOMPENSE_ETD. Rejouable : ne crée que ce qui manque.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF OBJECT_ID(N'dbo.T_PieceWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_PieceWU
    (
        DateActivite        DATE            NOT NULL,   -- Journée comptabilisée
        Ligne               INT             NOT NULL,   -- Rang de l'écriture dans la pièce

        Compte              NVARCHAR(50)    NOT NULL,
        Libelle             NVARCHAR(255)   NULL,

        -- Une écriture porte un débit OU un crédit, jamais les deux. En FCFA, donc entier :
        -- la monnaie n'a pas de décimale, et la pièce est déjà arrondie à l'unité.
        Debit               BIGINT          NOT NULL DEFAULT (0),
        Credit              BIGINT          NOT NULL DEFAULT (0),

        -- Agence de rattachement : elle ne figure pas sur la pièce imprimée, mais alimente
        -- la colonne ACBRN du fichier destiné au core banking, qui n'a aucun autre moyen de
        -- savoir à quelle agence rattacher l'écriture.
        CodeAgence          NVARCHAR(50)    NULL,

        DateEnregistrement  DATETIME        NOT NULL DEFAULT (GETDATE()),
        EnregistrePar       NVARCHAR(100)   NULL,

        CONSTRAINT PK_T_PieceWU PRIMARY KEY (DateActivite, Ligne),

        -- Une écriture sans montant n'a rien à faire dans une pièce, et une écriture des
        -- deux côtés n'aurait pas de sens comptable.
        CONSTRAINT CK_T_PieceWU_UnSeulSens CHECK
            ((Debit <> 0 AND Credit = 0) OR (Credit <> 0 AND Debit = 0))
    );

    -- La consultation se fait toujours par journée.
    CREATE INDEX IX_T_PieceWU_DateActivite ON dbo.T_PieceWU (DateActivite);

    PRINT 'Table T_PieceWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_PieceWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- Droits
--
-- DELETE est accordé au rôle de compense, comme sur l'historique : rejouer une journée
-- suppose d'effacer la version précédente, ce que fait EnregistrerJournee dans une
-- transaction. Le commercial lit, il ne comptabilise rien.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_compense' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT, INSERT, DELETE ON dbo.T_PieceWU TO wu_compense');
    PRINT 'Droits accordés à wu_compense sur T_PieceWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_commercial' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_PieceWU TO wu_commercial');
    PRINT 'Droits accordés à wu_commercial sur T_PieceWU.';
END
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_admin' AND type = 'R')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_PieceWU TO wu_admin');
    PRINT 'Droits accordés à wu_admin sur T_PieceWU.';
END
GO

-- =========================================================================
-- Contrôle
-- =========================================================================
SELECT  journees   = COUNT(DISTINCT DateActivite),
        ecritures  = COUNT(*),
        premiere   = MIN(DateActivite),
        derniere   = MAX(DateActivite)
FROM    dbo.T_PieceWU;
GO

PRINT N'';
PRINT N'Les journees deja comptabilisees AVANT cette table n''ont pas de piece conservee.';
PRINT N'L''ecran de consultation le dira, et proposera une reconstitution etiquetee comme telle.';
GO
