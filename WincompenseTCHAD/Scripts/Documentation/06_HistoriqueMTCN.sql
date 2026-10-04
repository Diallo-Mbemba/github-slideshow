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
