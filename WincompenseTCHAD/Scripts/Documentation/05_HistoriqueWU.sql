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
