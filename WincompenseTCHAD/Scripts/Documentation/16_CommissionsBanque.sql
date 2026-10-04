/*
    =========================================================================
    Part de la banque dans les commissions — colonnes ajoutées à l'historique
    =========================================================================

    CE QUI MANQUAIT

    L'historique garde les commissions TOTALES d'une journée — celles que Western Union
    verse — et le type de chaque point de vente. Il ne gardait ni la part qui revient à la
    banque, ni le taux du sous-agent ce jour-là.

    Or ces deux parts ne se valent pas : sur un sous-agent à 70 %, les sept dixièmes de la
    commission affichée ne sont pas à la banque. Aucun état ne pouvait donc répondre à la
    question « combien la banque a-t-elle gagné sur la période ? ».

    POURQUOI CONSERVER, ET NON RECALCULER

    La part de la banque vaut « commission × (1 − taux du sous-agent) ». Le taux, lui,
    change : un sous-agent passé de 70 % à 60 % ferait varier RÉTROACTIVEMENT ce que la
    banque a gagné le mois dernier, si on recalculait avec le paramétrage d'aujourd'hui.
    C'est ce que l'application refuse déjà de faire pour les pièces comptables, et pour la
    même raison : un état comptable se relit, il ne se recompose pas.

    POURQUOI NULL, ET NON DEFAULT 0

    Les journées comptabilisées AVANT ce script n'ont pas de part enregistrée. Un zéro
    affirmerait que la banque n'a rien gagné ce jour-là ; NULL dit qu'on ne sait pas, et
    l'application le dit à son tour à l'écran. Une case vide qui se lit comme un zéro est
    la pire façon de se tromper.

    CE QUI RESTE EXACT POUR LE PASSÉ

    Les agences propres. Leur taux vaut zéro par définition, et le type de point de vente
    est conservé depuis toujours : sur une ligne marquée EC, la banque a gardé la totalité.
    Seuls les sous-agents restent sans répartition avant ce script.

    À exécuter APRÈS 01 à 15, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. L'historique des journées
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueWU')
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'T_HistoriqueWU') AND name = N'CommissionEnvoiBanque')
BEGIN
    ALTER TABLE dbo.T_HistoriqueWU
        ADD CommissionEnvoiBanque     DECIMAL(18, 2) NULL,
            CommissionPaiementBanque  DECIMAL(18, 2) NULL,
            CommissionTransfertBanque DECIMAL(18, 2) NULL,

            -- Le taux appliqué CE JOUR-LÀ. Il ne sert pas au calcul de l'état — la part est
            -- déjà là — mais il permet de comprendre un chiffre sans rouvrir le paramétrage,
            -- et de retrouver la règle si elle est un jour contestée.
            TauxSA                    DECIMAL(4, 2)  NULL;

    PRINT 'Colonnes de part bancaire ajoutées à T_HistoriqueWU.';
END
ELSE
BEGIN
    PRINT 'T_HistoriqueWU : colonnes de part bancaire déjà présentes ou table absente.';
END
GO

-- =========================================================================
-- 2. L'archive des journées annulées
--
--    L'annulation déplace les lignes colonne par colonne. Sans ces quatre colonnes ici,
--    une journée annulée perdrait sa répartition en partant en archive — et une archive
--    qui perd une partie de ce qu'elle archive ne prouve plus grand-chose.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_HistoriqueAnnuleWU')
   AND NOT EXISTS (SELECT 1 FROM sys.columns
                   WHERE object_id = OBJECT_ID(N'T_HistoriqueAnnuleWU') AND name = N'CommissionEnvoiBanque')
BEGIN
    ALTER TABLE dbo.T_HistoriqueAnnuleWU
        ADD CommissionEnvoiBanque     DECIMAL(18, 2) NULL,
            CommissionPaiementBanque  DECIMAL(18, 2) NULL,
            CommissionTransfertBanque DECIMAL(18, 2) NULL,
            TauxSA                    DECIMAL(4, 2)  NULL;

    PRINT 'Colonnes de part bancaire ajoutées à T_HistoriqueAnnuleWU.';
END
ELSE
BEGIN
    PRINT 'T_HistoriqueAnnuleWU : colonnes déjà présentes ou table absente.';
END
GO

-- =========================================================================
-- 3. Vue d'audit : ce que la banque a gardé, par journée et par population
--
--    La même règle que l'application : la part enregistrée quand elle existe, la totalité
--    pour une agence propre — dont le taux vaut zéro par construction — et rien pour un
--    sous-agent dont la répartition n'a pas été conservée.
-- =========================================================================
IF OBJECT_ID(N'dbo.V_CommissionsBanque', N'V') IS NOT NULL DROP VIEW dbo.V_CommissionsBanque;
GO
CREATE VIEW dbo.V_CommissionsBanque
AS
    SELECT  h.DateActivite,
            population = CASE WHEN h.TypePdv = 'SA' THEN N'Sous-agents'
                              WHEN h.TypePdv = 'EC' THEN N'Agences propres'
                              ELSE N'Non paramétrés' END,

            documentee = CASE WHEN h.CommissionEnvoiBanque IS NOT NULL THEN 1
                              WHEN h.TypePdv <> 'SA' THEN 1
                              ELSE 0 END,

            partBanque = CASE WHEN h.CommissionEnvoiBanque IS NOT NULL
                              THEN h.CommissionEnvoiBanque + h.CommissionPaiementBanque
                                   + h.CommissionTransfertBanque
                              WHEN h.TypePdv <> 'SA'
                              THEN h.CommissionEnvoi + h.CommissionPaiement + h.CommissionTransfert
                              ELSE NULL END,

            commissionTotale = h.CommissionEnvoi + h.CommissionPaiement + h.CommissionTransfert,
            h.TauxSA,
            h.Account,
            h.Designation
    FROM    dbo.T_HistoriqueWU AS h;
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.V_CommissionsBanque TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.V_CommissionsBanque TO wu_commercial');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT ON dbo.V_CommissionsBanque TO wu_admin');
GO

-- =========================================================================
-- 4. Contrôle : à partir de quand la répartition est-elle documentée ?
-- =========================================================================
SELECT  premiereJourneeDocumentee = MIN(DateActivite),
        journeesDocumentees       = COUNT(DISTINCT DateActivite)
FROM    dbo.T_HistoriqueWU
WHERE   CommissionEnvoiBanque IS NOT NULL;
GO

SELECT  journeesSansRepartition = COUNT(DISTINCT DateActivite)
FROM    dbo.T_HistoriqueWU
WHERE   CommissionEnvoiBanque IS NULL
  AND   TypePdv = 'SA';
GO

PRINT N'';
PRINT N'La repartition commence a la premiere journee comptabilisee APRES ce script.';
PRINT N'Avant elle, l''etat affiche le total encaisse, reconstitue depuis les pieces';
PRINT N'conservees, et dit que la repartition n''est pas disponible.';
GO
