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
