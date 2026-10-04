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
