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
