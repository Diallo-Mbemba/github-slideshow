/*
    =========================================================================
    Les produits de transfert — T_ProduitTransfert
    =========================================================================

    POURQUOI CETTE TABLE

    Wincompense n'est plus « l'application Western Union » : elle compense les transferts
    d'argent de la banque, et Western Union n'en est que le premier produit. Ria est en projet ;
    MoneyGram, Wari ou un autre suivront, et la banque ne doit pas attendre une livraison pour
    déclarer un produit qu'elle vient de signer.

    CE QUE CETTE TABLE DIT, ET CE QU'ELLE NE DIT PAS

    Elle dit quels produits EXISTENT, comment ils s'appellent, de quelle couleur est leur
    emblème et dans quel ordre ils s'affichent.

    Elle ne dit PAS lesquels sont TRAITÉS. Cela, c'est le code de l'application qui le sait,
    et lui seul :

        LA BASE DIT QUELS PRODUITS EXISTENT.
        LE CODE DIT LESQUELS SONT TRAITÉS.

    Aucune colonne « Disponible » n'existe ici, et c'est délibéré. Si elle existait, un
    administrateur pourrait déclarer MoneyGram disponible : l'écran de traitement de la
    compense s'ouvrirait, lirait les rapports Western Union, appliquerait les taux Western
    Union, et produirait une pièce sur les comptes Western Union — SOUS UN NOM MONEYGRAM.
    Personne ne s'en apercevrait avant la comptabilisation.

    Un produit ajouté ici est donc « En attente » tant qu'une livraison n'apporte pas, pour
    lui, la lecture de ses rapports, ses taux, sa pièce comptable et son jeu de tables.

    UNE TABLE COMMUNE, ET NON UNE TABLE DE PRODUIT

    Elle les liste tous : elle vit donc avec les utilisateurs et les connexions, et non dans
    le jeu de tables d'un produit. Son nom ne porte pas le suffixe WU, qui n'aurait aucun sens
    ici — les tables communes actuelles le portent par héritage.

    LE CODE EST LA CLÉ, ET IL NE SE MODIFIE PAS

    Il nomme le fichier du logo (Logos\MGRAM.png) et nommera demain le schéma ou le suffixe des
    tables du produit. L'écran d'administration interdit de le changer après création, et
    n'offre aucune suppression : un produit se met HORS SERVICE, ce qui le retire de la fenêtre
    de choix sans rien effacer de son historique.

    À exécuter APRÈS 00 (ou 01 à 19), sur la base GWC_WINCOMPENSE_ETD.
    Ce script est REJOUABLE : il crée ce qui manque et ne réécrit rien de ce qui est saisi.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. La table
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_ProduitTransfert' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE dbo.T_ProduitTransfert
    (
        -- Le code est la clé : court, en majuscules, sans accent ni espace. Il nomme le
        -- fichier du logo et, demain, les tables du produit.
        Code                NVARCHAR(10)    NOT NULL PRIMARY KEY,

        Nom                 NVARCHAR(100)   NOT NULL,
        Description         NVARCHAR(500)   NULL,

        -- La couleur de l'emblème dessiné tant que le logo officiel n'est pas déposé.
        -- Trois composantes séparées : plus lisibles qu'un entier signé dans une requête,
        -- et bornées par une contrainte plutôt que par la confiance.
        CouleurRouge        TINYINT         NOT NULL DEFAULT (108),
        CouleurVert         TINYINT         NOT NULL DEFAULT (117),
        CouleurBleu         TINYINT         NOT NULL DEFAULT (125),

        -- Rang d'affichage dans la fenêtre de choix, croissant.
        Ordre               INT             NOT NULL DEFAULT (100),

        -- Hors service : le produit n'apparaît plus au choix. Rien n'est effacé.
        EnService           BIT             NOT NULL DEFAULT (1),

        CreePar             NVARCHAR(50)    NULL,
        DateCreation        DATETIME        NULL,
        ModifiePar          NVARCHAR(50)    NULL,
        DateModification    DATETIME        NULL,

        CONSTRAINT CK_T_ProduitTransfert_Code
            CHECK (LEN(LTRIM(RTRIM(Code))) >= 2)
    );

    PRINT 'Table T_ProduitTransfert créée.';
END
ELSE
BEGIN
    PRINT 'Table T_ProduitTransfert déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. Les deux produits que l'application connaît aujourd'hui
--
--    AMORÇAGE SEULEMENT. Les lignes déjà présentes ne sont pas réécrites : un administrateur
--    qui a corrigé un nom ou une couleur ne doit pas les voir revenir à chaque exécution du
--    script.
--
--    « installation » comme auteur, et non un identifiant d'utilisateur : la pose initiale
--    n'est pas une saisie de quelqu'un.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM dbo.T_ProduitTransfert WHERE Code = N'WU')
BEGIN
    INSERT INTO dbo.T_ProduitTransfert
        (Code, Nom, Description, CouleurRouge, CouleurVert, CouleurBleu, Ordre, EnService,
         CreePar, DateCreation)
    VALUES
        (N'WU', N'Western Union',
         N'La compensation Western Union : rapport d''activité, commissions sur envois et sur réceptions, pièce comptable et fichier core banking.',
         255, 182, 0, 10, 1, N'installation', GETDATE());

    PRINT 'Produit WU amorcé.';
END
ELSE
    PRINT 'Produit WU déjà présent : amorçage ignoré.';
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ProduitTransfert WHERE Code = N'RIA')
BEGIN
    INSERT INTO dbo.T_ProduitTransfert
        (Code, Nom, Description, CouleurRouge, CouleurVert, CouleurBleu, Ordre, EnService,
         CreePar, DateCreation)
    VALUES
        (N'RIA', N'Ria',
         N'Produit en projet à la banque. L''environnement est en place ; le traitement de sa compensation reste à écrire.',
         238, 118, 35, 20, 1, N'installation', GETDATE());

    PRINT 'Produit RIA amorcé.';
END
ELSE
    PRINT 'Produit RIA déjà présent : amorçage ignoré.';
GO

-- =========================================================================
-- 3. Droits
--
--    Tout le monde LIT : la fenêtre de choix s'affiche avant que le rôle n'ait la moindre
--    importance, et tout utilisateur doit pouvoir choisir son produit.
--
--    Seul l'administrateur ÉCRIT : déclarer un produit est un acte de configuration, au même
--    titre que les comptes systèmes et les taxes.
--
--    AUCUN DELETE, pour personne. Un produit supprimé laisserait son historique, ses pièces
--    et son référentiel orphelins, sans moyen de retrouver à quoi ils se rapportaient. La
--    colonne EnService est là pour cela.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.T_ProduitTransfert TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.T_ProduitTransfert TO wu_commercial');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_ProduitTransfert TO wu_admin');
GO

-- =========================================================================
-- 4. Compte rendu
-- =========================================================================
SELECT  Code,
        Nom,
        Ordre,
        EnService,
        CONCAT(CouleurRouge, ', ', CouleurVert, ', ', CouleurBleu) AS Couleur,
        CreePar,
        DateCreation
FROM    dbo.T_ProduitTransfert
ORDER BY Ordre, Nom;
GO

PRINT '';
PRINT '=========================================================================';
PRINT 'Produits de transfert : table posée.';
PRINT '';
PRINT 'RAPPEL : cette table dit quels produits EXISTENT, pas lesquels sont TRAITÉS.';
PRINT 'Un produit ajouté par l''administrateur reste « En attente » tant qu''une';
PRINT 'livraison n''apporte pas le traitement de sa compensation.';
PRINT '=========================================================================';
GO
