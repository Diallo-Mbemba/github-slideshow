/*
    =========================================================================
    Options de traitement — T_ParametreWU
    =========================================================================

    POURQUOI UNE TABLE, ET NON UN RÉGLAGE SUR LE POSTE

    Ces options décrivent la FAÇON DE TRAVAILLER de la banque, et non la configuration d'un
    poste. Si le visa est obligatoire avant de produire le fichier core banking, il l'est pour
    tout le monde — un agent ne doit pas pouvoir s'en affranchir en décochant une case chez
    lui. La règle vit donc en base, avec les données qu'elle protège.

    UNE TABLE CLÉ / VALEUR, ET NON UNE COLONNE PAR OPTION

    Une option de plus demanderait sinon une colonne de plus, donc une modification de schéma,
    donc un script, donc une reprise. Une ligne suffit. Le prix à payer est que la valeur est
    du texte : c'est l'application qui sait la lire, et elle le fait en un seul endroit.

    SystemeWU n'accueille PAS ces options : cette table porte les neuf comptes qui déterminent
    la pièce comptable, et y mêler des réglages de procédure rendrait sa lecture ambiguë pour
    qui l'ouvre dans Management Studio.

    À exécuter APRÈS 01 à 17, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. La table
-- =========================================================================
IF OBJECT_ID(N'dbo.T_ParametreWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_ParametreWU
    (
        Cle                 NVARCHAR(50)    NOT NULL,
        Valeur              NVARCHAR(255)   NOT NULL,

        -- Ce que l'option veut dire, en clair : la table se lit aussi hors de l'application.
        Libelle             NVARCHAR(255)   NULL,

        DateModification    DATETIME        NULL,
        ModifiePar          NVARCHAR(50)    NULL,

        CONSTRAINT PK_T_ParametreWU PRIMARY KEY (Cle)
    );

    PRINT 'Table T_ParametreWU créée.';
END
ELSE
BEGIN
    PRINT 'Table T_ParametreWU déjà présente : création ignorée.';
END
GO

-- =========================================================================
-- 2. L'option : le visa bloque-t-il la production du fichier core banking ?
--
--    NON par défaut, et ce n'est pas une facilité : imposer une règle bloquante que la
--    banque n'a pas demandée arrêterait la compense du jour au premier matin où le chef de
--    service est absent. Avec NON, l'application AVERTIT et laisse passer ; avec OUI, elle
--    refuse. Dans les deux cas l'agent sait où il en est.
--
--    La valeur existante n'est JAMAIS écrasée : rejouer ce script ne remet pas la banque
--    dans l'état d'origine.
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'VISA_AVANT_CORE_BANKING')
BEGIN
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'VISA_AVANT_CORE_BANKING', N'NON',
            N'OUI : le fichier core banking ne peut pas être produit tant que la journée n''est pas visée. NON : l''application avertit seulement.',
            GETDATE(), N'installation');

    PRINT 'Option VISA_AVANT_CORE_BANKING créée à NON.';
END
ELSE
BEGIN
    PRINT 'Option VISA_AVANT_CORE_BANKING déjà présente : valeur conservée.';
END
GO

-- =========================================================================
-- 3. Droits
--
--    Tout le monde LIT : l'option gouverne un contrôle que l'agent de la compense subit, il
--    faut donc qu'il puisse la connaître. Seul l'administrateur ÉCRIT : une règle de
--    procédure ne se change pas depuis le guichet.
--
--    Aucun DELETE : une option effacée redeviendrait silencieusement la valeur par défaut.
-- =========================================================================
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.T_ParametreWU TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.T_ParametreWU TO wu_commercial');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_ParametreWU TO wu_admin');
GO

-- =========================================================================
-- 4. Contrôle
-- =========================================================================
SELECT Cle, Valeur, ModifiePar, DateModification FROM dbo.T_ParametreWU ORDER BY Cle;
GO

PRINT N'';
PRINT N'L''option se change dans l''application : menu Parametrage > Options de traitement.';
PRINT N'Tant que cette table est absente, l''application se comporte comme avec NON : elle';
PRINT N'avertit, elle ne bloque pas.';
GO
