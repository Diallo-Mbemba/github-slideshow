/*
    Script de création des tables de paramétrage des points de vente Western Union
    pour la base GWC_WINCOMPENSE_ETD (SQL Server Express, instance .\SQLEXPRESS).

    IMPORTANT : l'Account (Code_Pdv / Codesite) est l'identifiant métier unique.
    Ne jamais utiliser codeagence comme clé d'identification.

    Types alignés sur le schéma réellement en place à Ecobank Tchad (NVARCHAR(255) partout,
    Taux en DECIMAL(4,2)). Le script étant protégé par IF NOT EXISTS, il ne modifie jamais
    une table déjà créée : il ne sert qu'à monter un environnement neuf.

    Ces deux tables se gèrent depuis l'application : boutons « Sous-agents... » et
    « Agences propres... » de l'écran principal.
*/

USE master;
GO

IF NOT EXISTS (SELECT 1 FROM sys.databases WHERE name = N'GWC_WINCOMPENSE_ETD')
BEGIN
    CREATE DATABASE GWC_WINCOMPENSE_ETD;
END
GO

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- Table des SOUS-AGENTS Western Union
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_Pdv_SA')
BEGIN
    CREATE TABLE dbo.T_Pdv_SA
    (
        Code_Pdv          NVARCHAR(255)   NOT NULL PRIMARY KEY, -- Account (identifiant métier unique)
        Designationagence NVARCHAR(255)   NOT NULL,
        GroupeStatistique NVARCHAR(255)   NOT NULL,
        Taux              DECIMAL(4, 2)   NOT NULL DEFAULT (0),  -- Quote-part du sous-agent (0.70 = 70 %), DEUX décimales
        CompteCompense    NVARCHAR(255)   NOT NULL,              -- Compte de compensation du sous-agent
        CompteCommission  NVARCHAR(255)   NOT NULL,              -- Compte de commission du sous-agent
        codeagence        NVARCHAR(255)   NOT NULL               -- Informatif uniquement, jamais utilisé comme clé
    );
END
GO

-- =========================================================================
-- Table des AGENCES PROPRES (Ecobank)
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_Pdv_EC')
BEGIN
    CREATE TABLE dbo.T_Pdv_EC
    (
        Codesite            NVARCHAR(255)   NOT NULL PRIMARY KEY, -- Account (identifiant métier unique)
        Designationagence   NVARCHAR(255)   NULL,
        [CodeAgenc-Voyager] NVARCHAR(255)   NOT NULL
    );
END
GO
