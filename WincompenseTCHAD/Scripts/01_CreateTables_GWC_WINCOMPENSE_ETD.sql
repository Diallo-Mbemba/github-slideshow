/*
    Script de création des tables de paramétrage des points de vente Western Union
    pour la base GWC_WINCOMPENSE_ETD (SQL Server Express, instance .\SQLEXPRESS).

    IMPORTANT : l'Account (Code_Pdv / Codesite) est l'identifiant métier unique.
    Ne jamais utiliser codeagence comme clé d'identification.
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
        Code_Pdv          VARCHAR(50)     NOT NULL PRIMARY KEY,   -- Account (identifiant métier unique)
        Designationagence NVARCHAR(200)   NOT NULL,
        GroupeStatistique NVARCHAR(100)   NULL,
        Taux              DECIMAL(9, 6)   NOT NULL DEFAULT (0),   -- Quote-part du sous-agent (ex : 0.70 = 70 %)
        CompteCompense    VARCHAR(30)     NOT NULL,               -- Compte de compensation du sous-agent
        CompteCommission  VARCHAR(30)     NOT NULL,               -- Compte de commission du sous-agent
        codeagence        VARCHAR(30)     NULL                    -- Informatif uniquement, jamais utilisé comme clé
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
        Codesite            VARCHAR(50)     NOT NULL PRIMARY KEY, -- Account (identifiant métier unique)
        Designationagence   NVARCHAR(200)   NOT NULL,
        [CodeAgenc-Voyager] VARCHAR(30)     NULL
    );
END
GO
