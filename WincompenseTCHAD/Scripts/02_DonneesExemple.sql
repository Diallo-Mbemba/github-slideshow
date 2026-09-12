/*
    Jeu de données d'exemple pour les tests de non-régression décrits en section 17
    du cahier des charges (AHB020200 : sous-agent, AHB020057 : agence propre Ecobank).
    À adapter/compléter avec les données réelles de production.
*/

USE GWC_WINCOMPENSE;
GO

-- Sous-agent de test (cas de référence AHB020200 - taux de rétrocession 70 %).
IF NOT EXISTS (SELECT 1 FROM dbo.T_Pdv_SA WHERE Code_Pdv = 'AHB020200')
BEGIN
    INSERT INTO dbo.T_Pdv_SA (Code_Pdv, Designationagence, GroupeStatistique, Taux, CompteCompense, CompteCommission, codeagence)
    VALUES ('AHB020200', 'SOUS-AGENT DE TEST AHB020200', 'SOUS-AGENTS', 0.70, '571100200', '706600200', 'AHB02');
END
GO

-- Agence propre Ecobank de test (cas de référence AHB020057 - 100% des commissions à la banque).
IF NOT EXISTS (SELECT 1 FROM dbo.T_Pdv_EC WHERE Codesite = 'AHB020057')
BEGIN
    INSERT INTO dbo.T_Pdv_EC (Codesite, Designationagence, [CodeAgenc-Voyager])
    VALUES ('AHB020057', 'AGENCE PROPRE ECOBANK AHB020057', 'AHB057');
END
GO
