/*
    =========================================================================
    Rôles SQL Server de la base GWC_WINCOMPENSE_ETD
    =========================================================================

    L'application contrôle déjà les droits par son propre mot de passe applicatif. Ces rôles
    de base de données en sont le second verrou : ils s'appliquent aux connexions SQL Server
    elles-mêmes, y compris à quelqu'un qui contournerait l'application — un poste relié à la
    base avec SQL Server Management Studio, par exemple.

    Trois rôles, calqués sur ceux de l'application :

      wu_compense    — lit le paramétrage, écrit l'historique et le journal des connexions.
                       Il ne peut PAS modifier les sous-agents, les agences, les groupes
                       ni les comptes comptables : c'est ce paramétrage qui détermine les
                       écritures, le laisser modifiable par celui qui les génère reviendrait
                       à supprimer le contrôle croisé voulu par la banque.

      wu_commercial  — écrit le paramétrage des points de vente, lit l'historique pour
                       consulter les rapports d'activité. Il ne peut pas écrire l'historique :
                       il ne comptabilise rien.

      wu_admin       — les deux, plus les comptes systèmes et la table des utilisateurs.

    AUCUN rôle ne reçoit DELETE sur le journal des connexions : un journal que ses propres
    utilisateurs peuvent effacer ne prouve rien. Seul le propriétaire de la base pourra le
    purger, et cela se verra.

    L'historique, lui, accepte DELETE pour le rôle de compense : rejouer une journée suppose
    d'effacer la précédente version, ce que fait EnregistrerJournee dans une transaction.

    Ce script est rejouable : il ne crée que ce qui manque.

    À exécuter APRÈS 01 à 07, sur la base GWC_WINCOMPENSE_ETD.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- =========================================================================
-- 1. Création des rôles
-- =========================================================================
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_compense' AND type = 'R')
BEGIN
    CREATE ROLE wu_compense;
    PRINT 'Rôle wu_compense créé.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_commercial' AND type = 'R')
BEGIN
    CREATE ROLE wu_commercial;
    PRINT 'Rôle wu_commercial créé.';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'wu_admin' AND type = 'R')
BEGIN
    CREATE ROLE wu_admin;
    PRINT 'Rôle wu_admin créé.';
END
GO

-- =========================================================================
-- 2. Droits du rôle wu_compense
--
--    Lecture du paramétrage, écriture de l'historique.
-- =========================================================================
GRANT SELECT ON dbo.T_Pdv_SA            TO wu_compense;
GRANT SELECT ON dbo.T_Pdv_EC            TO wu_compense;
GRANT SELECT ON dbo.SystemeWU           TO wu_compense;

GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueWU   TO wu_compense;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueMTCN TO wu_compense;

-- Se connecter suppose de lire sa propre fiche et d'écrire au journal ; le compteur d'échecs
-- et la date de dernière connexion sont mis à jour par l'application elle-même.
GRANT SELECT, UPDATE ON dbo.T_UtilisateurWU TO wu_compense;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU   TO wu_compense;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT ON dbo.T_GroupeStatistique TO wu_compense');
END
GO

-- =========================================================================
-- 3. Droits du rôle wu_commercial
--
--    Écriture du paramétrage des points de vente, lecture de l'historique.
-- =========================================================================
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_SA TO wu_commercial;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_EC TO wu_commercial;

GRANT SELECT ON dbo.T_HistoriqueWU   TO wu_commercial;
GRANT SELECT ON dbo.T_HistoriqueMTCN TO wu_commercial;

-- Les comptes comptables se consultent — un commercial doit pouvoir vérifier à quel compte
-- un groupe renvoie — mais ne se modifient pas.
GRANT SELECT ON dbo.SystemeWU TO wu_commercial;

GRANT SELECT, UPDATE ON dbo.T_UtilisateurWU TO wu_commercial;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU   TO wu_commercial;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_GroupeStatistique TO wu_commercial');
END
GO

-- =========================================================================
-- 4. Droits du rôle wu_admin
--
--    Tout ce que font les deux autres, plus les comptes systèmes et les utilisateurs.
-- =========================================================================
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_SA         TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_Pdv_EC         TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueWU   TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_HistoriqueMTCN TO wu_admin;
GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_UtilisateurWU  TO wu_admin;

GRANT SELECT, UPDATE ON dbo.SystemeWU     TO wu_admin;
GRANT SELECT, INSERT ON dbo.T_ConnexionWU TO wu_admin;
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_GroupeStatistique')
BEGIN
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_GroupeStatistique TO wu_admin');
END
GO

-- =========================================================================
-- 4 bis. Tables créées par les scripts 09 et 10
--
--    Elles n'existent pas encore lorsque ce script est exécuté dans l'ordre numérique :
--    les GRANT sont donc protégés par IF EXISTS et ne s'appliquent qu'au second passage,
--    ou lorsque ce script est exécuté en dernier (voir 00_InstallationComplete.sql).
--
--    SANS CES DROITS, l'application échoue là où on ne l'attend pas : le calendrier des
--    jours fériés devient illisible au moment de dater le fichier core banking, et la file
--    du double regard reste vide alors qu'elle contient des demandes.
--
--    RELANCEZ CE SCRIPT APRÈS 09 ET 10 si vous exécutez les scripts un par un.
-- =========================================================================

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_JourFerieWU')
BEGIN
    -- Lecture pour tous : dater une écriture suppose de connaître les jours chômés,
    -- quel que soit le rôle. L'écriture reste réservée à l'administrateur.
    EXEC('GRANT SELECT ON dbo.T_JourFerieWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.T_JourFerieWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_JourFerieWU TO wu_admin');
    PRINT 'Droits accordés sur T_JourFerieWU.';
END
ELSE
BEGIN
    PRINT 'T_JourFerieWU absente : relancez ce script après 10_JoursFeries.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_DemandeWU')
BEGIN
    -- La file du double regard portait d'abord le seul paramétrage, et l'agent de compense
    -- n'y avait aucun accès. Elle porte maintenant AUSSI les demandes d'annulation d'une
    -- journée comptabilisée, qui appartiennent, elles, à la compense : le rôle doit donc
    -- pouvoir y déposer, y lire, et y décider s'il porte la fonction d'authorizer.
    --
    -- Ce qui protège le référentiel n'est pas l'absence de ce droit SQL, mais la fonction
    -- INPUTER / AUTHORIZER portée par l'utilisateur et la contrainte
    -- CK_T_DemandeWU_PasSoiMeme, qu'un UPDATE fait à la main ne contourne pas davantage.
    --
    -- Aucun rôle ne reçoit DELETE : une demande rejetée se conserve, c'est elle qui prouve
    -- qu'un contrôle a eu lieu.
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_DemandeWU TO wu_admin');
    PRINT 'Droits accordés sur T_DemandeWU.';
END
ELSE
BEGIN
    PRINT 'T_DemandeWU absente : relancez ce script après 09_Demandes.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_AnnulationWU')
BEGIN
    -- Les archives d'annulation : SELECT et INSERT pour les trois rôles, JAMAIS de DELETE.
    -- L'INSERT est accordé largement parce que la fonction d'authorizer se porte
    -- indifféremment sur un compte de compense, de commercial ou d'administration ; le
    -- double regard, lui, est posé par la contrainte CK_T_AnnulationWU_PasSoiMeme et par
    -- l'application, non par les droits SQL.
    --
    -- Une archive que ses propres utilisateurs peuvent effacer ne prouve rien : c'est
    -- exactement quand une journée est annulée que l'auditeur veut la retrouver.
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_AnnulationWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueAnnuleWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_HistoriqueMTCNAnnuleWU TO wu_admin');

    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_compense');
    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT ON dbo.T_PieceAnnuleeWU TO wu_admin');

    PRINT 'Droits accordés sur les tables d''annulation.';
END
ELSE
BEGIN
    PRINT 'T_AnnulationWU absente : relancez ce script après 14_AnnulationComptabilisation.sql.';
END
GO

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_FichierCoreBankingWU')
BEGIN
    -- Le rôle de compense produit les fichiers : il écrit. Le commercial lit. Aucun DELETE :
    -- une trace que l'on peut effacer ne prouve rien.
    EXEC('GRANT SELECT, INSERT ON dbo.T_FichierCoreBankingWU TO wu_compense');
    EXEC('GRANT SELECT ON dbo.T_FichierCoreBankingWU TO wu_commercial');
    EXEC('GRANT SELECT, INSERT, UPDATE ON dbo.T_FichierCoreBankingWU TO wu_admin');
    PRINT 'Droits accordés sur T_FichierCoreBankingWU.';
END
ELSE
BEGIN
    PRINT 'T_FichierCoreBankingWU absente : relancez ce script après 15_FichierCoreBanking.sql.';
END
GO


-- =========================================================================
-- 5. Rattachement des utilisateurs SQL Server
--
--    À adapter : remplacez les noms ci-dessous par vos comptes réels, puis décommentez.
--    Les postes étant nominatifs, le plus simple est un compte Windows par agent :
--
--        CREATE LOGIN [DOMAINE\pnom] FROM WINDOWS;
--        CREATE USER  [DOMAINE\pnom] FOR LOGIN [DOMAINE\pnom];
--        ALTER ROLE wu_compense ADD MEMBER [DOMAINE\pnom];
--
--    Un compte SQL Server partagé par tout le service reste possible, mais le journal des
--    connexions de l'application redevient alors la seule trace nominative disponible.
-- =========================================================================

-- ALTER ROLE wu_compense   ADD MEMBER [DOMAINE\agent_compense];
-- ALTER ROLE wu_commercial ADD MEMBER [DOMAINE\agent_commercial];
-- ALTER ROLE wu_admin      ADD MEMBER [DOMAINE\administrateur_wu];
GO

-- =========================================================================
-- 6. Contrôle : droits effectivement accordés, par rôle et par table
-- =========================================================================
SELECT  principal    = dp.name,
        objet        = OBJECT_NAME(pe.major_id),
        droit        = pe.permission_name,
        etat         = pe.state_desc
FROM    sys.database_permissions AS pe
        INNER JOIN sys.database_principals AS dp ON dp.principal_id = pe.grantee_principal_id
WHERE   dp.name IN (N'wu_compense', N'wu_commercial', N'wu_admin')
ORDER BY dp.name, OBJECT_NAME(pe.major_id), pe.permission_name;
GO
