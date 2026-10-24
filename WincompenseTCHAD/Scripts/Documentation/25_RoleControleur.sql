/*
    =========================================================================
    Le rôle CONTROLEUR : un quatrième rôle, et ce qu'il ne peut pas faire
    =========================================================================

    CE SCRIPT NE CRÉE AUCUNE TABLE ET N'AJOUTE AUCUNE COLONNE. Il élargit UNE contrainte,
    et c'est tout. Il est reproduit dans Scripts\00_InstallationComplete.sql, qui est le
    seul à exécuter.

    POURQUOI UN QUATRIÈME RÔLE

    La banque place au siège des contrôleurs, chargés de suivre les comptes utilisateurs et
    l'activité du réseau. Ils créent des comptes, en activent et en désactivent — mais ils
    ne font ni la compense, ni le référentiel.

    LE RÉFLEXE AURAIT ÉTÉ DE LEUR DONNER LE RÔLE COMMERCIAL. C'est précisément ce qu'il ne
    fallait pas faire : COMMERCIAL ouvre la gestion du référentiel des points de vente. On
    aurait donné au contrôleur le pouvoir d'agir sur ce qu'il est chargé de contrôler.

    CE QUE LE CONTRÔLEUR PEUT

        - consulter les rapports d'activité, les pièces conservées, les commissions ;
        - consulter le contrôle des écarts de change, qui n'enregistre rien ;
        - ouvrir l'écran des utilisateurs et le journal des connexions ;
        - créer des comptes COMPENSE et COMMERCIAL ;
        - activer, désactiver et déverrouiller un compte qui n'est pas administrateur.

    CE QUE LE CONTRÔLEUR NE PEUT PAS — ET POURQUOI CHAQUE INTERDIT COMPTE

        - créer un ADMINISTRATEUR ou un autre CONTRÔLEUR. Celui qui crée un compte en
          choisit le mot de passe initial : un contrôleur qui créerait un administrateur
          se créerait le sien et s'y connecterait. Et un contrôleur créé par un contrôleur
          n'aurait plus d'administrateur au-dessus de lui.
        - MODIFIER un compte existant. Cet écran porte le RÔLE : le lui ouvrir reviendrait
          à le laisser promouvoir quelqu'un administrateur. Activer et désactiver, qui est
          son besoin, passe par une opération distincte qui ne touche que cette colonne.
        - RÉINITIALISER un mot de passe. Réinitialiser celui d'un administrateur, c'est
          prendre son compte — plus sûrement que le désactiver.
        - agir sur UN ADMINISTRATEUR, quel qu'il soit. La banque avait demandé « sauf
          l'administrateur qui l'a créé » ; la règle retenue est plus large, parce que la
          première se vérifie en remontant qui a créé qui, et qu'elle perd son objet dès
          qu'un contrôleur est créé par un contrôleur.
        - agir sur SON PROPRE COMPTE. Se désactiver soi-même n'est pas un droit, c'est une
          sortie sans retour.
        - traiter la compense, toucher au référentiel, aux comptes systèmes, aux libellés,
          au format du fichier core banking, ni changer le serveur de la base.

    CES RÈGLES VIVENT DANS LE CODE, PAS DANS CETTE CONTRAINTE. La base ne connaît que la
    liste des rôles acceptés ; qui peut quoi se lit dans Models\UtilisateurWU.vb, et se
    vérifie à deux endroits pour chaque opération : l'écran éteint le bouton, le dépôt
    refuse l'écriture. Un bouton éteint seul se contourne ; un refus seul se découvre trop
    tard.

    CE QUE FAIT CE SCRIPT

    T_UtilisateurWU porte une contrainte CHECK qui énumère les rôles acceptés. Tant qu'elle
    ignore CONTROLEUR, l'application proposera le rôle et SQL Server refusera la ligne —
    avec un message que l'utilisateur ne saura pas interpréter.

    Une contrainte CHECK ne s'étend pas : il faut la retirer et la reposer. C'est la seule
    opération de ce genre dans toute l'installation, et elle est sans risque : la contrainte
    est ÉLARGIE, jamais restreinte. Toutes les lignes existantes la satisfont déjà, et
    WITH CHECK le vérifie plutôt que de le supposer. Aucune donnée n'est touchée.

    POUR REVENIR EN ARRIÈRE : reposer la contrainte à trois rôles, après s'être assuré
    qu'aucun compte ne porte CONTROLEUR —

        SELECT Identifiant FROM dbo.T_UtilisateurWU WHERE Role = N'CONTROLEUR';

    LES RÔLES SQL SERVER NE CHANGENT PAS. wu_admin, wu_compense et wu_commercial sont les
    rôles de la BASE ; CONTROLEUR est un rôle de l'APPLICATION. Tant que l'application se
    connecte avec un compte SQL unique — c'est le cas aujourd'hui — un contrôleur emprunte
    la même connexion que les autres. Le jour où la banque passerait à l'authentification
    Windows, il faudrait dire dans quel rôle SQL ranger un contrôleur : la question est
    ouverte, et elle ne se pose pas avant.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

-- -------------------------------------------------------------------------
-- 1. Retirer l'ancienne contrainte, si elle ignore encore CONTROLEUR
-- -------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM sys.check_constraints
           WHERE name = N'CK_T_UtilisateurWU_Role'
             AND OBJECT_NAME(parent_object_id) = N'T_UtilisateurWU'
             AND definition NOT LIKE N'%CONTROLEUR%')
BEGIN
    ALTER TABLE dbo.T_UtilisateurWU DROP CONSTRAINT CK_T_UtilisateurWU_Role;
    PRINT 'Ancienne contrainte CK_T_UtilisateurWU_Role retirée (trois rôles).';
END
ELSE
    PRINT 'Aucune contrainte à retirer : elle accepte déjà CONTROLEUR, ou elle est absente.';
GO

-- -------------------------------------------------------------------------
-- 2. Reposer la contrainte, avec les quatre rôles
-- -------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'T_UtilisateurWU')
   AND NOT EXISTS (SELECT 1 FROM sys.check_constraints
                   WHERE name = N'CK_T_UtilisateurWU_Role'
                     AND OBJECT_NAME(parent_object_id) = N'T_UtilisateurWU')
BEGIN
    ALTER TABLE dbo.T_UtilisateurWU WITH CHECK
        ADD CONSTRAINT CK_T_UtilisateurWU_Role
        CHECK (Role IN ('COMPENSE', 'COMMERCIAL', 'ADMIN', 'CONTROLEUR'));

    PRINT 'Contrainte CK_T_UtilisateurWU_Role posée (quatre rôles, dont CONTROLEUR).';
END
ELSE
    PRINT 'T_UtilisateurWU : le rôle CONTROLEUR est déjà accepté.';
GO

-- -------------------------------------------------------------------------
-- 3. Constater, plutôt que supposer
-- -------------------------------------------------------------------------
SELECT  name        AS Contrainte,
        definition  AS Definition
FROM    sys.check_constraints
WHERE   name = N'CK_T_UtilisateurWU_Role';

SELECT  Role        AS Role,
        COUNT(*)    AS Comptes
FROM    dbo.T_UtilisateurWU
GROUP BY Role
ORDER BY Role;
GO
