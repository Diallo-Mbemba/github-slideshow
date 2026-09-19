/*
    =========================================================================
    Wincompense — accès du compte applicatif fourni par la banque
    =========================================================================

    À REMETTRE TEL QUEL À L'INFORMATIQUE DE LA BANQUE.

    UNE SEULE LIGNE À MODIFIER : celle de @compte, ci-dessous. Elle porte le nom du
    compte que la banque a communiqué avec la chaîne de connexion — le nom seul, jamais
    le mot de passe, qui n'a pas sa place dans un fichier.

    CE QUE CE SCRIPT FAIT, ET POURQUOI
    La banque a créé un LOGIN : le compte peut atteindre le serveur. Il lui manque deux
    choses, que seul un administrateur peut donner :

        1. un UTILISATEUR dans la base   — sans lui : « Cannot open database » (4060)
        2. un RÔLE                        — sans lui : « SELECT permission was denied » (229)

    Trois niveaux qu'on confond facilement : le login ouvre la porte du bâtiment,
    l'utilisateur celle du bureau, le rôle dit ce qu'on a le droit d'y faire.

    POURQUOI L'APPLICATION NE LE FAIT PAS ELLE-MÊME
    Elle essaie — bouton « Préparer la base… » de l'écran Sécurité > Connexion — et y
    parvient quand le compte administre le serveur. Ce n'est presque jamais le cas, et
    c'est heureux : une application comptable qui pourrait se donner des droits à
    elle-même n'aurait plus de contrôle d'accès du tout.

    PRÉREQUIS
    La base et ses tables doivent exister : Scripts\00_InstallationComplete.sql.
    Exécuter avec un compte membre de sysadmin, ou de db_owner sur cette base.

    REJOUABLE : chaque ordre ne s'exécute que si ce qu'il crée manque.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- LA SEULE LIGNE À MODIFIER
-- =========================================================================
DECLARE @compte SYSNAME = N'REMPLACER-PAR-LE-COMPTE-DE-LA-BANQUE';

-- Rôle accordé au compte.
--
--   wu_admin      un compte unique partagé par tous les postes : il doit porter la
--                 réunion des droits de tous les postes.
--   wu_compense   un compte par agent, poste de compense.
--   wu_commercial un compte par agent, saisie des points de vente.
--
-- Un compte par agent vaut mieux : les trois rôles retrouvent alors leur utilité, et un
-- accès direct à la base — par SQL Server Management Studio, hors de l'application —
-- reste borné au métier réel de chacun.
DECLARE @role SYSNAME = N'wu_admin';

DECLARE @ordre NVARCHAR(400);

IF @compte = N'REMPLACER-PAR-LE-COMPTE-DE-LA-BANQUE'
BEGIN
    RAISERROR(N'Renseignez @compte avant d''exécuter ce script.', 16, 1);
END
ELSE
BEGIN

    -- 1. Le login existe-t-il ? La banque l'a créé avec le mot de passe ; s'il manque,
    --    ce script ne peut pas le créer — il faudrait ce mot de passe.
    IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @compte)
    BEGIN
        PRINT N'ARRÊT : le login ' + @compte + N' n''existe pas sur ce serveur.';
        PRINT N'        Créez-le d''abord avec son mot de passe :';
        PRINT N'            CREATE LOGIN ' + QUOTENAME(@compte) + N' WITH PASSWORD = N''...'';';
    END
    ELSE
    BEGIN

        -- 2. L'utilisateur de base.
        IF DATABASE_PRINCIPAL_ID(@compte) IS NULL
        BEGIN
            SET @ordre = N'CREATE USER ' + QUOTENAME(@compte) + N' FOR LOGIN ' + QUOTENAME(@compte);
            EXEC sp_executesql @ordre;
            PRINT N'Utilisateur créé dans la base : ' + @compte;
        END
        ELSE
        BEGIN
            PRINT N'Déjà en place : utilisateur ' + @compte;
        END

        -- 3. Le rôle.
        IF DATABASE_PRINCIPAL_ID(@role) IS NULL
        BEGIN
            PRINT N'ARRÊT : le rôle ' + @role + N' n''existe pas.';
            PRINT N'        Exécutez Scripts\08_RolesSQLServer.sql, ou le script complet 00.';
        END
        -- ISNULL : si la création de l'utilisateur a échoué, IS_ROLEMEMBER rend NULL,
        -- et un NULL = 0 vaut « inconnu », donc faux : on passerait dans la branche
        -- « déjà en place » sans que rien ne le soit.
        ELSE IF ISNULL(IS_ROLEMEMBER(@role, @compte), 0) = 0
        BEGIN
            SET @ordre = N'ALTER ROLE ' + QUOTENAME(@role) + N' ADD MEMBER ' + QUOTENAME(@compte);
            EXEC sp_executesql @ordre;
            PRINT N'Rôle ' + @role + N' accordé à ' + @compte;
        END
        ELSE
        BEGIN
            PRINT N'Déjà en place : ' + @compte + N' dans ' + @role;
        END

        -- 4. Contrôle — les trois colonnes doivent répondre « oui ».
        SELECT  compte      = @compte,
                utilisateur = CASE WHEN DATABASE_PRINCIPAL_ID(@compte) IS NULL
                                   THEN N'non' ELSE N'oui' END,
                role_accorde = CASE WHEN IS_ROLEMEMBER(@role, @compte) = 1
                                    THEN N'oui' ELSE N'non' END,
                tables      = CASE WHEN OBJECT_ID(N'dbo.T_UtilisateurWU') IS NULL
                                   THEN N'non — exécutez 00_InstallationComplete.sql' ELSE N'oui' END;
    END
END
GO

PRINT N'';
PRINT N'=== Terminé ===';
PRINT N'Sur le poste : Sécurité > Connexion à la base de données, puis « Tester la connexion ».';
PRINT N'Le test doit répondre « Connexion réussie » en nommant la base et le compte.';
GO
