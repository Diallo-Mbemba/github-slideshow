/*
    =========================================================================
    Accès des utilisateurs à la base GWC_WINCOMPENSE_ETD
    =========================================================================

    CE SCRIPT RÉPOND À UNE ERREUR PRÉCISE

        « Login failed for user 'DOMAINE\utilisateur' »   (erreur 18456)
    « Login failed for user 'wincompense' »            (erreur 18456, compte SQL Server)
        « Cannot open database ... requested by the login »  (erreur 4060)

    Ces deux messages ne disent pas que le serveur est injoignable — au contraire,
    ils prouvent qu'il répond. Ils disent que le compte Windows du poste n'a pas
    encore le droit d'entrer.

    Les scripts 01 à 10 créent la base, les tables et les trois RÔLES. Ils ne créent
    aucun COMPTE : le script ne peut pas deviner les identifiants Windows de la
    banque. C'est ce script-ci qui fait le rattachement, et il est le seul à devoir
    être modifié avant d'être exécuté.

    TROIS NIVEAUX À NE PAS CONFONDRE

      1. Le LOGIN     — au niveau du serveur. Il ouvre la porte du bâtiment.
      2. L'UTILISATEUR — au niveau de la base. Il ouvre la porte du bureau.
      3. Le RÔLE      — ce qu'on a le droit d'y faire.

    Il faut les trois. Un login sans utilisateur donne l'erreur 4060 ; un utilisateur
    sans rôle donne « SELECT permission was denied » à la première lecture.

    QUI PEUT L'EXÉCUTER
    Un compte membre de sysadmin, ou à la fois de securityadmin (pour les logins) et
    db_owner sur la base (pour les utilisateurs et les rôles).

    CE SCRIPT EST REJOUABLE : il ne crée que ce qui manque, et n'enlève jamais rien.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. LA SEULE PARTIE À MODIFIER
--
--    Une ligne par compte. Les trois rôles possibles :
--
--      wu_compense    l'agent qui génère la compense et la comptabilise
--      wu_commercial  celui qui saisit les points de vente et consulte
--      wu_admin       l'administrateur : les deux, plus les comptes systèmes
--
--    Un même compte peut figurer sur plusieurs lignes s'il cumule deux rôles.
--
--    « windows » à 1 pour un compte Windows (DOMAINE\pnom) — le cas normal,
--    puisque l'application se connecte en authentification Windows.
--    À 0 pour un compte SQL Server : ce script ne le CRÉE alors pas (il faudrait
--    un mot de passe, qui n'a pas sa place dans un fichier), il se contente de le
--    rattacher à la base et au rôle.
-- =========================================================================
DECLARE @comptes TABLE
(
    rang     INT IDENTITY(1, 1) PRIMARY KEY,
    compte   SYSNAME NOT NULL,
    role_wu  SYSNAME NOT NULL,
    windows  BIT     NOT NULL
);

INSERT INTO @comptes (compte, role_wu, windows) VALUES
    (N'DOMAINE\agent_compense',        N'wu_compense',   1),
    (N'DOMAINE\agent_commercial',      N'wu_commercial', 1),
    (N'DOMAINE\administrateur_wu',     N'wu_admin',      1);

-- Un groupe Active Directory se rattache exactement comme un compte, et évite de
-- revenir sur le serveur à chaque arrivée ou départ :
--  (N'DOMAINE\GRP_WINCOMPENSE_COMPENSE', N'wu_compense', 1);

-- COMPTE SQL SERVER FOURNI PAR LA BANQUE : windows à 0, et le login existe déjà.
-- Ce script ne le crée pas - il faudrait son mot de passe, qui n'a pas sa place dans un
-- fichier - il lui ouvre la base et lui donne son rôle :
--  (N'wincompense', N'wu_admin', 0);
--
-- ATTENTION AU RÔLE dans ce cas. Un compte unique partagé par tout le service doit porter
-- la RÉUNION des droits de tous les postes, donc wu_admin. Les trois rôles cessent alors
-- de distinguer quoi que ce soit au niveau SQL Server : c'est l'application qui garde seule
-- la distinction entre l'agent de compense, le commercial et l'administrateur. Le second
-- verrou disparaît, celui qui s'opposait à une connexion faite hors de l'application.
-- Un compte SQL par agent, si la banque l'accepte, rend aux trois rôles leur utilité.

-- =========================================================================
-- 2. Rattachement
-- =========================================================================
DECLARE @rang     INT,
        @suivant  INT,
        @dernier  INT,
        @compte   SYSNAME,
        @role_wu  SYSNAME,
        @windows  BIT,
        @ordre    NVARCHAR(500);

SELECT @rang = MIN(rang), @dernier = MAX(rang) FROM @comptes;

WHILE @rang IS NOT NULL AND @rang <= @dernier
BEGIN
    SELECT  @compte  = compte,
            @role_wu = role_wu,
            @windows = windows
    FROM    @comptes
    WHERE   rang = @rang;

    IF @compte LIKE N'DOMAINE\%'
    BEGIN
        PRINT N'IGNORÉ : ' + @compte + N' — remplacez DOMAINE par le domaine réel de la banque.';
    END
    ELSE
    BEGIN
        -- --- 2.1 Le login, au niveau du serveur -------------------------------
        IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @compte)
        BEGIN
            IF @windows = 1
            BEGIN
                SET @ordre = N'CREATE LOGIN ' + QUOTENAME(@compte) + N' FROM WINDOWS';
                EXEC sp_executesql @ordre;
                PRINT N'Login créé : ' + @compte;
            END
            ELSE
            BEGIN
                PRINT N'MANQUANT : le login SQL Server ' + @compte + N' n''existe pas. ' +
                      N'Créez-le d''abord avec son mot de passe, puis relancez ce script.';
            END
        END

        -- --- 2.2 L'utilisateur, au niveau de la base --------------------------
        --  Le login peut exister sans que la base le connaisse : c'est exactement
        --  l'erreur 4060, « Cannot open database requested by the login ».
        IF EXISTS (SELECT 1 FROM sys.server_principals WHERE name = @compte)
           AND NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = @compte)
        BEGIN
            SET @ordre = N'CREATE USER ' + QUOTENAME(@compte) + N' FOR LOGIN ' + QUOTENAME(@compte);
            EXEC sp_executesql @ordre;
            PRINT N'Utilisateur créé dans la base : ' + @compte;
        END

        -- --- 2.3 Le rôle ------------------------------------------------------
        IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = @role_wu AND type = 'R')
        BEGIN
            PRINT N'MANQUANT : le rôle ' + @role_wu + N' n''existe pas. ' +
                  N'Exécutez d''abord Scripts\08_RolesSQLServer.sql.';
        END
        ELSE IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = @compte)
        BEGIN
            IF NOT EXISTS (SELECT  1
                           FROM    sys.database_role_members AS m
                           JOIN    sys.database_principals   AS r ON r.principal_id = m.role_principal_id
                           JOIN    sys.database_principals   AS u ON u.principal_id = m.member_principal_id
                           WHERE   r.name = @role_wu AND u.name = @compte)
            BEGIN
                SET @ordre = N'ALTER ROLE ' + QUOTENAME(@role_wu) + N' ADD MEMBER ' + QUOTENAME(@compte);
                EXEC sp_executesql @ordre;
                PRINT N'Rôle ' + @role_wu + N' accordé à ' + @compte;
            END
            ELSE
            BEGIN
                PRINT N'Déjà en place : ' + @compte + N' dans ' + @role_wu;
            END
        END
    END

    -- Variable distincte : affecter @rang depuis une requete qui le lit encore
    -- dans son WHERE se lirait mal, et demanderait de connaitre l'ordre
    -- d'evaluation de SQL Server pour etre sur de ce qu'on ecrit.
    SELECT @suivant = MIN(rang) FROM @comptes WHERE rang > @rang;
    SET @rang = @suivant;
END
GO

-- =========================================================================
-- 3. Contrôle : qui a accès à cette base, et à quel titre
--
--    Si cette liste est vide en dehors de dbo, personne ne pourra ouvrir
--    l'application : c'est que la section 1 n'a pas été modifiée.
-- =========================================================================
SELECT  compte      = u.name,
        genre       = u.type_desc,
        role_wu     = ISNULL(r.name, N'(aucun rôle Wincompense)'),
        login_serveur = CASE WHEN s.name IS NULL
                             THEN N'ABSENT — le compte ne pourra pas se connecter'
                             ELSE s.name END
FROM    sys.database_principals AS u
LEFT JOIN sys.database_role_members AS m ON m.member_principal_id = u.principal_id
LEFT JOIN sys.database_principals   AS r ON r.principal_id = m.role_principal_id
                                        AND r.name LIKE N'wu[_]%'
LEFT JOIN sys.server_principals     AS s ON s.sid = u.sid
WHERE   u.type IN ('U', 'S', 'G')
  AND   u.name NOT IN (N'dbo', N'guest', N'INFORMATION_SCHEMA', N'sys')
ORDER BY u.name, r.name;
GO

PRINT N'';
PRINT N'=== Rattachement terminé ===';
PRINT N'Sur le poste : Sécurité > Connexion à la base de données, puis « Tester la connexion ».';
PRINT N'Le test doit répondre « Connexion réussie » en nommant la base et le compte.';
GO
