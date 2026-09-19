Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Text

''' <summary>
''' Regarde si le compte de la chaîne de connexion a bien accès à la base, répare ce qu'il
''' peut réparer, et rédige pour la banque le reste.
'''
''' POURQUOI L'APPLICATION NE PEUT PAS TOUT FAIRE
'''
''' Créer un accès à SQL Server demande des droits d'administration du serveur. Or c'est
''' précisément ce que le compte fourni par la banque n'a presque jamais — et c'est heureux :
''' une application comptable qui pourrait se donner des droits à elle-même n'aurait plus de
''' contrôle d'accès du tout. Il y a là un cercle : pour créer l'accès, il faut déjà l'avoir.
'''
''' Ce service ne prétend donc pas le rompre. Il fait trois choses, dans cet ordre :
'''
'''   1. il CONSTATE — le compte entre-t-il ? la base existe-t-elle ? les tables ? les rôles ?
'''      le compte est-il membre de l'un d'eux ? Chaque réponse est lue sur le serveur, jamais
'''      supposée.
'''   2. il AGIT si le compte a les droits — ce qui arrive quand la banque donne un compte
'''      administrateur, ou quand l'agent prépare une base de test.
'''   3. il RÉDIGE sinon — le T-SQL exact, noms réels substitués, à remettre à la banque.
'''
''' Le point 3 n'est pas un pis-aller : c'est le cas normal en production. Mieux vaut un
''' script juste et prêt à exécuter qu'une tentative qui échoue en laissant l'agent deviner.
''' </summary>
Public NotInheritable Class PreparationBaseWU

    Private Sub New()
    End Sub

    ''' <summary>Les trois rôles de la base, dans l'ordre croissant de droits.</summary>
    Private Shared ReadOnly ROLES As String() = {"wu_compense", "wu_commercial", "wu_admin"}

    ''' <summary>
    ''' Rôle proposé par défaut à un compte SQL Server partagé.
    '''
    ''' wu_admin, et non wu_compense : un compte unique employé par tous les postes doit porter
    ''' la réunion des droits de tous les postes. Un compte par agent permettrait de descendre
    ''' au rôle réel de chacun — le script le rappelle en commentaire.
    ''' </summary>
    Public Const ROLE_PROPOSE As String = "wu_admin"

    ''' <summary>Table témoin : si elle manque, la base n'a jamais reçu les scripts.</summary>
    Private Const TABLE_TEMOIN As String = "dbo.T_UtilisateurWU"

#Region "Le constat"

    ''' <summary>Ce que le serveur a répondu. Aucun champ n'est supposé : tous sont lus.</summary>
    Public NotInheritable Class Rapport

        Public Property Joignable As Boolean
        Public Property MessageDeRefus As String = String.Empty

        Public Property Compte As String = String.Empty
        Public Property Base As String = String.Empty

        Public Property EstAdministrateurDuServeur As Boolean
        Public Property EstProprietaireDeLaBase As Boolean

        Public Property BaseExiste As Boolean
        Public Property AccesALaBase As Boolean
        Public Property UtilisateurDeLaBase As String = String.Empty
        Public Property TablesPresentes As Boolean

        Public Property RolesPresents As Integer
        Public Property RoleDejaAccorde As String = String.Empty

        ''' <summary>Vrai si le compte peut déjà travailler : rien à préparer.</summary>
        Public ReadOnly Property RienAFaire As Boolean
            Get
                Return Joignable AndAlso AccesALaBase AndAlso TablesPresentes AndAlso
                       (RoleDejaAccorde.Length > 0 OrElse EstProprietaireDeLaBase)
            End Get
        End Property

        ''' <summary>Vrai si l'application a les droits de faire elle-même ce qui manque.</summary>
        Public ReadOnly Property PeutAgir As Boolean
            Get
                If Not Joignable Then Return False
                If Not BaseExiste Then Return False
                If EstAdministrateurDuServeur Then Return True

                ' Sans être administrateur du serveur, on peut encore ajouter un rôle à
                ' quelqu'un qui entre déjà dans la base, si l'on en est propriétaire.
                Return AccesALaBase AndAlso EstProprietaireDeLaBase
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Interroge le serveur sur l'état du compte porté par cette chaîne.
    '''
    ''' L'interrogation se fait sur master, et non sur la base visée : une base qui refuse le
    ''' compte refuse aussi la connexion, et l'on ne saurait alors rien dire du tout. Le second
    ''' temps ouvre la base elle-même, quand elle est accessible.
    ''' </summary>
    Public Shared Function Analyser(chaine As String) As Rapport

        Dim rapport As New Rapport()

        Dim baseVisee As String = NomDeLaBase(chaine)
        rapport.Base = baseVisee

        Try
            Using connexion As New SqlConnection(VersLaBase(chaine, "master"))
                connexion.Open()
                rapport.Joignable = True

                Using commande As New SqlCommand(
                    "SELECT compte      = SUSER_SNAME()," &
                    "       sysadmin    = CONVERT(int, ISNULL(IS_SRVROLEMEMBER('sysadmin'), 0))," &
                    "       base_existe = CASE WHEN DB_ID(@base) IS NULL THEN 0 ELSE 1 END," &
                    "       acces       = CONVERT(int, ISNULL(HAS_DBACCESS(@base), 0))", connexion)

                    commande.Parameters.AddWithValue("@base", baseVisee)

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        If lecteur.Read() Then
                            rapport.Compte = lecteur.GetString(0)
                            rapport.EstAdministrateurDuServeur = lecteur.GetInt32(1) = 1
                            rapport.BaseExiste = lecteur.GetInt32(2) = 1
                            rapport.AccesALaBase = lecteur.GetInt32(3) = 1
                        End If
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            rapport.MessageDeRefus = DiagnosticSqlWU.Expliquer(ex, chaine)
            If rapport.MessageDeRefus.Length = 0 Then rapport.MessageDeRefus = ex.Message
            Return rapport

        Catch ex As InvalidOperationException
            rapport.MessageDeRefus = "Chaîne de connexion inexploitable : " & ex.Message
            Return rapport
        End Try

        If rapport.BaseExiste AndAlso rapport.AccesALaBase Then LireLaBase(chaine, rapport)

        Return rapport
    End Function

    ''' <summary>Second temps : ce qui ne se voit que de l'intérieur de la base.</summary>
    Private Shared Sub LireLaBase(chaine As String, rapport As Rapport)

        Try
            Using connexion As New SqlConnection(VersLaBase(chaine, rapport.Base))
                connexion.Open()

                Using commande As New SqlCommand(
                    "SELECT utilisateur = USER_NAME()," &
                    "       proprietaire = CONVERT(int, ISNULL(IS_MEMBER('db_owner'), 0))," &
                    "       tables      = CASE WHEN OBJECT_ID(@temoin) IS NULL THEN 0 ELSE 1 END," &
                    "       roles       = (SELECT COUNT(*) FROM sys.database_principals" &
                    "                      WHERE type = 'R' AND name IN ('wu_compense','wu_commercial','wu_admin'))," &
                    "       r_compense  = CONVERT(int, ISNULL(IS_ROLEMEMBER('wu_compense'), 0))," &
                    "       r_commercial = CONVERT(int, ISNULL(IS_ROLEMEMBER('wu_commercial'), 0))," &
                    "       r_admin     = CONVERT(int, ISNULL(IS_ROLEMEMBER('wu_admin'), 0))", connexion)

                    commande.Parameters.AddWithValue("@temoin", TABLE_TEMOIN)

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then Return

                        rapport.UtilisateurDeLaBase = lecteur.GetString(0)
                        rapport.EstProprietaireDeLaBase = lecteur.GetInt32(1) = 1
                        rapport.TablesPresentes = lecteur.GetInt32(2) = 1
                        rapport.RolesPresents = lecteur.GetInt32(3)

                        ' Le rôle le plus large l'emporte dans l'affichage : c'est lui qui
                        ' décide de ce que le compte peut faire.
                        If lecteur.GetInt32(6) = 1 Then
                            rapport.RoleDejaAccorde = "wu_admin"
                        ElseIf lecteur.GetInt32(5) = 1 Then
                            rapport.RoleDejaAccorde = "wu_commercial"
                        ElseIf lecteur.GetInt32(4) = 1 Then
                            rapport.RoleDejaAccorde = "wu_compense"
                        End If
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            ' La base s'est fermée entre les deux lectures, ou une vue système est interdite
            ' à ce compte : le constat reste ce qu'il est, et le script sera proposé.

        Catch ex As InvalidOperationException
            ' Même raison.
        End Try
    End Sub

#End Region

#Region "L'action, quand les droits le permettent"

    ''' <summary>
    ''' Exécute ce qui manque. Chaque ordre est tenté séparément et consigné : un refus sur
    ''' l'un ne doit pas empêcher les autres, et l'agent doit pouvoir lire ce qui a été fait.
    ''' </summary>
    ''' <param name="chaine">Chaîne de connexion de la banque.</param>
    ''' <param name="rapport">Constat établi par <see cref="Analyser"/>.</param>
    ''' <param name="journal">Compte rendu, ligne par ligne.</param>
    ''' <returns>Vrai si tout ce qui manquait a été fait.</returns>
    Public Shared Function Appliquer(chaine As String, rapport As Rapport, ByRef journal As String) As Boolean

        Dim lignes As New StringBuilder()
        Dim tout As Boolean = True

        If rapport Is Nothing OrElse Not rapport.PeutAgir Then
            journal = "Ce compte n'a pas les droits nécessaires : rien n'a été tenté."
            Return False
        End If

        ' L'utilisateur de base se crée depuis la base elle-même. Sans accès, il faut être
        ' administrateur du serveur — qui, lui, entre partout.
        If Not rapport.AccesALaBase Then
            If Executer(VersLaBase(chaine, rapport.Base),
                        $"CREATE USER {Crochets(rapport.Compte)} FOR LOGIN {Crochets(rapport.Compte)}",
                        $"Utilisateur {rapport.Compte} créé dans {rapport.Base}", lignes) Then
                rapport.AccesALaBase = True
            Else
                tout = False
            End If
        End If

        If rapport.RolesPresents < ROLES.Length Then
            lignes.AppendLine("Les trois rôles ne sont pas tous présents : exécutez " &
                              "Scripts\08_RolesSQLServer.sql, ou le script complet.")
            tout = False
        End If

        If rapport.RoleDejaAccorde.Length = 0 AndAlso rapport.AccesALaBase Then
            If Executer(VersLaBase(chaine, rapport.Base),
                        $"ALTER ROLE {ROLE_PROPOSE} ADD MEMBER {Crochets(rapport.Compte)}",
                        $"Rôle {ROLE_PROPOSE} accordé à {rapport.Compte}", lignes) Then
                rapport.RoleDejaAccorde = ROLE_PROPOSE
            Else
                tout = False
            End If
        End If

        If Not rapport.TablesPresentes Then
            lignes.AppendLine("La base ne porte pas les tables de l'application : exécutez " &
                              "Scripts\00_InstallationComplete.sql. Aucun accès ne remplace " &
                              "des tables absentes.")
            tout = False
        End If

        If lignes.Length = 0 Then lignes.AppendLine("Il n'y avait rien à faire.")

        journal = lignes.ToString().TrimEnd()
        Return tout
    End Function

    ''' <summary>Exécute un ordre et l'inscrit au journal, qu'il passe ou non.</summary>
    Private Shared Function Executer(chaine As String, ordre As String,
                                     libelle As String, lignes As StringBuilder) As Boolean
        Try
            Using connexion As New SqlConnection(chaine)
                connexion.Open()

                Using commande As New SqlCommand(ordre, connexion)
                    commande.ExecuteNonQuery()
                End Using
            End Using

            lignes.AppendLine("OK   — " & libelle)
            Return True

        Catch ex As SqlException
            lignes.AppendLine("REFUSÉ — " & libelle)
            lignes.AppendLine("         " & ex.Message)
            Return False

        Catch ex As InvalidOperationException
            lignes.AppendLine("REFUSÉ — " & libelle)
            lignes.AppendLine("         " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "Le script, quand les droits manquent"

    ''' <summary>
    ''' Le T-SQL à remettre à la banque, noms réels substitués, réduit à ce qui manque
    ''' réellement. Un script qui refait ce qui est déjà fait se fait renvoyer.
    ''' </summary>
    Public Shared Function ScriptPourLaBanque(rapport As Rapport) As String

        Dim compte As String = If(rapport IsNot Nothing AndAlso rapport.Compte.Length > 0,
                                  rapport.Compte, "COMPTE-FOURNI-PAR-LA-BANQUE")
        Dim base As String = If(rapport IsNot Nothing AndAlso rapport.Base.Length > 0,
                                rapport.Base, ConfigurationWU.BASE_PAR_DEFAUT)

        Dim texte As New StringBuilder()

        texte.AppendLine("-- =====================================================================")
        texte.AppendLine("-- Wincompense — accès du compte applicatif à la base " & base)
        texte.AppendLine("--")
        texte.AppendLine("-- À exécuter sur le serveur SQL Server par un compte membre de sysadmin,")
        texte.AppendLine("-- ou à la fois de securityadmin et db_owner sur cette base.")
        texte.AppendLine("--")
        texte.AppendLine("-- Rejouable : chaque ordre ne s'exécute que si ce qu'il crée manque.")
        texte.AppendLine("-- =====================================================================")
        texte.AppendLine()

        If rapport Is Nothing OrElse Not rapport.Joignable Then
            texte.AppendLine("-- Le compte n'a pas pu se connecter au serveur. Si le login n'existe pas")
            texte.AppendLine("-- encore, le créer d'abord — le mot de passe est celui de la banque :")
            texte.AppendLine("--")
            texte.AppendLine("-- CREATE LOGIN " & Crochets(compte) & " WITH PASSWORD = N'...';")
            texte.AppendLine()
        End If

        If rapport IsNot Nothing AndAlso Not rapport.BaseExiste Then
            texte.AppendLine("-- La base " & base & " n'existe pas sur ce serveur.")
            texte.AppendLine("-- Exécuter d'abord Scripts\00_InstallationComplete.sql.")
            texte.AppendLine()
        End If

        texte.AppendLine("USE " & base & ";")
        texte.AppendLine("GO")
        texte.AppendLine()
        texte.AppendLine("-- 1. Ouvrir la base au compte (erreur 4060 tant que c'est absent)")
        texte.AppendLine("IF DATABASE_PRINCIPAL_ID(N'" & Echapper(compte) & "') IS NULL")
        texte.AppendLine("    CREATE USER " & Crochets(compte) & " FOR LOGIN " & Crochets(compte) & ";")
        texte.AppendLine("GO")
        texte.AppendLine()
        texte.AppendLine("-- 2. Lui donner son rôle")
        texte.AppendLine("--")
        texte.AppendLine("--    " & ROLE_PROPOSE & " parce qu'un compte unique, partagé par tous les postes,")
        texte.AppendLine("--    doit porter la réunion des droits de tous les postes. Si la banque fournit")
        texte.AppendLine("--    un compte par agent, remplacer par wu_compense (agent de compense) ou")
        texte.AppendLine("--    wu_commercial (saisie des points de vente) : les rôles retrouvent alors")
        texte.AppendLine("--    leur utilité, et un accès direct à la base reste borné au métier réel.")
        texte.AppendLine("IF DATABASE_PRINCIPAL_ID(N'" & ROLE_PROPOSE & "') IS NULL")
        texte.AppendLine("    RAISERROR(N'Le rôle " & ROLE_PROPOSE &
                         " n''existe pas : exécutez d''abord 08_RolesSQLServer.sql.', 16, 1);")
        texte.AppendLine("ELSE IF IS_ROLEMEMBER(N'" & ROLE_PROPOSE & "', N'" & Echapper(compte) & "') = 0")
        texte.AppendLine("    ALTER ROLE " & ROLE_PROPOSE & " ADD MEMBER " & Crochets(compte) & ";")
        texte.AppendLine("GO")
        texte.AppendLine()
        texte.AppendLine("-- 3. Contrôle — les trois lignes doivent répondre « oui »")
        texte.AppendLine("SELECT  compte       = N'" & Echapper(compte) & "',")
        texte.AppendLine("        utilisateur  = CASE WHEN DATABASE_PRINCIPAL_ID(N'" & Echapper(compte) &
                         "') IS NULL THEN N'non' ELSE N'oui' END,")
        texte.AppendLine("        role_" & ROLE_PROPOSE.Replace("wu_", "") &
                         "   = CASE WHEN IS_ROLEMEMBER(N'" & ROLE_PROPOSE & "', N'" & Echapper(compte) &
                         "') = 1 THEN N'oui' ELSE N'non' END,")
        texte.AppendLine("        tables       = CASE WHEN OBJECT_ID(N'" & TABLE_TEMOIN &
                         "') IS NULL THEN N'non' ELSE N'oui' END;")
        texte.AppendLine("GO")

        Return texte.ToString()
    End Function

#End Region

#Region "Utilitaires"

    ''' <summary>La même chaîne, dirigée vers une autre base.</summary>
    Private Shared Function VersLaBase(chaine As String, base As String) As String

        Try
            Dim constructeur As New SqlConnectionStringBuilder(ConfigurationWU.ChaineEssayable(chaine)) With {
                .InitialCatalog = base
            }
            Return constructeur.ConnectionString

        Catch ex As ArgumentException
            Return chaine
        End Try
    End Function

    Private Shared Function NomDeLaBase(chaine As String) As String

        Try
            Dim nom As String = New SqlConnectionStringBuilder(chaine).InitialCatalog.Trim()
            Return If(nom.Length > 0, nom, ConfigurationWU.BASE_PAR_DEFAUT)

        Catch ex As ArgumentException
            Return ConfigurationWU.BASE_PAR_DEFAUT
        End Try
    End Function

    ''' <summary>Encadre un nom pour le T-SQL : sans cela, DOMAINE\pnom fait échouer l'ordre.</summary>
    Private Shared Function Crochets(nom As String) As String
        Return "[" & If(nom, String.Empty).Replace("]", "]]") & "]"
    End Function

    ''' <summary>Double les apostrophes d'un nom placé dans une chaîne littérale T-SQL.</summary>
    Private Shared Function Echapper(nom As String) As String
        Return If(nom, String.Empty).Replace("'", "''")
    End Function

#End Region

End Class
