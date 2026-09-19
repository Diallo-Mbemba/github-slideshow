Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.Text

''' <summary>
''' Traduit les refus de SQL Server en une consigne exécutable.
'''
''' Un message comme « Login failed for user 'ETD\wincompense' » est exact, mais il laisse
''' l'agent devant son écran sans rien à faire : il est en anglais, il ne dit pas si le tort
''' est au serveur, à la base ou au compte, et surtout il ne dit pas quoi demander à
''' l'informatique. Cette classe répond à ces trois questions, et donne le T-SQL à exécuter.
'''
''' Elle ne décide de rien et n'écrit nulle part : elle rédige.
''' </summary>
Public NotInheritable Class DiagnosticSqlWU

    Private Sub New()
    End Sub

#Region "Numéros d'erreur SQL Server"

    ''' <summary>Login refusé : le compte n'existe pas sur l'instance, ou le mot de passe est faux.</summary>
    Private Const REFUS_DE_CONNEXION As Integer = 18456

    ''' <summary>Le login existe sur le serveur, mais la base ne le connaît pas.</summary>
    Private Const BASE_INACCESSIBLE As Integer = 4060

    ''' <summary>Base par défaut du login introuvable ou hors ligne.</summary>
    Private Const BASE_PAR_DEFAUT_INTROUVABLE As Integer = 4064

    ''' <summary>Authentification Windows refusée : domaine non approuvé par le serveur.</summary>
    Private Const DOMAINE_NON_APPROUVE As Integer = 18452

    ''' <summary>Le compte est dans la base, mais pas sous un contexte qui lui donne accès.</summary>
    Private Const CONTEXTE_DE_SECURITE As Integer = 916

    ''' <summary>SELECT, INSERT ou UPDATE refusé : l'utilisateur existe mais n'a pas de rôle.</summary>
    Private Shared ReadOnly DROITS_REFUSES As Integer() = {229, 230, 262, 297}

    ''' <summary>Serveur injoignable : nom inconnu, instance arrêtée, pare-feu, réseau.</summary>
    Private Shared ReadOnly SERVEUR_INJOIGNABLE As Integer() = {-1, 2, 40, 53, 1231, 10060, 10061, 11001}

#End Region

#Region "Lecture de la chaîne de connexion"

    ''' <summary>
    ''' Le compte au nom duquel la connexion est tentée : l'utilisateur Windows du poste en
    ''' authentification intégrée, sinon le User ID de la chaîne.
    '''
    ''' On le prend dans la chaîne et dans Windows, et non dans le message d'erreur : ce
    ''' message est traduit dans la langue du serveur, et son découpage varierait avec elle.
    ''' </summary>
    Public Shared Function CompteDeConnexion(chaine As String) As String

        Dim lecteur As SqlConnectionStringBuilder = Lire(chaine)

        If lecteur IsNot Nothing AndAlso Not lecteur.IntegratedSecurity Then
            If lecteur.UserID.Trim().Length > 0 Then Return lecteur.UserID.Trim()
        End If

        Try
            Return WindowsIdentity.GetCurrent().Name
        Catch ex As SystemException
            ' Identité Windows illisible : le cas est rare, mais il ne doit pas
            ' empêcher d'afficher le reste du diagnostic.
            Return "(compte Windows du poste)"
        End Try
    End Function

    ''' <summary>Vrai si la chaîne demande l'authentification Windows — le cas normal ici.</summary>
    Public Shared Function EstAuthentificationWindows(chaine As String) As Boolean

        Dim lecteur As SqlConnectionStringBuilder = Lire(chaine)
        If lecteur Is Nothing Then Return True
        Return lecteur.IntegratedSecurity
    End Function

    Private Shared Function Lire(chaine As String) As SqlConnectionStringBuilder

        If String.IsNullOrWhiteSpace(chaine) Then Return Nothing

        Try
            Return New SqlConnectionStringBuilder(chaine)
        Catch ex As ArgumentException
            ' Chaîne mal formée : ce n'est pas à cette classe de le signaler, l'ouverture
            ' de la connexion l'aura déjà fait.
            Return Nothing
        End Try
    End Function

    Private Shared Function NomDeLaBase(chaine As String) As String

        Dim lecteur As SqlConnectionStringBuilder = Lire(chaine)
        If lecteur Is Nothing OrElse lecteur.InitialCatalog.Trim().Length = 0 Then
            Return "GWC_WINCOMPENSE_ETD"
        End If
        Return lecteur.InitialCatalog.Trim()
    End Function

#End Region

#Region "Le diagnostic"

    ''' <summary>
    ''' Explication en français, suivie de ce qu'il faut faire. Rend une chaîne vide si
    ''' l'erreur n'est pas de celles que cette classe sait expliquer : l'appelant garde
    ''' alors son propre message plutôt que d'en recevoir un vague.
    ''' </summary>
    Public Shared Function Expliquer(ex As SqlException, chaine As String) As String

        If ex Is Nothing Then Return String.Empty

        Select Case ex.Number

            Case REFUS_DE_CONNEXION, DOMAINE_NON_APPROUVE
                Return RefusDuCompte(chaine, ex)

            Case BASE_INACCESSIBLE, CONTEXTE_DE_SECURITE
                Return BaseFermeeAuCompte(chaine)

            Case BASE_PAR_DEFAUT_INTROUVABLE
                Return "Le compte est accepté par le serveur, mais sa base par défaut est " &
                       "introuvable ou hors ligne." & vbCrLf & vbCrLf &
                       "L'informatique doit lui donner une base par défaut valable :" & vbCrLf & vbCrLf &
                       "    ALTER LOGIN " & Crochets(CompteDeConnexion(chaine)) &
                       " WITH DEFAULT_DATABASE = " & NomDeLaBase(chaine) & ";"

            Case Else

                If Array.IndexOf(DROITS_REFUSES, ex.Number) >= 0 Then Return DroitManquant(chaine, ex)
                If Array.IndexOf(SERVEUR_INJOIGNABLE, ex.Number) >= 0 Then Return ServeurInjoignable(chaine)

                Return String.Empty
        End Select
    End Function

    ''' <summary>18456 et 18452 : la porte du bâtiment.</summary>
    Private Shared Function RefusDuCompte(chaine As String, ex As SqlException) As String

        Dim compte As String = CompteDeConnexion(chaine)
        Dim texte As New StringBuilder()

        texte.AppendLine("Le serveur répond, mais il refuse le compte « " & compte & " ».")
        texte.AppendLine()
        texte.AppendLine("Ce n'est donc ni le serveur ni la base qui sont en cause : la connexion " &
                         "arrive bien à destination. Ce compte n'a simplement pas encore le droit " &
                         "d'entrer sur cette instance SQL Server.")
        texte.AppendLine()

        If EstAuthentificationWindows(chaine) Then

            texte.AppendLine("Les scripts d'installation créent la base, les tables et les trois rôles. " &
                             "Ils ne créent aucun compte : ils ne peuvent pas deviner les identifiants " &
                             "Windows de la banque. Il reste ceci à exécuter sur le serveur :")
            texte.AppendLine()
            texte.AppendLine("    USE " & NomDeLaBase(chaine) & ";")
            texte.AppendLine("    CREATE LOGIN " & Crochets(compte) & " FROM WINDOWS;")
            texte.AppendLine("    CREATE USER  " & Crochets(compte) & " FOR LOGIN " & Crochets(compte) & ";")
            texte.AppendLine("    ALTER ROLE wu_compense ADD MEMBER " & Crochets(compte) & ";")
            texte.AppendLine()
            texte.AppendLine("Le rôle dépend du poste : wu_compense pour l'agent de compense, " &
                             "wu_commercial pour la saisie des points de vente, wu_admin pour " &
                             "l'administrateur.")
            texte.AppendLine()
            texte.AppendLine("Scripts\11_AccesUtilisateurs.sql fait ce rattachement pour toute une " &
                             "liste de comptes en une fois.")
        Else

            texte.AppendLine("La chaîne de connexion utilise un compte SQL Server, et non " &
                             "l'authentification Windows. Trois causes possibles, dans cet ordre :")
            texte.AppendLine()
            texte.AppendLine("    1. le mot de passe de la chaîne ne correspond pas ;")
            texte.AppendLine("    2. le compte « " & compte & " » n'existe pas sur cette instance ;")
            texte.AppendLine("    3. l'instance n'accepte que l'authentification Windows " &
                             "(le mode mixte n'y est pas activé).")
            texte.AppendLine()
            texte.AppendLine("Demandez confirmation à la banque de la chaîne fournie — en particulier " &
                             "du mot de passe, qui se recopie mal.")
        End If

        If ex.Number = DOMAINE_NON_APPROUVE Then
            texte.AppendLine()
            texte.AppendLine("Le serveur signale en outre que le domaine de ce compte ne lui est pas " &
                             "approuvé : le poste et le serveur ne sont peut-être pas dans le même " &
                             "domaine Active Directory.")
        End If

        Return texte.ToString().TrimEnd()
    End Function

    ''' <summary>4060 et 916 : la porte du bureau. Le login existe, la base l'ignore.</summary>
    Private Shared Function BaseFermeeAuCompte(chaine As String) As String

        Dim compte As String = CompteDeConnexion(chaine)
        Dim base As String = NomDeLaBase(chaine)

        Return "Le compte « " & compte & " » est accepté par le serveur, mais la base " &
               base & " ne le connaît pas." & vbCrLf & vbCrLf &
               "C'est la moitié du chemin : le login existe, l'utilisateur de base " &
               "reste à créer." & vbCrLf & vbCrLf &
               "    USE " & base & ";" & vbCrLf &
               "    CREATE USER " & Crochets(compte) & " FOR LOGIN " & Crochets(compte) & ";" & vbCrLf &
               "    ALTER ROLE wu_compense ADD MEMBER " & Crochets(compte) & ";" & vbCrLf & vbCrLf &
               "Vérifiez au passage que la base nommée dans la chaîne est la bonne : une base " &
               "inexistante donne le même message."
    End Function

    ''' <summary>229 et voisins : l'utilisateur est entré, mais n'a aucun rôle.</summary>
    Private Shared Function DroitManquant(chaine As String, ex As SqlException) As String

        Dim compte As String = CompteDeConnexion(chaine)

        Return "Le compte « " & compte & " » entre bien dans la base, mais il n'a le droit d'y " &
               "lire ni d'y écrire." & vbCrLf & vbCrLf &
               "Message du serveur : " & ex.Message & vbCrLf & vbCrLf &
               "Il lui manque son rôle. Un seul ordre suffit :" & vbCrLf & vbCrLf &
               "    USE " & NomDeLaBase(chaine) & ";" & vbCrLf &
               "    ALTER ROLE wu_compense ADD MEMBER " & Crochets(compte) & ";" & vbCrLf & vbCrLf &
               "Remplacez wu_compense par wu_commercial ou wu_admin selon le poste. Si le serveur " &
               "répond que le rôle n'existe pas, c'est que Scripts\08_RolesSQLServer.sql n'a pas " &
               "encore été exécuté."
    End Function

    ''' <summary>Réseau : rien n'a répondu.</summary>
    Private Shared Function ServeurInjoignable(chaine As String) As String

        Dim lecteur As SqlConnectionStringBuilder = Lire(chaine)
        Dim serveur As String = If(lecteur Is Nothing, "(serveur non lu)", lecteur.DataSource)

        Return "Le serveur « " & serveur & " » n'a pas répondu." & vbCrLf & vbCrLf &
               "Contrairement au refus de compte, rien n'est arrivé à destination. À vérifier, " &
               "dans cet ordre :" & vbCrLf & vbCrLf &
               "    1. le nom du serveur est-il exactement celui fourni par la banque ?" & vbCrLf &
               "    2. l'instance est-elle nommée (SERVEUR\INSTANCE) ou par défaut ?" & vbCrLf &
               "    3. le port 1433 est-il ouvert depuis ce poste ?" & vbCrLf &
               "    4. le service SQL Server tourne-t-il sur le serveur ?"
    End Function

    ''' <summary>
    ''' Encadre un nom de compte pour le T-SQL. Sans crochets, un compte de domaine
    ''' (DOMAINE\pnom) fait échouer l'ordre sur la barre oblique inverse.
    ''' </summary>
    Private Shared Function Crochets(compte As String) As String
        Return "[" & If(compte, String.Empty).Replace("]", "]]") & "]"
    End Function

#End Region

End Class
