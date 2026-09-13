Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Accès aux comptes utilisateurs (T_UtilisateurWU) et au journal des connexions
''' (T_ConnexionWU).
'''
''' L'authentification est portée ici et non dans l'interface : un écran de connexion ne doit
''' rien décider, il présente une demande et affiche un verdict.
''' </summary>
Public NotInheritable Class UtilisateurRepository

    Private Sub New()
    End Sub

    Private Const TABLE_UTILISATEUR As String = "T_UtilisateurWU"
    Private Const TABLE_CONNEXION As String = "T_ConnexionWU"

    Private Const ERREUR_TABLE_ABSENTE As Integer = 208
    Private Const ERREUR_CLE_DUPLIQUEE As Integer = 2627
    Private Const ERREUR_INDEX_UNIQUE As Integer = 2601

    ''' <summary>Nombre d'échecs consécutifs au-delà duquel le compte est verrouillé.</summary>
    Public Const ECHECS_AVANT_VERROUILLAGE As Integer = 5

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_UtilisateurWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\07_Utilisateurs.sql : il crée les tables des utilisateurs " &
        "et du journal des connexions, et ajoute les colonnes de traçabilité."

#Region "Authentification"

    ''' <summary>
    ''' Résultat d'une tentative de connexion. Le message est destiné à l'utilisateur ; il ne
    ''' distingue jamais « identifiant inconnu » de « mot de passe erroné », pour ne pas
    ''' renseigner sur l'existence d'un compte.
    ''' </summary>
    Public Class ResultatConnexion
        Public Property Reussi As Boolean = False
        Public Property Utilisateur As UtilisateurWU
        Public Property Message As String = String.Empty
        ''' <summary>Vrai si le mot de passe doit être changé avant toute autre action.</summary>
        Public Property ChangementExige As Boolean = False
    End Class

    ''' <summary>
    ''' Vérifie un identifiant et un mot de passe, met à jour le compteur d'échecs et journalise
    ''' la tentative — réussie ou non.
    '''
    ''' Toute tentative est journalisée AVANT d'être rendue à l'appelant : un refus non
    ''' enregistré serait précisément celui qu'un auditeur chercherait.
    ''' </summary>
    Public Shared Function Authentifier(identifiant As String, motDePasse As String) As ResultatConnexion

        Dim resultat As New ResultatConnexion()

        If String.IsNullOrWhiteSpace(identifiant) OrElse String.IsNullOrEmpty(motDePasse) Then
            resultat.Message = "Identifiant et mot de passe sont obligatoires."
            Return resultat
        End If

        Dim messageErreur As String = String.Empty
        Dim utilisateur As UtilisateurWU = Lire(identifiant.Trim(), messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then
            resultat.Message = messageErreur
            Return resultat
        End If

        ' Compte inconnu : même message qu'un mot de passe erroné, et journalisé comme tel.
        If utilisateur Is Nothing Then
            Journaliser(identifiant.Trim(), False, "Identifiant inconnu")
            resultat.Message = "Identifiant ou mot de passe incorrect."
            Return resultat
        End If

        If Not utilisateur.Actif Then
            Journaliser(utilisateur.Identifiant, False, "Compte désactivé")
            resultat.Message = "Ce compte est désactivé. Contactez l'administrateur de l'application."
            Return resultat
        End If

        If utilisateur.EchecsConsecutifs >= ECHECS_AVANT_VERROUILLAGE Then
            Journaliser(utilisateur.Identifiant, False, "Compte verrouillé")
            resultat.Message = $"Ce compte est verrouillé après {ECHECS_AVANT_VERROUILLAGE} tentatives " &
                               "infructueuses." & Environment.NewLine &
                               "Seul l'administrateur peut le déverrouiller."
            Return resultat
        End If

        If Not MotDePasseService.Verifier(motDePasse, utilisateur.Empreinte,
                                          utilisateur.Sel, utilisateur.Iterations) Then

            Dim echecs As Integer = utilisateur.EchecsConsecutifs + 1
            EnregistrerEchec(utilisateur.Identifiant, echecs)
            Journaliser(utilisateur.Identifiant, False, "Mot de passe incorrect")

            Dim restantes As Integer = ECHECS_AVANT_VERROUILLAGE - echecs

            resultat.Message = "Identifiant ou mot de passe incorrect." &
                               If(restantes > 0 AndAlso restantes <= 2,
                                  Environment.NewLine & $"Encore {restantes} tentative(s) avant verrouillage du compte.",
                                  String.Empty)
            Return resultat
        End If

        EnregistrerConnexionReussie(utilisateur.Identifiant)
        Journaliser(utilisateur.Identifiant, True, String.Empty)

        utilisateur.EchecsConsecutifs = 0
        resultat.Reussi = True
        resultat.Utilisateur = utilisateur
        resultat.ChangementExige = utilisateur.DoitChangerMotDePasse

        Return resultat
    End Function

    ''' <summary>
    ''' Indique s'il existe au moins un administrateur actif et utilisable.
    '''
    ''' Sert au tout premier démarrage : sans administrateur, personne ne pourrait créer de
    ''' compte et l'application serait inutilisable. La longueur de l'empreinte et du sel est
    ''' contrôlée, et pas seulement leur présence : une ligne posée à la main dans SQL Server
    ''' avec une empreinte de fortune ouvrirait un compte administrateur inutilisable, dont
    ''' l'existence empêcherait pourtant la création du vrai.
    ''' </summary>
    Public Shared Function ExisteAdministrateurUtilisable(ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Const requete As String =
            "SELECT COUNT(*) FROM " & TABLE_UTILISATEUR & " " &
            "WHERE Role = 'ADMIN' AND Actif = 1 AND LEN(Empreinte) > 40 AND LEN(Sel) > 10"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    Dim valeur As Object = commande.ExecuteScalar()
                    If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return False
                    Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) > 0
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des utilisateurs impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try
    End Function

#End Region

#Region "Lecture"

    ''' <summary>Lit un utilisateur par son identifiant. Nothing s'il n'existe pas.</summary>
    Public Shared Function Lire(identifiant As String, ByRef messageErreur As String) As UtilisateurWU

        messageErreur = String.Empty

        Const requete As String =
            "SELECT Identifiant, NomComplet, Role, Fonction, Empreinte, Sel, Iterations, Actif, " &
            "DoitChangerMotDePasse, EchecsConsecutifs, DateVerrouillage, DerniereConnexion, " &
            "DateCreation, CreePar, DateModification, ModifiePar " &
            "FROM " & TABLE_UTILISATEUR & " WHERE Identifiant = @identifiant"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        If lecteur.Read() Then Return Construire(lecteur)
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de l'utilisateur impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return Nothing
    End Function

    ''' <summary>Liste tous les utilisateurs, triés par rôle puis par identifiant.</summary>
    Public Shared Function Lister(ByRef messageErreur As String) As List(Of UtilisateurWU)

        messageErreur = String.Empty
        Dim resultat As New List(Of UtilisateurWU)

        Const requete As String =
            "SELECT Identifiant, NomComplet, Role, Fonction, Empreinte, Sel, Iterations, Actif, " &
            "DoitChangerMotDePasse, EchecsConsecutifs, DateVerrouillage, DerniereConnexion, " &
            "DateCreation, CreePar, DateModification, ModifiePar " &
            "FROM " & TABLE_UTILISATEUR & " ORDER BY Role, Identifiant"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(Construire(lecteur))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des utilisateurs impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    Private Shared Function Construire(lecteur As SqlDataReader) As UtilisateurWU

        Return New UtilisateurWU() With {
            .Identifiant = LireChaine(lecteur, "Identifiant"),
            .NomComplet = LireChaine(lecteur, "NomComplet"),
            .Role = UtilisateurWU.RoleDepuisLibelle(LireChaine(lecteur, "Role")),
            .Fonction = UtilisateurWU.FonctionDepuisLibelle(LireChaine(lecteur, "Fonction")),
            .Empreinte = LireChaine(lecteur, "Empreinte"),
            .Sel = LireChaine(lecteur, "Sel"),
            .Iterations = LireEntier(lecteur, "Iterations"),
            .Actif = LireBooleen(lecteur, "Actif"),
            .DoitChangerMotDePasse = LireBooleen(lecteur, "DoitChangerMotDePasse"),
            .EchecsConsecutifs = LireEntier(lecteur, "EchecsConsecutifs"),
            .DateVerrouillage = LireDate(lecteur, "DateVerrouillage"),
            .DerniereConnexion = LireDate(lecteur, "DerniereConnexion"),
            .DateCreation = LireDate(lecteur, "DateCreation"),
            .CreePar = LireChaine(lecteur, "CreePar"),
            .DateModification = LireDate(lecteur, "DateModification"),
            .ModifiePar = LireChaine(lecteur, "ModifiePar")
        }
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Crée un utilisateur avec un mot de passe initial, qu'il devra changer à sa première
    ''' connexion : celui qui crée le compte n'a pas à connaître durablement le mot de passe
    ''' de quelqu'un d'autre.
    ''' </summary>
    Public Shared Function Ajouter(utilisateur As UtilisateurWU, motDePasseInitial As String,
                                   ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If utilisateur Is Nothing Then
            messageErreur = "Aucun utilisateur à enregistrer."
            Return False
        End If

        Dim sel As String = String.Empty
        Dim iterations As Integer = 0
        Dim empreinte As String = MotDePasseService.Hacher(motDePasseInitial, sel, iterations)

        Const requete As String =
            "INSERT INTO " & TABLE_UTILISATEUR & " (Identifiant, NomComplet, Role, Fonction, " &
            "Empreinte, Sel, Iterations, Actif, DoitChangerMotDePasse, EchecsConsecutifs, " &
            "DateCreation, CreePar) " &
            "VALUES (@identifiant, @nom, @role, @fonction, @empreinte, @sel, @iterations, " &
            "@actif, 1, 0, GETDATE(), @creePar)"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = utilisateur.Identifiant
                    commande.Parameters.Add("@nom", SqlDbType.NVarChar, 150).Value = utilisateur.NomComplet
                    commande.Parameters.Add("@role", SqlDbType.NVarChar, 20).Value = UtilisateurWU.LibelleDepuisRole(utilisateur.Role)
                    commande.Parameters.Add("@fonction", SqlDbType.NVarChar, 20).Value = ValeurFonction(utilisateur)
                    commande.Parameters.Add("@empreinte", SqlDbType.NVarChar, 256).Value = empreinte
                    commande.Parameters.Add("@sel", SqlDbType.NVarChar, 128).Value = sel
                    commande.Parameters.Add("@iterations", SqlDbType.Int).Value = iterations
                    commande.Parameters.Add("@actif", SqlDbType.Bit).Value = utilisateur.Actif
                    commande.Parameters.Add("@creePar", SqlDbType.NVarChar, 50).Value = ParOuScript()

                    commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_CLE_DUPLIQUEE OrElse ex.Number = ERREUR_INDEX_UNIQUE,
                               $"L'identifiant « {utilisateur.Identifiant} » est déjà utilisé.",
                               If(ex.Number = ERREUR_TABLE_ABSENTE, MESSAGE_TABLE_ABSENTE,
                                  $"Création de l'utilisateur impossible : {ex.Message}"))
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Modifie le nom, le rôle et l'état d'activation. Le mot de passe n'est PAS touché ici :
    ''' le réinitialiser est une opération distincte, qui doit se voir.
    ''' </summary>
    Public Shared Function Modifier(utilisateur As UtilisateurWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If utilisateur Is Nothing Then
            messageErreur = "Aucun utilisateur à enregistrer."
            Return False
        End If

        Dim perdSesDroits As Boolean = utilisateur.Role <> RoleWU.Administrateur OrElse Not utilisateur.Actif

        If perdSesDroits AndAlso AutresAdministrateursUtilisables(utilisateur.Identifiant) = 0 Then
            messageErreur = "Cette modification laisserait la base sans aucun administrateur actif." & Environment.NewLine &
                            "Créez d’abord un autre administrateur."
            Return False
        End If

        Const requete As String =
            "UPDATE " & TABLE_UTILISATEUR & " SET NomComplet = @nom, Role = @role, " &
            "Fonction = @fonction, Actif = @actif, DateModification = GETDATE(), " &
            "ModifiePar = @modifiePar WHERE Identifiant = @identifiant"

        Return Executer(requete, messageErreur,
                        Sub(commande)
                            commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = utilisateur.Identifiant
                            commande.Parameters.Add("@nom", SqlDbType.NVarChar, 150).Value = utilisateur.NomComplet
                            commande.Parameters.Add("@role", SqlDbType.NVarChar, 20).Value = UtilisateurWU.LibelleDepuisRole(utilisateur.Role)
                            commande.Parameters.Add("@fonction", SqlDbType.NVarChar, 20).Value = ValeurFonction(utilisateur)
                            commande.Parameters.Add("@actif", SqlDbType.Bit).Value = utilisateur.Actif
                            commande.Parameters.Add("@modifiePar", SqlDbType.NVarChar, 50).Value = ParOuScript()
                        End Sub)
    End Function

    ''' <summary>
    ''' Fixe un nouveau mot de passe. Le verrouillage éventuel est levé au passage : un compte
    ''' dont le mot de passe vient d'être changé n'a plus de raison d'être bloqué.
    ''' </summary>
    ''' <param name="changementExige">
    ''' Vrai pour une réinitialisation par l'administrateur — l'utilisateur devra en choisir un
    ''' autre ; faux lorsque l'utilisateur choisit lui-même son mot de passe.
    ''' </param>
    Public Shared Function ChangerMotDePasse(identifiant As String, nouveauMotDePasse As String,
                                             changementExige As Boolean,
                                             ByRef messageErreur As String) As Boolean

        Dim sel As String = String.Empty
        Dim iterations As Integer = 0
        Dim empreinte As String = MotDePasseService.Hacher(nouveauMotDePasse, sel, iterations)

        Const requete As String =
            "UPDATE " & TABLE_UTILISATEUR & " SET Empreinte = @empreinte, Sel = @sel, " &
            "Iterations = @iterations, DoitChangerMotDePasse = @exige, EchecsConsecutifs = 0, " &
            "DateVerrouillage = NULL, DateModification = GETDATE(), ModifiePar = @modifiePar " &
            "WHERE Identifiant = @identifiant"

        Return Executer(requete, messageErreur,
                        Sub(commande)
                            commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant
                            commande.Parameters.Add("@empreinte", SqlDbType.NVarChar, 256).Value = empreinte
                            commande.Parameters.Add("@sel", SqlDbType.NVarChar, 128).Value = sel
                            commande.Parameters.Add("@iterations", SqlDbType.Int).Value = iterations
                            commande.Parameters.Add("@exige", SqlDbType.Bit).Value = changementExige
                            commande.Parameters.Add("@modifiePar", SqlDbType.NVarChar, 50).Value = ParOuScript()
                        End Sub)
    End Function

    ''' <summary>
    ''' Nombre d'administrateurs actifs et utilisables AUTRES que celui indiqué.
    '''
    ''' Sert à empêcher qu'on retire son rôle ou son activation au dernier administrateur :
    ''' la base se retrouverait sans personne pour créer ou débloquer un compte, et il faudrait
    ''' rouvrir SQL Server à la main pour s'en sortir.
    ''' </summary>
    Private Shared Function AutresAdministrateursUtilisables(identifiant As String) As Integer

        Const requete As String =
            "SELECT COUNT(*) FROM " & TABLE_UTILISATEUR & " " &
            "WHERE Role = 'ADMIN' AND Actif = 1 AND LEN(Empreinte) > 40 AND LEN(Sel) > 10 " &
            "AND Identifiant <> @identifiant"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant

                    Dim valeur As Object = commande.ExecuteScalar()
                    If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return 0
                    Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture)
                End Using
            End Using

        Catch ex As SqlException
            ' Comptage impossible : on répond 0, donc la modification sera refusée. Mieux vaut
            ' refuser une modification légitime que laisser la base sans administrateur.
            Return 0
        Catch ex As InvalidOperationException
            Return 0
        End Try
    End Function

    ''' <summary>Déverrouille un compte bloqué par des échecs répétés.</summary>
    Public Shared Function Deverrouiller(identifiant As String, ByRef messageErreur As String) As Boolean

        Const requete As String =
            "UPDATE " & TABLE_UTILISATEUR & " SET EchecsConsecutifs = 0, DateVerrouillage = NULL, " &
            "DateModification = GETDATE(), ModifiePar = @modifiePar WHERE Identifiant = @identifiant"

        Return Executer(requete, messageErreur,
                        Sub(commande)
                            commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant
                            commande.Parameters.Add("@modifiePar", SqlDbType.NVarChar, 50).Value = ParOuScript()
                        End Sub)
    End Function

    Private Shared Sub EnregistrerEchec(identifiant As String, echecs As Integer)

        Const requete As String =
            "UPDATE " & TABLE_UTILISATEUR & " SET EchecsConsecutifs = @echecs, " &
            "DateVerrouillage = CASE WHEN @echecs >= @seuil THEN GETDATE() ELSE DateVerrouillage END " &
            "WHERE Identifiant = @identifiant"

        Dim messageIgnore As String = String.Empty

        Executer(requete, messageIgnore,
                 Sub(commande)
                     commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant
                     commande.Parameters.Add("@echecs", SqlDbType.Int).Value = echecs
                     commande.Parameters.Add("@seuil", SqlDbType.Int).Value = ECHECS_AVANT_VERROUILLAGE
                 End Sub)
    End Sub

    Private Shared Sub EnregistrerConnexionReussie(identifiant As String)

        Const requete As String =
            "UPDATE " & TABLE_UTILISATEUR & " SET DerniereConnexion = GETDATE(), " &
            "EchecsConsecutifs = 0, DateVerrouillage = NULL WHERE Identifiant = @identifiant"

        Dim messageIgnore As String = String.Empty

        Executer(requete, messageIgnore,
                 Sub(commande)
                     commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = identifiant
                 End Sub)
    End Sub

#End Region

#Region "Journal des connexions"

    ''' <summary>
    ''' Enregistre une tentative de connexion, réussie ou non.
    '''
    ''' Un échec de journalisation n'empêche jamais la connexion : le journal sert à l'audit,
    ''' pas au contrôle d'accès, et priver un agent de son outil parce qu'une écriture de trace
    ''' a échoué serait disproportionné.
    ''' </summary>
    Public Shared Sub Journaliser(identifiant As String, succes As Boolean, motif As String)

        Const requete As String =
            "INSERT INTO " & TABLE_CONNEXION & " (Identifiant, Poste, CompteWindows, Succes, Motif) " &
            "VALUES (@identifiant, @poste, @windows, @succes, @motif)"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@identifiant", SqlDbType.NVarChar, 50).Value = If(identifiant, String.Empty)
                    commande.Parameters.Add("@poste", SqlDbType.NVarChar, 100).Value = SessionWU.Poste
                    commande.Parameters.Add("@windows", SqlDbType.NVarChar, 100).Value = SessionWU.CompteWindows
                    commande.Parameters.Add("@succes", SqlDbType.Bit).Value = succes
                    commande.Parameters.Add("@motif", SqlDbType.NVarChar, 200).Value = If(motif, String.Empty)

                    commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            ' Journal indisponible : sans conséquence sur la connexion elle-même.
        Catch ex As InvalidOperationException
        End Try
    End Sub

    ''' <summary>Dernières tentatives de connexion, les plus récentes d'abord.</summary>
    Public Shared Function ListerConnexions(nombreMaximum As Integer, ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty
        Dim table As New DataTable("Connexions")

        Dim requete As String =
            $"SELECT TOP {Math.Max(nombreMaximum, 1)} DateConnexion, Identifiant, Poste, " &
            "CompteWindows, Succes, Motif FROM " & TABLE_CONNEXION & " ORDER BY DateConnexion DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using adaptateur As New SqlDataAdapter(requete, connexion)
                    adaptateur.Fill(table)
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture du journal impossible : {ex.Message}")
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return table
    End Function

#End Region

#Region "Utilitaires internes"

    ''' <summary>
    ''' Auteur de l'écriture : l'utilisateur connecté, ou « script » lorsque personne ne l'est —
    ''' cas de la création du tout premier administrateur.
    ''' </summary>
    Private Shared Function ParOuScript() As String
        Dim identifiant As String = SessionWU.Identifiant
        Return If(String.IsNullOrEmpty(identifiant), "installation", identifiant)
    End Function

    ''' <summary>
    ''' Valeur à écrire dans la colonne Fonction. « Aucune » s'enregistre en NULL et non en
    ''' chaîne vide : la contrainte CHECK de la base n'accepte que INPUTER, AUTHORIZER ou NULL.
    ''' </summary>
    Private Shared Function ValeurFonction(utilisateur As UtilisateurWU) As Object

        Dim libelle As String = UtilisateurWU.LibelleDepuisFonction(utilisateur.Fonction)
        If libelle.Length = 0 Then Return DBNull.Value

        Return libelle
    End Function

    ''' <summary>Exécute une commande d'écriture et traduit les erreurs courantes.</summary>
    Private Shared Function Executer(requete As String, ByRef messageErreur As String,
                                     preparer As Action(Of SqlCommand)) As Boolean

        messageErreur = String.Empty

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    preparer(commande)

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = "Aucune modification : le compte n'existe plus."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Écriture impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Function LireChaine(lecteur As SqlDataReader, colonne As String) As String
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return String.Empty
        Return lecteur.GetValue(index).ToString().Trim()
    End Function

    Private Shared Function LireEntier(lecteur As SqlDataReader, colonne As String) As Integer
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0
        Return Convert.ToInt32(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireBooleen(lecteur As SqlDataReader, colonne As String) As Boolean
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return False
        Return Convert.ToBoolean(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireDate(lecteur As SqlDataReader, colonne As String) As Date?
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return Nothing
        Return Convert.ToDateTime(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

#End Region

End Class
