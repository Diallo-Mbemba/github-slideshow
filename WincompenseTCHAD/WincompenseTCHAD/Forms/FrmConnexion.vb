Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Écran de connexion : première fenêtre affichée, avant toute autre.
'''
''' L'application retient un mot de passe applicatif propre à chaque utilisateur, et non la
''' session Windows du poste : les postes sont nominatifs, mais un agent qui laisse sa session
''' ouverte ne doit pas pour autant laisser l'accès à la compense. Le compte Windows est
''' néanmoins relevé et journalisé, ce qui permet de repérer un identifiant utilisé depuis un
''' poste qui n'est pas le sien.
'''
''' Cet écran ne décide rien : il transmet une demande à UtilisateurRepository et affiche le
''' verdict. Il ne sait pas distinguer un identifiant inconnu d'un mot de passe erroné, et
''' c'est voulu.
''' </summary>
Public Class FrmConnexion

    ''' <summary>Utilisateur authentifié, ou Nothing si la connexion a été abandonnée.</summary>
    Public ReadOnly Property UtilisateurConnecte As UtilisateurWU
        Get
            Return _utilisateur
        End Get
    End Property
    Private _utilisateur As UtilisateurWU

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture"

    Private Sub FrmConnexion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AfficherBase()
        VerifierPremierDemarrage()
        txtIdentifiant.Focus()
    End Sub

    ''' <summary>
    ''' Rappelle la base visée. Une erreur d'environnement — base de test prise pour la
    ''' production — se voit dès l'écran de connexion, avant la moindre saisie.
    ''' </summary>
    Private Sub AfficherBase()

        Try
            Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder(
                WURepository.ObtenirChaineConnexion())

            lblBase.Text = $"Serveur {constructeur.DataSource} — base {constructeur.InitialCatalog}"

        Catch ex As ArgumentException
            lblBase.Text = "Chaîne de connexion non exploitable."
        End Try
    End Sub

    ''' <summary>
    ''' Demande s'il faut créer le premier administrateur — en nommant la base visée.
    '''
    ''' Un compte administrateur créé dans la mauvaise base est une faute silencieuse : tout
    ''' semble avoir fonctionné, et personne ne revient jamais voir cette base. Le serveur, la
    ''' base et la provenance de la chaîne sont donc affichés avant le geste, et une troisième
    ''' réponse permet d'aller corriger la connexion plutôt que de renoncer.
    ''' </summary>
    ''' <returns>Yes pour créer, No pour régler le serveur d'abord, Cancel pour renoncer.</returns>
    Private Function DemanderLaCreationDuPremierAdministrateur() As DialogResult

        Dim constructeur As System.Data.SqlClient.SqlConnectionStringBuilder = Nothing

        Try
            constructeur = New System.Data.SqlClient.SqlConnectionStringBuilder(
                WURepository.ObtenirChaineConnexion())
        Catch
            ' Chaîne illisible : l'écran affiche alors ce qu'il peut, sans échouer ici.
        End Try

        Dim serveur As String = If(constructeur Is Nothing, "(inconnu)", constructeur.DataSource)
        Dim base As String = If(constructeur Is Nothing, "(inconnue)", constructeur.InitialCatalog)

        Return MessageBox.Show(
            "Aucun administrateur n'est encore défini dans cette base." & Environment.NewLine & Environment.NewLine &
            "Le compte va être créé ici :" & Environment.NewLine &
            $"        serveur : {serveur}" & Environment.NewLine &
            $"        base    : {base}" & Environment.NewLine &
            $"        origine : {ConfigurationWU.Origine()}" & Environment.NewLine & Environment.NewLine &
            "Est-ce bien la base de production ?" & Environment.NewLine & Environment.NewLine &
            "    Oui     — créer le premier administrateur ici" & Environment.NewLine &
            "    Non     — régler d'abord le serveur" & Environment.NewLine &
            "    Annuler — ne rien faire",
            "Premier démarrage", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button3)
    End Function

    ''' <summary>
    ''' Fait apparaître le bouton de réglage du serveur, et invite à s'en servir.
    '''
    ''' Il reste masqué tant que la base répond : régler le serveur n'est pas un geste
    ''' quotidien, et un bouton toujours visible inviterait à y toucher sans raison.
    '''
    ''' Il est indispensable ici. L'écran de réglage vit normalement dans le menu Sécurité,
    ''' donc DERRIÈRE la connexion — or c'est précisément quand la connexion échoue qu'il faut
    ''' l'atteindre. Sans ce bouton, corriger un serveur imposerait de passer par le fichier de
    ''' configuration à la main, sur chaque poste concerné.
    ''' </summary>
    Private Sub OffrirDeReglerLeServeur()

        btnParametres.Visible = True

        lblMessage.Text &= Environment.NewLine & Environment.NewLine &
                           "Le serveur est peut-être mal indiqué : bouton « Serveur... » ci-dessous."
    End Sub

    ''' <summary>
    ''' Ouvre le réglage du serveur, puis retente la lecture. Si la base répond enfin, le
    ''' message d'erreur et le bouton disparaissent : l'utilisateur voit que c'est réglé sans
    ''' avoir à relancer l'application.
    ''' </summary>
    Private Sub btnParametres_Click(sender As Object, e As EventArgs) Handles btnParametres.Click

        ReglerLeServeur()
        VerifierPremierDemarrage()
    End Sub

    ''' <summary>
    ''' Au tout premier démarrage, la base ne contient aucun administrateur utilisable : sans
    ''' lui, personne ne pourrait créer de compte et l'application resterait inaccessible. On
    ''' propose alors de créer ce premier administrateur.
    '''
    ''' La proposition n'apparaît que si la lecture a réellement abouti : une table absente ou
    ''' une base injoignable ne doit pas être confondue avec « aucun administrateur », sous
    ''' peine de proposer une création vouée à l'échec.
    '''
    ''' La boucle permet de corriger le serveur autant de fois qu'il le faut sans empiler les
    ''' appels : un technicien qui se trompe deux fois de serveur ne doit pas creuser la pile
    ''' d'exécution à chaque essai.
    ''' </summary>
    Private Sub VerifierPremierDemarrage()

        Do
            Dim messageErreur As String = String.Empty

            If UtilisateurRepository.ExisteAdministrateurUtilisable(messageErreur) Then Return

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblMessage.Text = messageErreur
                OffrirDeReglerLeServeur()
                Return
            End If

            Select Case DemanderLaCreationDuPremierAdministrateur()

                Case DialogResult.Yes
                    CreerLePremierAdministrateur()
                    Return

                Case DialogResult.No
                    ' Ce n'est pas la bonne base : on règle le serveur, puis on recommence.
                    ' Créer un administrateur ailleurs que dans la base de production donnerait
                    ' un compte que personne ne retrouverait jamais.
                    If Not ReglerLeServeur() Then Return

                Case Else
                    lblMessage.Text = "Aucun administrateur défini : la connexion est impossible tant qu'un compte n'a pas été créé."
                    Return
            End Select
        Loop
    End Sub

    ''' <summary>
    ''' Ouvre l'écran de réglage du serveur. Retourne False si la chaîne en service n'a pas
    ''' changé — inutile alors de reposer la même question sur la même base.
    ''' </summary>
    Private Function ReglerLeServeur() As Boolean

        Dim avant As String = WURepository.ObtenirChaineConnexion()

        Using parametres As New FrmParametresConnexion()
            parametres.ShowDialog(Me)
        End Using

        lblMessage.Text = String.Empty
        btnParametres.Visible = False
        AfficherBase()

        If String.Equals(avant, WURepository.ObtenirChaineConnexion(), StringComparison.OrdinalIgnoreCase) Then
            lblMessage.Text = "La connexion n'a pas été modifiée."
            Return False
        End If

        Return True
    End Function

    ''' <summary>Ouvre la création du premier administrateur et prépare la saisie qui suit.</summary>
    Private Sub CreerLePremierAdministrateur()

        Using edition As New FrmUtilisateurEdition(Nothing, True)

            If edition.ShowDialog(Me) <> DialogResult.OK OrElse edition.UtilisateurEnregistre Is Nothing Then
                lblMessage.Text = "Création abandonnée : aucun compte n'a été enregistré."
                Return
            End If

            ' L'identifiant est prérempli, mais pas le mot de passe : celui qui vient d'être
            ' choisi doit être ressaisi, ce qui vérifie au passage qu'il a bien été mémorisé.
            txtIdentifiant.Text = edition.UtilisateurEnregistre.Identifiant
            lblMessage.Text = String.Empty
        End Using
    End Sub

#End Region

#Region "Connexion"

    Private Sub btnConnexion_Click(sender As Object, e As EventArgs) Handles btnConnexion.Click

        lblMessage.Text = String.Empty
        Cursor = Cursors.WaitCursor
        btnConnexion.Enabled = False

        Try
            Dim resultat As UtilisateurRepository.ResultatConnexion =
                UtilisateurRepository.Authentifier(txtIdentifiant.Text, txtMotDePasse.Text)

            If Not resultat.Reussi Then
                lblMessage.Text = resultat.Message
                txtMotDePasse.Clear()
                txtMotDePasse.Focus()
                Return
            End If

            ' Mot de passe imposé par l'administrateur : il doit être remplacé AVANT d'ouvrir
            ' la session, sans quoi le renvoyer à plus tard reviendrait à ne jamais le changer.
            If resultat.ChangementExige AndAlso Not ChangementAccepte(resultat.Utilisateur.Identifiant) Then
                lblMessage.Text = "Le mot de passe doit être changé pour accéder à l'application."
                txtMotDePasse.Clear()
                txtMotDePasse.Focus()
                Return
            End If

            _utilisateur = resultat.Utilisateur
            DialogResult = DialogResult.OK
            Close()

        Finally
            btnConnexion.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Ouvre le changement de mot de passe imposé. Retourne vrai si l'utilisateur a bien
    ''' enregistré un nouveau mot de passe.
    ''' </summary>
    Private Function ChangementAccepte(identifiant As String) As Boolean

        Using changement As New FrmChangerMotDePasse(identifiant, txtMotDePasse.Text)
            Return changement.ShowDialog(Me) = DialogResult.OK
        End Using
    End Function

    Private Sub btnAnnuler_Click(sender As Object, e As EventArgs) Handles btnAnnuler.Click
        _utilisateur = Nothing
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

#End Region

End Class
