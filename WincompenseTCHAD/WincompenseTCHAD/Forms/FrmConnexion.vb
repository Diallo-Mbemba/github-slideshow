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
    ''' Au tout premier démarrage, la base ne contient aucun administrateur utilisable : sans
    ''' lui, personne ne pourrait créer de compte et l'application resterait inaccessible. On
    ''' propose alors de créer ce premier administrateur.
    '''
    ''' La proposition n'apparaît que si la lecture a réellement abouti : une table absente ou
    ''' une base injoignable ne doit pas être confondue avec « aucun administrateur », sous
    ''' peine de proposer une création vouée à l'échec.
    ''' </summary>
    Private Sub VerifierPremierDemarrage()

        Dim messageErreur As String = String.Empty

        If UtilisateurRepository.ExisteAdministrateurUtilisable(messageErreur) Then Return

        If Not String.IsNullOrEmpty(messageErreur) Then
            lblMessage.Text = messageErreur
            Return
        End If

        Dim reponse As DialogResult = MessageBox.Show(
            "Aucun administrateur n'est encore défini dans la base." & Environment.NewLine & Environment.NewLine &
            "Souhaitez-vous créer maintenant le premier compte administrateur ?" & Environment.NewLine &
            "Sans lui, aucun utilisateur ne pourra être créé et l'application restera inaccessible.",
            "Premier démarrage", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then
            lblMessage.Text = "Aucun administrateur défini : la connexion est impossible tant qu'un compte n'a pas été créé."
            Return
        End If

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
