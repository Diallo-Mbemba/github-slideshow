Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Changement de mot de passe, dans deux situations :
'''
''' — à la première connexion, ou après une réinitialisation par l'administrateur : le mot de
'''   passe actuel vient d'être saisi à l'écran de connexion et validé, il n'est donc pas
'''   redemandé ;
''' — à la demande de l'utilisateur, depuis le menu : le mot de passe actuel est alors exigé,
'''   faute de quoi un poste laissé ouvert permettrait à n'importe qui de verrouiller le compte
'''   de son titulaire.
''' </summary>
Public Class FrmChangerMotDePasse

    Private ReadOnly _identifiant As String

    ''' <summary>
    ''' Mot de passe actuel déjà vérifié, ou Nothing lorsqu'il reste à saisir et à contrôler.
    ''' </summary>
    Private ReadOnly _ancienConnu As String

    ''' <summary>Décalage vertical appliqué lorsque la ligne « mot de passe actuel » est retirée.</summary>
    Private Const DECALAGE_SANS_ANCIEN As Integer = 36

    ''' <summary>Changement demandé par l'utilisateur : le mot de passe actuel sera exigé.</summary>
    Public Sub New(identifiant As String)
        InitializeComponent()
        IconesWU.Habiller(Me)
        _identifiant = If(identifiant, String.Empty).Trim()
        _ancienConnu = Nothing
    End Sub

    ''' <summary>
    ''' Changement imposé à la connexion. Le mot de passe actuel vient d'être validé par
    ''' l'authentification ; le redemander n'apporterait aucune garantie supplémentaire.
    ''' </summary>
    Public Sub New(identifiant As String, ancienMotDePasseValide As String)
        InitializeComponent()
        IconesWU.Habiller(Me)
        _identifiant = If(identifiant, String.Empty).Trim()
        _ancienConnu = If(ancienMotDePasseValide, String.Empty)
    End Sub

#Region "Ouverture"

    Private Sub FrmChangerMotDePasse_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblUtilisateur.Text = $"Compte « {_identifiant} »"

        lblRegles.Text =
            $"Le mot de passe comporte au moins {MotDePasseService.LONGUEUR_MINIMALE} caractères, " &
            "dont une majuscule, une minuscule et un chiffre. Il ne doit pas contenir l'identifiant."

        If _ancienConnu Is Nothing Then
            txtAncien.Focus()
            Return
        End If

        MasquerAncienMotDePasse()
        lblUtilisateur.Text &= " — un nouveau mot de passe doit être choisi avant d'accéder à l'application."
        txtNouveau.Focus()
    End Sub

    ''' <summary>
    ''' Retire la ligne « mot de passe actuel » et remonte ce qui la suit, pour ne pas laisser
    ''' un vide inexpliqué au milieu du formulaire.
    ''' </summary>
    Private Sub MasquerAncienMotDePasse()

        lblAncien.Visible = False
        txtAncien.Visible = False

        For Each controle As Control In New Control() {lblNouveau, txtNouveau, lblConfirmation,
                                                       txtConfirmation, lblRegles, lblMessage,
                                                       btnValider, btnAnnuler}
            controle.Top -= DECALAGE_SANS_ANCIEN
        Next

        Height -= DECALAGE_SANS_ANCIEN
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnValider_Click(sender As Object, e As EventArgs) Handles btnValider.Click

        lblMessage.Text = String.Empty

        Dim ancien As String = If(_ancienConnu, txtAncien.Text)

        If _ancienConnu Is Nothing AndAlso Not AncienMotDePasseValide(ancien) Then Return

        Dim anomalies As List(Of String) = MotDePasseService.Anomalies(txtNouveau.Text, _identifiant)

        If txtNouveau.Text <> txtConfirmation.Text Then
            anomalies.Add("Le nouveau mot de passe et sa confirmation sont différents.")
        End If

        If txtNouveau.Text = ancien Then
            anomalies.Add("Le nouveau mot de passe doit être différent de l'ancien.")
        End If

        If anomalies.Count > 0 Then
            lblMessage.Text = String.Join(Environment.NewLine, anomalies)
            Return
        End If

        Dim messageErreur As String = String.Empty

        ' changementExige à faux : l'utilisateur vient de choisir lui-même son mot de passe, il
        ' n'y a plus lieu de le lui redemander à la prochaine connexion.
        If Not UtilisateurRepository.ChangerMotDePasse(_identifiant, txtNouveau.Text, False, messageErreur) Then
            lblMessage.Text = messageErreur
            Return
        End If

        UtilisateurRepository.Journaliser(_identifiant, True, "Changement de mot de passe")

        MessageBox.Show("Le mot de passe a été modifié.", "Mot de passe",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)

        DialogResult = DialogResult.OK
        Close()
    End Sub

    ''' <summary>
    ''' Vérifie le mot de passe actuel saisi par l'utilisateur. La vérification se fait sur
    ''' l'empreinte enregistrée, sans passer par l'authentification complète : un essai
    ''' malheureux ici ne doit pas verrouiller le compte de quelqu'un déjà connecté.
    ''' </summary>
    Private Function AncienMotDePasseValide(ancien As String) As Boolean

        Dim messageErreur As String = String.Empty
        Dim utilisateur As UtilisateurWU = UtilisateurRepository.Lire(_identifiant, messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then
            lblMessage.Text = messageErreur
            Return False
        End If

        If utilisateur Is Nothing Then
            lblMessage.Text = $"Le compte « {_identifiant} » n'existe plus."
            Return False
        End If

        If Not MotDePasseService.Verifier(ancien, utilisateur.Empreinte,
                                          utilisateur.Sel, utilisateur.Iterations) Then
            lblMessage.Text = "Le mot de passe actuel est incorrect."
            txtAncien.SelectAll()
            txtAncien.Focus()
            Return False
        End If

        Return True
    End Function

    Private Sub btnAnnuler_Click(sender As Object, e As EventArgs) Handles btnAnnuler.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

#End Region

End Class
