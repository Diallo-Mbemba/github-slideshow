Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Création ou modification d'un compte utilisateur.
'''
''' Le même écran sert au tout premier administrateur, créé avant toute connexion : le rôle est
''' alors imposé et le compte forcément actif, pour qu'un premier démarrage ne puisse pas
''' produire un compte sans droits — l'application serait alors définitivement inaccessible.
'''
''' En modification, le mot de passe n'est pas touché : le réinitialiser est une opération
''' distincte, déclenchée explicitement depuis la liste des utilisateurs, afin qu'elle laisse
''' une trace claire au lieu de se glisser dans une simple correction de nom.
''' </summary>
Public Class FrmUtilisateurEdition

    ''' <summary>Compte modifié, ou Nothing en création.</summary>
    Private ReadOnly _existant As UtilisateurWU

    ''' <summary>Vrai pour la création du premier administrateur, avant toute connexion.</summary>
    Private ReadOnly _premierAdministrateur As Boolean

    ''' <summary>Compte effectivement enregistré, ou Nothing si l'écran a été abandonné.</summary>
    Public ReadOnly Property UtilisateurEnregistre As UtilisateurWU
        Get
            Return _enregistre
        End Get
    End Property
    Private _enregistre As UtilisateurWU

    ''' <param name="utilisateurAModifier">Compte à modifier, ou Nothing pour une création.</param>
    ''' <param name="premierAdministrateur">Vrai pour l'amorçage du premier compte administrateur.</param>
    Public Sub New(utilisateurAModifier As UtilisateurWU, premierAdministrateur As Boolean)
        InitializeComponent()
        IconesWU.Habiller(Me)
        _existant = utilisateurAModifier
        _premierAdministrateur = premierAdministrateur
    End Sub

#Region "Ouverture"

    Private Sub FrmUtilisateurEdition_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cboRole.Items.AddRange(UtilisateurWU.RolesProposes)
        cboFonction.Items.AddRange(UtilisateurWU.FonctionsProposees)
        cboFonction.SelectedIndex = 0

        lblRegles.Text =
            $"Le mot de passe initial comporte au moins {MotDePasseService.LONGUEUR_MINIMALE} caractères, " &
            "dont une majuscule, une minuscule et un chiffre." & Environment.NewLine &
            "L'utilisateur devra le remplacer par le sien à sa première connexion."

        If _existant Is Nothing Then
            PreparerCreation()
        Else
            PreparerModification()
        End If

        AfficherDroits()
        AccorderLaFonctionAuRole()
    End Sub

    Private Sub PreparerCreation()

        If _premierAdministrateur Then
            lblTitre.Text = "Premier administrateur"
            Text = "Création du premier administrateur"

            ' Rôle imposé : un premier compte sans droit d'administration laisserait la base
            ' sans personne pour créer les suivants.
            cboRole.SelectedItem = UtilisateurWU.LibelleLisibleDepuisRole(RoleWU.Administrateur)
            cboRole.Enabled = False

            chkActif.Checked = True
            chkActif.Enabled = False
        Else
            lblTitre.Text = "Nouvel utilisateur"
            Text = "Nouvel utilisateur"
            cboRole.SelectedIndex = 0
        End If

        txtIdentifiant.Focus()
    End Sub

    Private Sub PreparerModification()

        lblTitre.Text = "Modification du compte"
        Text = $"Compte « {_existant.Identifiant} »"

        txtIdentifiant.Text = _existant.Identifiant
        txtIdentifiant.ReadOnly = True
        txtIdentifiant.TabStop = False

        txtNomComplet.Text = _existant.NomComplet
        cboRole.SelectedItem = _existant.LibelleRole
        cboFonction.SelectedItem = _existant.LibelleFonction
        chkActif.Checked = _existant.Actif

        ' Le mot de passe ne se modifie pas ici : les champs disparaissent plutôt que d'être
        ' grisés, pour ne pas laisser croire qu'ils pourraient être remplis.
        lblMotDePasse.Visible = False
        txtMotDePasse.Visible = False
        lblConfirmation.Visible = False
        txtConfirmation.Visible = False
        lblRegles.Text = "Le mot de passe se réinitialise depuis la liste des utilisateurs."

        txtNomComplet.Focus()
    End Sub

    ''' <summary>
    ''' Rappelle ce que le rôle choisi autorise. L'écran de création est le seul endroit où ce
    ''' choix se fait ; il doit donc dire ce qu'il engage, plutôt que de renvoyer à une note.
    ''' </summary>
    Private Sub AfficherDroits()

        Select Case UtilisateurWU.RoleDepuisLibelleLisible(Convert.ToString(cboRole.SelectedItem))

            Case RoleWU.Compense
                lblDroits.Text = "Charge les rapports, calcule la compensation, génère et historise " &
                                 "les pièces comptables. Consulte les rapports d'activité. " &
                                 "Aucun accès au paramétrage."

            Case RoleWU.Commercial
                lblDroits.Text = "Consulte les sous-agents, les agences propres et les groupes " &
                                 "statistiques, et les rapports d'activité. Ce qu'il peut y faire " &
                                 "dépend de la fonction ci-dessous. Aucun accès à la compense."

            Case RoleWU.Administrateur
                lblDroits.Text = "Tous les droits des deux autres rôles, plus les comptes systèmes " &
                                 "de la pièce comptable et la gestion des utilisateurs. Sur le " &
                                 "référentiel, il est soumis au double regard comme les autres."

            Case Else
                lblDroits.Text = "Aucun rôle sélectionné : le compte n'aurait accès à rien."
        End Select
    End Sub

    Private Sub cboRole_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRole.SelectedIndexChanged
        AfficherDroits()
        AccorderLaFonctionAuRole()
    End Sub

    ''' <summary>
    ''' Un agent de la compense n'a aucun accès au référentiel : lui attribuer une fonction
    ''' n'aurait aucun effet, et laisser le choix ouvert laisserait croire le contraire.
    ''' </summary>
    Private Sub AccorderLaFonctionAuRole()

        Dim role As RoleWU = UtilisateurWU.RoleDepuisLibelleLisible(Convert.ToString(cboRole.SelectedItem))
        Dim concerne As Boolean = role = RoleWU.Commercial OrElse role = RoleWU.Administrateur

        cboFonction.Enabled = concerne
        If Not concerne Then cboFonction.SelectedIndex = 0

        If Not concerne Then
            lblAideFonction.Text = "Sans objet : ce rôle n'a aucun accès au référentiel."
        ElseIf _existant Is Nothing AndAlso _premierAdministrateur Then
            ' Le premier administrateur peut se passer de fonction : il commencera par créer
            ' les comptes et leur attribuer inputer et authorizer.
            lblAideFonction.Text = "Facultative. Vous pourrez l'attribuer ensuite, à vous comme aux autres."
        Else
            lblAideFonction.Text = "Facultative. Un inputer et un authorizer doivent être deux personnes."
        End If
    End Sub

    ''' <summary>Fonction retenue dans la liste déroulante.</summary>
    Private Function FonctionChoisie() As FonctionWU

        If Not cboFonction.Enabled Then Return FonctionWU.Aucune
        Return UtilisateurWU.FonctionDepuisLibelleLisible(Convert.ToString(cboFonction.SelectedItem))
    End Function

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        lblMessage.Text = String.Empty

        Dim anomalies As List(Of String) = ControlerSaisie()

        If anomalies.Count > 0 Then
            lblMessage.Text = String.Join(Environment.NewLine, anomalies)
            Return
        End If

        Dim utilisateur As New UtilisateurWU() With {
            .Identifiant = txtIdentifiant.Text.Trim(),
            .NomComplet = txtNomComplet.Text.Trim(),
            .Role = UtilisateurWU.RoleDepuisLibelleLisible(Convert.ToString(cboRole.SelectedItem)),
            .Fonction = FonctionChoisie(),
            .Actif = chkActif.Checked
        }

        Dim messageErreur As String = String.Empty
        Dim reussi As Boolean

        If _existant Is Nothing Then
            reussi = UtilisateurRepository.Ajouter(utilisateur, txtMotDePasse.Text, messageErreur)
        Else
            reussi = UtilisateurRepository.Modifier(utilisateur, messageErreur)
        End If

        If Not reussi Then
            lblMessage.Text = messageErreur
            Return
        End If

        _enregistre = utilisateur
        DialogResult = DialogResult.OK
        Close()
    End Sub

    ''' <summary>Contrôle la saisie et retourne les anomalies constatées.</summary>
    Private Function ControlerSaisie() As List(Of String)

        Dim anomalies As New List(Of String)
        Dim identifiant As String = txtIdentifiant.Text.Trim()

        If identifiant.Length = 0 Then
            anomalies.Add("L'identifiant est obligatoire.")
        ElseIf identifiant.Any(Function(caractere) Char.IsWhiteSpace(caractere)) Then
            ' Un identifiant à espaces se saisit mal et se compare mal : autant l'interdire.
            anomalies.Add("L'identifiant ne doit pas contenir d'espace.")
        End If

        If txtNomComplet.Text.Trim().Length = 0 Then
            anomalies.Add("Le nom et le prénom sont obligatoires : le journal doit désigner une personne.")
        End If

        If UtilisateurWU.RoleDepuisLibelleLisible(Convert.ToString(cboRole.SelectedItem)) = RoleWU.Inconnu Then
            anomalies.Add("Le rôle est obligatoire.")
        End If

        If _existant IsNot Nothing Then Return anomalies

        anomalies.AddRange(MotDePasseService.Anomalies(txtMotDePasse.Text, identifiant))

        If txtMotDePasse.Text <> txtConfirmation.Text Then
            anomalies.Add("Le mot de passe et sa confirmation sont différents.")
        End If

        Return anomalies
    End Function

    Private Sub btnAnnuler_Click(sender As Object, e As EventArgs) Handles btnAnnuler.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

#End Region

End Class
