Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Déclaration des produits de transfert — écran d'administration.
'''
''' CE QUE CET ÉCRAN PEUT, ET CE QU'IL NE PEUT PAS
'''
''' Il déclare qu'un produit EXISTE : son code, son nom, sa description, sa couleur, son rang
''' d'affichage, et s'il est en service. Un produit ajouté ici apparaît aussitôt dans la
''' fenêtre de choix que voit l'utilisateur après sa connexion.
'''
''' Il ne déclare JAMAIS qu'un produit est TRAITÉ. Cela, c'est le code de l'application qui le
''' sait. Si cet écran pouvait le faire, un administrateur déclarant MoneyGram disponible
''' ouvrirait l'écran de traitement de la compense — qui lirait les rapports Western Union,
''' appliquerait les taux Western Union et produirait une pièce sur les comptes Western Union,
''' SOUS UN NOM MONEYGRAM. Personne ne s'en apercevrait avant la comptabilisation.
'''
''' TROIS REFUS, ET AUCUN N'EST DE LA PRUDENCE DÉCORATIVE
'''
'''   - Le CODE ne se modifie pas après création : il nomme le fichier du logo et, demain, les
'''     tables du produit. Le changer orphelinerait les uns et les autres.
'''   - AUCUNE SUPPRESSION n'est offerte : un produit supprimé laisserait son historique, ses
'''     pièces et son référentiel sans rien pour dire à quoi ils se rapportaient.
'''   - LE DERNIER produit en service ne peut pas être retiré : il n'y aurait plus rien à
'''     choisir après la connexion, et plus d'écran pour revenir en arrière.
''' </summary>
Public Class FrmProduits

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

    ''' <summary>Côté des vignettes de la liste, en pixels.</summary>
    Private Const TAILLE_VIGNETTE As Integer = 24

    ''' <summary>
    ''' Vrai quand l'écran compose un produit qui n'existe pas encore. Le code est alors
    ''' saisissable ; il ne l'est plus une fois le produit créé.
    ''' </summary>
    Private _creation As Boolean

    ''' <summary>
    ''' La couleur en cours de composition. Elle vit ici et non dans la couleur de fond du
    ''' panneau : un panneau peut être grisé par Windows, et la couleur saisie serait perdue.
    ''' </summary>
    Private _couleur As Color = Color.FromArgb(108, 117, 125)

    ''' <summary>Vrai pendant qu'on remplit les champs : les événements de saisie se taisent.</summary>
    Private _remplissage As Boolean

    Private Sub FrmProduits_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesComptesSystemes Then
            MessageBox.Show("La déclaration des produits de transfert est réservée aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        RelireLesProduits()
    End Sub

#Region "La liste"

    ''' <summary>Relit la table, réinstalle la liste de l'application, et rafraîchit l'écran.</summary>
    Private Sub RelireLesProduits(Optional codeASelectionner As String = Nothing)

        Dim messageErreur As String = String.Empty
        Dim deLaBase As Boolean = ProduitRepository.Charger(messageErreur)

        lblStatut.Text = If(deLaBase,
                            $"{ProduitTransfert.Tous.Count} produit(s) déclaré(s). " &
                            "Les modifications sont visibles dès le prochain retour au choix du produit.",
                            messageErreur)

        ' Sans table, rien ne peut être enregistré : mieux vaut fermer les commandes que
        ' laisser l'administrateur saisir un produit qui disparaîtra à la fermeture.
        btnNouveau.Enabled = deLaBase
        btnEnregistrer.Enabled = deLaBase

        RemplirLaListe()
        SelectionnerLeProduit(codeASelectionner)
    End Sub

    Private Sub RemplirLaListe()

        lvProduits.BeginUpdate()

        Try
            lvProduits.Items.Clear()
            imgLogos.Images.Clear()

            For Each produit As ProduitTransfert In ProduitTransfert.Tous

                Dim vignette As Bitmap = LogosProduits.Obtenir(produit, TAILLE_VIGNETTE)
                If vignette IsNot Nothing Then imgLogos.Images.Add(produit.Code, vignette)

                Dim ligne As New ListViewItem(produit.Code, produit.Code)
                ligne.SubItems.Add(produit.Nom)
                ligne.SubItems.Add(produit.Ordre.ToString(Globalization.CultureInfo.CurrentCulture))
                ligne.SubItems.Add(produit.Etat)
                ligne.Tag = produit

                If Not produit.EnService Then ligne.ForeColor = SystemColors.GrayText

                lvProduits.Items.Add(ligne)
            Next

        Finally
            lvProduits.EndUpdate()
        End Try
    End Sub

    Private Sub SelectionnerLeProduit(code As String)

        If lvProduits.Items.Count = 0 Then
            Composer(Nothing)
            Return
        End If

        For indice As Integer = 0 To lvProduits.Items.Count - 1

            Dim produit As ProduitTransfert = TryCast(lvProduits.Items(indice).Tag, ProduitTransfert)
            If produit Is Nothing Then Continue For

            If code Is Nothing OrElse produit.PorteLeCode(code) Then
                lvProduits.Items(indice).Selected = True
                lvProduits.Items(indice).Focused = True
                lvProduits.EnsureVisible(indice)
                Return
            End If
        Next

        lvProduits.Items(0).Selected = True
    End Sub

    Private Function ProduitSelectionne() As ProduitTransfert

        If lvProduits.SelectedItems.Count = 0 Then Return Nothing
        Return TryCast(lvProduits.SelectedItems(0).Tag, ProduitTransfert)
    End Function

    Private Sub lvProduits_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvProduits.SelectedIndexChanged

        If _remplissage Then Return
        Composer(ProduitSelectionne())
    End Sub

#End Region

#Region "Le formulaire"

    ''' <summary>
    ''' Prépare les champs pour ce produit, ou pour un produit neuf si l'argument est Nothing.
    ''' </summary>
    Private Sub Composer(produit As ProduitTransfert)

        _remplissage = True

        Try
            _creation = produit Is Nothing

            txtCode.Text = If(_creation, String.Empty, produit.Code)
            txtCode.ReadOnly = Not _creation

            txtNom.Text = If(_creation, String.Empty, produit.Nom)
            txtDescription.Text = If(_creation, String.Empty, produit.Description)

            _couleur = If(_creation, Color.FromArgb(108, 117, 125), produit.Couleur)
            pnlCouleur.BackColor = _couleur

            numOrdre.Value = If(_creation, ProchainOrdre(), CDec(produit.Ordre))
            chkEnService.Checked = If(_creation, True, produit.EnService)

            AfficherLEtat(produit)

        Finally
            _remplissage = False
        End Try
    End Sub

    ''' <summary>Le rang du produit suivant : dix de plus que le dernier, pour laisser la place.</summary>
    Private Shared Function ProchainOrdre() As Decimal

        Dim dernier As Integer = 0

        For Each produit As ProduitTransfert In ProduitTransfert.Tous
            If produit.Ordre > dernier Then dernier = produit.Ordre
        Next

        Return CDec(Math.Min(dernier + 10, 9999))
    End Function

    ''' <summary>
    ''' Dit l'état du produit, et surtout POURQUOI. « En attente » sans explication ressemble à
    ''' une panne ; accompagné de sa raison, c'est une information.
    ''' </summary>
    Private Sub AfficherLEtat(produit As ProduitTransfert)

        If _creation Then
            lblEtat.Text = "Ce produit naîtra « En attente » : l'application ne sait pas encore traiter sa compensation."
            lblTracabilite.Text = String.Empty
            Return
        End If

        If produit Is Nothing Then
            lblEtat.Text = String.Empty
            lblTracabilite.Text = String.Empty
            Return
        End If

        If Not produit.EnService Then
            lblEtat.Text = "Hors service : ce produit n'apparaît pas dans la fenêtre de choix. Rien n'est effacé."

        ElseIf Not produit.EstTraite Then
            lblEtat.Text = "En attente : son espace de travail s'ouvre, mais ses écrans métier " &
                           "annoncent qu'ils ne sont pas encore disponibles."
        Else
            lblEtat.Text = "Disponible : l'application sait traiter la compensation de ce produit."
        End If

        lblTracabilite.Text = Tracabilite(produit)
    End Sub

    ''' <summary>Qui a créé ce produit, et qui l'a modifié en dernier.</summary>
    Private Shared Function Tracabilite(produit As ProduitTransfert) As String

        Dim morceaux As New List(Of String)

        If produit.DateCreation.HasValue Then
            morceaux.Add($"Créé le {produit.DateCreation.Value:dd/MM/yyyy}" &
                         If(String.IsNullOrEmpty(produit.CreePar), String.Empty, $" par {produit.CreePar}"))
        End If

        If produit.DateModification.HasValue Then
            morceaux.Add($"modifié le {produit.DateModification.Value:dd/MM/yyyy}" &
                         If(String.IsNullOrEmpty(produit.ModifiePar), String.Empty, $" par {produit.ModifiePar}"))
        End If

        Return String.Join(" · ", morceaux)
    End Function

    Private Sub btnCouleur_Click(sender As Object, e As EventArgs) Handles btnCouleur.Click

        dlgCouleur.Color = _couleur
        dlgCouleur.FullOpen = True

        If dlgCouleur.ShowDialog(Me) <> DialogResult.OK Then Return

        _couleur = dlgCouleur.Color
        pnlCouleur.BackColor = _couleur
    End Sub

    Private Sub btnNouveau_Click(sender As Object, e As EventArgs) Handles btnNouveau.Click

        _remplissage = True
        Try
            lvProduits.SelectedItems.Clear()
        Finally
            _remplissage = False
        End Try

        Composer(Nothing)
        txtCode.Focus()
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim refus As String = String.Empty
        If Not LaSaisieTient(refus) Then
            MessageBox.Show(refus, "Saisie incomplète", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim code As String = txtCode.Text.Trim().ToUpperInvariant()

        ' L'état de création est retenu MAINTENANT : la relecture qui suit recompose le
        ' formulaire et le remet à faux, et le compte rendu final dirait alors « enregistré »
        ' pour un produit qui vient d'être créé.
        Dim creation As Boolean = _creation

        If Not creation AndAlso ProduitTransfert.ParCode(code) Is Nothing Then
            MessageBox.Show($"Le produit « {code} » n'est plus dans la liste. Rouvrez l'écran.",
                            "Produit introuvable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Un objet NEUF, et non celui de la liste : si l'enregistrement échoue, la liste
        ' affichée ne doit pas porter des valeurs que la base n'a pas.
        Dim produit As New ProduitTransfert(code)

        produit.Nom = txtNom.Text
        produit.Description = txtDescription.Text
        produit.Couleur = _couleur
        produit.Ordre = CInt(numOrdre.Value)
        produit.EnService = chkEnService.Checked

        Dim messageErreur As String = String.Empty
        Dim pose As Boolean = If(creation,
                                 ProduitRepository.Creer(produit, messageErreur),
                                 ProduitRepository.Enregistrer(produit, messageErreur))

        If Not pose Then
            MessageBox.Show(messageErreur, "Enregistrement impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        RelireLesProduits(code)

        lblStatut.Text = If(creation,
                            $"Produit « {code} » créé. Il apparaît dès maintenant dans la fenêtre de choix.",
                            $"Produit « {code} » enregistré.")
    End Sub

    ''' <summary>
    ''' Vérifie la saisie, et refuse ce qui doit l'être. Chaque refus a une conséquence
    ''' concrète derrière lui, aucun n'est une formalité.
    ''' </summary>
    Private Function LaSaisieTient(ByRef refus As String) As Boolean

        refus = String.Empty

        Dim code As String = txtCode.Text.Trim().ToUpperInvariant()

        If Not ProduitTransfert.CodeValide(code, refus) Then Return False

        If String.IsNullOrWhiteSpace(txtNom.Text) Then
            refus = "Le nom du produit est obligatoire : c'est lui que l'utilisateur lit dans la fenêtre de choix."
            Return False
        End If

        ' --- un code déjà pris
        If _creation AndAlso ProduitTransfert.ParCode(code) IsNot Nothing Then
            refus = $"Le code « {code} » est déjà employé par un autre produit. Un code ne peut " &
                    "pas servir deux fois : il nomme le logo et les tables du produit."
            Return False
        End If

        If chkEnService.Checked Then Return True

        ' --- le produit en cours d'utilisation
        If Not _creation AndAlso ProduitTransfert.EstLeProduitActif(code) Then

            refus = $"« {code} » est le produit que vous traitez en ce moment : il ne peut pas être " &
                    "mis hors service depuis son propre espace de travail." & Environment.NewLine &
                    Environment.NewLine &
                    "Changez de produit, puis revenez ici."
            Return False
        End If

        ' --- le dernier en service
        Dim autresEnService As Integer = 0

        For Each produit As ProduitTransfert In ProduitTransfert.Tous
            If produit.EnService AndAlso Not produit.PorteLeCode(code) Then autresEnService += 1
        Next

        If autresEnService = 0 Then
            refus = "C'est le dernier produit en service. Le mettre hors service laisserait la " &
                    "fenêtre de choix vide, et plus personne ne pourrait travailler."
            Return False
        End If

        Return True
    End Function

#End Region

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

End Class
