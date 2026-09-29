Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Choix du produit de transfert, entre la connexion et l'espace de travail.
'''
''' L'identification dit QUI travaille ; cette fenêtre dit SUR QUOI. Les deux précèdent
''' l'espace de travail, qui n'est même pas construit tant que le produit n'est pas choisi :
''' son titre le nomme, et ses écrans en dépendent.
'''
''' TOUS LES PRODUITS SONT OUVRABLES, Y COMPRIS CEUX QUI NE SONT PAS PRÊTS
'''
''' Un produit en préparation n'est ni grisé ni refusé : son espace de travail s'ouvre, avec
''' ses menus, et chaque écran métier annonce qu'il n'est pas encore disponible. C'est ce qui
''' permet de voir le cadre avant que le métier n'existe — et c'était la demande.
'''
''' AUCUN DROIT PAR PRODUIT
'''
''' Un utilisateur habilité à traiter la compense la traite pour tous les produits. Le produit
''' n'est pas un droit : les droits restent ceux du rôle, et ils sont les mêmes partout.
''' </summary>
Public Class FrmChoixProduit

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

    ''' <summary>
    ''' Le produit retenu, ou Nothing si l'utilisateur a quitté. N'a de sens que si la fenêtre
    ''' a rendu DialogResult.OK.
    ''' </summary>
    Public ReadOnly Property ProduitChoisi As ProduitTransfert
        Get
            Return _choisi
        End Get
    End Property

    Private _choisi As ProduitTransfert

    ''' <summary>Cote des logos, en pixels : la vignette des lignes, et le logo en exergue.</summary>
    Private Const TAILLE_VIGNETTE As Integer = 32
    Private Const TAILLE_LOGO As Integer = 96

    Private Sub FrmChoixProduit_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblUtilisateur.Text = SessionWU.Description

        RemplirLaListe()
        PreselectionnerLeProduit()
    End Sub

    ''' <summary>
    ''' Inscrit un produit par ligne. L'objet lui-même est rangé dans Tag : la sélection le
    ''' retrouve ainsi tel quel, sans chercher son code dans la liste.
    ''' </summary>
    Private Sub RemplirLaListe()

        lvProduits.BeginUpdate()

        Try
            lvProduits.Items.Clear()
            imgLogos.Images.Clear()

            ' Seuls les produits EN SERVICE : l'administrateur peut en retirer un de la vue
            ' sans rien effacer de son historique.
            For Each produit As ProduitTransfert In ProduitTransfert.ProduitsEnService()

                ' Le logo est range dans la liste d'images sous le code du produit, et la
                ' ligne le designe par cette cle : deux produits ne peuvent pas se tromper
                ' d'image, meme si l'ordre d'affichage change un jour.
                Dim vignette As Bitmap = LogosProduits.Obtenir(produit, TAILLE_VIGNETTE)
                If vignette IsNot Nothing Then imgLogos.Images.Add(produit.Code, vignette)

                Dim ligne As New ListViewItem(produit.Nom, produit.Code)
                ligne.SubItems.Add(produit.Etat)
                ligne.Tag = produit

                ' Un produit en préparation s'écrit en gris. Il reste ouvrable : la couleur
                ' informe, elle n'interdit pas.
                If Not produit.Disponible Then ligne.ForeColor = SystemColors.GrayText

                lvProduits.Items.Add(ligne)
            Next

        Finally
            lvProduits.EndUpdate()
        End Try
    End Sub

    ''' <summary>
    ''' Place la sélection sur le produit déjà en cours — l'utilisateur revient au choix par le
    ''' menu, et retrouve alors sa ligne sous le curseur — ou, à défaut, sur le premier produit
    ''' disponible. Ouvrir sur une liste sans sélection laisserait le bouton sans effet.
    ''' </summary>
    Private Sub PreselectionnerLeProduit()

        If lvProduits.Items.Count = 0 Then
            btnOuvrir.Enabled = False
            Return
        End If

        Dim voulu As Integer = -1

        For indice As Integer = 0 To lvProduits.Items.Count - 1

            Dim produit As ProduitTransfert = TryCast(lvProduits.Items(indice).Tag, ProduitTransfert)
            If produit Is Nothing Then Continue For

            If ProduitTransfert.Actif IsNot Nothing AndAlso produit Is ProduitTransfert.Actif Then
                voulu = indice
                Exit For
            End If

            If voulu < 0 AndAlso produit.Disponible Then voulu = indice
        Next

        If voulu < 0 Then voulu = 0

        lvProduits.Items(voulu).Selected = True
        lvProduits.Items(voulu).Focused = True
        lvProduits.Select()
    End Sub

    ''' <summary>Le produit de la ligne sélectionnée, ou Nothing.</summary>
    Private Function ProduitSelectionne() As ProduitTransfert

        If lvProduits.SelectedItems.Count = 0 Then Return Nothing
        Return TryCast(lvProduits.SelectedItems(0).Tag, ProduitTransfert)
    End Function

    Private Sub lvProduits_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvProduits.SelectedIndexChanged

        Dim produit As ProduitTransfert = ProduitSelectionne()

        lblDescription.Text = If(produit Is Nothing, String.Empty, produit.Description)
        btnOuvrir.Enabled = produit IsNot Nothing

        AfficherLeLogo(produit)
    End Sub

    ''' <summary>
    ''' Met le logo du produit selectionne en evidence, et dit dans l'infobulle ou deposer le
    ''' fichier officiel tant qu'il n'y est pas.
    '''
    ''' POURQUOI L'INFOBULLE ET NON UNE LIGNE DE TEXTE
    '''
    ''' Cet ecran s'ouvre a chaque lancement, devant l'agent de compense : une ligne lui
    ''' annoncant tous les matins qu'un fichier manque serait un reproche qu'il ne peut pas
    ''' lever. L'information est pour celui qui installe, et elle l'attend la ou il la
    ''' cherchera -- sur le logo lui-meme.
    ''' </summary>
    Private Sub AfficherLeLogo(produit As ProduitTransfert)

        picLogo.Image = LogosProduits.Obtenir(produit, TAILLE_LOGO)

        If produit Is Nothing Then
            tipLogo.SetToolTip(picLogo, String.Empty)
            Return
        End If

        If LogosProduits.EstFourni(produit) Then
            tipLogo.SetToolTip(picLogo, $"Logo {produit.Nom}")
            Return
        End If

        tipLogo.SetToolTip(picLogo,
                           $"Embleme provisoire. Pour afficher le logo officiel de " &
                           $"{produit.Nom}, deposer le fichier fourni par la banque sous :" &
                           Environment.NewLine & LogosProduits.CheminAttendu(produit))
    End Sub

    ''' <summary>Le double clic ouvre : c'est le geste attendu sur une liste de choix.</summary>
    Private Sub lvProduits_DoubleClick(sender As Object, e As EventArgs) Handles lvProduits.DoubleClick
        Ouvrir()
    End Sub

    Private Sub btnOuvrir_Click(sender As Object, e As EventArgs) Handles btnOuvrir.Click
        Ouvrir()
    End Sub

    Private Sub Ouvrir()

        Dim produit As ProduitTransfert = ProduitSelectionne()
        If produit Is Nothing Then Return

        _choisi = produit
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnQuitter_Click(sender As Object, e As EventArgs) Handles btnQuitter.Click

        _choisi = Nothing
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
