Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' L'avis qui remplace un écran métier quand le produit en cours n'est pas encore traité.
'''
''' POURQUOI UN ÉCRAN, ET NON UNE BOÎTE DE MESSAGE
'''
''' Une boîte de message se claque et ne laisse rien. Cette fenêtre s'ouvre dans la zone MDI
''' comme les autres, se range dans le menu Fenêtres, et reste sous les yeux : l'utilisateur
''' voit qu'il est bien dans l'espace Ria, que le menu a répondu, et que c'est le métier qui
''' manque — non son geste qui a échoué.
'''
''' UNE SEULE INSTANCE
'''
''' L'espace de travail n'en ouvre qu'une, et la réemploie en changeant son texte. Dix clics
''' sur dix menus donneraient sinon dix fenêtres disant la même chose.
''' </summary>
Public Class FrmEcranIndisponible

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

    ''' <summary>
    ''' Annonce que cet écran-là n'est pas disponible pour le produit en cours.
    ''' </summary>
    ''' <param name="nomDeLEcran">L'écran demandé, tel que le menu le nomme.</param>
    Public Sub Annoncer(nomDeLEcran As String)

        Dim ecran As String = If(String.IsNullOrWhiteSpace(nomDeLEcran), "Cet écran", nomDeLEcran)
        Dim produit As String = ProduitTransfert.NomDuProduitActif

        If String.IsNullOrEmpty(produit) Then

            ' Aucun produit en cours : c'est un enchaînement anormal, et le dire vaut mieux que
            ' d'écrire une phrase à trou.
            Me.Text = "Écran non disponible"
            lblTitre.Text = $"« {ecran} » ne peut pas s'ouvrir."
            lblExplication.Text = "Aucun produit de transfert n'est en cours. Revenez au choix du produit " &
                                  "par le menu « Changer de produit »."
            Return
        End If

        Me.Text = $"{ecran} — {produit}"
        lblTitre.Text = $"« {ecran} » n'est pas encore disponible pour {produit}."

        lblExplication.Text =
            $"L'environnement de {produit} est en place — ses menus, son espace de travail et son " &
            "paramétrage à venir — mais le traitement de sa compensation reste à écrire : ses " &
            "rapports, ses taux, sa pièce comptable et son jeu de tables."
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

End Class
