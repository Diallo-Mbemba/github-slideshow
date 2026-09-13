Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Fenêtre principale de l'application, conteneur MDI.
'''
''' Elle est le point d'accès unique aux écrans : traitement de la compense, rapport
''' d'activité, paramétrage des points de vente et des comptes systèmes. Chaque écran s'ouvre
''' comme fenêtre fille, ce qui permet d'en consulter plusieurs à la fois — comparer un rapport
''' d'activité et le paramétrage d'un sous-agent, par exemple — là où des boîtes de dialogue
''' modales l'interdisaient.
'''
''' Un écran déjà ouvert n'est jamais dupliqué : il est ramené au premier plan. Sans cette
''' règle, dix clics sur un menu produiraient dix copies de la même liste, chacune avec ses
''' propres données, et l'utilisateur ne saurait plus laquelle fait foi.
''' </summary>
Public Class FrmPrincipal

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture des écrans"

    ''' <summary>
    ''' Ouvre un écran en fenêtre fille, ou ramène au premier plan celui qui est déjà ouvert.
    ''' </summary>
    ''' <typeparam name="T">Type du formulaire à ouvrir.</typeparam>
    Public Function AfficherEnfant(Of T As {Form, New})() As T

        For Each enfant As Form In Me.MdiChildren

            If Not TypeOf enfant Is T Then Continue For

            ' Une fenêtre réduite doit être rétablie, sinon « l'ouvrir » ne montrerait rien.
            If enfant.WindowState = FormWindowState.Minimized Then
                enfant.WindowState = FormWindowState.Normal
            End If

            enfant.Activate()
            Return DirectCast(enfant, T)
        Next

        Dim nouveau As New T()
        nouveau.MdiParent = Me

        ' Une fenêtre fille ne se centre pas sur son parent : le réglage prévu pour un affichage
        ' modal placerait la fenêtre hors de la zone MDI.
        nouveau.StartPosition = FormStartPosition.WindowsDefaultLocation

        nouveau.Show()
        Return nouveau
    End Function

    Private Sub mnuTraitement_Click(sender As Object, e As EventArgs) Handles mnuTraitement.Click
        AfficherEnfant(Of FrmCompensationWU)()
    End Sub

    Private Sub mnuRapport_Click(sender As Object, e As EventArgs) Handles mnuRapport.Click
        AfficherEnfant(Of FrmRapportActivite)()
    End Sub

    Private Sub mnuSousAgents_Click(sender As Object, e As EventArgs) Handles mnuSousAgents.Click
        AfficherEnfant(Of FrmSousAgents)()
    End Sub

    Private Sub mnuAgences_Click(sender As Object, e As EventArgs) Handles mnuAgences.Click
        AfficherEnfant(Of FrmAgences)()
    End Sub

    Private Sub mnuGroupes_Click(sender As Object, e As EventArgs) Handles mnuGroupes.Click
        AfficherEnfant(Of FrmGroupesStatistiques)()
    End Sub

    Private Sub mnuComptes_Click(sender As Object, e As EventArgs) Handles mnuComptes.Click
        AfficherEnfant(Of FrmComptesSysteme)()
    End Sub

#End Region

#Region "Disposition des fenêtres"

    Private Sub mnuCascade_Click(sender As Object, e As EventArgs) Handles mnuCascade.Click
        LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub mnuMosaiqueH_Click(sender As Object, e As EventArgs) Handles mnuMosaiqueH.Click
        LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub mnuMosaiqueV_Click(sender As Object, e As EventArgs) Handles mnuMosaiqueV.Click
        LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub mnuFermerTout_Click(sender As Object, e As EventArgs) Handles mnuFermerTout.Click

        ' La collection est parcourue sur une copie : fermer une fenêtre la retire de MdiChildren,
        ' ce qui fausserait l'énumération en cours.
        For Each enfant As Form In Me.MdiChildren.Clone()
            enfant.Close()
        Next
    End Sub

#End Region

#Region "Démarrage et barre d'état"

    Private Sub FrmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AfficherEtatBase()

        ' L'écran de traitement est ouvert d'emblée : c'est l'usage quotidien de l'application.
        AfficherEnfant(Of FrmCompensationWU)()
    End Sub

    ''' <summary>
    ''' Indique dans la barre d'état la base à laquelle l'application est reliée. Une erreur de
    ''' base — environnement de test pris pour la production — se voit alors immédiatement.
    ''' </summary>
    Private Sub AfficherEtatBase()

        Try
            Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder(
                WURepository.ObtenirChaineConnexion())

            tsslBase.Text = $"Serveur {constructeur.DataSource} — base {constructeur.InitialCatalog}"

        Catch ex As ArgumentException
            ' Chaîne de connexion illisible : le diagnostic viendra de la première requête.
            tsslBase.Text = "Chaîne de connexion non exploitable."
        End Try
    End Sub

    Private Sub mnuQuitter_Click(sender As Object, e As EventArgs) Handles mnuQuitter.Click
        Close()
    End Sub

#End Region

End Class
