Option Strict On
Option Explicit On

Imports System.Drawing
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

#Region "Sécurité"

    Private Sub mnuUtilisateurs_Click(sender As Object, e As EventArgs) Handles mnuUtilisateurs.Click
        AfficherEnfant(Of FrmUtilisateurs)()
    End Sub

    Private Sub mnuMonMotDePasse_Click(sender As Object, e As EventArgs) Handles mnuMonMotDePasse.Click

        If SessionWU.Utilisateur Is Nothing Then Return

        ' Le mot de passe actuel sera exigé : un poste laissé ouvert ne doit pas permettre à un
        ' tiers de changer le mot de passe de son titulaire et de le priver de son outil.
        Using changement As New FrmChangerMotDePasse(SessionWU.Identifiant)
            changement.ShowDialog(Me)
        End Using
    End Sub

    ''' <summary>
    ''' Masque les écrans que le rôle connecté n'a pas le droit d'ouvrir.
    '''
    ''' Les entrées sont masquées et non grisées : une entrée grisée invite à demander le droit,
    ''' alors que la banque a tranché — l'agent de compense n'a AUCUN accès au paramétrage, pas
    ''' même en lecture. Le contrôle est de toute façon refait dans chaque écran : un menu n'est
    ''' qu'un confort de navigation, jamais une barrière.
    '''
    ''' C'est Available et non Visible qui est employé : sur un élément de menu, Visible répond
    ''' « faux » tant que le menu déroulant qui le contient n'est pas ouvert, si bien qu'une
    ''' entrée qu'on vient d'afficher se relit comme masquée. Les deux droits seraient alors
    ''' calculés à partir d'une lecture fausse.
    ''' </summary>
    Private Sub AppliquerLesDroits()

        Dim traite As Boolean = SessionWU.PeutTraiterLaCompense
        Dim rapports As Boolean = SessionWU.PeutVoirLesRapports
        Dim pointsDeVente As Boolean = SessionWU.PeutGererLesPointsDeVente
        Dim comptes As Boolean = SessionWU.PeutGererLesComptesSystemes
        Dim utilisateurs As Boolean = SessionWU.PeutGererLesUtilisateurs

        mnuTraitement.Available = traite
        mnuRapport.Available = rapports
        SEP1.Available = traite AndAlso rapports
        mnuCompensation.Available = traite OrElse rapports

        mnuSousAgents.Available = pointsDeVente
        mnuAgences.Available = pointsDeVente
        mnuGroupes.Available = pointsDeVente
        mnuComptes.Available = comptes
        SEP2.Available = pointsDeVente AndAlso comptes

        ' Un menu dont toutes les entrées sont masquées resterait affiché, et s'ouvrirait sur
        ' un vide : il disparaît avec elles.
        mnuParametrage.Available = pointsDeVente OrElse comptes

        mnuUtilisateurs.Available = utilisateurs
        SEP4.Available = utilisateurs
    End Sub

    ''' <summary>
    ''' Ouvre l'écran correspondant à l'usage quotidien du rôle connecté : le traitement pour
    ''' l'agent de compense, le rapport d'activité pour le commercial. Ouvrir un écran interdit
    ''' au démarrage n'accueillerait l'utilisateur qu'avec un refus.
    ''' </summary>
    Private Sub OuvrirEcranDAccueil()

        If SessionWU.PeutTraiterLaCompense Then
            AfficherEnfant(Of FrmCompensationWU)()
        ElseIf SessionWU.PeutVoirLesRapports Then
            AfficherEnfant(Of FrmRapportActivite)()
        End If
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

        ' La liste est relevée avant la première fermeture : fermer une fenêtre la retire de
        ' MdiChildren, et énumérer directement cette collection pendant qu'elle se vide en
        ' sauterait une sur deux.
        Dim ouvertes As Form() = Me.MdiChildren

        For Each enfant As Form In ouvertes
            enfant.Close()
        Next
    End Sub

#End Region

#Region "Filigrane de la zone MDI"

    ''' <summary>Mot inscrit en fond de la zone de travail.</summary>
    Private Const FILIGRANE As String = "WINCOMPENSE"

    ''' <summary>
    ''' Inscrit le nom de l'application en fond de la zone MDI.
    '''
    ''' Le MdiClient n'existe pas au moment où le concepteur travaille : Windows Forms le crée
    ''' lui-même parmi les enfants du formulaire dès que IsMdiContainer passe à vrai. Il faut
    ''' donc aller le chercher à l'ouverture plutôt que le poser dans le .Designer.vb.
    '''
    ''' Le mot est peint et non posé dans une étiquette : une étiquette resterait au premier plan
    ''' et les fenêtres filles passeraient dessous.
    ''' </summary>
    Private Sub PoserFiligrane()

        For Each controle As Control In Me.Controls

            Dim zone As MdiClient = TryCast(controle, MdiClient)
            If zone Is Nothing Then Continue For

            AddHandler zone.Paint, AddressOf ZoneMdi_Paint

            ' Sans redessin au redimensionnement, le mot resterait centré sur l'ancienne largeur
            ' et se retrouverait de travers dès qu'on agrandit la fenêtre.
            AddHandler zone.Resize, AddressOf ZoneMdi_Resize
            Return
        Next
    End Sub

    Private Sub ZoneMdi_Resize(sender As Object, e As EventArgs)
        DirectCast(sender, Control).Invalidate()
    End Sub

    Private Sub ZoneMdi_Paint(sender As Object, e As PaintEventArgs)

        Dim zone As Control = DirectCast(sender, Control)

        Dim largeur As Integer = zone.ClientSize.Width
        Dim hauteur As Integer = zone.ClientSize.Height
        If largeur <= 0 OrElse hauteur <= 0 Then Return

        ' Taille proportionnelle à la largeur, entre deux bornes : le filigrane garde la même
        ' présence sur un poste en 1366 pixels et sur un écran large, sans jamais déborder.
        Dim taille As Single = Math.Min(Math.Max(CSng(largeur) / 11.0F, 28.0F), 200.0F)

        Using police As New Font("Segoe UI", taille, FontStyle.Bold, GraphicsUnit.Pixel)
            Using alignement As New StringFormat()

                alignement.Alignment = StringAlignment.Center
                alignement.LineAlignment = StringAlignment.Center

                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

                Dim cadre As New RectangleF(0.0F, 0.0F, CSng(largeur), CSng(hauteur))

                ' Une ombre portée d'abord : du blanc sur le gris de la zone de travail manquerait
                ' de contraste, et le mot se lirait mal.
                Using pinceauOmbre As New SolidBrush(Color.FromArgb(100, 0, 0, 0))
                    e.Graphics.DrawString(FILIGRANE, police, pinceauOmbre,
                                          New RectangleF(cadre.X + 3.0F, cadre.Y + 3.0F,
                                                         cadre.Width, cadre.Height),
                                          alignement)
                End Using

                Using pinceau As New SolidBrush(Color.White)
                    e.Graphics.DrawString(FILIGRANE, police, pinceau, cadre, alignement)
                End Using

                RayerLeFiligrane(e.Graphics, police, zone.BackColor, largeur, hauteur, taille)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Fait passer des rayures en travers du mot.
    '''
    ''' Elles reprennent exactement la couleur de fond de la zone MDI : sur le fond, elles sont
    ''' donc invisibles, et n'apparaissent qu'en travers des lettres blanches, qu'elles découpent
    ''' en lames. Peindre par-dessus avec la couleur du fond revient à effacer par bandes.
    '''
    ''' Elles sont tracées en dernier, après l'ombre et après le mot, faute de quoi l'ombre
    ''' resterait pleine là où la lettre est coupée.
    ''' </summary>
    Private Shared Sub RayerLeFiligrane(surface As Graphics, police As Font, couleurFond As Color,
                                        largeur As Integer, hauteur As Integer, taille As Single)

        Dim mesure As SizeF = surface.MeasureString(FILIGRANE, police)

        ' La bande déborde un peu du mot : l'ombre portée est décalée de trois pixels, et les
        ' rayures doivent la traverser elle aussi.
        Dim bande As New RectangleF(
            (CSng(largeur) - mesure.Width) / 2.0F - 2.0F,
            (CSng(hauteur) - mesure.Height) / 2.0F - 2.0F,
            mesure.Width + 8.0F,
            mesure.Height + 8.0F)

        ' Sept rayures environ sur la hauteur des lettres : assez pour que l'effet se voie,
        ' assez peu pour que le mot reste lisible.
        Dim periode As Single = taille / 7.0F
        Dim epaisseur As Single = periode * 0.38F

        Using pinceauRayure As New SolidBrush(couleurFond)

            Dim y As Single = bande.Y

            While y < bande.Bottom
                surface.FillRectangle(pinceauRayure, bande.X, y, bande.Width, epaisseur)
                y += periode
            End While
        End Using
    End Sub

#End Region

#Region "Démarrage et barre d'état"

    Private Sub FrmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PoserFiligrane()
        AppliquerLesDroits()
        AfficherEtatBase()

        tsslUtilisateur.Text = SessionWU.Description

        OuvrirEcranDAccueil()
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
