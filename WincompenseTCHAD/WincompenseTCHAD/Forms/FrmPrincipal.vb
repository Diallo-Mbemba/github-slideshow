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

        ' CenterParent et CenterScreen ne conviennent pas à une fenêtre fille : le premier ne
        ' vaut que pour un affichage modal, le second la placerait au centre de l'écran, donc
        ' à cheval sur les bords de la zone MDI. Le placement est donc calculé à la main.
        nouveau.StartPosition = FormStartPosition.Manual
        CentrerDansLaZoneMdi(nouveau)

        nouveau.Show()
        Return nouveau
    End Function

    ''' <summary>
    ''' Centre une fenêtre fille dans la zone de travail MDI, en la rétrécissant d'abord si
    ''' elle y est trop grande.
    '''
    ''' Une fenêtre ouverte en haut à gauche, à sa taille de conception, débordait de la zone
    ''' sur les postes à petit écran : ses boutons du bas devenaient inaccessibles sans faire
    ''' défiler. Le dimensionnement précède donc le centrage, faute de quoi on centrerait une
    ''' fenêtre plus large que la place disponible.
    ''' </summary>
    Private Sub CentrerDansLaZoneMdi(enfant As Form)

        If enfant Is Nothing Then Return

        Dim zone As MdiClient = ZoneMdi()
        If zone Is Nothing Then Return

        Dim disponible As Size = zone.ClientSize
        If disponible.Width <= 0 OrElse disponible.Height <= 0 Then Return

        ' MinimumSize peut refuser le rétrécissement : la taille réellement obtenue est donc
        ' relue sur la fenêtre, et non supposée égale à celle qu'on vient de lui demander.
        enfant.Size = New Size(Math.Min(enfant.Width, disponible.Width),
                               Math.Min(enfant.Height, disponible.Height))

        enfant.Location = New Point(Math.Max(0, (disponible.Width - enfant.Width) \ 2),
                                    Math.Max(0, (disponible.Height - enfant.Height) \ 2))
    End Sub

    ''' <summary>
    ''' Zone de travail MDI, c'est-à-dire le fond gris sur lequel s'ouvrent les fenêtres filles.
    ''' Windows Forms l'ajoute lui-même aux contrôles du conteneur, sans l'exposer autrement que
    ''' par ce parcours. Retourne Nothing tant que la fenêtre n'est pas construite.
    ''' </summary>
    Private Function ZoneMdi() As MdiClient

        For Each controle As Control In Me.Controls

            Dim zone As MdiClient = TryCast(controle, MdiClient)
            If zone IsNot Nothing Then Return zone
        Next

        Return Nothing
    End Function

    Private Sub mnuTraitement_Click(sender As Object, e As EventArgs) Handles mnuTraitement.Click
        AfficherEnfant(Of FrmCompensationWU)()
    End Sub

    Private Sub mnuRapport_Click(sender As Object, e As EventArgs) Handles mnuRapport.Click
        AfficherEnfant(Of FrmRapportSousAgents)()
    End Sub

    ''' <summary>
    ''' Le rapport du réseau propre. Deux entrées et non une avec un filtre : les deux
    ''' populations n'ont pas les mêmes axes, et une fenêtre qui bascule de l'une à l'autre
    ''' donne l'occasion de citer les chiffres des sous-agents en croyant parler des agences.
    ''' </summary>
    Private Sub mnuRapportAgences_Click(sender As Object, e As EventArgs) Handles mnuRapportAgences.Click
        AfficherEnfant(Of FrmRapportAgences)()
    End Sub

    ''' <summary>
    ''' Consultation des pièces déjà produites. Rangée avec le rapport d'activité, et sous le
    ''' même droit : ce sont deux façons de regarder ce qui a été comptabilisé, l'une qui cumule
    ''' sur une période, l'autre qui restitue un justificatif daté.
    ''' </summary>
    Private Sub mnuPiecesArchivees_Click(sender As Object, e As EventArgs) Handles mnuPiecesArchivees.Click
        AfficherEnfant(Of FrmPiecesArchivees)()
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

    Private Sub mnuDemandes_Click(sender As Object, e As EventArgs) Handles mnuDemandes.Click
        AfficherEnfant(Of FrmDemandes)()
        RafraichirLeCompteurDeDemandes()
    End Sub

#End Region

#Region "Sécurité"

    Private Sub mnuUtilisateurs_Click(sender As Object, e As EventArgs) Handles mnuUtilisateurs.Click
        AfficherEnfant(Of FrmUtilisateurs)()
    End Sub

    ''' <summary>
    ''' Regarde si le partage annonce une version plus récente, et le dit.
    '''
    ''' Rien n'avertissait un agent qu'il travaillait sur une version dépassée : une correction
    ''' livrée un lundi pouvait rester ignorée d'un poste pendant des semaines, et deux agents
    ''' produire des pièces différentes à partir des mêmes rapports.
    '''
    ''' La mention reste dans la barre d'état tant que la version n'a pas changé ; le message,
    ''' lui, ne paraît qu'une fois par version. Répété chaque matin, il serait fermé sans être
    ''' lu — et le suivant, celui qui compte vraiment, le serait aussi.
    '''
    ''' Rien n'est bloquant, et rien n'est téléchargé : l'informatique garde la main sur le
    ''' moment de la mise à jour. Pendant la marche en parallèle, deux versions différentes
    ''' fausseraient la comparaison avec la pièce manuelle.
    ''' </summary>
    Private Sub SignalerUneVersionPlusRecente()

        Dim publiee As MiseAJourWU.VersionPubliee = MiseAJourWU.Verifier()
        If publiee Is Nothing Then Return

        _setupDeLaMiseAJour = publiee.CheminSetup

        tsslMiseAJour.Text = $"Version {publiee.Numero} disponible — cliquez ici"
        tsslMiseAJour.ToolTipText = If(publiee.Note.Length > 0, publiee.Note, "Mise à jour publiée sur le partage.")
        tsslMiseAJour.Visible = True

        If MiseAJourWU.DejaSignalee(publiee.Numero) Then Return

        MessageBox.Show(
            $"La version {publiee.Numero} est publiée. Ce poste utilise la version {MiseAJourWU.VersionCourante()}." &
            Environment.NewLine & Environment.NewLine &
            If(publiee.Note.Length > 0, publiee.Note & Environment.NewLine & Environment.NewLine, String.Empty) &
            If(_setupDeLaMiseAJour.Length > 0,
               "Programme d'installation :" & Environment.NewLine & "    " & _setupDeLaMiseAJour,
               "Rapprochez-vous de l'informatique pour l'installer.") & Environment.NewLine & Environment.NewLine &
            "Vous pouvez continuer à travailler : la mise à jour n'est pas obligatoire aujourd'hui.",
            "Mise à jour disponible", MessageBoxButtons.OK, MessageBoxIcon.Information)

        MiseAJourWU.MarquerSignalee(publiee.Numero)
    End Sub

    ''' <summary>Chemin du programme d'installation annoncé, pour le clic sur la barre d'état.</summary>
    Private _setupDeLaMiseAJour As String = String.Empty

    ''' <summary>
    ''' Ouvre le dossier du programme d'installation dans l'Explorateur, le fichier sélectionné.
    '''
    ''' On n'exécute rien : installer une application est le geste de l'informatique, et le
    ''' lancer au nom d'un agent qui n'a pas les droits échouerait de toute façon.
    ''' </summary>
    Private Sub tsslMiseAJour_Click(sender As Object, e As EventArgs) Handles tsslMiseAJour.Click

        If _setupDeLaMiseAJour.Length = 0 Then
            MessageBox.Show("Aucun programme d'installation n'est indiqué sur le partage." & Environment.NewLine &
                            "Rapprochez-vous de l'informatique.",
                            "Mise à jour", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ' Les guillemets encadrent le chemin : un dossier au nom espacé serait sinon lu
            ' comme plusieurs arguments, et l'Explorateur ouvrirait la mauvaise fenêtre.
            Dim argument As String = "/select," & Chr(34) & _setupDeLaMiseAJour & Chr(34)
            Process.Start("explorer.exe", argument)

        Catch ex As Exception
            MessageBox.Show("Impossible d'ouvrir l'emplacement :" & Environment.NewLine &
                            _setupDeLaMiseAJour & Environment.NewLine & Environment.NewLine & ex.Message,
                            "Mise à jour", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Ouvre l'écran de connexion à la base, réservé à l'administrateur.
    '''
    ''' Il est modal, et non fenêtre fille : changer de serveur pendant qu'un traitement est en
    ''' cours dans une autre fenêtre donnerait une journée moitié lue sur un serveur, moitié sur
    ''' l'autre. L'état de la base est relu ensuite, pour que la barre d'état dise la vérité.
    ''' </summary>
    Private Sub mnuConnexionBase_Click(sender As Object, e As EventArgs) Handles mnuConnexionBase.Click

        Using parametres As New FrmParametresConnexion()
            parametres.ShowDialog(Me)
        End Using

        AfficherEtatBase()
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
        mnuRapportAgences.Available = rapports
        mnuPiecesArchivees.Available = rapports
        SEP1.Available = traite AndAlso rapports
        mnuCompensation.Available = traite OrElse rapports

        mnuSousAgents.Available = pointsDeVente
        mnuAgences.Available = pointsDeVente
        mnuGroupes.Available = pointsDeVente
        mnuComptes.Available = comptes
        mnuDemandes.Available = pointsDeVente
        SEP5.Available = pointsDeVente
        SEP2.Available = pointsDeVente AndAlso comptes

        ' Un menu dont toutes les entrées sont masquées resterait affiché, et s'ouvrirait sur
        ' un vide : il disparaît avec elles.
        mnuParametrage.Available = pointsDeVente OrElse comptes

        mnuUtilisateurs.Available = utilisateurs

        ' Changer de serveur engage toute la banque : seul l'administrateur, qui gère déjà les
        ' utilisateurs, peut ouvrir cet écran.
        mnuConnexionBase.Available = utilisateurs
        SEP6.Available = utilisateurs
        SEP4.Available = utilisateurs
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

        Dim zone As MdiClient = ZoneMdi()
        If zone Is Nothing Then Return

        AddHandler zone.Paint, AddressOf ZoneMdi_Paint

        ' Sans redessin au redimensionnement, le mot resterait centré sur l'ancienne largeur
        ' et se retrouverait de travers dès qu'on agrandit la fenêtre.
        AddHandler zone.Resize, AddressOf ZoneMdi_Resize
    End Sub

    Private Sub ZoneMdi_Resize(sender As Object, e As EventArgs)
        DirectCast(sender, Control).Invalidate()
    End Sub

    ''' <summary>
    ''' Numéro de version, lu sur l'assemblage plutôt que recopié en dur : il suit ainsi
    ''' automatiquement l'AssemblyVersion et ne peut pas annoncer une version que le binaire
    ''' n'est pas.
    ''' </summary>
    Private Shared ReadOnly Property NumeroDeVersion As String
        Get
            Dim identite As Reflection.AssemblyName = Reflection.Assembly.GetExecutingAssembly().GetName()
            If identite Is Nothing OrElse identite.Version Is Nothing Then Return String.Empty

            Return $"Version {identite.Version.Major}.{identite.Version.Minor}.{identite.Version.Build}"
        End Get
    End Property

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

                Dim mesure As SizeF = e.Graphics.MeasureString(FILIGRANE, police)

                RayerLeFiligrane(e.Graphics, zone.BackColor, largeur, hauteur, taille, mesure)
                EcrireLaVersion(e.Graphics, largeur, hauteur, taille, mesure, alignement)
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
    Private Shared Sub RayerLeFiligrane(surface As Graphics, couleurFond As Color,
                                        largeur As Integer, hauteur As Integer,
                                        taille As Single, mesure As SizeF)

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

    ''' <summary>
    ''' Inscrit le numéro de version juste sous le filigrane.
    '''
    ''' Il est écrit après les rayures, et non avant : elles le hacheraient, alors qu'un numéro
    ''' de version doit rester lisible — c'est ce qu'on demandera à l'agent au téléphone le jour
    ''' où quelque chose ne tournera pas rond.
    ''' </summary>
    Private Shared Sub EcrireLaVersion(surface As Graphics, largeur As Integer, hauteur As Integer,
                                       taille As Single, mesure As SizeF, alignement As StringFormat)

        Dim texte As String = NumeroDeVersion
        If String.IsNullOrEmpty(texte) Then Return

        Dim tailleVersion As Single = Math.Max(taille / 6.0F, 12.0F)

        Using police As New Font("Segoe UI", tailleVersion, FontStyle.Regular, GraphicsUnit.Pixel)

            ' La boîte de ligne mesurée descend nettement plus bas que les capitales, à cause
            ' des jambages et de l'interligne. S'en tenir à elle laisserait un blanc de la moitié
            ' d'une hauteur de police entre le mot et son numéro : on remonte d'autant.
            Const RETRAIT_INTERLIGNE As Single = 0.22F

            Dim cadre As New RectangleF(
                0.0F,
                (CSng(hauteur) + mesure.Height) / 2.0F - taille * RETRAIT_INTERLIGNE,
                CSng(largeur),
                tailleVersion * 1.8F)

            ' Moins opaque que le mot : le numéro accompagne le filigrane, il ne rivalise pas
            ' avec lui.
            Using pinceau As New SolidBrush(Color.FromArgb(170, 255, 255, 255))
                surface.DrawString(texte, police, pinceau, cadre, alignement)
            End Using
        End Using
    End Sub

#End Region

#Region "Démarrage et barre d'état"

    Private Sub FrmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PoserFiligrane()
        AppliquerLesDroits()
        AfficherEtatBase()

        tsslUtilisateur.Text = SessionWU.Description
        RafraichirLeCompteurDeDemandes()
        SignalerUneVersionPlusRecente()

        ' Aucun écran n'est ouvert d'office : l'application s'ouvre sur sa zone de travail, et
        ' c'est l'utilisateur qui choisit par où commencer. Ouvrir le traitement de la compense
        ' d'emblée imposait cet écran au commercial, qui n'y a pas accès, et faisait attendre
        ' l'agent de compense les jours où il venait seulement consulter un rapport.
    End Sub

    ''' <summary>
    ''' Inscrit le nombre de demandes en attente dans l'entrée de menu.
    '''
    ''' Sans ce rappel, une demande peut dormir une semaine parce que personne ne sait qu'elle
    ''' existe — et pendant ce temps un sous-agent créé n'est rattaché à rien.
    ''' </summary>
    Private Sub RafraichirLeCompteurDeDemandes()

        If Not mnuDemandes.Available Then Return

        Dim enAttente As Integer = DemandeRepository.CompterEnAttente()

        mnuDemandes.Text = If(enAttente > 0,
                              $"&Autorisations du référentiel ({enAttente})",
                              "&Autorisations du référentiel")
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
