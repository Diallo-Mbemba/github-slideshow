Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Charte visuelle de l'application : couleurs, polices, métriques, et leur application à un
''' formulaire.
'''
''' Tout passe par ce fichier. Changer l'habillage se fait ici, à un seul endroit, au lieu de
''' reprendre treize fichiers .Designer.vb — que le concepteur de Visual Studio réécrit d'ailleurs
''' à sa façon dès qu'on y ouvre un formulaire, ce qui défairait une partie du travail.
'''
''' En contrepartie, l'aperçu du concepteur ne montre pas le thème : il ne se voit qu'à
''' l'exécution. C'est le prix de la réversibilité.
'''
''' Variante A retenue par la banque : le noir et l'or de la fenêtre de paramétrage d'origine
''' sont conservés, l'ardoise sert aux en-têtes de grille.
''' </summary>
Public NotInheritable Class ThemeWU

    Private Sub New()
    End Sub

#Region "Couleurs"

    ''' <summary>Fond du bandeau de titre.</summary>
    Public Shared ReadOnly NOIR As Color = Color.FromArgb(&H10, &H10, &H10)

    ''' <summary>Texte du bandeau de titre.</summary>
    Public Shared ReadOnly OR_WU As Color = Color.FromArgb(&HFF, &HD2, &H0)

    ''' <summary>En-têtes de grille et bouton d'action principal.</summary>
    Public Shared ReadOnly ARDOISE As Color = Color.FromArgb(&H2E, &H3A, &H46)

    Public Shared ReadOnly ARDOISE_TRAIT As Color = Color.FromArgb(&H3D, &H4B, &H59)
    Public Shared ReadOnly FOND_ECRAN As Color = Color.FromArgb(&HF5, &HF6, &HF8)
    Public Shared ReadOnly LIGNE_ALTERNEE As Color = Color.FromArgb(&HFA, &HFA, &HFA)
    Public Shared ReadOnly QUADRILLAGE As Color = Color.FromArgb(&HE3, &HE3, &HE3)
    Public Shared ReadOnly ENCRE As Color = Color.FromArgb(&H1B, &H1B, &H1B)
    Public Shared ReadOnly ENCRE_DISCRETE As Color = Color.FromArgb(&H5A, &H5A, &H5A)

    ''' <summary>Fond de la zone MDI, et son filigrane.</summary>
    Public Shared ReadOnly FOND_MDI As Color = Color.FromArgb(&H20, &H26, &H2C)

#End Region

#Region "Polices et métriques"

    Public Const POLICE As String = "Segoe UI"
    Public Const TAILLE_COURANTE As Single = 9.0F
    Public Const TAILLE_BANDEAU As Single = 15.0F

    ''' <summary>Hauteur du bandeau de titre, identique sur les treize écrans.</summary>
    Public Const HAUTEUR_BANDEAU As Integer = 44

    ''' <summary>Hauteur d'une ligne de grille : 24 px au lieu des 22 px par défaut.</summary>
    Public Const HAUTEUR_LIGNE As Integer = 24

    Public Const HAUTEUR_ENTETE As Integer = 28

    ''' <summary>Nom du bandeau posé par le thème. Sert aussi à ne pas l'appliquer deux fois.</summary>
    Private Const NOM_BANDEAU As String = "panelBandeauWU"

    Private Const NOM_HOTE As String = "panelHoteWU"

    ''' <summary>Nom du bandeau d'origine, sur les quatre écrans qui en portaient déjà un.</summary>
    Private Const NOM_BANDEAU_ORIGINE As String = "panelTitre"

    ''' <summary>
    ''' Police ancienne à remplacer. Deux libellés la portaient explicitement et n'auraient donc
    ''' pas hérité de celle du formulaire.
    ''' </summary>
    Private Const POLICE_ANCIENNE As String = "Microsoft Sans Serif"

#End Region

#Region "Application à un formulaire"

    ''' <summary>Applique la charte à un formulaire et lui pose son bandeau de titre.</summary>
    ''' <param name="formulaire">Formulaire à habiller.</param>
    ''' <param name="titre">Titre du bandeau. Vide : le texte de la barre de titre est repris.</param>
    Public Shared Sub Appliquer(formulaire As Form, titre As String)
        Appliquer(formulaire, titre, Nothing)
    End Sub

    ''' <summary>
    ''' Applique la charte, pose le bandeau et met un bouton en avant.
    ''' </summary>
    ''' <param name="boutonPrincipal">
    ''' Action de l'écran — « Calculer », « Enregistrer ». Un seul par écran : si tous les boutons
    ''' sont mis en avant, aucun ne l'est.
    ''' </param>
    Public Shared Sub Appliquer(formulaire As Form, titre As String, boutonPrincipal As Button)

        If formulaire Is Nothing Then Return

        ' Déjà habillé : réappliquer poserait un second bandeau.
        If formulaire.Controls.Find(NOM_BANDEAU, False).Length > 0 Then Return

        formulaire.SuspendLayout()

        Try
            formulaire.Font = New Font(POLICE, TAILLE_COURANTE, FontStyle.Regular, GraphicsUnit.Point)
            formulaire.BackColor = FOND_ECRAN

            If Not formulaire.IsMdiContainer Then
                PoserBandeau(formulaire, TitreRetenu(formulaire, titre))
            End If

            Habiller(formulaire)

            If boutonPrincipal IsNot Nothing Then MettreEnAvant(boutonPrincipal)

        Finally
            formulaire.ResumeLayout(True)
        End Try
    End Sub

    ''' <summary>Titre à inscrire dans le bandeau.</summary>
    Private Shared Function TitreRetenu(formulaire As Form, titre As String) As String

        If Not String.IsNullOrWhiteSpace(titre) Then Return titre.Trim()

        ' À défaut, la barre de titre de la fenêtre : elle dit déjà de quel écran il s'agit.
        Return If(formulaire.Text, String.Empty)
    End Function

#End Region

#Region "Bandeau de titre"

    ''' <summary>
    ''' Pose le bandeau noir et or en haut du formulaire.
    '''
    ''' Les contrôles existants sont déplacés dans un panneau d'accueil occupant tout l'espace
    ''' restant, plutôt que décalés un par un : leurs coordonnées et leurs ancrages restent
    ''' valables tels quels, et aucun calcul de position ne peut se tromper.
    ''' </summary>
    Private Shared Sub PoserBandeau(formulaire As Form, titre As String)

        Dim bandeauOrigine As Panel = TrouverBandeauOrigine(formulaire)
        Dim titreFinal As String = titre
        Dim decalage As Integer = 0

        If bandeauOrigine IsNot Nothing Then
            ' Quatre écrans portaient déjà un bandeau, encadré et de trois tailles différentes.
            ' Il cède la place au bandeau commun, et ce qui le suivait remonte d'autant.
            titreFinal = TexteDuBandeau(bandeauOrigine, titre)
            decalage = bandeauOrigine.Bottom - HAUTEUR_BANDEAU
            formulaire.Controls.Remove(bandeauOrigine)
            bandeauOrigine.Dispose()
        End If

        Dim aDeplacer As New List(Of Control)

        For Each controle As Control In formulaire.Controls
            ' Les barres d'état et de menus gardent leur ancrage sur le formulaire lui-même.
            If TypeOf controle Is StatusStrip OrElse TypeOf controle Is MenuStrip Then Continue For
            If TypeOf controle Is MdiClient Then Continue For
            aDeplacer.Add(controle)
        Next

        Dim hote As New Panel() With {
            .Name = NOM_HOTE,
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent
        }

        For Each controle As Control In aDeplacer
            formulaire.Controls.Remove(controle)
            If decalage <> 0 Then controle.Top -= decalage
            hote.Controls.Add(controle)
        Next

        Dim bandeau As Panel = ConstruireBandeau(titreFinal)

        formulaire.Controls.Add(hote)
        formulaire.Controls.Add(bandeau)

        ' L'ordre de rattachement décide de l'ordre d'ancrage : Windows Forms ancre du dernier
        ' enfant vers le premier. Le panneau d'accueil doit passer en dernier pour recevoir
        ' l'espace restant, donc figurer en tête de la collection.
        formulaire.Controls.SetChildIndex(hote, 0)
        formulaire.Controls.SetChildIndex(bandeau, 1)

        ' Le formulaire grandit de la hauteur du bandeau, pour que la zone de travail reste
        ' celle que le concepteur avait prévue. Lorsqu'un bandeau était déjà présent, il occupait
        ' déjà cette place : la fenêtre rétrécit au contraire de ce qu'il prenait en trop.
        If bandeauOrigine Is Nothing Then
            formulaire.Height += HAUTEUR_BANDEAU
        Else
            formulaire.Height -= decalage
        End If
    End Sub

    ''' <summary>Bandeau d'origine du formulaire, ou Nothing s'il n'en portait pas.</summary>
    Private Shared Function TrouverBandeauOrigine(formulaire As Form) As Panel

        Dim trouves As Control() = formulaire.Controls.Find(NOM_BANDEAU_ORIGINE, False)
        If trouves.Length = 0 Then Return Nothing

        Return TryCast(trouves(0), Panel)
    End Function

    ''' <summary>Texte porté par le bandeau d'origine, à reprendre dans le nouveau.</summary>
    Private Shared Function TexteDuBandeau(bandeauOrigine As Panel, defaut As String) As String

        For Each controle As Control In bandeauOrigine.Controls
            Dim libelle As Label = TryCast(controle, Label)
            If libelle IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(libelle.Text) Then
                Return libelle.Text
            End If
        Next

        Return defaut
    End Function

    Private Shared Function ConstruireBandeau(titre As String) As Panel

        Dim libelle As New Label() With {
            .Name = "lblBandeauWU",
            .Dock = DockStyle.Fill,
            .BackColor = NOIR,
            .ForeColor = OR_WU,
            .Font = New Font(POLICE, TAILLE_BANDEAU, FontStyle.Bold, GraphicsUnit.Point),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Text = titre
        }

        Dim bandeau As New Panel() With {
            .Name = NOM_BANDEAU,
            .Dock = DockStyle.Top,
            .Height = HAUTEUR_BANDEAU,
            .BackColor = NOIR
        }

        bandeau.Controls.Add(libelle)
        Return bandeau
    End Function

#End Region

#Region "Habillage des contrôles"

    ''' <summary>Parcourt l'arbre des contrôles et applique la charte à chacun.</summary>
    Private Shared Sub Habiller(parent As Control)

        For Each controle As Control In parent.Controls

            RemplacerPoliceAncienne(controle)

            Dim grille As DataGridView = TryCast(controle, DataGridView)
            If grille IsNot Nothing Then
                HabillerGrille(grille)
                Continue For
            End If

            Dim onglets As TabControl = TryCast(controle, TabControl)
            If onglets IsNot Nothing Then
                For Each page As TabPage In onglets.TabPages
                    page.BackColor = Color.White
                    Habiller(page)
                Next
                Continue For
            End If

            Dim barre As StatusStrip = TryCast(controle, StatusStrip)
            If barre IsNot Nothing Then
                barre.BackColor = Color.White
                barre.ForeColor = ENCRE_DISCRETE
                Continue For
            End If

            Dim menu As MenuStrip = TryCast(controle, MenuStrip)
            If menu IsNot Nothing Then
                menu.BackColor = Color.White
                menu.ForeColor = ENCRE
                Continue For
            End If

            If controle.HasChildren Then Habiller(controle)
        Next
    End Sub

    ''' <summary>
    ''' Remplace la police héritée de Windows Forms là où elle avait été fixée explicitement.
    ''' Ailleurs, le contrôle hérite de celle du formulaire et il n'y a rien à faire.
    ''' </summary>
    Private Shared Sub RemplacerPoliceAncienne(controle As Control)

        If controle.Font Is Nothing Then Return
        If Not String.Equals(controle.Font.Name, POLICE_ANCIENNE, StringComparison.OrdinalIgnoreCase) Then Return

        controle.Font = New Font(POLICE, controle.Font.SizeInPoints, controle.Font.Style, GraphicsUnit.Point)
    End Sub

    ''' <summary>
    ''' Habille une grille : en-têtes ardoise, lignes alternées, quadrillage clair.
    '''
    ''' La colonne d'en-tête de ligne disparaît : elle était vide, ne servait à rien ici, et
    ''' coûtait vingt-cinq pixels sur chaque grille.
    '''
    ''' La sélection ne repeint plus toute la ligne. Elle l'effaçait jusqu'ici en bleu, y compris
    ''' la couleur d'anomalie — celle-là même qu'on venait de cliquer pour l'examiner.
    ''' </summary>
    Public Shared Sub HabillerGrille(grille As DataGridView)

        If grille Is Nothing Then Return

        grille.EnableHeadersVisualStyles = False
        grille.BackgroundColor = Color.White
        grille.BorderStyle = BorderStyle.FixedSingle
        grille.CellBorderStyle = DataGridViewCellBorderStyle.Single
        grille.GridColor = QUADRILLAGE
        grille.RowHeadersVisible = False
        grille.AllowUserToResizeRows = False
        grille.ColumnHeadersHeight = HAUTEUR_ENTETE
        grille.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

        With grille.ColumnHeadersDefaultCellStyle
            .BackColor = ARDOISE
            .ForeColor = Color.White
            .SelectionBackColor = ARDOISE
            .SelectionForeColor = Color.White
            .Font = New Font(POLICE, TAILLE_COURANTE, FontStyle.Bold, GraphicsUnit.Point)
            .Padding = New Padding(6, 0, 6, 0)
        End With

        With grille.DefaultCellStyle
            .BackColor = Color.White
            .ForeColor = ENCRE
            .Padding = New Padding(5, 0, 5, 0)
            ' La sélection garde le fond de la ligne : seul le texte change, ce qui suffit à
            ' montrer où l'on est sans masquer la couleur d'état.
            .SelectionBackColor = Color.FromArgb(&HD6, &HE4, &HF0)
            .SelectionForeColor = ENCRE
        End With

        grille.AlternatingRowsDefaultCellStyle.BackColor = LIGNE_ALTERNEE
        grille.RowTemplate.Height = HAUTEUR_LIGNE

        For Each ligne As DataGridViewRow In grille.Rows
            ligne.Height = HAUTEUR_LIGNE
        Next
    End Sub

    ''' <summary>Met un bouton en avant : aplat ardoise, texte blanc.</summary>
    Public Shared Sub MettreEnAvant(bouton As Button)

        If bouton Is Nothing Then Return

        bouton.FlatStyle = FlatStyle.Flat
        bouton.FlatAppearance.BorderSize = 1
        bouton.FlatAppearance.BorderColor = Color.FromArgb(&H22, &H2C, &H35)
        bouton.FlatAppearance.MouseOverBackColor = Color.FromArgb(&H3D, &H4B, &H59)
        bouton.FlatAppearance.MouseDownBackColor = Color.FromArgb(&H1B, &H23, &H2B)
        bouton.BackColor = ARDOISE
        bouton.ForeColor = Color.White
        bouton.UseVisualStyleBackColor = False
        bouton.Font = New Font(POLICE, TAILLE_COURANTE, FontStyle.Bold, GraphicsUnit.Point)
    End Sub

#End Region

#Region "Zone MDI"

    ''' <summary>
    ''' Habille le fond de la zone MDI, gris Windows par défaut.
    '''
    ''' Le MdiClient n'existe pas au moment où le concepteur travaille : Windows Forms le crée
    ''' lui-même parmi les enfants du formulaire. Il faut donc aller le chercher.
    ''' </summary>
    Public Shared Sub HabillerZoneMdi(formulaire As Form)

        If formulaire Is Nothing Then Return

        For Each controle As Control In formulaire.Controls
            Dim zone As MdiClient = TryCast(controle, MdiClient)
            If zone Is Nothing Then Continue For

            zone.BackColor = FOND_MDI
            Return
        Next
    End Sub

#End Region

End Class
