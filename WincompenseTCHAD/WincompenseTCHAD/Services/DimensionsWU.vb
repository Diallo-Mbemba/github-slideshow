Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Adapte la taille des fenêtres à l'écran du poste.
'''
''' Les dimensions posées par le concepteur — 1200 × 700 pour la fenêtre principale, 1184 × 629
''' pour le traitement de la compense, 1040 × 680 pour le rapport d'activité — supposaient un
''' écran large. Sur un poste en 1366 × 768, la fenêtre du rapport d'activité ne tenait pas dans
''' la zone MDI une fois le menu et la barre d'état déduits : une partie de la grille sortait de
''' l'écran, hors d'atteinte.
'''
''' Plutôt que de viser une résolution particulière, chaque fenêtre se mesure ici à la place
''' réellement disponible au moment où elle s'ouvre. Elle s'y réduit si elle est trop grande, et
''' l'occupe si elle est plus petite alors que son contenu gagne à s'étendre — le cas des écrans
''' de liste.
'''
''' Les formulaires ancrent et arriment déjà leurs contrôles : c'est ce qui permet de les
''' redimensionner sans que leur mise en page se défasse.
''' </summary>
Public NotInheritable Class DimensionsWU

    Private Sub New()
    End Sub

    ''' <summary>Marge laissée autour d'une fenêtre fille dans la zone MDI.</summary>
    Private Const MARGE_MDI As Integer = 16

    ''' <summary>Marge laissée autour d'une fenêtre indépendante sur le bureau.</summary>
    Private Const MARGE_ECRAN As Integer = 48

    ''' <summary>
    ''' Part de l'espace disponible qu'occupe un écran de liste. Pas la totalité : une fenêtre
    ''' fille collée aux bords de la zone MDI ne se distingue plus de son conteneur.
    ''' </summary>
    Private Const PART_OCCUPEE As Double = 0.97

    ''' <summary>
    ''' Part de sa taille d'origine sous laquelle une fenêtre ne descend pas : en deçà, les
    ''' contrôles ancrés se chevauchent.
    ''' </summary>
    Private Const PART_MINIMALE As Double = 0.72

    ''' <summary>Adapte une fenêtre à l'espace disponible, sans chercher à l'occuper.</summary>
    Public Shared Sub Adapter(formulaire As Form)
        Adapter(formulaire, False)
    End Sub

    ''' <summary>Adapte une fenêtre à l'espace disponible.</summary>
    ''' <param name="remplirEspace">
    ''' Vrai pour un écran de liste, dont la grille gagne à s'étendre ; faux pour une boîte de
    ''' dialogue, qu'il serait absurde d'étirer sur tout l'écran autour de quatre champs.
    ''' </param>
    Public Shared Sub Adapter(formulaire As Form, remplirEspace As Boolean)

        If formulaire Is Nothing Then Return
        If formulaire.WindowState <> FormWindowState.Normal Then Return

        Dim disponible As Size = ZoneDisponible(formulaire)
        If disponible.Width <= 0 OrElse disponible.Height <= 0 Then Return

        Dim marge As Integer = If(formulaire.MdiParent Is Nothing, MARGE_ECRAN, MARGE_MDI)

        Dim largeurMax As Integer = Math.Max(disponible.Width - marge, 1)
        Dim hauteurMax As Integer = Math.Max(disponible.Height - marge, 1)

        Dim largeur As Integer = formulaire.Width
        Dim hauteur As Integer = formulaire.Height

        If remplirEspace Then
            largeur = Math.Max(largeur, CInt(largeurMax * PART_OCCUPEE))
            hauteur = Math.Max(hauteur, CInt(hauteurMax * PART_OCCUPEE))
        End If

        largeur = Math.Min(largeur, largeurMax)
        hauteur = Math.Min(hauteur, hauteurMax)

        ' Le plancher est calculé sur la taille d'origine, puis lui-même ramené à ce que l'écran
        ' peut afficher : une taille minimale plus grande que l'écran rendrait la fenêtre
        ' impossible à replacer.
        formulaire.MinimumSize = New Size(
            Math.Min(CInt(formulaire.Width * PART_MINIMALE), largeurMax),
            Math.Min(CInt(formulaire.Height * PART_MINIMALE), hauteurMax))

        formulaire.Size = New Size(largeur, hauteur)

        If formulaire.MdiParent IsNot Nothing Then Centrer(formulaire, disponible)
    End Sub

    ''' <summary>
    ''' Espace réellement offert à la fenêtre : la zone MDI pour une fenêtre fille, la zone de
    ''' travail de son écran pour une fenêtre indépendante — barre des tâches déduite.
    ''' </summary>
    Private Shared Function ZoneDisponible(formulaire As Form) As Size

        If formulaire.MdiParent IsNot Nothing Then

            For Each controle As Control In formulaire.MdiParent.Controls
                Dim zone As MdiClient = TryCast(controle, MdiClient)
                If zone IsNot Nothing Then Return zone.ClientSize
            Next

            ' Zone MDI introuvable : on se rabat sur la surface utile du parent.
            Return formulaire.MdiParent.ClientSize
        End If

        Try
            Return Screen.FromControl(formulaire).WorkingArea.Size
        Catch ex As InvalidOperationException
            ' Fenêtre sans poignée : l'écran principal fait un repli acceptable.
            Return Screen.PrimaryScreen.WorkingArea.Size
        End Try
    End Function

    ''' <summary>Centre une fenêtre fille dans la zone MDI.</summary>
    Private Shared Sub Centrer(formulaire As Form, disponible As Size)

        formulaire.StartPosition = FormStartPosition.Manual

        formulaire.Location = New Point(
            Math.Max((disponible.Width - formulaire.Width) \ 2, 0),
            Math.Max((disponible.Height - formulaire.Height) \ 2, 0))
    End Sub

    ''' <summary>
    ''' Cadre la fenêtre principale sur la zone de travail de l'écran.
    '''
    ''' Elle s'ouvre agrandie, mais l'utilisateur peut la restaurer : sans taille minimale, il
    ''' pourrait la réduire au point que la zone MDI ne puisse plus rien contenir.
    ''' </summary>
    Public Shared Sub AdapterFenetrePrincipale(formulaire As Form)

        If formulaire Is Nothing Then Return

        Dim travail As Size = ZoneDisponible(formulaire)

        formulaire.MinimumSize = New Size(
            Math.Min(1024, travail.Width),
            Math.Min(640, travail.Height))

        ' Taille de restauration : la quasi-totalité de la zone de travail, quelle que soit la
        ' résolution du poste.
        If formulaire.WindowState = FormWindowState.Normal Then
            formulaire.Size = New Size(
                CInt(travail.Width * PART_OCCUPEE),
                CInt(travail.Height * PART_OCCUPEE))
        End If
    End Sub

End Class
