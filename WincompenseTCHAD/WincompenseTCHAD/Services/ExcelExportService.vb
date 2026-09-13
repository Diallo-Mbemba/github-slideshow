' Liaison tardive vers Microsoft Excel : Option Strict Off est INDISPENSABLE ici, et
' volontairement circonscrit à ce seul fichier. Il évite d'imposer au projet une référence COM
' « Microsoft Excel XX.0 Object Library », dont le numéro de version varie d'un poste à l'autre
' et empêcherait la compilation sur un poste équipé d'une autre version d'Office.
' Tout le reste de l'application demeure sous Option Strict On.
Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.IO

''' <summary>
''' Un tableau à écrire dans la feuille Excel : son intitulé et ses données. Plusieurs blocs
''' peuvent se succéder dans une même feuille (par exemple un récapitulatif puis un détail).
''' </summary>
Public Class BlocExcel

    ''' <summary>Intitulé affiché au-dessus du tableau. Facultatif.</summary>
    Public Property Titre As String = String.Empty

    ''' <summary>Données du tableau. Les noms de colonnes servent d'en-têtes par défaut.</summary>
    Public Property Donnees As DataTable

    ''' <summary>
    ''' Intitulés de colonnes à substituer aux noms techniques, par nom de colonne.
    ''' Une colonne absente de ce dictionnaire conserve son nom.
    ''' </summary>
    Public Property Entetes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Formats de nombre Excel à appliquer, par nom de colonne (ex. "0 %", "# ##0").
    ''' Une colonne absente garde le format général.
    ''' </summary>
    Public Property Formats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Pose un filtre automatique sur l'en-tête de ce bloc. Excel n'admettant qu'un seul filtre
    ''' par feuille, une seule valeur True est prise en compte — la première rencontrée.
    ''' </summary>
    Public Property AvecFiltre As Boolean = False

    Public Sub New()
    End Sub

    Public Sub New(titreBloc As String, donneesBloc As DataTable)
        Titre = titreBloc
        Donnees = donneesBloc
    End Sub

End Class

''' <summary>
''' Export générique de tableaux vers Microsoft Excel, avec titre, sous-titres, en-têtes mis en
''' forme et mise en page d'impression.
'''
''' Volontairement indépendant de tout métier : il ne connaît que des DataTable. La pièce
''' comptable conserve son propre export (PieceComptableService), dont la mise en forme répond
''' à un format imposé par la Direction Comptable.
'''
''' Excel n'est jamais requis pour faire fonctionner l'application : son absence est signalée
''' par une exception explicite, au seul moment où l'utilisateur demande un export.
''' </summary>
Public NotInheritable Class ExcelExportService

    Private Sub New()
    End Sub

#Region "Constantes Excel (liaison tardive : les énumérations Interop ne sont pas disponibles)"

    Private Const XL_CENTRE As Integer = -4108
    Private Const XL_GAUCHE As Integer = -4131
    Private Const XL_TRAIT_CONTINU As Integer = 1
    Private Const XL_PAYSAGE As Integer = 2
    Private Const XL_MAXIMISE As Integer = -4137

#End Region

    ''' <summary>
    ''' Écrit les blocs dans un classeur Excel, l'enregistre et le laisse OUVERT à l'écran pour
    ''' consultation, impression ou enregistrement sous un autre nom.
    ''' </summary>
    ''' <param name="titre">Titre principal, en tête de feuille.</param>
    ''' <param name="sousTitres">Lignes d'information sous le titre (période, filtre, date d'édition...).</param>
    ''' <param name="blocs">Tableaux à écrire, dans l'ordre.</param>
    ''' <param name="nomFeuille">Nom de l'onglet.</param>
    ''' <param name="cheminFichier">Chemin du fichier .xlsx à créer.</param>
    ''' <exception cref="InvalidOperationException">Excel absent du poste, ou aucune donnée.</exception>
    Public Shared Sub ExporterEtOuvrir(titre As String,
                                       sousTitres As IEnumerable(Of String),
                                       blocs As IEnumerable(Of BlocExcel),
                                       nomFeuille As String,
                                       cheminFichier As String)

        Dim listeBlocs As List(Of BlocExcel) = BlocsExploitables(blocs)

        If listeBlocs.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter.")
        End If

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing
        Dim feuille As Object = Nothing

        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            If typeExcel Is Nothing Then
                Throw New InvalidOperationException(
                    "Microsoft Excel n'est pas installé sur ce poste : l'export est impossible." &
                    Environment.NewLine &
                    "La liste reste consultable à l'écran.")
            End If

            excelApp = Activator.CreateInstance(typeExcel)
            excelApp.Visible = True
            excelApp.DisplayAlerts = False

            classeur = excelApp.Workbooks.Add()
            feuille = classeur.Worksheets(1)

            RemplirFeuille(feuille, titre, sousTitres, listeBlocs, nomFeuille)

            classeur.SaveAs(cheminFichier)

            ' Mise au premier plan : purement cosmétique, et protégée par un Try/Catch silencieux.
            ' Windows peut refuser le focus à un processus d'arrière-plan ; le classeur est de
            ' toute façon déjà ouvert, l'échec ne doit donc pas remonter à l'utilisateur.
            Try
                excelApp.WindowState = XL_MAXIMISE
                classeur.Activate()
            Catch
            End Try

        Catch
            ' Échec en cours de route : Excel ne doit pas rester ouvert sur un classeur incomplet.
            Try
                If classeur IsNot Nothing Then classeur.Close(False)
            Catch
            End Try
            Try
                If excelApp IsNot Nothing Then excelApp.Quit()
            Catch
            End Try
            Throw

        Finally
            ' En cas de succès le classeur reste ouvert : seules les références COM intermédiaires
            ' sont libérées, excelApp restant actif tant que sa fenêtre est affichée.
            If feuille IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(feuille)
            If classeur IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(classeur)
        End Try
    End Sub

    ''' <summary>Ne retient que les blocs porteurs d'au moins une ligne.</summary>
    Private Shared Function BlocsExploitables(blocs As IEnumerable(Of BlocExcel)) As List(Of BlocExcel)

        Dim retenus As New List(Of BlocExcel)

        If blocs Is Nothing Then Return retenus

        For Each bloc As BlocExcel In blocs
            If bloc IsNot Nothing AndAlso bloc.Donnees IsNot Nothing AndAlso bloc.Donnees.Columns.Count > 0 Then
                retenus.Add(bloc)
            End If
        Next

        Return retenus
    End Function

#Region "Écriture de la feuille"

    Private Shared Sub RemplirFeuille(feuille As Object, titre As String,
                                      sousTitres As IEnumerable(Of String),
                                      blocs As List(Of BlocExcel), nomFeuille As String)

        If Not String.IsNullOrWhiteSpace(nomFeuille) Then
            ' Excel limite le nom d'onglet à 31 caractères et en interdit certains.
            feuille.Name = NettoyerNomFeuille(nomFeuille)
        End If

        ' Largeur du bandeau de titre : celle du bloc le plus large.
        Dim largeur As Integer = 1
        For Each bloc As BlocExcel In blocs
            largeur = Math.Max(largeur, bloc.Donnees.Columns.Count)
        Next

        Dim ligne As Integer = 1
        ligne = EcrireTitre(feuille, titre, sousTitres, largeur, ligne)

        ' Hauteur réelle du bandeau : c'est elle qui sera répétée en haut de chaque page
        ' imprimée, et non un nombre de lignes supposé.
        Dim hauteurBandeau As Integer = ligne - 1

        Dim filtrePose As Boolean = False

        For Each bloc As BlocExcel In blocs
            ligne = EcrireBloc(feuille, bloc, ligne, filtrePose)
            ligne += 1 ' ligne vide entre deux tableaux
        Next

        feuille.Columns.AutoFit()
        PreparerImpression(feuille, titre, hauteurBandeau)
    End Sub

    ''' <summary>Écrit le bandeau de titre et les sous-titres. Retourne la première ligne libre.</summary>
    Private Shared Function EcrireTitre(feuille As Object, titre As String,
                                        sousTitres As IEnumerable(Of String),
                                        largeur As Integer, ligne As Integer) As Integer

        If Not String.IsNullOrWhiteSpace(titre) Then

            Dim plage As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, largeur))
            feuille.Cells(ligne, 1).Value = titre

            plage.Merge()
            plage.Font.Bold = True
            plage.Font.Size = 14
            plage.HorizontalAlignment = XL_CENTRE
            plage.Interior.Color = RGB(31, 78, 120)
            plage.Font.Color = RGB(255, 255, 255)
            feuille.Rows(ligne).RowHeight = 24

            ligne += 1
        End If

        If sousTitres IsNot Nothing Then
            For Each sousTitre As String In sousTitres

                If String.IsNullOrWhiteSpace(sousTitre) Then Continue For

                Dim plage As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, largeur))
                feuille.Cells(ligne, 1).Value = sousTitre

                plage.Merge()
                plage.HorizontalAlignment = XL_GAUCHE
                plage.Font.Italic = True

                ligne += 1
            Next
        End If

        Return ligne + 1 ' une ligne vide sépare l'en-tête des tableaux
    End Function

    ''' <summary>Écrit un tableau (intitulé, en-têtes, données). Retourne la première ligne libre.</summary>
    Private Shared Function EcrireBloc(feuille As Object, bloc As BlocExcel,
                                       ligne As Integer, ByRef filtrePose As Boolean) As Integer

        Dim nombreColonnes As Integer = bloc.Donnees.Columns.Count

        If Not String.IsNullOrWhiteSpace(bloc.Titre) Then
            feuille.Cells(ligne, 1).Value = bloc.Titre
            Dim plageTitre As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, nombreColonnes))
            plageTitre.Font.Bold = True
            plageTitre.Font.Size = 11
            ligne += 1
        End If

        ' --- En-têtes ---
        Dim ligneEntete As Integer = ligne

        For c As Integer = 0 To nombreColonnes - 1
            Dim colonne As DataColumn = bloc.Donnees.Columns(c)
            Dim intitule As String = colonne.ColumnName

            If bloc.Entetes IsNot Nothing AndAlso bloc.Entetes.ContainsKey(colonne.ColumnName) Then
                intitule = bloc.Entetes(colonne.ColumnName)
            End If

            feuille.Cells(ligneEntete, c + 1).Value = intitule
        Next

        Dim plageEntete As Object = feuille.Range(feuille.Cells(ligneEntete, 1),
                                                  feuille.Cells(ligneEntete, nombreColonnes))
        plageEntete.Font.Bold = True
        plageEntete.Interior.Color = RGB(217, 226, 243)
        plageEntete.HorizontalAlignment = XL_CENTRE
        plageEntete.Borders.LineStyle = XL_TRAIT_CONTINU

        ligne += 1

        ' --- Données ---
        Dim premiereLigneDonnees As Integer = ligne

        For Each enregistrement As DataRow In bloc.Donnees.Rows
            For c As Integer = 0 To nombreColonnes - 1
                Dim valeur As Object = enregistrement(c)
                If valeur Is DBNull.Value Then Continue For

                ' Les valeurs numériques sont écrites en tant que nombres, pour rester
                ' calculables dans Excel ; tout le reste part en texte.
                If TypeOf valeur Is Decimal OrElse TypeOf valeur Is Double OrElse
                   TypeOf valeur Is Integer OrElse TypeOf valeur Is Long Then
                    feuille.Cells(ligne, c + 1).Value = valeur
                Else
                    feuille.Cells(ligne, c + 1).Value = Convert.ToString(valeur, CultureInfo.CurrentCulture)
                End If
            Next
            ligne += 1
        Next

        Dim derniereLigne As Integer = ligne - 1

        If derniereLigne >= premiereLigneDonnees Then

            feuille.Range(feuille.Cells(premiereLigneDonnees, 1),
                          feuille.Cells(derniereLigne, nombreColonnes)).Borders.LineStyle = XL_TRAIT_CONTINU

            ' --- Formats de nombre ---
            For c As Integer = 0 To nombreColonnes - 1
                Dim nom As String = bloc.Donnees.Columns(c).ColumnName
                If bloc.Formats Is Nothing OrElse Not bloc.Formats.ContainsKey(nom) Then Continue For

                feuille.Range(feuille.Cells(premiereLigneDonnees, c + 1),
                              feuille.Cells(derniereLigne, c + 1)).NumberFormat = bloc.Formats(nom)
            Next
        End If

        ' Excel n'admet qu'un filtre automatique par feuille : seul le premier demandé est posé.
        If bloc.AvecFiltre AndAlso Not filtrePose AndAlso derniereLigne >= premiereLigneDonnees Then
            Try
                feuille.Range(feuille.Cells(ligneEntete, 1),
                              feuille.Cells(derniereLigne, nombreColonnes)).AutoFilter()
                filtrePose = True
            Catch
                ' Le filtre est un confort : son échec ne compromet pas l'export.
            End Try
        End If

        Return ligne
    End Function

    ''' <summary>
    ''' Mise en page : paysage, ajusté à la largeur d'une page, titre répété en haut de chaque
    ''' page et pied de page numéroté. Protégée par un Try/Catch : certaines de ces propriétés
    ''' échouent lorsque aucune imprimante n'est installée sur le poste, ce qui ne doit pas
    ''' compromettre un export par ailleurs réussi.
    ''' </summary>
    Private Shared Sub PreparerImpression(feuille As Object, titre As String, hauteurBandeau As Integer)

        Try
            With feuille.PageSetup
                .Orientation = XL_PAYSAGE
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .PrintTitleRows = $"$1:${Math.Max(hauteurBandeau, 1)}"
                .CenterFooter = "Page &P / &N"
                .RightFooter = "Edite le " & Date.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture)
                .LeftFooter = If(titre, String.Empty)
            End With
        Catch
            ' Aucune imprimante configurée : l'export reste valable, seule la mise en page échoue.
        End Try
    End Sub

    ''' <summary>
    ''' Rend un nom d'onglet acceptable par Excel : 31 caractères au plus, sans les caractères
    ''' que le tableur interdit.
    ''' </summary>
    Private Shared Function NettoyerNomFeuille(nom As String) As String

        Dim propre As String = nom.Trim()

        For Each interdit As Char In New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
            propre = propre.Replace(interdit, " "c)
        Next

        If propre.Length > 31 Then propre = propre.Substring(0, 31)
        If propre.Length = 0 Then propre = "Feuil1"

        Return propre
    End Function

#End Region

End Class
