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
''' Une ligne d'information sous le titre principal. Elle peut être mise en exergue, pour
''' faire ressortir ce qui caractérise l'extraction — le groupe retenu, par exemple.
''' </summary>
Public Class SousTitreExcel

    Public Property Texte As String = String.Empty

    ''' <summary>Vrai pour un fond jaune : la ligne saute alors aux yeux sur un état imprimé.</summary>
    Public Property EnExergue As Boolean = False

    Public Sub New(texteSousTitre As String)
        Texte = texteSousTitre
    End Sub

    Public Sub New(texteSousTitre As String, exergue As Boolean)
        Texte = texteSousTitre
        EnExergue = exergue
    End Sub

End Class

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
    ''' Met le titre du bloc en exergue (fond jaune), au même titre que les sous-titres et les
    ''' lignes désignées par ExergueColonne / ExergueValeur.
    ''' </summary>
    Public Property TitreEnExergue As Boolean = False

    ''' <summary>
    ''' Nom de la colonne servant à repérer les lignes à mettre en exergue. Vide : aucune mise
    ''' en exergue. Utilisée avec ExergueValeur.
    ''' </summary>
    Public Property ExergueColonne As String = String.Empty

    ''' <summary>
    ''' Valeur recherchée dans ExergueColonne : toute ligne qui la porte reçoit un fond jaune.
    ''' Comparaison insensible à la casse.
    ''' </summary>
    Public Property ExergueValeur As String = String.Empty

    ''' <summary>
    ''' Pose un filtre automatique sur l'en-tête de ce bloc. Excel n'admettant qu'un seul filtre
    ''' par feuille, une seule valeur True est prise en compte — la première rencontrée.
    ''' </summary>
    Public Property AvecFiltre As Boolean = False

    ''' <summary>
    ''' Hauteur imposée aux lignes de données, en points. Zéro laisse la hauteur normale.
    '''
    ''' Elle n'existe que pour une chose : les cartouches de signature. Une case où l'on
    ''' doit écrire un nom et signer à la main a besoin de place, et la hauteur d'une ligne
    ''' de tableau ne suffit pas.
    ''' </summary>
    Public Property HauteurLignes As Double = 0R

    Public Sub New()
    End Sub

    Public Sub New(titreBloc As String, donneesBloc As DataTable)
        Titre = titreBloc
        Donnees = donneesBloc
    End Sub

End Class

''' <summary>
''' Export générique de tableaux vers Microsoft Excel ou PDF, avec titre, sous-titres, en-têtes
''' mis en forme et mise en page d'impression. Le PDF est produit par Excel à partir du même
''' classeur, qui n'est alors ni affiché ni enregistré : seul le PDF sort.
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

    ''' <summary>xlTypePDF : format d'export de ExportAsFixedFormat.</summary>
    Private Const XL_TYPE_PDF As Integer = 0

    ''' <summary>Jaune de mise en exergue, conforme au modèle fourni par la Direction Comptable.</summary>
    Private Shared ReadOnly JAUNE_EXERGUE As Integer = RGB(255, 255, 0)

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
                                       sousTitres As IEnumerable(Of SousTitreExcel),
                                       blocs As IEnumerable(Of BlocExcel),
                                       nomFeuille As String,
                                       cheminFichier As String,
                                       Optional progression As ProgressionWU = Nothing)

        Dim listeBlocs As List(Of BlocExcel) = BlocsExploitables(blocs)

        If listeBlocs.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter.")
        End If

        ' Ouverture d'Excel, un état par bloc, mise en page, enregistrement.
        If progression IsNot Nothing Then progression.Commencer(3 + listeBlocs.Count)

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing
        Dim feuille As Object = Nothing

        Try
            Annoncer(progression, "Ouverture de Microsoft Excel…")

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

            RemplirFeuille(feuille, titre, sousTitres, listeBlocs, nomFeuille, progression)

            Annoncer(progression, "Enregistrement du classeur…")
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

    ''' <summary>
    ''' Écrit les blocs dans un classeur Excel invisible, l'exporte en PDF, puis referme tout :
    ''' aucun fichier Excel intermédiaire n'est laissé sur le disque, et le classeur lui-même
    ''' n'est jamais enregistré.
    '''
    ''' L'intérêt du PDF est de figer l'état édité : il ne s'ouvre pas dans un tableur et ne se
    ''' retouche pas au fil de l'eau. Ce n'est pas pour autant un document infalsifiable — un PDF
    ''' reste modifiable avec l'outil adéquat. Pour une valeur probante, il faudrait le signer
    ''' électroniquement, ce qui relève d'un dispositif de la banque et non de cette application.
    '''
    ''' La mise en page d'impression déjà posée (paysage, ajusté à la largeur d'une page, bandeau
    ''' de titre répété, pied de page numéroté) est reprise telle quelle par l'export.
    ''' </summary>
    ''' <param name="cheminFichier">Chemin du fichier .pdf à créer.</param>
    ''' <param name="ouvrirApres">Ouvre le PDF dans le lecteur par défaut du poste.</param>
    ''' <exception cref="InvalidOperationException">Excel absent du poste, ou aucune donnée.</exception>
    Public Shared Sub ExporterEnPdf(titre As String,
                                    sousTitres As IEnumerable(Of SousTitreExcel),
                                    blocs As IEnumerable(Of BlocExcel),
                                    nomFeuille As String,
                                    cheminFichier As String,
                                    Optional ouvrirApres As Boolean = True,
                                    Optional progression As ProgressionWU = Nothing)

        Dim listeBlocs As List(Of BlocExcel) = BlocsExploitables(blocs)

        If listeBlocs.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter.")
        End If

        ' Ouverture d'Excel, un état par bloc, mise en page, conversion en PDF.
        If progression IsNot Nothing Then progression.Commencer(3 + listeBlocs.Count)

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing
        Dim feuille As Object = Nothing

        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            If typeExcel Is Nothing Then
                Throw New InvalidOperationException(
                    "Microsoft Excel n'est pas installé sur ce poste : l'export PDF est impossible." &
                    Environment.NewLine &
                    "Excel sert ici à la mise en page ; le rapport reste consultable à l'écran.")
            End If

            Annoncer(progression, "Ouverture de Microsoft Excel…")

            excelApp = Activator.CreateInstance(typeExcel)

            ' Classeur invisible : l'utilisateur ne doit jamais voir passer un tableur qu'il
            ' pourrait prendre pour le document livrable.
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            classeur = excelApp.Workbooks.Add()
            feuille = classeur.Worksheets(1)

            RemplirFeuille(feuille, titre, sousTitres, listeBlocs, nomFeuille, progression)

            Annoncer(progression, "Conversion en PDF…")

            ' xlTypePDF = 0. Le classeur n'est pas enregistré : seul le PDF sort.
            classeur.ExportAsFixedFormat(XL_TYPE_PDF, cheminFichier)

        Catch ex As Runtime.InteropServices.COMException
            Throw New InvalidOperationException(
                "Microsoft Excel n'a pas pu produire le PDF : " & ex.Message & Environment.NewLine & Environment.NewLine &
                "Vérifiez que le fichier n'est pas déjà ouvert dans un lecteur PDF, et que la " &
                "version d'Excel installée prend en charge l'export PDF (Excel 2007 et ultérieurs).", ex)

        Finally
            ' Contrairement à l'export Excel, rien ne doit rester ouvert : le livrable est le PDF.
            Try
                If classeur IsNot Nothing Then classeur.Close(False)
            Catch
            End Try
            Try
                If excelApp IsNot Nothing Then excelApp.Quit()
            Catch
            End Try

            If feuille IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(feuille)
            If classeur IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(classeur)
            If excelApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        End Try

        If ouvrirApres Then
            OuvrirDocument(cheminFichier)
        End If
    End Sub

    ''' <summary>
    ''' Ouvre un document dans l'application associée du poste. L'échec — aucun lecteur PDF
    ''' installé, par exemple — n'est pas remonté comme une erreur : le fichier est produit, et
    ''' c'est ce qui était demandé.
    ''' </summary>
#Region "Export brut : un tableau, sans habillage"

    ''' <summary>
    ''' Écrit une table telle quelle dans un classeur : les en-têtes en première ligne, les
    ''' données dessous, et rien d'autre.
    '''
    ''' C'est l'opposé de ExporterEtOuvrir, qui compose un état destiné à être lu par un humain.
    ''' Ici le lecteur est un automate : un titre, un sous-titre ou une ligne vide décaleraient
    ''' les colonnes et feraient rejeter le fichier.
    '''
    ''' Toutes les colonnes sont mises au format Texte avant écriture, sauf celles désignées
    ''' comme numériques. Sans cela Excel réinterprète ce qu'il croit reconnaître : un numéro de
    ''' compte perdrait ses zéros de tête, et un numéro de lot comme « 07p1 » resterait du texte
    ''' quand « 0741 » deviendrait le nombre 741.
    ''' </summary>
    ''' <param name="table">Données à écrire, en-têtes compris.</param>
    ''' <param name="colonnesNumeriques">Noms des colonnes à écrire en nombres.</param>
    ''' <param name="nomFeuille">Nom de la feuille.</param>
    ''' <param name="cheminFichier">Chemin complet du classeur à créer.</param>
    ''' <param name="ouvrirApres">
    ''' Vrai pour présenter le classeur à l'utilisateur une fois écrit, comme le fait la pièce
    ''' comptable. Un fichier produit sans être montré laisse toujours un doute sur son contenu.
    ''' </param>
    Public Shared Sub ExporterTableBrute(table As DataTable,
                                         colonnesNumeriques As IEnumerable(Of String),
                                         nomFeuille As String,
                                         cheminFichier As String,
                                         ouvrirApres As Boolean,
                                         Optional progression As ProgressionWU = Nothing)

        If table Is Nothing OrElse table.Rows.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter.")
        End If

        Dim numeriques As New List(Of String)
        If colonnesNumeriques IsNot Nothing Then
            For Each nom As String In colonnesNumeriques
                numeriques.Add(nom.ToUpperInvariant())
            Next
        End If

        ' Ouverture d'Excel, écriture du tableau, enregistrement.
        If progression IsNot Nothing Then progression.Commencer(3)

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing
        Dim feuille As Object = Nothing

        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            If typeExcel Is Nothing Then
                Throw New InvalidOperationException(
                    "Microsoft Excel n'est pas installé sur ce poste : le fichier ne peut pas être produit.")
            End If

            Annoncer(progression, "Ouverture de Microsoft Excel…")

            excelApp = Activator.CreateInstance(typeExcel)
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            classeur = excelApp.Workbooks.Add()
            feuille = classeur.Worksheets(1)
            feuille.Name = NettoyerNomFeuille(nomFeuille)

            Dim nbColonnes As Integer = table.Columns.Count
            Dim nbLignes As Integer = table.Rows.Count

            Annoncer(progression, $"Écriture de {nbLignes:N0} ligne(s)…")

            ' Le format des colonnes est posé AVANT l'écriture : appliqué après, Excel aurait
            ' déjà converti les valeurs, et reformater n'aurait rien rendu.
            For index As Integer = 1 To nbColonnes
                Dim estNumerique As Boolean = numeriques.Contains(table.Columns(index - 1).ColumnName.ToUpperInvariant())
                feuille.Columns(index).NumberFormat = If(estNumerique, "0", "@")
            Next

            ' Écriture en un seul bloc : cellule par cellule, chaque affectation est un appel COM,
            ' et une pièce de plusieurs milliers de lignes prendrait plusieurs minutes.
            Dim valeurs(nbLignes, nbColonnes - 1) As Object

            For colonne As Integer = 0 To nbColonnes - 1
                valeurs(0, colonne) = table.Columns(colonne).ColumnName
            Next

            For ligne As Integer = 0 To nbLignes - 1
                For colonne As Integer = 0 To nbColonnes - 1

                    Dim brut As String = Convert.ToString(table.Rows(ligne)(colonne))

                    If numeriques.Contains(table.Columns(colonne).ColumnName.ToUpperInvariant()) Then
                        Dim nombre As Long
                        valeurs(ligne + 1, colonne) = If(Long.TryParse(brut, nombre), CObj(nombre), CObj(brut))
                    Else
                        valeurs(ligne + 1, colonne) = brut
                    End If
                Next
            Next

            Dim plage As Object = feuille.Range(feuille.Cells(1, 1), feuille.Cells(nbLignes + 1, nbColonnes))
            plage.Value = valeurs

            feuille.Range(feuille.Cells(1, 1), feuille.Cells(1, nbColonnes)).Font.Bold = True
            feuille.Columns.AutoFit()
            feuille.Range("A1").Select()

            Annoncer(progression, "Enregistrement du classeur…")

            classeur.SaveAs(cheminFichier)

        Finally
            If classeur IsNot Nothing Then
                Try
                    classeur.Close(False)
                Catch ex As Runtime.InteropServices.COMException
                End Try
            End If

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch ex As Runtime.InteropServices.COMException
                End Try
            End If

            LibererObjet(feuille)
            LibererObjet(classeur)
            LibererObjet(excelApp)
        End Try

        ' L'ouverture se fait APRÈS la libération : l'instance d'Excel qui a écrit le fichier
        ' travaillait masquée, et c'est celle de l'utilisateur qui doit le présenter.
        If ouvrirApres Then OuvrirDocument(cheminFichier)
    End Sub

    ''' <summary>
    ''' Relâche une référence COM. Sans cela, une instance d'Excel resterait en mémoire après
    ''' chaque export, invisible et jamais fermée.
    ''' </summary>
    Private Shared Sub LibererObjet(objet As Object)

        If objet Is Nothing Then Return

        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(objet)
        Catch ex As ArgumentException
            ' L'objet n'était pas une référence COM : rien à relâcher.
        End Try
    End Sub

#End Region

    Private Shared Sub OuvrirDocument(cheminFichier As String)

        Try
            Process.Start(cheminFichier)
        Catch ex As ComponentModel.Win32Exception
            ' Aucune application associée à ce type de fichier sur ce poste.
        Catch ex As IO.FileNotFoundException
            ' Fichier disparu entre-temps : sans conséquence sur l'export lui-même.
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
                                      sousTitres As IEnumerable(Of SousTitreExcel),
                                      blocs As List(Of BlocExcel), nomFeuille As String,
                                      Optional progression As ProgressionWU = Nothing)

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
            Annoncer(progression, bloc.Titre)
            ligne = EcrireBloc(feuille, bloc, ligne, filtrePose)
            ligne += 1 ' ligne vide entre deux tableaux
        Next

        Annoncer(progression, "Mise en page…")

        feuille.Columns.AutoFit()
        PreparerImpression(feuille, titre, hauteurBandeau)
    End Sub

    ''' <summary>Écrit le bandeau de titre et les sous-titres. Retourne la première ligne libre.</summary>
    Private Shared Function EcrireTitre(feuille As Object, titre As String,
                                        sousTitres As IEnumerable(Of SousTitreExcel),
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
            For Each sousTitre As SousTitreExcel In sousTitres

                If sousTitre Is Nothing OrElse String.IsNullOrWhiteSpace(sousTitre.Texte) Then Continue For

                Dim plage As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, largeur))
                feuille.Cells(ligne, 1).Value = sousTitre.Texte

                plage.Merge()
                plage.HorizontalAlignment = XL_GAUCHE
                plage.Font.Italic = True

                If sousTitre.EnExergue Then
                    plage.Interior.Color = JAUNE_EXERGUE
                    plage.Font.Bold = True
                End If

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

            If bloc.TitreEnExergue Then
                plageTitre.Interior.Color = JAUNE_EXERGUE
            End If

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

        ' Lignes à mettre en exergue, repérées pendant l'écriture et mises en forme ensuite :
        ' un seul aller-retour de mise en forme par ligne concernée, plutôt qu'un par cellule.
        Dim lignesEnExergue As New List(Of Integer)

        ' La variable ne doit pas porter le nom de la fonction : VB étant insensible à la casse,
        ' une locale nommée « indexColonneExergue » masquerait IndexColonneExergue et l'appel
        ' serait lu comme une indexation de cette locale.
        Dim positionExergue As Integer = IndexColonneExergue(bloc)

        For Each enregistrement As DataRow In bloc.Donnees.Rows

            If positionExergue >= 0 AndAlso
               String.Equals(Convert.ToString(enregistrement(positionExergue)),
                             bloc.ExergueValeur, StringComparison.OrdinalIgnoreCase) Then
                lignesEnExergue.Add(ligne)
            End If

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

            ' Hauteur imposée, s'il y en a une : c'est l'espace où l'on signe.
            If bloc.HauteurLignes > 0R Then
                feuille.Range(feuille.Cells(premiereLigneDonnees, 1),
                              feuille.Cells(derniereLigne, nombreColonnes)).RowHeight = bloc.HauteurLignes
            End If

            ' --- Formats de nombre ---
            For c As Integer = 0 To nombreColonnes - 1
                Dim nom As String = bloc.Donnees.Columns(c).ColumnName
                If bloc.Formats Is Nothing OrElse Not bloc.Formats.ContainsKey(nom) Then Continue For

                feuille.Range(feuille.Cells(premiereLigneDonnees, c + 1),
                              feuille.Cells(derniereLigne, c + 1)).NumberFormat = bloc.Formats(nom)
            Next

            ' --- Mise en exergue ---
            ' Appliquée APRÈS les formats de nombre : ceux-ci portent sur des colonnes entières
            ' et effaceraient sinon le fond des lignes concernées.
            For Each ligneExergue As Integer In lignesEnExergue
                Dim plage As Object = feuille.Range(feuille.Cells(ligneExergue, 1),
                                                    feuille.Cells(ligneExergue, nombreColonnes))
                plage.Interior.Color = JAUNE_EXERGUE
                plage.Font.Bold = True
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
    ''' Position, dans le tableau, de la colonne servant à repérer les lignes en exergue.
    ''' Retourne -1 si aucune mise en exergue n'est demandée, ou si la colonne nommée n'existe
    ''' pas : une demande portant sur une colonne absente est ignorée, jamais bloquante.
    ''' </summary>
    Private Shared Function IndexColonneExergue(bloc As BlocExcel) As Integer

        If String.IsNullOrWhiteSpace(bloc.ExergueColonne) OrElse String.IsNullOrWhiteSpace(bloc.ExergueValeur) Then
            Return -1
        End If

        If Not bloc.Donnees.Columns.Contains(bloc.ExergueColonne) Then
            Return -1
        End If

        Return bloc.Donnees.Columns(bloc.ExergueColonne).Ordinal
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
    ''' <summary>Annonce une étape, s'il y a quelqu'un pour l'entendre.</summary>
    Private Shared Sub Annoncer(progression As ProgressionWU, libelle As String)

        If progression Is Nothing Then Return
        progression.Avancer(libelle)
    End Sub

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
