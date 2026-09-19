' ATTENTION : Option Strict Off est INDISPENSABLE dans ce fichier, et nulle part ailleurs.
'
' Excel est piloté par liaison tardive (Type.GetTypeFromProgID), seul moyen de ne pas imposer
' une référence à une version précise d'Office sur les postes de la banque. Toutes les
' expressions du genre feuille.Cells(1, 1).Value sont donc résolues à l'exécution, ce que
' Option Strict On interdit.
Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

''' <summary>
''' Écrit la pièce comptable dans le formulaire de la banque — celui qui est visé et signé à
''' la main — avec une feuille par point de vente en plus de la pièce globale.
'''
''' POURQUOI UNE FEUILLE PAR POINT DE VENTE
''' La pièce globale équilibre la journée entière. Mais c'est point de vente par point de
''' vente que la comptabilité contrôle, que le sous-agent conteste, et que l'inspection
''' remonte. Extraire ces pièces à la main d'un tableau de trois cents lignes est un travail
''' de recopie, et la recopie se trompe.
'''
''' POURQUOI CE FORMULAIRE, ET PAS UN TABLEAU
''' Le modèle fourni par la banque (classe_bis.xlsx) n'est pas une présentation : c'est le
''' document que le guichet de la comptabilité accepte ou refuse. En-tête, blocs DÉBIT et
''' CRÉDIT, RAISON, puis les quatre cartouches de signature. Une pièce qui ne lui ressemble
''' pas se fait renvoyer, quelles que soient ses écritures.
'''
''' CE QUI DIFFÈRE DU MODÈLE, ET POURQUOI
''' Le modèle réserve un nombre fixe de lignes (25 environ) parce qu'il est fait pour être
''' rempli à la main. Une pièce générée en a autant que la journée en produit : les blocs
''' grandissent, et tout ce qui suit descend d'autant. La géométrie des colonnes, les polices,
''' les hauteurs de ligne et l'enchaînement des cartouches, eux, sont repris tels quels.
''' </summary>
Public NotInheritable Class PieceExcelWU

    Private Sub New()
    End Sub

#Region "Constantes Excel"

    ' Valeurs de l'énumération Excel, écrites en clair : la liaison tardive ne donne pas accès
    ' aux constantes nommées d'Office.
    Private Const XL_CONTINU As Integer = 1
    Private Const XL_FIN As Integer = 2        ' xlThin
    Private Const XL_BORD_GAUCHE As Integer = 7
    Private Const XL_BORD_HAUT As Integer = 8
    Private Const XL_BORD_BAS As Integer = 9
    Private Const XL_BORD_DROIT As Integer = 10
    Private Const XL_INTERIEUR_VERTICAL As Integer = 11
    Private Const XL_INTERIEUR_HORIZONTAL As Integer = 12
    Private Const XL_GAUCHE As Integer = -4131
    Private Const XL_CENTRE As Integer = -4108
    Private Const XL_DROITE As Integer = -4152
    Private Const XL_MAXIMISE As Integer = -4137
    Private Const XL_PAYSAGE As Integer = 2

    ''' <summary>Format des montants : séparateur de milliers, pas de décimale — le FCFA n'en a pas.</summary>
    Private Const FORMAT_MONTANT As String = "#,##0"

#End Region

#Region "Le contexte d'une feuille"

    ''' <summary>
    ''' Ce qui distingue une feuille d'une autre : son en-tête. Les écritures, elles, viennent
    ''' de la DataTable.
    ''' </summary>
    Public NotInheritable Class ContexteFeuille

        ''' <summary>Nom de l'onglet.</summary>
        Public Property NomFeuille As String = "PIECE"

        ''' <summary>Ligne d'identification, au-dessus du tableau.</summary>
        Public Property Intitule As String = String.Empty

        ''' <summary>Journée comptabilisée — et non la date d'impression.</summary>
        Public Property DateActivite As Date = Date.Today

        ''' <summary>Numéro d'ordre de la pièce dans le classeur.</summary>
        Public Property Numero As Integer = 1

        ''' <summary>Texte de la ligne RAISON.</summary>
        Public Property Raison As String = String.Empty
    End Class

#End Region

#Region "Écriture du classeur"

    ''' <summary>
    ''' Écrit le classeur complet : la pièce globale en première feuille, puis une feuille par
    ''' point de vente comptabilisé. Excel reste ouvert sur le résultat.
    ''' </summary>
    ''' <param name="dtGlobale">Pièce comptable de la journée entière.</param>
    ''' <param name="listeCalculs">Points de vente de la journée. Nothing pour n'écrire que la pièce globale.</param>
    ''' <param name="dateActivite">Journée comptabilisée.</param>
    ''' <param name="cheminFichier">Chemin complet du .xlsx à produire.</param>
    ''' <param name="nomPremiereFeuille">Onglet de la première feuille.</param>
    ''' <param name="intitulePremiereFeuille">Ligne d'identification de la première feuille.
    ''' Vide pour l'intitulé de la pièce globale.</param>
    ''' <returns>Le chemin écrit.</returns>
    Public Shared Function Ecrire(dtGlobale As DataTable,
                                  listeCalculs As IEnumerable(Of CalculWU),
                                  dateActivite As Date,
                                  cheminFichier As String,
                                  Optional nomPremiereFeuille As String = "PIECE GLOBALE",
                                  Optional intitulePremiereFeuille As String = "") As String

        If dtGlobale Is Nothing OrElse dtGlobale.Rows.Count = 0 Then
            Throw New InvalidOperationException(
                "Aucune donnée à exporter : générez la pièce comptable au préalable.")
        End If

        If String.IsNullOrWhiteSpace(cheminFichier) Then
            Throw New ArgumentException("Chemin de fichier non renseigné.", NameOf(cheminFichier))
        End If

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing

        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            If typeExcel Is Nothing Then
                Throw New InvalidOperationException("Microsoft Excel n'est pas installé sur ce poste.")
            End If

            excelApp = Activator.CreateInstance(typeExcel)
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' L'affichage est coupé le temps de l'écriture : sur cinquante feuilles, laisser
            ' Excel redessiner chaque cellule multiplie la durée par cinq ou six.
            excelApp.ScreenUpdating = False

            classeur = excelApp.Workbooks.Add()
            NeGarderQueLaPremiereFeuille(classeur)

            EcrireLaPremiereFeuille(classeur, dtGlobale, dateActivite,
                                    nomPremiereFeuille, intitulePremiereFeuille)
            EcrireLesPiecesIndividuelles(classeur, listeCalculs, dateActivite)

            ' La première feuille est celle qu'on veut voir en ouvrant le classeur.
            classeur.Worksheets(1).Activate()

            classeur.SaveAs(cheminFichier)

            excelApp.ScreenUpdating = True
            excelApp.Visible = True

            Try
                excelApp.WindowState = XL_MAXIMISE
                classeur.Activate()
            Catch
                ' Mise au premier plan refusée par Windows : le classeur est ouvert, c'est
                ' l'essentiel. Voir la même remarque dans PieceComptableService.
            End Try

            Return cheminFichier

        Catch
            ' Un classeur incomplet ne doit pas rester à l'écran : on referme avant de relancer.
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
            If classeur IsNot Nothing Then Marshal.ReleaseComObject(classeur)
        End Try
    End Function

    ''' <summary>
    ''' Supprime les feuilles vides d'un classeur neuf.
    '''
    ''' Excel en crée une, deux ou trois selon le réglage du poste. Les laisser donnerait un
    ''' classeur où les pièces sont noyées entre des « Feuil2 » vides — et le comptable qui
    ''' imprime tout le classeur sortirait des pages blanches au milieu de ses pièces.
    ''' </summary>
    Private Shared Sub NeGarderQueLaPremiereFeuille(classeur As Object)

        While CInt(classeur.Worksheets.Count) > 1

            Dim superflue As Object = classeur.Worksheets(CInt(classeur.Worksheets.Count))
            Try
                superflue.Delete()
            Finally
                Marshal.ReleaseComObject(superflue)
            End Try
        End While
    End Sub

    Private Shared Sub EcrireLaPremiereFeuille(classeur As Object, dtGlobale As DataTable,
                                               dateActivite As Date,
                                               nomFeuille As String, intitule As String)

        Dim contexte As New ContexteFeuille() With {
            .NomFeuille = If(String.IsNullOrWhiteSpace(nomFeuille), "PIECE GLOBALE", nomFeuille),
            .Intitule = If(String.IsNullOrWhiteSpace(intitule),
                           $"PIECE GLOBALE — journée du {dateActivite:dd/MM/yyyy}",
                           intitule),
            .DateActivite = dateActivite,
            .Numero = 1,
            .Raison = $"Compensation Western Union — activité du {dateActivite:dd/MM/yyyy}"
        }

        Dim feuille As Object = classeur.Worksheets(1)
        Try
            EcrireFeuille(feuille, dtGlobale, contexte)
        Finally
            Marshal.ReleaseComObject(feuille)
        End Try
    End Sub

    ''' <summary>
    ''' Une feuille par point de vente comptabilisé. La pièce de chacun est REGÉNÉRÉE à partir
    ''' de son seul CalculWU, et non découpée dans la pièce globale : c'est le même code qui
    ''' produit les deux, donc les mêmes écritures, aux mêmes comptes, par construction.
    ''' </summary>
    Private Shared Sub EcrireLesPiecesIndividuelles(classeur As Object,
                                                    listeCalculs As IEnumerable(Of CalculWU),
                                                    dateActivite As Date)

        If listeCalculs Is Nothing Then Return

        ' Les onglets déjà posés, pour ne pas en produire deux du même nom — Excel refuse
        ' alors d'écrire le classeur entier, après vingt feuilles déjà remplies.
        Dim nomsPris As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For indice As Integer = 1 To CInt(classeur.Worksheets.Count)
            nomsPris.Add(CStr(classeur.Worksheets(indice).Name))
        Next
        Dim numero As Integer = 1

        For Each calc As CalculWU In listeCalculs

            ' Un Account non comptabilisé n'a pas de pièce — il n'a pas d'écriture non plus.
            ' Lui donner un onglet vide laisserait croire à une pièce à zéro.
            If calc Is Nothing OrElse Not calc.EstComptabilisable Then Continue For

            Dim dtPdv As DataTable = PieceComptableService.GenererPieceComptable(New CalculWU() {calc})
            If dtPdv Is Nothing OrElse dtPdv.Rows.Count = 0 Then Continue For

            numero += 1

            Dim contexte As New ContexteFeuille() With {
                .NomFeuille = NomDOnglet(calc, nomsPris),
                .Intitule = IntituleDe(calc),
                .DateActivite = dateActivite,
                .Numero = numero,
                .Raison = $"Compensation Western Union — {calc.Designation} ({calc.Account}) — " &
                          $"activité du {dateActivite:dd/MM/yyyy}"
            }

            ' Ajoutée APRÈS la dernière : sans cela les onglets sortiraient à l'envers, et
            ' l'ordre des pièces est celui de la grille de contrôle.
            Dim feuille As Object = classeur.Worksheets.Add(After:=classeur.Worksheets(classeur.Worksheets.Count))
            Try
                EcrireFeuille(feuille, dtPdv, contexte)
            Finally
                Marshal.ReleaseComObject(feuille)
            End Try
        Next
    End Sub

#End Region

#Region "Mise en page d'une feuille"

    ''' <summary>
    ''' Pose le formulaire de la banque sur une feuille : en-tête, bloc DÉBIT, bloc CRÉDIT,
    ''' RAISON, puis les cartouches de signature.
    ''' </summary>
    Private Shared Sub EcrireFeuille(feuille As Object, dt As DataTable, contexte As ContexteFeuille)

        feuille.Name = contexte.NomFeuille

        LargeursEtPolice(feuille)
        EnTete(feuille, contexte)

        ' Le tableau commence sous ses en-têtes, ligne 13 comme dans le modèle.
        Dim premiereLigne As Integer = 14
        Dim ligne As Integer = premiereLigne

        Dim debits As List(Of DataRow) = Lignes(dt, "Debit")
        Dim credits As List(Of DataRow) = Lignes(dt, "Credit")

        Dim premierDebit As Integer = ligne
        ligne = EcrireBloc(feuille, debits, "Debit", ligne, ConstantesWU.PIECE_DEBIT)

        Dim premierCredit As Integer = ligne
        ligne = EcrireBloc(feuille, credits, "Credit", ligne, ConstantesWU.PIECE_CREDIT)

        Dim derniereLigne As Integer = ligne - 1

        Encadrer(feuille, premiereLigne, derniereLigne)
        Controle(feuille, premierDebit, debits.Count, premierCredit, credits.Count)

        Dim ligneRaison As Integer = derniereLigne + 1
        Raison(feuille, ligneRaison, contexte.Raison)

        Cartouches(feuille, ligneRaison + 2)
        MiseEnPageImpression(feuille)
    End Sub

    ''' <summary>Colonnes et police du modèle. Les largeurs sont celles du classeur fourni.</summary>
    Private Shared Sub LargeursEtPolice(feuille As Object)

        feuille.Cells.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
        feuille.Cells.Font.Size = 12

        feuille.Columns("A").ColumnWidth = 24.86
        feuille.Columns("B").ColumnWidth = 45.86

        ' 114 dans le modèle : le libellé d'une écriture WU est long, et le comptable lit la
        ' pièce sans élargir la colonne.
        feuille.Columns("C").ColumnWidth = 114.57
        feuille.Columns("D").ColumnWidth = 48.14
    End Sub

    Private Shared Sub EnTete(feuille As Object, contexte As ContexteFeuille)

        Ecrire(feuille, 4, 1, ConstantesWU.PIECE_BANQUE, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_GAUCHE)
        feuille.Rows(4).RowHeight = 20.25

        Ecrire(feuille, 6, 2, ConstantesWU.PIECE_TITRE, ConstantesWU.PIECE_POLICE_CORPS, 20, False, XL_GAUCHE)
        feuille.Rows(6).RowHeight = 39.75

        Ecrire(feuille, 9, 1, ConstantesWU.PIECE_DATE, ConstantesWU.PIECE_POLICE_TITRE, 14, True, XL_GAUCHE)

        ' La date est celle de la JOURNÉE COMPTABILISÉE, pas celle de l'impression. Le modèle
        ' porte =TODAY() parce qu'il est vierge ; une pièce rejouée trois jours plus tard doit
        ' rester datée du jour qu'elle comptabilise.
        Dim celluleDate As Object = feuille.Cells(9, 2)
        Try
            celluleDate.Value = contexte.DateActivite
            celluleDate.NumberFormat = "dd/mm/yyyy"
            celluleDate.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            celluleDate.Font.Size = 24
        Finally
            Marshal.ReleaseComObject(celluleDate)
        End Try
        feuille.Rows(9).RowHeight = 30.0

        Ecrire(feuille, 10, 1, ConstantesWU.PIECE_DE, ConstantesWU.PIECE_POLICE_TITRE, 14, True, XL_GAUCHE)
        Ecrire(feuille, 10, 2, ConstantesWU.PIECE_SERVICE_EMETTEUR, ConstantesWU.PIECE_POLICE_CORPS, 14, False, XL_GAUCHE)
        Ecrire(feuille, 10, 3, ConstantesWU.PIECE_AGENCE, ConstantesWU.PIECE_POLICE_TITRE, 24, True, XL_DROITE)
        Ecrire(feuille, 10, 4, ConstantesWU.PIECE_VILLE, ConstantesWU.PIECE_POLICE_TITRE, 28, True, XL_DROITE)
        feuille.Rows(10).RowHeight = 42.75

        Ecrire(feuille, 11, 1, ConstantesWU.PIECE_POUR, ConstantesWU.PIECE_POLICE_TITRE, 14, True, XL_GAUCHE)
        Ecrire(feuille, 11, 2, ConstantesWU.PIECE_SERVICE_DESTINATAIRE, ConstantesWU.PIECE_POLICE_CORPS, 14, False, XL_GAUCHE)
        Ecrire(feuille, 11, 4, contexte.Numero, ConstantesWU.PIECE_POLICE_CORPS, 26, True, XL_DROITE)
        feuille.Rows(11).RowHeight = 33.0

        Ecrire(feuille, 12, 3, contexte.Intitule, ConstantesWU.PIECE_POLICE_CORPS, 26, True, XL_CENTRE)
        feuille.Rows(12).RowHeight = 33.75

        Ecrire(feuille, 13, 2, ConstantesWU.PIECE_COMPTES, ConstantesWU.PIECE_POLICE_TITRE, 16, True, XL_CENTRE)
        Ecrire(feuille, 13, 3, ConstantesWU.PIECE_LIBELLES, ConstantesWU.PIECE_POLICE_TITRE, 16, True, XL_CENTRE)
        Ecrire(feuille, 13, 4, ConstantesWU.PIECE_MONTANTS, ConstantesWU.PIECE_POLICE_TITRE, 16, True, XL_DROITE)
        feuille.Rows(13).RowHeight = 31.5

        Dim entetes As Object = feuille.Range("B13:D13")
        Try
            entetes.Borders.LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(entetes)
        End Try
    End Sub

    ''' <summary>
    ''' Écrit un bloc (débits ou crédits) et rend la première ligne libre après lui.
    ''' L'intitulé du bloc est posé en colonne A, en face de sa première écriture.
    ''' </summary>
    ''' <remarks>
    ''' Le paramètre s'appelle « ecritures » et non « lignes » : VB ne distingue pas la casse,
    ''' et « lignes » masquerait la fonction Lignes() de cette même classe.
    ''' </remarks>
    Private Shared Function EcrireBloc(feuille As Object, ecritures As List(Of DataRow), colonne As String,
                                       premiereLigne As Integer, intitule As String) As Integer

        If ecritures.Count = 0 Then Return premiereLigne

        Ecrire(feuille, premiereLigne, 1, intitule, ConstantesWU.PIECE_POLICE_TITRE, 20, True, XL_CENTRE)

        ' Le tableau part en un seul bloc plutôt que cellule par cellule : chaque appel COM
        ' coûte un aller-retour, et une journée de trois cents écritures en ferait neuf cents.
        Dim valeurs(ecritures.Count - 1, 2) As Object

        For index As Integer = 0 To ecritures.Count - 1
            valeurs(index, 0) = Convert.ToString(ecritures(index)("Compte"))
            valeurs(index, 1) = Convert.ToString(ecritures(index)("Libelle"))
            valeurs(index, 2) = Convert.ToInt64(ecritures(index)(colonne))
        Next

        Dim derniere As Integer = premiereLigne + ecritures.Count - 1
        Dim zone As Object = feuille.Range(feuille.Cells(premiereLigne, 2), feuille.Cells(derniere, 4))

        Try
            zone.Value = valeurs
            zone.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            zone.Font.Size = 16
            zone.RowHeight = 33.0
        Finally
            Marshal.ReleaseComObject(zone)
        End Try

        Dim montants As Object = feuille.Range(feuille.Cells(premiereLigne, 4), feuille.Cells(derniere, 4))
        Try
            montants.NumberFormat = FORMAT_MONTANT
            montants.HorizontalAlignment = XL_DROITE
        Finally
            Marshal.ReleaseComObject(montants)
        End Try

        Return derniere + 1
    End Function

    ''' <summary>Quadrille le tableau, et ferme sa colonne A par un trait à gauche comme le modèle.</summary>
    Private Shared Sub Encadrer(feuille As Object, premiereLigne As Integer, derniereLigne As Integer)

        If derniereLigne < premiereLigne Then Return

        Dim corps As Object = feuille.Range(feuille.Cells(premiereLigne, 2), feuille.Cells(derniereLigne, 4))
        Try
            corps.Borders.LineStyle = XL_CONTINU
            corps.Borders.Weight = XL_FIN
        Finally
            Marshal.ReleaseComObject(corps)
        End Try

        Dim marge As Object = feuille.Range(feuille.Cells(premiereLigne, 1), feuille.Cells(derniereLigne + 1, 1))
        Try
            marge.Borders(XL_BORD_GAUCHE).LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(marge)
        End Try
    End Sub

    ''' <summary>
    ''' La case de contrôle du modèle, en D6 : total des débits moins total des crédits.
    '''
    ''' Le modèle y porte une soustraction écrite à la main, ligne par ligne. Une formule
    ''' sur les deux plages dit la même chose et survit à l'ajout d'une écriture — c'est
    ''' une case de contrôle, elle doit rester vraie quand la pièce change.
    ''' </summary>
    Private Shared Sub Controle(feuille As Object, premierDebit As Integer, nbDebits As Integer,
                                premierCredit As Integer, nbCredits As Integer)

        Dim formule As String = "="

        If nbDebits > 0 Then
            formule &= $"SUM(D{premierDebit}:D{premierDebit + nbDebits - 1})"
        Else
            formule &= "0"
        End If

        If nbCredits > 0 Then
            formule &= $"-SUM(D{premierCredit}:D{premierCredit + nbCredits - 1})"
        End If

        Dim cellule As Object = feuille.Cells(6, 4)
        Try
            cellule.Formula = formule
            cellule.NumberFormat = FORMAT_MONTANT
            cellule.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            cellule.Font.Size = 20
            cellule.Font.Bold = True
            cellule.HorizontalAlignment = XL_DROITE
            cellule.Borders.LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(cellule)
        End Try
    End Sub

    Private Shared Sub Raison(feuille As Object, ligne As Integer, texte As String)

        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_RAISON, ConstantesWU.PIECE_POLICE_TITRE, 20, True, XL_CENTRE)

        Dim zone As Object = feuille.Range(feuille.Cells(ligne, 2), feuille.Cells(ligne, 4))
        Try
            zone.Merge()
            zone.Value = texte
            zone.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            zone.Font.Size = 24
            zone.Font.Bold = True
            zone.HorizontalAlignment = XL_CENTRE
            zone.Borders.LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(zone)
        End Try

        feuille.Rows(ligne).RowHeight = 32.25
    End Sub

    ''' <summary>
    ''' Les quatre cartouches de signature, dans l'ordre et aux écartements du modèle : quatre
    ''' lignes vides entre chacun, parce que c'est l'espace où l'on signe.
    ''' </summary>
    Private Shared Sub Cartouches(feuille As Object, premiereLigne As Integer)

        Dim ligne As Integer = premiereLigne + 3

        Souligner(feuille, ligne, ConstantesWU.PIECE_SIGNATURES, ConstantesWU.PIECE_FCU)
        ligne += 5
        TroisIntitules(feuille, ligne, ConstantesWU.PIECE_INITIE, ConstantesWU.PIECE_CONTROLE, ConstantesWU.PIECE_APPROUVE)
        ligne += 5
        Souligner(feuille, ligne, ConstantesWU.PIECE_ECRITURE, ConstantesWU.PIECE_OPS)
        ligne += 5
        TroisIntitules(feuille, ligne, ConstantesWU.PIECE_PASSEE, String.Empty, ConstantesWU.PIECE_AUTORISEE)
        ligne += 4
        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_DATE_ENREGISTREMENT, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_GAUCHE)
        ligne += 4
        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_NUMERO_SEQUENCE, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_GAUCHE)

        Dim pied As Object = feuille.Range(feuille.Cells(premiereLigne, 1), feuille.Cells(ligne, 4))
        Try
            pied.RowHeight = 20.25
        Finally
            Marshal.ReleaseComObject(pied)
        End Try
    End Sub

    ''' <summary>Un cartouche à trait continu : intitulé à gauche, service à droite.</summary>
    Private Shared Sub Souligner(feuille As Object, ligne As Integer, gauche As String, droite As String)

        Ecrire(feuille, ligne, 1, gauche, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_GAUCHE)
        Ecrire(feuille, ligne, 4, droite, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_DROITE)

        Dim zone As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, 4))
        Try
            zone.Borders(XL_BORD_BAS).LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(zone)
        End Try
    End Sub

    Private Shared Sub TroisIntitules(feuille As Object, ligne As Integer,
                                      gauche As String, milieu As String, droite As String)

        Ecrire(feuille, ligne, 1, gauche, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_GAUCHE)
        If milieu.Length > 0 Then Ecrire(feuille, ligne, 2, milieu, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_DROITE)
        Ecrire(feuille, ligne, 3, droite, ConstantesWU.PIECE_POLICE_CORPS, 16, True, XL_CENTRE)
    End Sub

    ''' <summary>
    ''' Une pièce se signe sur papier : elle doit tenir en largeur, sous peine de ressortir de
    ''' l'imprimante en deux morceaux qu'il faut ensuite raccorder à la main.
    ''' </summary>
    Private Shared Sub MiseEnPageImpression(feuille As Object)

        Try
            feuille.PageSetup.Orientation = XL_PAYSAGE
            feuille.PageSetup.Zoom = False
            feuille.PageSetup.FitToPagesWide = 1
            feuille.PageSetup.FitToPagesTall = False
        Catch
            ' Aucune imprimante installée : Excel refuse alors PageSetup. La pièce est écrite,
            ' et c'est ce qui compte — l'échec ne doit pas faire perdre le classeur.
        End Try
    End Sub

#End Region

#Region "Utilitaires"

    ''' <summary>Écrit une valeur et l'habille. Ignore les valeurs vides.</summary>
    Private Shared Sub Ecrire(feuille As Object, ligne As Integer, colonne As Integer, valeur As Object,
                              police As String, taille As Integer, gras As Boolean, alignement As Integer)

        If valeur Is Nothing Then Return
        If TypeOf valeur Is String AndAlso CStr(valeur).Length = 0 Then Return

        Dim cellule As Object = feuille.Cells(ligne, colonne)
        Try
            cellule.Value = valeur
            cellule.Font.Name = police
            cellule.Font.Size = taille
            cellule.Font.Bold = gras
            cellule.HorizontalAlignment = alignement
        Finally
            Marshal.ReleaseComObject(cellule)
        End Try
    End Sub

    ''' <summary>Les lignes de la pièce dont la colonne indiquée porte un montant non nul.</summary>
    Private Shared Function Lignes(dt As DataTable, colonne As String) As List(Of DataRow)

        Dim retenues As New List(Of DataRow)()

        For Each ligne As DataRow In dt.Rows
            If Convert.ToInt64(ligne(colonne)) <> 0L Then retenues.Add(ligne)
        Next

        Return retenues
    End Function

    ''' <summary>
    ''' Ligne d'identification d'une pièce de point de vente, à la forme du modèle :
    ''' « Agence  001: SS-AGENCE_BOLOLO SIEGE ».
    '''
    ''' Le code retenu est celui de l'agence de rattachement, celui-là même qui alimente la
    ''' colonne ACBRN du fichier destiné au core banking : une pièce et l'écriture qu'elle
    ''' justifie doivent porter la même agence, sans quoi le rapprochement est à refaire.
    ''' </summary>
    Public Shared Function IntituleDe(calc As CalculWU) As String

        Dim code As String = If(String.IsNullOrWhiteSpace(calc.CodeAgence), calc.Account, calc.CodeAgence)
        Return $"Agence  {code}: {calc.Designation}".Trim()
    End Function

    ''' <summary>
    ''' Nom d'onglet acceptable par Excel : 31 caractères au plus, sans \ / ? * [ ] : et
    ''' distinct de ceux déjà posés.
    '''
    ''' Deux points de vente peuvent porter des désignations voisines, et un onglet en double
    ''' fait échouer l'écriture du classeur entier — après vingt feuilles déjà écrites.
    ''' </summary>
    Private Shared Function NomDOnglet(calc As CalculWU, dejaPris As HashSet(Of String)) As String

        Dim brut As String = If(String.IsNullOrWhiteSpace(calc.Account), calc.Designation, calc.Account)
        brut = Regex.Replace(If(brut, String.Empty), "[\\/\?\*\[\]:]", " ").Trim()

        If brut.Length = 0 Then brut = "PIECE"
        If brut.Length > 31 Then brut = brut.Substring(0, 31)

        Dim nom As String = brut
        Dim suffixe As Integer = 1

        While dejaPris.Contains(nom)
            suffixe += 1
            Dim marque As String = " (" & suffixe.ToString() & ")"
            Dim garde As Integer = Math.Min(brut.Length, 31 - marque.Length)
            nom = brut.Substring(0, garde) & marque
        End While

        dejaPris.Add(nom)
        Return nom
    End Function

#End Region

End Class
