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
    Private Const XL_PORTRAIT As Integer = 1

    ''' <summary>Format des montants : séparateur de milliers, pas de décimale — le FCFA n'en a pas.</summary>
    Private Const FORMAT_MONTANT As String = "#,##0"

#End Region

#Region "Géométrie du formulaire"

    ' POURQUOI CES TAILLES, ET NON CELLES DU MODÈLE
    '
    ' Le modèle fourni est dessiné pour être REMPLI À LA MAIN : colonnes de 24 à 114
    ' caractères, corps de 20 à 28 points, lignes de 33. Il faut cela pour écrire au stylo
    ' dans une case.
    '
    ' Mesuré, ce modèle fait 1 241 points de large et 997 de haut pour une pièce de six
    ' écritures. Une page A4 en offre 487 sur 734. Imprimé « une page », il sortirait donc à
    ' 39 % : les libellés à 4 points, illisibles ; imprimé à l'échelle, il sortirait sur deux
    ' pages et demie, et une pièce comptable en deux morceaux ne se signe pas.
    '
    ' Les tailles ci-dessous gardent les PROPORTIONS du modèle — les intitulés plus gros que
    ' le corps, Arial Black pour les uns, Century Schoolbook pour les autres — mais ramenées
    ' à ce qui tient sur une page : 519 points de large, et 662 de haut pour six écritures.
    ' Une pièce de point de vente sort ainsi à 94 %, et à 80 % au-delà de vingt écritures.
    '
    ' La structure, les libellés et l'enchaînement des cartouches, eux, ne bougent pas : c'est
    ' cela que le guichet de la comptabilité reconnaît, pas le corps de la police.

    Private Const LARGEUR_A As Double = 14.0   ' DEBIT : / CREDIT : / RAISON :
    Private Const LARGEUR_B As Double = 18.0   ' N° de comptes
    Private Const LARGEUR_C As Double = 48.0   ' libellés — le plus long fait 46 caractères
    Private Const LARGEUR_D As Double = 16.0   ' montants

    Private Const TAILLE_BANQUE As Integer = 14
    Private Const TAILLE_TITRE As Integer = 16
    Private Const TAILLE_CONTROLE As Integer = 14
    Private Const TAILLE_ETIQUETTE As Integer = 10   ' DATE : / DE : / POUR :
    Private Const TAILLE_VALEUR As Integer = 12      ' la date, le numéro
    Private Const TAILLE_AGENCE As Integer = 14
    Private Const TAILLE_INTITULE As Integer = 13
    Private Const TAILLE_COLONNE As Integer = 11     ' en-têtes du tableau
    Private Const TAILLE_ECRITURE As Integer = 11
    Private Const TAILLE_BLOC As Integer = 11        ' DEBIT : / CREDIT : / RAISON :
    Private Const TAILLE_RAISON As Integer = 12
    Private Const TAILLE_PIED As Integer = 11

    Private Const HAUTEUR_BANQUE As Double = 18.0
    Private Const HAUTEUR_TITRE As Double = 22.0
    Private Const HAUTEUR_DATE As Double = 18.0
    Private Const HAUTEUR_AGENCE As Double = 22.0
    Private Const HAUTEUR_POUR As Double = 18.0
    Private Const HAUTEUR_INTITULE As Double = 20.0
    Private Const HAUTEUR_COLONNES As Double = 18.0
    Private Const HAUTEUR_ECRITURE As Double = 18.0
    Private Const HAUTEUR_RAISON As Double = 20.0
    Private Const HAUTEUR_PIED As Double = 14.0

    ''' <summary>Marges d'impression, en pouces. 1 cm de chaque côté.</summary>
    Private Const MARGE As Double = 0.4

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

        ''' <summary>
        ''' Numéro de la pièce, engendré par l'application : numéro de lot de la journée, puis
        ''' le rang de la pièce dans le classeur. Voir <see cref="NumeroDePiece"/>.
        ''' </summary>
        Public Property Numero As String = String.Empty

        ''' <summary>
        ''' Agence émettrice, portée en face de « AGENCE: ». Elle varie d'une pièce à l'autre :
        ''' c'est l'agence de rattachement du point de vente.
        ''' </summary>
        Public Property AgenceEmettrice As String = ConstantesWU.PIECE_AGENCE_DEFAUT

        ''' <summary>Texte de la ligne RAISON.</summary>
        Public Property Raison As String = String.Empty

        ''' <summary>
        ''' Vrai pour une pièce qui doit tenir sur UNE page, quitte à réduire l'échelle.
        '''
        ''' C'est le cas d'une pièce de point de vente : elle se signe, et un document qui se
        ''' signe ne se signe pas en deux morceaux. La pièce globale, elle, peut compter trois
        ''' cents écritures — l'y forcer la rendrait illisible ; elle s'imprime sur plusieurs
        ''' pages, en-tête répété en haut de chacune.
        ''' </summary>
        Public Property TientSurUnePage As Boolean = True
    End Class

    ''' <summary>
    ''' Numéro d'une pièce, engendré sans rien demander à personne :
    ''' le numéro de lot de la journée, un tiret, le rang de la pièce dans le classeur.
    '''
    ''' Le numéro de lot est celui-là même que porte le fichier destiné au core banking
    ''' (<see cref="CoreBankingService.NumeroDeLot"/>) : quatre caractères tirés de la date,
    ''' donc identiques d'une exécution à l'autre pour une même journée, et différents d'une
    ''' journée à la suivante. Une pièce et l'écriture qu'elle justifie se retrouvent ainsi
    ''' l'une par l'autre, ce qu'un simple compteur repartant de 1 chaque matin ne permettrait
    ''' pas : deux pièces de deux journées porteraient le même numéro 2.
    ''' </summary>
    Public Shared Function NumeroDePiece(dateActivite As Date, rang As Integer) As String
        Return $"{CoreBankingService.NumeroDeLot(dateActivite)}-{rang:000}"
    End Function

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
    ''' <param name="agencePremiereFeuille">Agence émettrice de la première feuille. Vide pour
    ''' l'agence par défaut — la pièce globale n'appartient à aucune agence.</param>
    ''' <param name="progression">
    ''' Rend compte de l'avancement, onglet par onglet. Nothing pour ne rien annoncer : ce
    ''' service reste appelable sans interface.
    ''' </param>
    ''' <returns>Le chemin écrit.</returns>
    Public Shared Function Ecrire(dtGlobale As DataTable,
                                  listeCalculs As IEnumerable(Of CalculWU),
                                  dateActivite As Date,
                                  cheminFichier As String,
                                  Optional nomPremiereFeuille As String = "PIECE GLOBALE",
                                  Optional intitulePremiereFeuille As String = "",
                                  Optional agencePremiereFeuille As String = "",
                                  Optional progression As ProgressionWU = Nothing) As String

        If dtGlobale Is Nothing OrElse dtGlobale.Rows.Count = 0 Then
            Throw New InvalidOperationException(
                "Aucune donnée à exporter : générez la pièce comptable au préalable.")
        End If

        If String.IsNullOrWhiteSpace(cheminFichier) Then
            Throw New ArgumentException("Chemin de fichier non renseigné.", NameOf(cheminFichier))
        End If

        ' Les points de vente sont dénombrés AVANT d'ouvrir Excel : la barre doit connaître son
        ' total dès la première étape, sinon elle repart en arrière quand il se précise.
        Dim aDetailler As List(Of CalculWU) = PointsDeVenteADetailler(listeCalculs)

        ' Quatre étapes fixes — ouverture d'Excel, pièce globale, enregistrement, affichage —
        ' plus un onglet par point de vente. Le total doit être JUSTE : une barre qui s'arrête
        ' à quatre-vingt-dix pour cent laisse croire à un blocage, et une qui sature avant la
        ' fin laisse croire que c'est fini.
        If progression IsNot Nothing Then progression.Commencer(4 + aDetailler.Count)

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing

        Try
            Annoncer(progression, "Ouverture de Microsoft Excel…")

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

            ' Le référentiel des agences est lu UNE FOIS pour tout le classeur : une lecture
            ' par feuille ferait cinquante allers-retours vers SQL Server pour une information
            ' qui ne change pas pendant l'export.
            Dim agences As Dictionary(Of String, String) = ChargerLesAgences()

            Annoncer(progression, $"Pièce globale — journée du {dateActivite:dd/MM/yyyy}")

            EcrireLaPremiereFeuille(classeur, dtGlobale, dateActivite,
                                    nomPremiereFeuille, intitulePremiereFeuille,
                                    agencePremiereFeuille)
            EcrireLesPiecesIndividuelles(classeur, aDetailler, dateActivite, agences, progression)

            ' La première feuille est celle qu'on veut voir en ouvrant le classeur.
            classeur.Worksheets(1).Activate()

            Annoncer(progression, "Enregistrement du classeur…")
            classeur.SaveAs(cheminFichier)

            Annoncer(progression, "Ouverture du classeur à l'écran…")

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
                                               nomFeuille As String, intitule As String,
                                               agence As String)

        Dim contexte As New ContexteFeuille() With {
            .NomFeuille = If(String.IsNullOrWhiteSpace(nomFeuille), "PIECE GLOBALE", nomFeuille),
            .Intitule = If(String.IsNullOrWhiteSpace(intitule),
                           $"PIECE GLOBALE — journée du {dateActivite:dd/MM/yyyy}",
                           intitule),
            .DateActivite = dateActivite,
            .Numero = NumeroDePiece(dateActivite, 1),
            .TientSurUnePage = False,
            .AgenceEmettrice = If(String.IsNullOrWhiteSpace(agence),
                                  ConstantesWU.PIECE_AGENCE_DEFAUT, agence),
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
                                                    aDetailler As List(Of CalculWU),
                                                    dateActivite As Date,
                                                    agences As Dictionary(Of String, String),
                                                    progression As ProgressionWU)

        If aDetailler Is Nothing Then Return

        ' Les onglets déjà posés, pour ne pas en produire deux du même nom — Excel refuse
        ' alors d'écrire le classeur entier, après vingt feuilles déjà remplies.
        Dim nomsPris As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For indice As Integer = 1 To CInt(classeur.Worksheets.Count)
            nomsPris.Add(CStr(classeur.Worksheets(indice).Name))
        Next
        Dim numero As Integer = 1

        For Each calc As CalculWU In aDetailler

            Annoncer(progression, $"{calc.Account} — {calc.Designation}")

            Dim dtPdv As DataTable = PieceComptableService.GenererPieceComptable(New CalculWU() {calc})

            ' Un point de vente dont tous les montants s'arrondissent à zéro ne produit aucune
            ' écriture. Son étape a déjà été annoncée : la barre avance quand même, sans quoi
            ' elle n'atteindrait jamais son total.
            If dtPdv Is Nothing OrElse dtPdv.Rows.Count = 0 Then Continue For

            numero += 1

            Dim contexte As New ContexteFeuille() With {
                .NomFeuille = NomDOnglet(calc, nomsPris),
                .Intitule = IntituleDe(calc),
                .DateActivite = dateActivite,
                .Numero = NumeroDePiece(dateActivite, numero),
                .AgenceEmettrice = AgenceDe(calc, agences),
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

        Dim derniereDuPied As Integer = Cartouches(feuille, ligneRaison + 2)

        MiseEnPageImpression(feuille, derniereDuPied, contexte.TientSurUnePage)
    End Sub

    ''' <summary>Colonnes et police du modèle. Les largeurs sont celles du classeur fourni.</summary>
    Private Shared Sub LargeursEtPolice(feuille As Object)

        feuille.Cells.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
        feuille.Cells.Font.Size = TAILLE_ECRITURE

        feuille.Columns("A").ColumnWidth = LARGEUR_A
        feuille.Columns("B").ColumnWidth = LARGEUR_B

        ' La colonne des libellés reste la plus large : « COMMISSION TRANSFERT SA
        ' SS-AGENCE_BOLOLO SIEGE » fait quarante-six caractères, et le comptable lit la pièce
        ' sans élargir la colonne.
        feuille.Columns("C").ColumnWidth = LARGEUR_C
        feuille.Columns("D").ColumnWidth = LARGEUR_D
    End Sub

    Private Shared Sub EnTete(feuille As Object, contexte As ContexteFeuille)

        Ecrire(feuille, 4, 1, ConstantesWU.PIECE_BANQUE, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_BANQUE, True, XL_GAUCHE)
        feuille.Rows(4).RowHeight = HAUTEUR_BANQUE

        Ecrire(feuille, 6, 2, ConstantesWU.PIECE_TITRE, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_TITRE, False, XL_GAUCHE)
        feuille.Rows(6).RowHeight = HAUTEUR_TITRE

        Ecrire(feuille, 9, 1, ConstantesWU.PIECE_DATE, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_ETIQUETTE, True, XL_GAUCHE)

        ' La date est celle de la JOURNÉE COMPTABILISÉE, pas celle de l'impression. Le modèle
        ' porte =TODAY() parce qu'il est vierge ; une pièce rejouée trois jours plus tard doit
        ' rester datée du jour qu'elle comptabilise.
        Dim celluleDate As Object = feuille.Cells(9, 2)
        Try
            celluleDate.Value = contexte.DateActivite
            celluleDate.NumberFormat = "dd/mm/yyyy"
            celluleDate.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            celluleDate.Font.Size = TAILLE_VALEUR
        Finally
            Marshal.ReleaseComObject(celluleDate)
        End Try
        feuille.Rows(9).RowHeight = HAUTEUR_DATE

        Ecrire(feuille, 10, 1, ConstantesWU.PIECE_DE, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_ETIQUETTE, True, XL_GAUCHE)
        Ecrire(feuille, 10, 2, Emetteur(), ConstantesWU.PIECE_POLICE_CORPS, TAILLE_VALEUR, False, XL_GAUCHE)
        Ecrire(feuille, 10, 3, ConstantesWU.PIECE_AGENCE, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_VALEUR, True, XL_DROITE)
        Ecrire(feuille, 10, 4, contexte.AgenceEmettrice, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_AGENCE, True, XL_DROITE)
        feuille.Rows(10).RowHeight = HAUTEUR_AGENCE

        Ecrire(feuille, 11, 1, ConstantesWU.PIECE_POUR, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_ETIQUETTE, True, XL_GAUCHE)
        Ecrire(feuille, 11, 2, ConstantesWU.PIECE_SERVICE_DESTINATAIRE, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_VALEUR, False, XL_GAUCHE)
        Ecrire(feuille, 11, 4, contexte.Numero, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_VALEUR, True, XL_DROITE)
        feuille.Rows(11).RowHeight = HAUTEUR_POUR

        Ecrire(feuille, 12, 3, contexte.Intitule, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_INTITULE, True, XL_CENTRE)
        feuille.Rows(12).RowHeight = HAUTEUR_INTITULE

        Ecrire(feuille, 13, 2, ConstantesWU.PIECE_COMPTES, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_COLONNE, True, XL_CENTRE)
        Ecrire(feuille, 13, 3, ConstantesWU.PIECE_LIBELLES, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_COLONNE, True, XL_CENTRE)
        Ecrire(feuille, 13, 4, ConstantesWU.PIECE_MONTANTS, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_COLONNE, True, XL_DROITE)
        feuille.Rows(13).RowHeight = HAUTEUR_COLONNES

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

        Ecrire(feuille, premiereLigne, 1, intitule, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_BLOC, True, XL_CENTRE)

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
            zone.Font.Size = TAILLE_ECRITURE
            zone.RowHeight = HAUTEUR_ECRITURE
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
            cellule.Font.Size = TAILLE_CONTROLE
            cellule.Font.Bold = True
            cellule.HorizontalAlignment = XL_DROITE
            cellule.Borders.LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(cellule)
        End Try
    End Sub

    Private Shared Sub Raison(feuille As Object, ligne As Integer, texte As String)

        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_RAISON, ConstantesWU.PIECE_POLICE_TITRE, TAILLE_BLOC, True, XL_CENTRE)

        Dim zone As Object = feuille.Range(feuille.Cells(ligne, 2), feuille.Cells(ligne, 4))
        Try
            zone.Merge()
            zone.Value = texte
            zone.Font.Name = ConstantesWU.PIECE_POLICE_CORPS
            zone.Font.Size = TAILLE_RAISON
            zone.Font.Bold = True
            zone.HorizontalAlignment = XL_CENTRE
            zone.Borders.LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(zone)
        End Try

        feuille.Rows(ligne).RowHeight = HAUTEUR_RAISON
    End Sub

    ''' <summary>
    ''' Les quatre cartouches de signature, dans l'ordre et aux écartements du modèle : quatre
    ''' lignes vides entre chacun, parce que c'est l'espace où l'on signe.
    ''' </summary>
    Private Shared Function Cartouches(feuille As Object, premiereLigne As Integer) As Integer

        Dim ligne As Integer = premiereLigne + 3

        Souligner(feuille, ligne, ConstantesWU.PIECE_SIGNATURES, ConstantesWU.PIECE_FCU)
        ligne += 5
        TroisIntitules(feuille, ligne, ConstantesWU.PIECE_INITIE, ConstantesWU.PIECE_CONTROLE, ConstantesWU.PIECE_APPROUVE)
        ligne += 5
        Souligner(feuille, ligne, ConstantesWU.PIECE_ECRITURE, ConstantesWU.PIECE_OPS)
        ligne += 5
        TroisIntitules(feuille, ligne, ConstantesWU.PIECE_PASSEE, String.Empty, ConstantesWU.PIECE_AUTORISEE)
        ligne += 4
        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_DATE_ENREGISTREMENT, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_GAUCHE)
        ligne += 4
        Ecrire(feuille, ligne, 1, ConstantesWU.PIECE_NUMERO_SEQUENCE, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_GAUCHE)

        Dim pied As Object = feuille.Range(feuille.Cells(premiereLigne, 1), feuille.Cells(ligne, 4))
        Try
            pied.RowHeight = HAUTEUR_PIED
        Finally
            Marshal.ReleaseComObject(pied)
        End Try

        Return ligne
    End Function

    ''' <summary>Un cartouche à trait continu : intitulé à gauche, service à droite.</summary>
    Private Shared Sub Souligner(feuille As Object, ligne As Integer, gauche As String, droite As String)

        Ecrire(feuille, ligne, 1, gauche, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_GAUCHE)
        Ecrire(feuille, ligne, 4, droite, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_DROITE)

        Dim zone As Object = feuille.Range(feuille.Cells(ligne, 1), feuille.Cells(ligne, 4))
        Try
            zone.Borders(XL_BORD_BAS).LineStyle = XL_CONTINU
        Finally
            Marshal.ReleaseComObject(zone)
        End Try
    End Sub

    Private Shared Sub TroisIntitules(feuille As Object, ligne As Integer,
                                      gauche As String, milieu As String, droite As String)

        Ecrire(feuille, ligne, 1, gauche, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_GAUCHE)
        If milieu.Length > 0 Then Ecrire(feuille, ligne, 2, milieu, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_DROITE)
        Ecrire(feuille, ligne, 3, droite, ConstantesWU.PIECE_POLICE_CORPS, TAILLE_PIED, True, XL_CENTRE)
    End Sub

    ''' <summary>
    ''' Règle l'impression : une pièce se signe sur papier, et un document qui se signe ne se
    ''' signe pas en deux morceaux.
    '''
    ''' Trois réglages font tout le travail :
    '''
    '''   — la ZONE D'IMPRESSION s'arrête à la dernière ligne écrite. Sans elle, Excel décide
    '''     lui-même de ce qu'il imprime, et une cellule touchée par mégarde en colonne H
    '''     ajoute une page blanche ;
    '''   — une seule page EN LARGEUR, toujours. Une pièce coupée verticalement oblige à
    '''     raccorder les montants à leur libellé au scotch ;
    '''   — une seule page EN HAUTEUR pour une pièce de point de vente. La pièce globale, elle,
    '''     s'étale, en-tête répété en haut de chaque page : à trois cents écritures, la forcer
    '''     sur une page la réduirait à un timbre-poste.
    ''' </summary>
    ''' <param name="derniereLigne">Dernière ligne écrite, cartouches compris.</param>
    ''' <param name="unePage">Vrai pour une pièce qui doit tenir sur une seule page.</param>
    Private Shared Sub MiseEnPageImpression(feuille As Object, derniereLigne As Integer, unePage As Boolean)

        Try
            Dim reglage As Object = feuille.PageSetup

            Try
                reglage.PrintArea = "$A$1:$D$" & derniereLigne.ToString()

                ' Portrait : la pièce est plus haute que large, et c'est le format sous lequel
                ' la comptabilité classe ses justificatifs.
                reglage.Orientation = XL_PORTRAIT

                reglage.LeftMargin = Pouces(feuille, MARGE)
                reglage.RightMargin = Pouces(feuille, MARGE)
                reglage.TopMargin = Pouces(feuille, MARGE)
                reglage.BottomMargin = Pouces(feuille, MARGE)

                reglage.CenterHorizontally = True

                ' Zoom doit passer à False AVANT FitToPages : les deux réglages s'excluent, et
                ' Excel ignore silencieusement le second tant que le premier vaut un nombre.
                reglage.Zoom = False
                reglage.FitToPagesWide = 1

                If unePage Then
                    reglage.FitToPagesTall = 1
                Else
                    reglage.FitToPagesTall = False

                    ' L'en-tête se répète : sans cela, la page 2 d'une pièce globale arrive
                    ' sans date, sans agence et sans nom de colonne — une colonne de chiffres
                    ' dont on ne sait plus ce qu'ils sont.
                    reglage.PrintTitleRows = "$4:$13"
                End If

            Finally
                Marshal.ReleaseComObject(reglage)
            End Try

        Catch
            ' Aucune imprimante installée : Excel refuse alors PageSetup en bloc. La pièce est
            ' écrite, et c'est ce qui compte — l'échec ne doit pas faire perdre le classeur.
        End Try
    End Sub

    ''' <summary>
    ''' Convertit des pouces en points, unité des marges d'Excel.
    '''
    ''' Passe par InchesToPoints plutôt que par une multiplication par 72 : c'est Excel qui
    ''' décide de son unité, et elle a déjà changé entre deux versions.
    ''' </summary>
    Private Shared Function Pouces(feuille As Object, valeur As Double) As Double

        Try
            Return CDbl(feuille.Application.InchesToPoints(valeur))
        Catch
            Return valeur * 72.0
        End Try
    End Function

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

    ''' <summary>
    ''' Les points de vente qui auront leur onglet, dénombrés une fois pour toutes.
    ''' 
    ''' Un Account non comptabilisé n'a pas de pièce — il n'a pas d'écriture non plus — et lui
    ''' donner un onglet vide laisserait croire à une pièce à zéro.
    ''' </summary>
    Private Shared Function PointsDeVenteADetailler(listeCalculs As IEnumerable(Of CalculWU)) As List(Of CalculWU)

        Dim retenus As New List(Of CalculWU)()

        If listeCalculs Is Nothing Then Return retenus

        For Each calc As CalculWU In listeCalculs
            If calc IsNot Nothing AndAlso calc.EstComptabilisable Then retenus.Add(calc)
        Next

        Return retenus
    End Function

    ''' <summary>Annonce une étape, s'il y a quelqu'un pour l'entendre.</summary>
    Private Shared Sub Annoncer(progression As ProgressionWU, libelle As String)

        If progression Is Nothing Then Return
        progression.Avancer(libelle)
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
    ''' Qui établit la pièce : le service, puis la personne connectée.
    ''' 
    ''' « DE : » dit DE LA PART DE QUI. Le service seul ne le dit qu'à moitié — trois agents se
    ''' relaient sur la compense, et c'est le cartouche « INITIE PAR » que le comptable regarde
    ''' ensuite pour savoir à qui la renvoyer. Autant que la pièce le porte dès son en-tête.
    ''' </summary>
    Private Shared Function Emetteur() As String

        Dim personne As String = String.Empty

        If SessionWU.Utilisateur IsNot Nothing Then
            personne = If(SessionWU.Utilisateur.NomComplet, String.Empty).Trim()
            If personne.Length = 0 Then personne = If(SessionWU.Utilisateur.Identifiant, String.Empty).Trim()
        End If

        If personne.Length = 0 Then Return ConstantesWU.PIECE_SERVICE_EMETTEUR

        Return ConstantesWU.PIECE_SERVICE_EMETTEUR & " — " & personne
    End Function

    ''' <summary>
    ''' Le référentiel des agences, indexé par tout ce qui peut servir à les retrouver : leur
    ''' Account (Codesite) ET leur code agence Voyager.
    ''' 
    ''' Un sous-agent porte le code agence de son rattachement ; une agence propre porte le sien.
    ''' Les deux tombent dans le même dictionnaire, et une seule recherche suffit ensuite.
    ''' 
    ''' Une erreur SQL rend un dictionnaire vide plutôt que de faire échouer l'export : une pièce
    ''' qui porte un code d'agence au lieu de son nom reste une pièce juste. Une pièce qu'on n'a
    ''' pas pu produire, non.
    ''' </summary>
    Private Shared Function ChargerLesAgences() As Dictionary(Of String, String)

        Dim repertoire As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Try
            Dim messageErreur As String = String.Empty
            Dim liste As List(Of PointDeVenteEC) = PdvRepository.ListerAgences(String.Empty, messageErreur)

            If liste Is Nothing Then Return repertoire

            For Each agence As PointDeVenteEC In liste

                Dim nom As String = If(agence.Designation, String.Empty).Trim()
                If nom.Length = 0 Then Continue For

                Dim codeSite As String = If(agence.CodeSite, String.Empty).Trim()
                If codeSite.Length > 0 Then repertoire(codeSite) = nom

                Dim codeVoyager As String = If(agence.CodeAgenceVoyager, String.Empty).Trim()
                If codeVoyager.Length > 0 Then repertoire(codeVoyager) = nom
            Next

        Catch ex As Exception
            ' Base injoignable ou droits manquants : voir la remarque ci-dessus.
        End Try

        Return repertoire
    End Function

    ''' <summary>Agence émettrice d'un point de vente, référentiel lu à la volée.</summary>
    Public Shared Function AgenceDe(calc As CalculWU) As String
        Return AgenceDe(calc, ChargerLesAgences())
    End Function

    ''' <summary>
    ''' Agence émettrice d'une pièce de point de vente, du plus précis au plus vague : le nom de
    ''' l'agence de rattachement, sinon son code, sinon l'agence par défaut.
    ''' 
    ''' Une agence propre est sa propre agence émettrice ; un sous-agent relève de celle qui le
    ''' porte dans ses livres — la même que la colonne ACBRN du fichier core banking.
    ''' </summary>
    Private Shared Function AgenceDe(calc As CalculWU, agences As Dictionary(Of String, String)) As String

        Dim nom As String = String.Empty

        If agences IsNot Nothing Then

            Dim code As String = If(calc.CodeAgence, String.Empty).Trim()
            If code.Length > 0 Then agences.TryGetValue(code, nom)

            ' Une agence propre se retrouve aussi par son Account, quand son code agence
            ' Voyager n'est pas renseigné dans le référentiel.
            If String.IsNullOrEmpty(nom) Then
                Dim account As String = If(calc.Account, String.Empty).Trim()
                If account.Length > 0 Then agences.TryGetValue(account, nom)
            End If
        End If

        If Not String.IsNullOrEmpty(nom) Then Return nom
        If Not String.IsNullOrWhiteSpace(calc.CodeAgence) Then Return calc.CodeAgence.Trim()

        Return ConstantesWU.PIECE_AGENCE_DEFAUT
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
