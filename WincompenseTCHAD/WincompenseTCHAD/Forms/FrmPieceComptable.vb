Option Strict On
Option Explicit On

Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Windows.Forms

''' <summary>
''' Consultation d'une pièce comptable — globale ou limitée à un seul point de vente — avant
''' de l'exporter.
'''
''' L'écran est purement présentationnel : il reçoit la DataTable dtPiece déjà construite par
''' PieceComptableService, l'affiche, en présente les totaux et n'écrit rien sur le disque tant
''' que l'export n'est pas demandé. La pièce engage la comptabilité de la banque : elle doit
''' pouvoir être regardée avant de sortir, exactement comme le fichier destiné au core banking.
''' </summary>
Public Class FrmPieceComptable

    Private ReadOnly _dtPiece As DataTable

    ''' <summary>Vrai si l'utilisateur a effectivement exporté la pièce.</summary>
    Public ReadOnly Property PieceExportee As Boolean
        Get
            Return _exportee
        End Get
    End Property
    Private _exportee As Boolean = False

    ''' <summary>Chemin du classeur produit, ou chaîne vide.</summary>
    Public ReadOnly Property CheminExporte As String
        Get
            Return _chemin
        End Get
    End Property
    Private _chemin As String = String.Empty

    ''' <summary>
    ''' Nom de fichier proposé par défaut dans la boîte d'enregistrement. L'appelant le
    ''' renseigne avant d'afficher l'écran ; à défaut, un nom daté du jour est employé.
    ''' </summary>
    Public Property NomFichierPropose As String = PieceComptableService.NomDeFichier(Date.Today)

    ''' <summary>
    ''' Journée comptabilisée, portée par l'en-tête du formulaire exporté.
    ''' 
    ''' Ce n'est pas la date d'impression : une pièce rejouée trois jours plus tard doit rester
    ''' datée du jour qu'elle comptabilise, sans quoi elle ne se rapproche plus de rien.
    ''' </summary>
    Public Property DateActivite As Date = Date.Today

    ''' <summary>
    ''' Points de vente à détailler, un onglet chacun, dans le classeur exporté.
    ''' 
    ''' Nothing — le cas par défaut — n'écrit que la pièce affichée, en une feuille. L'appelant
    ''' qui dispose de la liste la renseigne : c'est ce qui donne au comptable la pièce de chaque
    ''' sous-agent sans l'extraire à la main d'un tableau de trois cents lignes.
    ''' </summary>
    Public Property Calculs As IEnumerable(Of CalculWU) = Nothing

    ''' <summary>
    ''' Ligne d'identification de la première feuille, à la forme du modèle de la banque
    ''' (« Agence  001: … »). Vide pour l'intitulé de la pièce globale.
    ''' </summary>
    Public Property IntitulePiece As String = String.Empty

    ''' <summary>Onglet de la première feuille du classeur exporté.</summary>
    Public Property NomPremiereFeuille As String = "PIECE GLOBALE"

    ''' <summary>
    ''' Agence émettrice de la première feuille. Vide pour l'agence par défaut — la pièce
    ''' globale n'appartient à aucune agence en particulier.
    ''' </summary>
    Public Property AgencePiece As String = String.Empty

    ''' <summary>Constructeur sans paramètre requis par le Concepteur Windows Forms.</summary>
    Public Sub New()
        InitializeComponent()
        _dtPiece = New DataTable()
    End Sub

    ''' <summary>
    ''' Construit le formulaire pour une pièce comptable donnée.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable générée par PieceComptableService.GenererPieceComptable.</param>
    ''' <param name="titre">Titre affiché en haut du formulaire (ex. Account et désignation du PDV).</param>
    ''' <param name="sousTitre">Ligne d'information complémentaire (type de PDV, taux, etc.).</param>
    Public Sub New(dtPiece As DataTable, titre As String, sousTitre As String)

        InitializeComponent()

        _dtPiece = If(dtPiece, New DataTable())
        Text = titre
        lblTitre.Text = titre
        lblSousTitre.Text = sousTitre

        AfficherPiece()
    End Sub

#Region "Affichage"

    ''' <summary>
    ''' La liste des portées se remplit ICI, et non dans le constructeur : l'appelant pose la
    ''' propriété Calculs APRÈS avoir construit la fenêtre, et un filtre bâti trop tôt serait
    ''' vide sans que rien ne le dise.
    ''' </summary>
    Private Sub FrmPieceComptable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PreparerLaPortee()
    End Sub

    Private Sub AfficherPiece()

        dgvPiece.DataSource = _dtPiece
        FormaterColonnes()
        AfficherTotaux()
    End Sub

    Private Sub FormaterColonnes()

        ' Format à trois sections "positif;négatif;zéro" : la section « zéro » étant vide, une
        ' écriture sans montant de ce côté laisse la cellule vide, ce qui rend le sens de chaque
        ' ligne (débit ou crédit) immédiatement lisible.
        For Each nomColonne As String In {"Debit", "Credit"}
            If dgvPiece.Columns.Contains(nomColonne) Then
                dgvPiece.Columns(nomColonne).DefaultCellStyle.Format = "#,##0;-#,##0;"
                dgvPiece.Columns(nomColonne).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvPiece.Columns(nomColonne).Width = 110
            End If
        Next

        If dgvPiece.Columns.Contains("Compte") Then dgvPiece.Columns("Compte").Width = 140
        If dgvPiece.Columns.Contains("Libelle") Then
            dgvPiece.Columns("Libelle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvPiece.Columns("Libelle").HeaderText = "Libellé"
        End If

        ' Le code agence ne fait pas partie de la pièce comptable : il ne sert qu'à alimenter la
        ' colonne ACBRN du fichier destiné au core banking. L'afficher ici ferait croire à une
        ' colonne comptable de plus.
        If dgvPiece.Columns.Contains("CodeAgence") Then
            dgvPiece.Columns("CodeAgence").Visible = False
        End If
    End Sub

    Private Sub AfficherTotaux()

        If _dtPiece Is Nothing OrElse _dtPiece.Rows.Count = 0 Then
            lblTotaux.Text = "Aucune écriture."
            lblEcart.Text = String.Empty
            btnExporter.Enabled = False
            Return
        End If

        Dim totalDebit As Decimal = _dtPiece.AsEnumerable().Sum(Function(r) Convert.ToDecimal(r("Debit"), CultureInfo.InvariantCulture))
        Dim totalCredit As Decimal = _dtPiece.AsEnumerable().Sum(Function(r) Convert.ToDecimal(r("Credit"), CultureInfo.InvariantCulture))
        Dim ecart As Decimal = totalDebit - totalCredit

        lblTotaux.Text = $"{_dtPiece.Rows.Count} écritures     TOTAL DÉBIT : {totalDebit:N0} FCFA     TOTAL CRÉDIT : {totalCredit:N0} FCFA"

        If ecart = 0D Then
            lblEcart.ForeColor = Color.DarkGreen
            lblEcart.Text = "Pièce équilibrée (écart nul)."
        Else
            ' Un écart de quelques francs est normal : il provient de l'arrondi au FCFA de chaque
            ' ligne. Il n'est jamais corrigé ici — le compte d'attente ne s'applique qu'à la pièce
            ' globale, jamais point de vente par point de vente (règle métier, section 14).
            lblEcart.ForeColor = If(Math.Abs(ecart) > ConstantesWU.SEUIL_ECART_LIGNE_ANORMAL, Color.Firebrick, Color.DarkGoldenrod)
            lblEcart.Text = $"Écart Débit − Crédit : {ecart:N0} FCFA (arrondi au FCFA de chaque ligne)." &
                            Environment.NewLine &
                            "Cet écart n'est pas corrigé ici : le compte d'attente ne s'applique qu'à la pièce globale."
        End If
    End Sub

#End Region

#Region "Portée de l'export"

    ''' <summary>
    ''' Un choix de la liste : son libellé, et le type de point de vente qu'il retient.
    ''' </summary>
    Private NotInheritable Class ChoixPortee

        Public Sub New(libelle As String, typePdv As String, nombre As Integer)
            _libelle = libelle
            TypePdv = typePdv
            Nombre = nombre
        End Sub

        Private ReadOnly _libelle As String

        ''' <summary>Type retenu : "SA", "EC", ou chaîne vide pour tout prendre.</summary>
        Public ReadOnly Property TypePdv As String

        ''' <summary>Nombre de pièces que ce choix produira.</summary>
        Public ReadOnly Property Nombre As Integer

        ''' <summary>Ce que la liste affiche. Le nombre y figure : on choisit mieux en sachant combien.</summary>
        Public Overrides Function ToString() As String
            Return $"{_libelle} ({Nombre})"
        End Function
    End Class

    ''' <summary>
    ''' Remplit la liste des portées, ou la masque s'il n'y a pas de pièces individuelles à
    ''' produire — la fenêtre sert aussi à présenter la pièce d'un seul point de vente, et un
    ''' filtre y serait un choix entre une chose et elle-même.
    ''' </summary>
    Private Sub PreparerLaPortee()

        Dim pieces As List(Of CalculWU) = Comptabilisables()

        lblPortee.Visible = pieces.Count > 0
        cboPortee.Visible = pieces.Count > 0
        If pieces.Count = 0 Then Return

        Dim sousAgents As Integer = pieces.Where(Function(calcul) EstDuType(calcul, "SA")).Count()
        Dim agences As Integer = pieces.Count - sousAgents

        cboPortee.Items.Clear()
        cboPortee.Items.Add(New ChoixPortee("Tous les points de vente", String.Empty, pieces.Count))
        cboPortee.Items.Add(New ChoixPortee("Sous-agents seulement", "SA", sousAgents))
        cboPortee.Items.Add(New ChoixPortee("Agences propres seulement", "EC", agences))
        cboPortee.SelectedIndex = 0
    End Sub

    ''' <summary>Les points de vente qui produiront une pièce : les autres n'ont pas d'écriture.</summary>
    Private Function Comptabilisables() As List(Of CalculWU)

        If Calculs Is Nothing Then Return New List(Of CalculWU)()

        Return Calculs.Where(Function(calcul) calcul IsNot Nothing AndAlso calcul.EstComptabilisable).ToList()
    End Function

    Private Shared Function EstDuType(calcul As CalculWU, typePdv As String) As Boolean
        Return String.Equals(calcul.TypePdv, typePdv, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>La portée retenue, ou « tout » tant que rien n'a été choisi.</summary>
    Private Function PorteeChoisie() As ChoixPortee

        Dim choix As ChoixPortee = TryCast(cboPortee.SelectedItem, ChoixPortee)
        If choix IsNot Nothing Then Return choix

        Return New ChoixPortee("Tous les points de vente", String.Empty, Comptabilisables().Count)
    End Function

    ''' <summary>Les points de vente dont le classeur portera une feuille.</summary>
    Private Function CalculsAExporter() As IEnumerable(Of CalculWU)

        Dim choix As ChoixPortee = PorteeChoisie()
        If choix.TypePdv.Length = 0 Then Return Calculs

        Return Comptabilisables().Where(Function(calcul) EstDuType(calcul, choix.TypePdv)).ToList()
    End Function

    ''' <summary>
    ''' Suffixe porté par le nom du fichier proposé, pour qu'un dossier reste lisible sans ouvrir
    ''' les classeurs : PieceWU_20260530_SA.xlsx à côté de PieceWU_20260530.xlsx.
    ''' </summary>
    Private Function NomFichierSelonLaPortee() As String

        Dim choix As ChoixPortee = PorteeChoisie()
        If choix.TypePdv.Length = 0 Then Return NomFichierPropose

        Dim sansExtension As String = IO.Path.GetFileNameWithoutExtension(NomFichierPropose)
        Return $"{sansExtension}_{choix.TypePdv}{IO.Path.GetExtension(NomFichierPropose)}"
    End Function

#End Region

#Region "Export"

    ''' <summary>
    ''' Demande où enregistrer, écrit le classeur, puis l'ouvre immédiatement dans Excel :
    ''' l'utilisateur voit ce qu'il vient de produire sans avoir à le retrouver sur son disque.
    ''' </summary>
    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        Dim chemin As String = DemanderLeChemin()
        If chemin.Length = 0 Then Return

        Try
            Cursor = Cursors.WaitCursor

            PieceComptableService.ExporterEtOuvrirPieceExcel(
                _dtPiece, CalculsAExporter(), DateActivite, chemin,
                NomPremiereFeuille, IntitulePiece, AgencePiece)

            _exportee = True
            _chemin = chemin

            lblSousTitre.Text = $"Pièce exportée : {chemin} — elle s'ouvre dans Excel." &
                                NombreDeFeuilles()

        Catch ex As Exception
            MessageBox.Show(
                "Impossible d'exporter la pièce comptable : " & ex.Message & Environment.NewLine &
                "La pièce reste consultable dans cette fenêtre.",
                "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Rappelle, après coup, combien d'onglets le classeur porte. Un comptable qui attend une
    ''' pièce par sous-agent doit pouvoir constater qu'elles y sont toutes.
    ''' </summary>
    Private Function NombreDeFeuilles() As String

        If Calculs Is Nothing Then Return String.Empty

        ' Where(...).Count() plutôt que Count(...) : le jour où cette liste serait typée
        ' List(Of CalculWU), Count deviendrait une propriété et Count(...) se lirait comme un
        ' indexeur — BC32016, à la compilation, loin d'ici.
        Dim pieces As Integer =
            CalculsAExporter().Where(Function(calcul) calcul IsNot Nothing AndAlso calcul.EstComptabilisable).Count()
        If pieces = 0 Then Return String.Empty

        Return $" — {pieces} pièce{If(pieces > 1, "s", String.Empty)} de point de vente, en onglets séparés."
    End Function

    ''' <summary>Demande où enregistrer. Chaîne vide si l'utilisateur renonce.</summary>
    Private Function DemanderLeChemin() As String

        Using dialogue As New SaveFileDialog()

            dialogue.Title = "Enregistrer la pièce comptable"
            dialogue.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
            dialogue.FileName = NomFichierSelonLaPortee()
            dialogue.OverwritePrompt = True

            If dialogue.ShowDialog(Me) <> DialogResult.OK Then Return String.Empty
            Return dialogue.FileName
        End Using
    End Function

#End Region

End Class
