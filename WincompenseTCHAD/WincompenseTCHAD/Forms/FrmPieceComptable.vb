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

            PieceComptableService.ExporterEtOuvrirPieceExcel(_dtPiece, chemin)

            _exportee = True
            _chemin = chemin

            lblSousTitre.Text = $"Pièce exportée : {chemin} — elle s'ouvre dans Excel."

        Catch ex As Exception
            MessageBox.Show(
                "Impossible d'exporter la pièce comptable : " & ex.Message & Environment.NewLine &
                "La pièce reste consultable dans cette fenêtre.",
                "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>Demande où enregistrer. Chaîne vide si l'utilisateur renonce.</summary>
    Private Function DemanderLeChemin() As String

        Using dialogue As New SaveFileDialog()

            dialogue.Title = "Enregistrer la pièce comptable"
            dialogue.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
            dialogue.FileName = NomFichierPropose
            dialogue.OverwritePrompt = True

            If dialogue.ShowDialog(Me) <> DialogResult.OK Then Return String.Empty
            Return dialogue.FileName
        End Using
    End Function

#End Region

End Class
