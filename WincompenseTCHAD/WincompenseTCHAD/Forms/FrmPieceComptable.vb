Option Strict On
Option Explicit On

Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Linq

''' <summary>
''' Formulaire d'affichage d'une pièce comptable (globale ou limitée à un seul point de vente).
''' Purement présentationnel : il reçoit la DataTable dtPiece déjà construite par
''' PieceComptableService et se contente de l'afficher, d'en présenter les totaux et
''' d'en proposer l'ouverture dans Excel.
''' </summary>
Public Class FrmPieceComptable

    Private ReadOnly _dtPiece As DataTable

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

    Private Sub FrmPieceComptable_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Le bandeau nomme l'écran ; lblTitre continue d'annoncer de quelle pièce il s’agit.
        ' Les deux ne font pas double emploi : l’un est fixe, l’autre change à chaque pièce.
        ThemeWU.Appliquer(Me, "Pièce comptable", btnOuvrirExcel)
        ' Pas de remplissage : cette pièce s'ouvre en boîte de dialogue modale. L'étirer sur
        ' tout l'écran masquerait l'écran de traitement d'où elle vient.
        DimensionsWU.Adapter(Me, False)
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
    End Sub

    Private Sub AfficherTotaux()

        If _dtPiece Is Nothing OrElse _dtPiece.Rows.Count = 0 Then
            lblTotaux.Text = "Aucune écriture."
            lblEcart.Text = String.Empty
            btnOuvrirExcel.Enabled = False
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

#Region "Actions"

    Private Sub btnOuvrirExcel_Click(sender As Object, e As EventArgs) Handles btnOuvrirExcel.Click

        Try
            Cursor = Cursors.WaitCursor
            Dim cheminTemp As String = PieceComptableService.OuvrirPieceComptableExcel(_dtPiece)
            lblSousTitre.Text = $"Classeur ouvert dans Excel : {cheminTemp}"

        Catch ex As Exception
            MessageBox.Show(
                "Impossible d'ouvrir la pièce comptable dans Excel : " & ex.Message & Environment.NewLine &
                "La pièce reste consultable dans ce formulaire.",
                "Ouverture Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

#End Region

End Class
