Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>
''' Consultation du fichier destiné au core banking, avant de le produire.
'''
''' Ce fichier impacte réellement les comptes : il ne doit pas sortir sans qu'on ait pu le
''' regarder. L'écran présente les treize colonnes telles qu'elles seront écrites — pas une
''' version arrangée pour la lecture — et n'écrit rien tant que l'export n'est pas demandé.
''' </summary>
Public Class FrmFichierCoreBanking

    Private ReadOnly _fichier As DataTable
    Private ReadOnly _dateActivite As Date
    Private ReadOnly _dateValeur As Date
    Private ReadOnly _numeroLot As String

    ''' <summary>Vrai si l'utilisateur a effectivement produit le fichier.</summary>
    Public ReadOnly Property FichierProduit As Boolean
        Get
            Return _produit
        End Get
    End Property
    Private _produit As Boolean = False

    ''' <summary>Chemin du fichier produit, ou chaîne vide.</summary>
    Public ReadOnly Property CheminProduit As String
        Get
            Return _chemin
        End Get
    End Property
    Private _chemin As String = String.Empty

    ''' <summary>Constructeur requis par le Concepteur Windows Forms.</summary>
    Public Sub New()
        InitializeComponent()
        _fichier = New DataTable()
        _numeroLot = String.Empty
    End Sub

    Public Sub New(fichier As DataTable, dateActivite As Date, dateValeur As Date, numeroLot As String)

        InitializeComponent()

        _fichier = If(fichier, New DataTable())
        _dateActivite = dateActivite
        _dateValeur = dateValeur
        _numeroLot = If(numeroLot, String.Empty)
    End Sub

#Region "Ouverture"

    Private Sub FrmFichierCoreBanking_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dgvFichier.DataSource = _fichier
        AjusterColonnes()
        AfficherRecapitulatif()
    End Sub

    Private Sub AjusterColonnes()

        For Each colonne As DataGridViewColumn In dgvFichier.Columns
            colonne.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        ' AMOUNT est le seul nombre du fichier : il se lit aligné à droite, comme un montant.
        If dgvFichier.Columns.Contains("AMOUNT") Then
            dgvFichier.Columns("AMOUNT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    ''' <summary>
    ''' Rappelle ce que le fichier contient et vérifie son équilibre sous les yeux de
    ''' l'utilisateur : un total affiché vaut mieux qu'une promesse.
    ''' </summary>
    Private Sub AfficherRecapitulatif()

        Dim debit As Long = TotalParSens(ConstantesWU.CB_SENS_DEBIT)
        Dim credit As Long = TotalParSens(ConstantesWU.CB_SENS_CREDIT)

        lblTitre.Text = $"Journée du {_dateActivite:dd/MM/yyyy} — lot {_numeroLot}"

        lblRecapitulatif.Text =
            $"{_fichier.Rows.Count} ligne(s)     date de valeur : {_dateValeur:dd/MM/yyyy}" &
            Environment.NewLine &
            CalendrierWU.Explication(_dateActivite, _dateValeur)

        If debit = credit Then
            lblEquilibre.ForeColor = Drawing.Color.DarkGreen
            lblEquilibre.Text = $"Équilibré : {debit:N0} FCFA au débit comme au crédit."
        Else
            lblEquilibre.ForeColor = Drawing.Color.Firebrick
            lblEquilibre.Text = $"DÉSÉQUILIBRÉ : {debit:N0} au débit, {credit:N0} au crédit."
            btnExporter.Enabled = False
        End If
    End Sub

    Private Function TotalParSens(sens As String) As Long

        Dim total As Long = 0L

        If Not _fichier.Columns.Contains("DRCR") OrElse Not _fichier.Columns.Contains("AMOUNT") Then
            Return total
        End If

        For Each ligne As DataRow In _fichier.Rows

            If Not String.Equals(Convert.ToString(ligne("DRCR"), CultureInfo.InvariantCulture),
                                 sens, StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim montant As Long
            If Long.TryParse(Convert.ToString(ligne("AMOUNT"), CultureInfo.InvariantCulture),
                             NumberStyles.Integer, CultureInfo.InvariantCulture, montant) Then
                total += montant
            End If
        Next

        Return total
    End Function

#End Region

#Region "Export"

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        Dim chemin As String = DemanderLeChemin()
        If chemin.Length = 0 Then Return

        Dim avancement As FrmProgression = Nothing

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, "Production du fichier core banking")

            ' Seul AMOUNT est écrit en nombre : tout le reste est du texte, sans quoi Excel
            ' réinterpréterait les numéros de compte et les numéros de lot.
            ExcelExportService.ExporterTableBrute(_fichier, New String() {"AMOUNT"},
                                                  "CoreBanking", chemin, ouvrirApres:=True,
                                                  progression:=avancement.Progression)

            _produit = True
            _chemin = chemin

            TracerLaProduction(chemin)

            btnExporter.Enabled = False
            lblEquilibre.ForeColor = Drawing.Color.DarkGreen
            lblEquilibre.Text = $"Fichier produit : {IO.Path.GetFileName(chemin)} — il s'ouvre dans Excel."

        Catch ex As Exception
            If avancement IsNot Nothing Then avancement.Fermer()

            MessageBox.Show("Production du fichier impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Enregistre que ce fichier est sorti : journée, date de valeur, lot, auteur, totaux.
    '''
    ''' C'est le seul moment où la journée quitte Wincompense pour entrer dans les livres
    ''' de la banque. Sans cette trace, l'application ne peut ni avertir avant de
    ''' recomptabiliser une journée déjà partie, ni prévenir avant de l'annuler.
    '''
    ''' L'échec n'est JAMAIS bloquant : le fichier existe et s'ouvre dans Excel. Refuser
    ''' la production parce que le carnet n'a pas pu s'écrire ferait perdre le travail
    ''' pour sauver le carnet. L'agent est simplement averti, dans la barre d'état.
    ''' </summary>
    Private Sub TracerLaProduction(chemin As String)

        Dim messageErreur As String = String.Empty

        Dim trace As Boolean = CoreBankingRepository.Enregistrer(
            _dateActivite, _dateValeur, _numeroLot, chemin, _fichier.Rows.Count,
            TotalParSens(ConstantesWU.CB_SENS_DEBIT),
            TotalParSens(ConstantesWU.CB_SENS_CREDIT), messageErreur)

        If trace Then Return

        lblRecapitulatif.Text &= Environment.NewLine &
                                 "Production NON tracée : " & PremiereLigne(messageErreur)
    End Sub

    ''' <summary>La première ligne d'un message, pour un libellé qui n'en tient qu'une.</summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim fin As Integer = texte.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If fin < 0 Then Return texte
        Return texte.Substring(0, fin)
    End Function

    ''' <summary>Demande où enregistrer. Chaîne vide si l'utilisateur renonce.</summary>
    Private Function DemanderLeChemin() As String

        Using dialogue As New SaveFileDialog()

            dialogue.Title = "Enregistrer le fichier destiné au core banking"
            dialogue.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
            dialogue.FileName = CoreBankingService.NomDeFichier(_dateActivite)
            dialogue.OverwritePrompt = True

            If dialogue.ShowDialog(Me) <> DialogResult.OK Then Return String.Empty
            Return dialogue.FileName
        End Using
    End Function

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
