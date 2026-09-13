Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Rapport d'activité Western Union sur une période, en quatre états : synthèse, jour par
''' jour, par point de vente et par groupe statistique.
'''
''' La source est l'historique des journées COMPTABILISÉES (table T_HistoriqueWU), alimentée à
''' chaque génération de pièce comptable. Le rapport restitue donc ce qui a réellement été
''' comptabilisé, et non un recalcul a posteriori qui, lui, dépendrait du paramétrage du jour.
'''
''' Les quatre états sont bâtis sur la même lecture, agrégée différemment : leurs totaux sont
''' donc nécessairement identiques d'une page à l'autre.
''' </summary>
Public Class FrmRapportActivite

    ''' <summary>Lignes d'historique de la période affichée.</summary>
    Private _lignes As New List(Of LigneHistoriqueWU)

    ''' <summary>Les quatre états, conservés pour l'export : ce qui part dans Excel est ce qui est à l'écran.</summary>
    Private _synthese As DataTable
    Private _parJour As DataTable
    Private _parPdv As DataTable
    Private _parGroupe As DataTable

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture"

    ''' <summary>
    ''' À l'ouverture, la période proposée est celle effectivement disponible dans l'historique :
    ''' présenter un intervalle vide obligerait l'utilisateur à deviner ce qui a été comptabilisé.
    ''' </summary>
    Private Sub FrmRapportActivite_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim premiere As Date = Date.Today
        Dim derniere As Date = Date.Today
        Dim messageErreur As String = String.Empty

        If Not HistoriqueRepository.ObtenirBornes(premiere, derniere, messageErreur) Then

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblDisponible.Text = "Historique illisible."
                MessageBox.Show(messageErreur, "Historique", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                lblDisponible.Text = "Aucune journée comptabilisée à ce jour."
                lblStatut.Text = "L'historique se remplit à chaque génération de pièce comptable."
            End If

            btnExporter.Enabled = False
            Return
        End If

        dtpDebut.Value = premiere
        dtpFin.Value = derniere
        lblDisponible.Text = $"Journées disponibles : du {premiere:dd/MM/yyyy} au {derniere:dd/MM/yyyy}"

        AfficherRapport()
    End Sub

#End Region

#Region "Affichage"

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click
        AfficherRapport()
    End Sub

    Private Sub AfficherRapport()

        If dtpDebut.Value.Date > dtpFin.Value.Date Then
            MessageBox.Show(
                "La date de début est postérieure à la date de fin.",
                "Période incorrecte", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            _lignes = HistoriqueRepository.ListerPeriode(dtpDebut.Value.Date, dtpFin.Value.Date, messageErreur)

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = "Lecture impossible : voir le message affiché."
                MessageBox.Show(messageErreur, "Historique", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _synthese = RapportActiviteService.ConstruireSynthese(_lignes)
            _parJour = RapportActiviteService.ConstruireParJour(_lignes)
            _parPdv = RapportActiviteService.ConstruireParPointDeVente(_lignes)
            _parGroupe = RapportActiviteService.ConstruireParGroupe(_lignes)

            AfficherSynthese()
            AfficherDetail(dgvParJour, _parJour, "Date", "Date")
            AfficherDetail(dgvParPdv, _parPdv, "Account", "Account")
            AfficherDetail(dgvParGroupe, _parGroupe, "Groupe", "Groupe")

            btnExporter.Enabled = _lignes.Count > 0

            If _lignes.Count = 0 Then
                lblStatut.Text = "Aucune journée comptabilisée sur cette période."
            Else
                lblStatut.Text = $"{RapportActiviteService.CompterJours(_lignes)} journée(s), " &
                                 $"{RapportActiviteService.CompterPointsDeVente(_lignes)} point(s) de vente."
            End If

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>Synthèse : deux colonnes, les intitulés de section en gras et sans valeur.</summary>
    Private Sub AfficherSynthese()

        dgvSynthese.DataSource = Nothing
        dgvSynthese.DataSource = _synthese

        If dgvSynthese.Columns.Contains("Valeur") Then
            dgvSynthese.Columns("Valeur").DefaultCellStyle.Format = "N0"
            dgvSynthese.Columns("Valeur").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvSynthese.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' Une rubrique sans valeur est un intitulé de section : mise en gras sur fond clair.
        For Each ligne As DataGridViewRow In dgvSynthese.Rows

            Dim valeur As Object = ligne.Cells("Valeur").Value

            If valeur Is Nothing OrElse valeur Is DBNull.Value Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(dgvSynthese.Font, Drawing.FontStyle.Bold)
                ligne.DefaultCellStyle.BackColor = Drawing.Color.Gainsboro
            ElseIf Convert.ToString(ligne.Cells("Rubrique").Value).Trim().StartsWith("TOTAL", StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(dgvSynthese.Font, Drawing.FontStyle.Bold)
            End If
        Next
    End Sub

    ''' <summary>Pages de détail : intitulés lisibles, formats numériques, ligne de total en gras.</summary>
    Private Sub AfficherDetail(grille As DataGridView, table As DataTable,
                               colonneCle As String, intituleCle As String)

        grille.DataSource = Nothing
        grille.DataSource = table

        DefinirEntete(grille, colonneCle, intituleCle)
        DefinirEntete(grille, "Designation", "Désignation")
        DefinirEntete(grille, "PointsDeVente", "Points de vente")
        DefinirEntete(grille, "NbEnvois", "Envois")
        DefinirEntete(grille, "PrincipalEnvoi", "Principal envoyé")
        DefinirEntete(grille, "NbPaiements", "Paiements")
        DefinirEntete(grille, "PrincipalPaye", "Principal payé")
        DefinirEntete(grille, "TotalTaxes", "Total taxes")

        For Each nom As String In New String() {"NbEnvois", "NbPaiements", "PointsDeVente",
                                                "PrincipalEnvoi", "PrincipalPaye", "Commissions",
                                                "TVA", "TTA", "TotalTaxes"}
            If Not grille.Columns.Contains(nom) Then Continue For
            grille.Columns(nom).DefaultCellStyle.Format = "N0"
            grille.Columns(nom).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next

        grille.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' La ligne de total ferme chaque page : elle doit se distinguer au premier coup d'oeil.
        For Each ligne As DataGridViewRow In grille.Rows
            If String.Equals(Convert.ToString(ligne.Cells(0).Value),
                             RapportActiviteService.LIBELLE_TOTAL, StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(grille.Font, Drawing.FontStyle.Bold)
                ligne.DefaultCellStyle.BackColor = Drawing.Color.Gainsboro
            End If
        Next
    End Sub

    Private Shared Sub DefinirEntete(grille As DataGridView, nomColonne As String, intitule As String)
        If grille.Columns.Contains(nomColonne) Then
            grille.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

#End Region

#Region "Export Excel"

    ''' <summary>
    ''' Exporte les quatre états dans un même classeur, l'un sous l'autre, avec le même titre
    ''' et la même mise en page que les autres états de l'application.
    ''' </summary>
    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        If _lignes Is Nothing OrElse _lignes.Count = 0 Then
            MessageBox.Show("Aucune donnée à exporter.", "Export Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        sfdExport.FileName = $"RapportActivite_{dtpDebut.Value:yyyyMMdd}_{dtpFin.Value:yyyyMMdd}.xlsx"
        If sfdExport.ShowDialog(Me) <> DialogResult.OK Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim sousTitres As New List(Of SousTitreExcel) From {
                New SousTitreExcel($"Période du {dtpDebut.Value:dd/MM/yyyy} au {dtpFin.Value:dd/MM/yyyy}", True),
                New SousTitreExcel($"{RapportActiviteService.CompterJours(_lignes)} journée(s) comptabilisée(s) — " &
                                   $"{RapportActiviteService.CompterPointsDeVente(_lignes)} point(s) de vente"),
                New SousTitreExcel($"Édité le {Date.Now:dd/MM/yyyy à HH:mm}")
            }

            ExcelExportService.ExporterEtOuvrir(
                "ECOBANK TCHAD — RAPPORT D'ACTIVITÉ WESTERN UNION",
                sousTitres,
                ConstruireBlocs(),
                "Rapport activité",
                sfdExport.FileName)

            lblStatut.Text = $"Rapport exporté vers {sfdExport.FileName}."

        Catch ex As InvalidOperationException
            MessageBox.Show(ex.Message, "Export Excel impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As IO.IOException
            MessageBox.Show(
                $"Écriture du fichier impossible : {ex.Message}" & Environment.NewLine & Environment.NewLine &
                "Le classeur est peut-être déjà ouvert dans Excel.",
                "Export Excel impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As Runtime.InteropServices.COMException
            MessageBox.Show($"Microsoft Excel a signalé une erreur : {ex.Message}",
                            "Export Excel impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>Les quatre états, dans l'ordre des onglets, avec leurs intitulés et formats.</summary>
    Private Function ConstruireBlocs() As List(Of BlocExcel)

        Dim formatsChiffres As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"NbEnvois", "# ##0"}, {"NbPaiements", "# ##0"}, {"PointsDeVente", "# ##0"},
            {"PrincipalEnvoi", "# ##0"}, {"PrincipalPaye", "# ##0"}, {"Commissions", "# ##0"},
            {"TVA", "# ##0"}, {"TTA", "# ##0"}, {"TotalTaxes", "# ##0"}
        }

        Dim entetes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Designation", "Désignation"}, {"PointsDeVente", "Points de vente"},
            {"NbEnvois", "Envois"}, {"PrincipalEnvoi", "Principal envoyé"},
            {"NbPaiements", "Paiements"}, {"PrincipalPaye", "Principal payé"},
            {"TotalTaxes", "Total taxes"}
        }

        Dim synthese As New BlocExcel("1. Synthèse de la période", _synthese)
        synthese.Formats = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {{"Valeur", "# ##0"}}

        Dim parJour As New BlocExcel("2. Jour par jour", _parJour) With {
            .Entetes = entetes, .Formats = formatsChiffres,
            .ExergueColonne = "Date", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        ' Le filtre automatique va à la page des points de vente : c'est celle que l'on fouille.
        Dim parPdv As New BlocExcel("3. Par point de vente", _parPdv) With {
            .Entetes = entetes, .Formats = formatsChiffres, .AvecFiltre = True,
            .ExergueColonne = "Account", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        Dim parGroupe As New BlocExcel("4. Par groupe statistique", _parGroupe) With {
            .Entetes = entetes, .Formats = formatsChiffres,
            .ExergueColonne = "Groupe", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        Return New List(Of BlocExcel) From {synthese, parJour, parPdv, parGroupe}
    End Function

#End Region

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

End Class
