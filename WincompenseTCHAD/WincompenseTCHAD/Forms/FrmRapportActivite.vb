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

    ''' <summary>Entrée du filtre affichant l'ensemble des groupes.</summary>
    Private Const TOUS_LES_GROUPES As String = "(tous les groupes)"

    ''' <summary>
    ''' Nombre de transactions au-delà duquel l'inclusion du détail dans le PDF est soumise à
    ''' confirmation : une vingtaine de pages, soit environ une semaine d'activité.
    ''' </summary>
    Private Const SEUIL_DETAIL_VOLUMINEUX As Integer = 2000

    ''' <summary>Lignes d'historique de la période, TOUS groupes confondus.</summary>
    Private _lignes As New List(Of LigneHistoriqueWU)

    ''' <summary>
    ''' Lignes effectivement restituées : les précédentes, restreintes au groupe choisi.
    ''' Les états et l'export travaillent sur elles, jamais sur _lignes : l'écran et le
    ''' classeur portent ainsi toujours sur le même périmètre.
    ''' </summary>
    Private _lignesAffichees As New List(Of LigneHistoriqueWU)

    ''' <summary>Groupe statistique retenu, ou chaîne vide pour tous les groupes.</summary>
    Private _groupeChoisi As String = String.Empty

    ''' <summary>
    ''' Agences propres du paramétrage, lues une fois à l'ouverture : elles complètent l'état
    ''' par point de vente avec celles restées sans activité sur la période.
    ''' </summary>
    Private _agences As New List(Of PointDeVenteEC)

    ''' <summary>Les cinq états, conservés pour l'export : ce qui part dans Excel est ce qui est à l'écran.</summary>
    Private _synthese As DataTable
    Private _parJour As DataTable
    Private _parPdv As DataTable
    Private _parGroupe As DataTable
    Private _commissions As DataTable

    ''' <summary>
    ''' Détail des transactions de la période, rattaché aux points de vente dans l'état
    ''' correspondant : dérouler un Account fait apparaître ses envois et ses paiements.
    ''' </summary>
    Private _operations As New List(Of TransactionWU)

    ''' <summary>Empêche le filtre de relancer l'affichage pendant qu'on le remplit.</summary>
    Private _chargementEnCours As Boolean = False

    ''' <summary>L'absence de détail n'est expliquée qu'une fois par ouverture du rapport.</summary>
    Private _detailSignale As Boolean = False

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
        ' Nouvelle période demandée : l'absence de détail mérite d'être signalée à nouveau.
        _detailSignale = False
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

            ' Le paramétrage des agences propres est relu avec la période : une agence créée
            ' entre-temps doit apparaître, même sans activité.
            Dim erreurAgences As String = String.Empty
            _agences = PdvRepository.ListerAgences(String.Empty, erreurAgences)

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = "Lecture impossible : voir le message affiché."
                MessageBox.Show(messageErreur, "Historique", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            RemplirListeGroupes()
            ConstruireEtats()

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Explique l'absence de détail plutôt que de laisser l'utilisateur devant des points de
    ''' vente qui ne se déroulent pas.
    '''
    ''' Deux causes, très différentes : la table du détail n'existe pas encore, ou les journées
    ''' ont été comptabilisées AVANT la mise en place du suivi des MTCN — leur agrégat est bien
    ''' là, mais leur détail n'a jamais été écrit. Le message dit laquelle, et quoi faire.
    ''' </summary>
    Private Sub SignalerDetailIndisponible(erreurDetail As String)

        ' Le détail est là : rien à signaler.
        If _operations.Count > 0 Then Return

        ' Aucune activité du tout : l'absence de détail va de soi.
        If _lignesAffichees.Count = 0 Then Return

        ' L'avertissement n'est donné qu'une fois par ouverture du rapport : le répéter à chaque
        ' changement de groupe deviendrait vite insupportable.
        If _detailSignale Then
            lblStatut.Text &= "  |  Aucun détail de transaction sur cette période."
            Return
        End If

        _detailSignale = True

        If Not String.IsNullOrEmpty(erreurDetail) Then
            MessageBox.Show(erreurDetail, "Détail des transactions indisponible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            lblStatut.Text &= "  |  Détail des transactions indisponible."
            Return
        End If

        MessageBox.Show(
            "Aucun détail de transaction n'est enregistré pour cette période." &
            Environment.NewLine & Environment.NewLine &
            "Les points de vente ne pourront donc pas être déroulés : leurs cumuls sont bien là, " &
            "mais les MTCN qui les composent n'ont jamais été écrits." & Environment.NewLine &
            Environment.NewLine &
            "Cause la plus fréquente : ces journées ont été comptabilisées AVANT la mise en place " &
            "du suivi des MTCN." & Environment.NewLine & Environment.NewLine &
            "Pour l'obtenir :" & Environment.NewLine &
            "   1. exécuter le script Scripts\06_HistoriqueMTCN.sql s'il ne l'a pas encore été ;" & Environment.NewLine &
            "   2. recharger les rapports de ces journées et regénérer leur pièce comptable." & Environment.NewLine &
            Environment.NewLine &
            "Une journée regénérée remplace proprement la précédente : l'opération est sans risque.",
            "Aucun détail de transaction", MessageBoxButtons.OK, MessageBoxIcon.Information)

        lblStatut.Text &= "  |  Aucun détail : journées comptabilisées avant le suivi des MTCN."
    End Sub

    ''' <summary>
    ''' Alimente le filtre avec les groupes réellement présents dans la période — lus dans
    ''' l'historique et non dans le paramétrage courant : un groupe supprimé depuis reste
    ''' présent dans l'historique et doit rester consultable.
    ''' </summary>
    Private Sub RemplirListeGroupes()

        Dim selectionPrecedente As String = _groupeChoisi

        _chargementEnCours = True
        Try
            cboGroupe.Items.Clear()
            cboGroupe.Items.Add(TOUS_LES_GROUPES)

            For Each groupe As String In RapportActiviteService.ListerGroupesPresents(_lignes)
                cboGroupe.Items.Add(groupe)
            Next

            ' La sélection est conservée d'une période à l'autre lorsque le groupe y figure encore.
            Dim index As Integer = If(String.IsNullOrEmpty(selectionPrecedente), 0, cboGroupe.Items.IndexOf(selectionPrecedente))
            cboGroupe.SelectedIndex = If(index >= 0, index, 0)
            _groupeChoisi = If(cboGroupe.SelectedIndex = 0, String.Empty, Convert.ToString(cboGroupe.SelectedItem))

        Finally
            _chargementEnCours = False
        End Try
    End Sub

    Private Sub cboGroupe_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGroupe.SelectedIndexChanged

        If _chargementEnCours Then Return

        _groupeChoisi = If(cboGroupe.SelectedIndex <= 0, String.Empty, Convert.ToString(cboGroupe.SelectedItem))
        ConstruireEtats()
    End Sub

    ''' <summary>Construit et affiche les cinq états sur le périmètre retenu.</summary>
    Private Sub ConstruireEtats()

        _lignesAffichees = RapportActiviteService.Filtrer(_lignes, _groupeChoisi)

        _synthese = RapportActiviteService.ConstruireSynthese(_lignesAffichees)
        _parJour = RapportActiviteService.ConstruireParJour(_lignesAffichees)
        ' Les agences propres sans activité doivent figurer à zéro : elles sont lues dans le
        ' paramétrage, l'historique ne pouvant évidemment pas contenir ce qui n'a pas eu lieu.
        ' Un filtre par groupe ne les concerne pas : une agence propre n'appartient à aucun
        ' groupe statistique, ceux-ci ne s'appliquant qu'aux sous-agents.
        Dim agences As List(Of PointDeVenteEC) = If(String.IsNullOrEmpty(_groupeChoisi), _agences, Nothing)

        ' Détail des transactions : il ne se déduit pas de l'agrégat, et le filtre par groupe
        ' est appliqué par la base plutôt que de rapatrier toute la période pour la trier.
        Dim erreurDetail As String = String.Empty
        _operations = HistoriqueRepository.ListerTransactions(dtpDebut.Value.Date, dtpFin.Value.Date,
                                                             _groupeChoisi, erreurDetail)

        If Not String.IsNullOrEmpty(erreurDetail) Then
            _operations = New List(Of TransactionWU)
        End If

        _parPdv = RapportActiviteService.ConstruireParPointDeVente(_lignesAffichees, agences, _operations)
        _parGroupe = RapportActiviteService.ConstruireParGroupe(_lignesAffichees)
        _commissions = RapportActiviteService.ConstruireEvolutionCommissions(_lignesAffichees)

        AfficherSynthese()
        AfficherDetail(dgvParJour, _parJour, "Date", "Date")
        AfficherDetail(dgvParPdv, _parPdv, "Account", "Account")
        AfficherDetail(dgvParGroupe, _parGroupe, "Groupe", "Groupe")
        AfficherDetail(dgvCommissions, _commissions, "Date", "Date")

        ReplierToutLeDetail()

        btnExporter.Enabled = _lignesAffichees.Count > 0

        If _lignesAffichees.Count = 0 Then
            lblStatut.Text = If(String.IsNullOrEmpty(_groupeChoisi),
                                "Aucune journée comptabilisée sur cette période.",
                                $"Aucune activité du groupe « {_groupeChoisi} » sur cette période.")
            Return
        End If

        lblStatut.Text = $"{RapportActiviteService.CompterJours(_lignesAffichees)} journée(s), " &
                         $"{RapportActiviteService.CompterPointsDeVente(_lignesAffichees)} point(s) de vente" &
                         If(String.IsNullOrEmpty(_groupeChoisi), String.Empty, $" — groupe « {_groupeChoisi} »") &
                         $", {_operations.Count} transaction(s) détaillée(s)."

        SignalerDetailIndisponible(erreurDetail)
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
        DefinirEntete(grille, "NbAnnulations", "Annulations")
        DefinirEntete(grille, "Montant", "Montant")
        DefinirEntete(grille, "TotalTaxes", "Total taxes")
        DefinirEntete(grille, "CommissionEnvoi", "Commission Envoi")
        DefinirEntete(grille, "CommissionPaiement", "Commission Paiement")
        DefinirEntete(grille, "CommissionTransfert", "Commission Transfert")
        DefinirEntete(grille, "TotalCommissions", "Total commissions")
        DefinirEntete(grille, "Variation", "Variation / veille")
        DefinirEntete(grille, "Cumul", "Cumul période")

        For Each nom As String In New String() {"NbEnvois", "NbPaiements", "NbAnnulations", "PointsDeVente",
                                                "PrincipalEnvoi", "PrincipalPaye", "Commissions",
                                                "TVA", "TTA", "TotalTaxes",
                                                "CommissionEnvoi", "CommissionPaiement",
                                                "CommissionTransfert", "TotalCommissions", "Cumul",
                                                "Montant"}
            If Not grille.Columns.Contains(nom) Then Continue For
            grille.Columns(nom).DefaultCellStyle.Format = "N0"
            grille.Columns(nom).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next

        ' La variation est une proportion : elle s'affiche en pourcentage signé.
        If grille.Columns.Contains("Variation") Then
            grille.Columns("Variation").DefaultCellStyle.Format = "+0.0 %;-0.0 %;0.0 %"
            grille.Columns("Variation").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        DefinirEntete(grille, "MTCN", "MTCN")
        DefinirEntete(grille, "Statut", "Statut")

        ' Colonne technique : elle pilote le pliage, elle n'a rien à faire à l'écran.
        If grille.Columns.Contains("Niveau") Then
            grille.Columns("Niveau").Visible = False
        End If

        grille.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' Le tri brouillerait l'ordre hiérarchique : une transaction se retrouverait détachée
        ' de son point de vente.
        If grille.Columns.Contains("Deroule") Then
            For Each colonne As DataGridViewColumn In grille.Columns
                colonne.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            With grille.Columns("Deroule")
                .Visible = True
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .MinimumWidth = 28
                .Width = 28
                .HeaderText = "+/-"
                .ToolTipText = "Cliquez un point de vente pour dérouler ses transactions"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Font = New Drawing.Font(grille.Font, Drawing.FontStyle.Bold)
            End With
        End If

        MettreEnEvidenceTotaux(grille)
        MettreEnEvidenceNiveaux(grille)
    End Sub

    ''' <summary>
    ''' Distingue les lignes de détail des lignes de point de vente, et signale en rose les
    ''' transactions annulées : ce sont elles que l'on cherche le plus souvent.
    ''' </summary>
    Private Shared Sub MettreEnEvidenceNiveaux(grille As DataGridView)

        If Not grille.Columns.Contains("Niveau") Then Return

        For Each ligne As DataGridViewRow In grille.Rows

            Dim valeur As Object = ligne.Cells("Niveau").Value
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Continue For

            If Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) <>
               RapportActiviteService.NIVEAU_TRANSACTION Then Continue For

            ligne.DefaultCellStyle.ForeColor = Drawing.SystemColors.GrayText

            If String.Equals(Convert.ToString(ligne.Cells("Statut").Value), "ANNULÉE", StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.BackColor = Drawing.Color.MistyRose
                ligne.DefaultCellStyle.ForeColor = Drawing.SystemColors.ControlText
            End If
        Next
    End Sub

    ''' <summary>
    ''' Distingue les lignes de structure : total général, sous-totaux par nature de point de
    ''' vente, et intitulés de catégorie — ces derniers étant des lignes sans chiffres.
    ''' </summary>
    Private Shared Sub MettreEnEvidenceTotaux(grille As DataGridView)

        For Each ligne As DataGridViewRow In grille.Rows

            Dim cle As String = Convert.ToString(ligne.Cells(0).Value)

            If String.Equals(cle, RapportActiviteService.LIBELLE_TOTAL, StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(grille.Font, Drawing.FontStyle.Bold)
                ligne.DefaultCellStyle.BackColor = Drawing.Color.Gainsboro

            ElseIf String.Equals(cle, RapportActiviteService.LIBELLE_SOUS_TOTAL, StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(grille.Font, Drawing.FontStyle.Bold)
                ligne.DefaultCellStyle.BackColor = Drawing.Color.WhiteSmoke

            ElseIf String.Equals(cle, RapportActiviteService.TYPE_SOUS_AGENT, StringComparison.Ordinal) OrElse
                   String.Equals(cle, RapportActiviteService.TYPE_AGENCE, StringComparison.Ordinal) OrElse
                   String.Equals(cle, RapportActiviteService.TYPE_NON_PARAMETRE, StringComparison.Ordinal) Then
                ligne.DefaultCellStyle.Font = New Drawing.Font(grille.Font, Drawing.FontStyle.Bold)
                ligne.DefaultCellStyle.BackColor = Drawing.Color.LightSteelBlue
            End If
        Next
    End Sub

    Private Shared Sub DefinirEntete(grille As DataGridView, nomColonne As String, intitule As String)
        If grille.Columns.Contains(nomColonne) Then
            grille.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

#End Region

#Region "Déroulé d'un point de vente"

    ''' <summary>
    ''' Replie tout le détail à l'affichage : un état de gestion se lit d'abord au niveau des
    ''' points de vente. Le détail ne s'ouvre que lorsqu'on le demande.
    ''' </summary>
    Private Sub ReplierToutLeDetail()

        ' Une ligne courante ne peut pas être masquée : on relâche la sélection avant de replier.
        dgvParPdv.CurrentCell = Nothing

        For Each ligne As DataGridViewRow In dgvParPdv.Rows
            If NiveauDeLaLigne(ligne) = RapportActiviteService.NIVEAU_TRANSACTION Then
                ligne.Visible = False
            End If
        Next
    End Sub

    ''' <summary>
    ''' Déroule ou replie un point de vente lorsqu'on clique dessus : ses envois et ses
    ''' paiements apparaissent alors sous lui, chacun avec son MTCN.
    ''' </summary>
    Private Sub dgvParPdv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvParPdv.CellClick

        If e.RowIndex < 0 Then Return

        Dim ligne As DataGridViewRow = dgvParPdv.Rows(e.RowIndex)
        If NiveauDeLaLigne(ligne) <> RapportActiviteService.NIVEAU_POINT_DE_VENTE Then Return

        Dim signe As String = Convert.ToString(ligne.Cells("Deroule").Value)
        If String.IsNullOrEmpty(signe) Then Return ' Point de vente sans détail : rien à dérouler.

        Dim ouvrir As Boolean = (signe = RapportActiviteService.SIGNE_PLIE)

        ' Une ligne courante ne peut pas être masquée : on ramène la sélection sur le point de
        ' vente lui-même avant de replier son détail.
        If Not ouvrir Then
            dgvParPdv.CurrentCell = ligne.Cells("Account")
        End If

        ' Le détail d'un point de vente est la suite ininterrompue des lignes de transaction
        ' qui le suivent : la suivante d'un autre niveau ferme le bloc.
        For index As Integer = e.RowIndex + 1 To dgvParPdv.Rows.Count - 1

            Dim suivante As DataGridViewRow = dgvParPdv.Rows(index)
            If NiveauDeLaLigne(suivante) <> RapportActiviteService.NIVEAU_TRANSACTION Then Exit For

            suivante.Visible = ouvrir
        Next

        ligne.Cells("Deroule").Value = If(ouvrir,
                                          RapportActiviteService.SIGNE_DEPLIE,
                                          RapportActiviteService.SIGNE_PLIE)
    End Sub

    ''' <summary>Niveau porté par une ligne, ou -1 si la colonne technique est absente.</summary>
    Private Shared Function NiveauDeLaLigne(ligne As DataGridViewRow) As Integer

        If ligne Is Nothing OrElse Not ligne.DataGridView.Columns.Contains("Niveau") Then Return -1

        Dim valeur As Object = ligne.Cells("Niveau").Value
        If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return -1

        Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture)
    End Function

#End Region

#Region "Export PDF"


    ''' <summary>
    ''' Exporte les cinq états en PDF, l'un sous l'autre, avec le même titre et la même mise en
    ''' page que les autres états de l'application.
    '''
    ''' Le PDF a été retenu contre le classeur Excel pour figer l'état édité : il ne s'ouvre pas
    ''' dans un tableur et ne se retouche pas au fil de l'eau. Il est produit PAR Excel, à partir
    ''' d'un classeur invisible et jamais enregistré : aucun fichier intermédiaire ne subsiste.
    ''' </summary>
    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        If _lignesAffichees Is Nothing OrElse _lignesAffichees.Count = 0 Then
            MessageBox.Show("Aucune donnée à exporter.", "Export PDF",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Le détail des transactions peut représenter des milliers de lignes, soit des dizaines
        ' de pages imprimées. Sur un mois complet, l'inclure sans le dire produirait un document
        ' que personne n'attendait : la question est posée, et seulement quand le volume le mérite.
        Dim inclureDetail As Boolean = _operations IsNot Nothing AndAlso _operations.Count > 0

        If inclureDetail AndAlso _operations.Count > SEUIL_DETAIL_VOLUMINEUX Then

            Dim reponse As DialogResult = MessageBox.Show(
                $"Le détail des transactions compte {_operations.Count:N0} lignes, " &
                "soit plusieurs dizaines de pages." & Environment.NewLine & Environment.NewLine &
                "L'inclure sous chaque point de vente dans le document PDF ?" & Environment.NewLine & Environment.NewLine &
                "Répondre « Non » produit le rapport au niveau des points de vente seulement ; " &
                "le détail reste consultable à l'écran.",
                "Détail des transactions volumineux", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            inclureDetail = (reponse = DialogResult.Yes)
        End If

        sfdExport.FileName = NomFichierPropose()
        If sfdExport.ShowDialog(Me) <> DialogResult.OK Then Return

        Cursor = Cursors.WaitCursor
        Try
            ' La période est toujours mise en exergue ; le groupe l'est aussi lorsqu'il en
            ' restreint le périmètre — c'est ce qui caractérise l'extraction.
            Dim sousTitres As New List(Of SousTitreExcel) From {
                New SousTitreExcel($"Période du {dtpDebut.Value:dd/MM/yyyy} au {dtpFin.Value:dd/MM/yyyy}", True),
                New SousTitreExcel(If(String.IsNullOrEmpty(_groupeChoisi),
                                      "Tous les groupes statistiques",
                                      $"Groupe statistique : {_groupeChoisi}"),
                                   Not String.IsNullOrEmpty(_groupeChoisi)),
                New SousTitreExcel($"{RapportActiviteService.CompterJours(_lignesAffichees)} journée(s) comptabilisée(s) — " &
                                   $"{RapportActiviteService.CompterPointsDeVente(_lignesAffichees)} point(s) de vente"),
                New SousTitreExcel($"Édité le {Date.Now:dd/MM/yyyy à HH:mm}")
            }

            ExcelExportService.ExporterEnPdf(
                "ECOBANK TCHAD — RAPPORT D'ACTIVITÉ WESTERN UNION",
                sousTitres,
                ConstruireBlocs(inclureDetail),
                "Rapport activité",
                sfdExport.FileName)

            lblStatut.Text = $"Rapport exporté vers {sfdExport.FileName}."

        Catch ex As InvalidOperationException
            ' Excel absent, ou refus d'Excel de produire le PDF : message déjà explicite.
            MessageBox.Show(ex.Message, "Export PDF impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As IO.IOException
            MessageBox.Show(
                $"Écriture du fichier impossible : {ex.Message}" & Environment.NewLine & Environment.NewLine &
                "Le document est peut-être déjà ouvert dans un lecteur PDF.",
                "Export PDF impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As UnauthorizedAccessException
            MessageBox.Show(
                $"Accès refusé au fichier : {ex.Message}",
                "Export PDF impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>Les états, dans l'ordre des onglets, avec leurs intitulés et formats.</summary>
    ''' <param name="inclureDetail">
    ''' Inclut, sous chaque point de vente, le détail de ses transactions — dont le volume peut
    ''' être élevé. Sinon l'état s'arrête au niveau des points de vente.
    ''' </param>
    Private Function ConstruireBlocs(inclureDetail As Boolean) As List(Of BlocExcel)

        Dim formatsChiffres As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"NbEnvois", "# ##0"}, {"NbPaiements", "# ##0"}, {"NbAnnulations", "# ##0"},
            {"PointsDeVente", "# ##0"},
            {"PrincipalEnvoi", "# ##0"}, {"PrincipalPaye", "# ##0"}, {"Commissions", "# ##0"},
            {"TVA", "# ##0"}, {"TTA", "# ##0"}, {"TotalTaxes", "# ##0"}
        }

        Dim entetes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Designation", "Désignation"}, {"PointsDeVente", "Points de vente"},
            {"NbEnvois", "Envois"}, {"PrincipalEnvoi", "Principal envoyé"},
            {"NbPaiements", "Paiements"}, {"PrincipalPaye", "Principal payé"},
            {"NbAnnulations", "Annulations"}, {"TotalTaxes", "Total taxes"},
            {"MTCN", "MTCN"}, {"Statut", "Statut"}
        }

        Dim synthese As New BlocExcel("1. Synthèse de la période", _synthese)
        synthese.Formats = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {{"Valeur", "# ##0"}}

        Dim parJour As New BlocExcel("2. Jour par jour", _parJour) With {
            .Entetes = entetes, .Formats = formatsChiffres,
            .ExergueColonne = "Date", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        ' Le filtre automatique va à la page des points de vente : c'est celle que l'on fouille.
        Dim parPdv As New BlocExcel("3. Par point de vente", TablePointDeVentePourExport(inclureDetail)) With {
            .Entetes = entetes, .Formats = formatsChiffres, .AvecFiltre = True,
            .ExergueColonne = "Account", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        Dim parGroupe As New BlocExcel("4. Par groupe statistique", _parGroupe) With {
            .Entetes = entetes, .Formats = formatsChiffres,
            .ExergueColonne = "Groupe", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        Dim formatsCommissions As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"CommissionEnvoi", "# ##0"}, {"CommissionPaiement", "# ##0"},
            {"CommissionTransfert", "# ##0"}, {"TotalCommissions", "# ##0"},
            {"Cumul", "# ##0"}, {"Variation", "+0,0 %;-0,0 %;0,0 %"}
        }

        Dim commissions As New BlocExcel("5. Évolution des commissions", _commissions) With {
            .Entetes = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"CommissionEnvoi", "Commission Envoi"}, {"CommissionPaiement", "Commission Paiement"},
                {"CommissionTransfert", "Commission Transfert"}, {"TotalCommissions", "Total commissions"},
                {"Variation", "Variation / veille"}, {"Cumul", "Cumul période"}
            },
            .Formats = formatsCommissions,
            .ExergueColonne = "Date", .ExergueValeur = RapportActiviteService.LIBELLE_TOTAL
        }

        Return New List(Of BlocExcel) From {synthese, parJour, parPdv, parGroupe, commissions}
    End Function

    ''' <summary>
    ''' État par point de vente destiné à l'export : les colonnes techniques sont retirées, et
    ''' les lignes de transaction le sont aussi lorsque le détail n'est pas demandé.
    ''' </summary>
    Private Function TablePointDeVentePourExport(inclureDetail As Boolean) As DataTable

        Dim exportable As DataTable = _parPdv.Copy()

        If Not inclureDetail Then
            For index As Integer = exportable.Rows.Count - 1 To 0 Step -1

                Dim valeur As Object = exportable.Rows(index)("Niveau")
                If valeur Is DBNull.Value Then Continue For

                If Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) =
                   RapportActiviteService.NIVEAU_TRANSACTION Then
                    exportable.Rows.RemoveAt(index)
                End If
            Next
        End If

        ' Colonnes techniques : elles pilotent l'affichage, elles n'ont rien à faire dans un état.
        For Each nom As String In New String() {"Niveau", "Deroule"}
            If exportable.Columns.Contains(nom) Then exportable.Columns.Remove(nom)
        Next

        Return exportable
    End Function

    ''' <summary>
    ''' Nom de fichier proposé : il porte la période et, le cas échéant, le groupe retenu, de
    ''' sorte que deux extractions successives ne s'écrasent pas.
    ''' </summary>
    Private Function NomFichierPropose() As String

        Dim partieGroupe As String = String.Empty

        If Not String.IsNullOrEmpty(_groupeChoisi) Then
            partieGroupe = _groupeChoisi
            ' Un libellé de groupe peut contenir des caractères interdits dans un nom de fichier.
            For Each interdit As Char In IO.Path.GetInvalidFileNameChars()
                partieGroupe = partieGroupe.Replace(interdit, "_"c)
            Next
            partieGroupe = "_" & partieGroupe
        End If

        Return $"RapportActivite{partieGroupe}_{dtpDebut.Value:yyyyMMdd}_{dtpFin.Value:yyyyMMdd}.pdf"
    End Function

#End Region

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

End Class
