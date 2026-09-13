Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Gestion des groupes statistiques (table T_GroupeStatistique) : consultation, création,
''' modification et suppression.
'''
''' Le groupe est l'unité de paramétrage : il porte le compte d'activité, le compte de
''' commission et le taux dont chaque sous-agent hérite. Les deux index uniques posés sur les
''' comptes rendent impossible qu'un compte appartienne à deux groupes — la règle est désormais
''' garantie par la base, et non plus seulement vérifiée par l'application.
'''
''' Les colonnes correspondantes de T_Pdv_SA sont conservées et tenues synchronisées : la
''' comptabilisation quotidienne continue de les lire, comme d'autres applications peuvent le
''' faire. Toute modification d'un groupe est donc reportée sur ses sous-agents.
''' </summary>
Public Class FrmGroupesStatistiques

    Private _liste As List(Of GroupeStatistiqueWU)
    Private _enCreation As Boolean = False
    Private _chargementEnCours As Boolean = False

    ''' <summary>Libellé à pré-remplir à l'ouverture, lorsque le formulaire est appelé pour créer un groupe précis.</summary>
    Private ReadOnly _nomAPreremplir As String

    ''' <summary>Vrai si au moins une demande a été déposée pendant la session du formulaire.</summary>
    Public ReadOnly Property ModificationEnregistree As Boolean
        Get
            Return _modifie
        End Get
    End Property
    Private _modifie As Boolean = False

    Public Sub New()
        Me.New(String.Empty)
    End Sub

    ''' <summary>
    ''' Ouvre le formulaire directement en création, avec un libellé pré-rempli. Utilisé depuis
    ''' l'écran des sous-agents lorsque le groupe saisi n'existe pas encore : l'utilisateur n'a
    ''' plus qu'à renseigner les comptes et le taux.
    ''' </summary>
    Public Sub New(nomAPreremplir As String)
        InitializeComponent()
        _nomAPreremplir = If(nomAPreremplir, String.Empty).Trim()
    End Sub

#Region "Chargement et affichage de la liste"

    Private Sub FrmGroupesStatistiques_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesPointsDeVente Then
            MessageBox.Show("Le paramétrage des groupes statistiques est réservé aux commerciaux et aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        AjusterAuxDroits()


        ChargerListe()

        If _nomAPreremplir.Length > 0 Then
            PreparerCreation(_nomAPreremplir)
        End If
    End Sub

    ''' <summary>
    ''' N'ouvre la saisie qu'à un inputer. Les autres gardent l'écran en consultation :
    ''' pouvoir relire le référentiel n'est pas pouvoir le changer.
    ''' </summary>
    Private Sub AjusterAuxDroits()

        Dim saisit As Boolean = SessionWU.PeutSaisirLesPointsDeVente

        btnNouveau.Enabled = saisit
        btnEnregistrer.Enabled = saisit
        btnSupprimer.Enabled = saisit
        btnSynchroniser.Enabled = saisit

        If saisit Then
            lblStatut.Text = "Toute saisie part en attente : elle prendra effet après autorisation."
        Else
            lblStatut.Text = "Consultation seule : la saisie est réservée à la fonction « inputer »."
        End If
    End Sub

    Private Sub ChargerListe(Optional nomASelectionner As String = Nothing)

        Cursor = Cursors.WaitCursor
        _chargementEnCours = True

        Try
            Dim messageErreur As String = String.Empty
            _liste = PdvRepository.ListerGroupes(txtRecherche.Text, messageErreur)

            dgvListe.DataSource = Nothing
            dgvListe.DataSource = _liste
            RenommerColonnes()

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = "Groupes non lus : voir le message affiché."
                lblNombre.Text = String.Empty
                MessageBox.Show(messageErreur, "Groupes statistiques", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                lblNombre.Text = $"{_liste.Count} groupe(s)"
                lblStatut.Text = If(_liste.Count = 0,
                                    "Aucun groupe ne correspond à la recherche.",
                                    "Sélectionnez un groupe pour le modifier, ou cliquez sur « Nouveau ».")
                SignalerDesynchronisation()
            End If

        Finally
            _chargementEnCours = False
            Cursor = Cursors.Default
        End Try

        If Not String.IsNullOrEmpty(nomASelectionner) Then
            SelectionnerNom(nomASelectionner)
        End If

        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>
    ''' Signale les groupes dont des sous-agents ne portent plus les valeurs du groupe. La
    ''' comptabilisation lisant les colonnes de T_Pdv_SA, une telle dérive produirait des
    ''' écritures sur des comptes qui ne sont plus ceux du groupe.
    ''' </summary>
    Private Sub SignalerDesynchronisation()

        Dim desynchronises As Integer = 0
        Dim herites As Integer = 0

        For Each groupe As GroupeStatistiqueWU In _liste
            If Not groupe.EstEnregistre Then
                herites += 1
            ElseIf groupe.EstDesynchronise Then
                desynchronises += 1
            End If
        Next

        If herites > 0 Then
            lblStatut.Text = $"{herites} groupe(s) hérité(s) de l'ancien paramétrage (en jaune), pas encore " &
                             "enregistré(s) : sélectionnez-les et cliquez sur « Enregistrer » pour les reprendre."
            Return
        End If

        If desynchronises = 0 Then Return

        lblStatut.Text = $"ATTENTION — {desynchronises} groupe(s) dont des sous-agents ne portent plus " &
                         "les valeurs du groupe. Sélectionnez-les et cliquez sur « Synchroniser les sous-agents »."
    End Sub

    Private Sub RenommerColonnes()

        DefinirEntete("Nom", "Groupe")
        DefinirEntete("CompteActivite", "Compte d'activité")
        DefinirEntete("CompteCommission", "Compte de commission")
        DefinirEntete("Taux", "Taux")
        DefinirEntete("NombreSousAgents", "Sous-agents")
        DefinirEntete("NombreDesynchronises", "Désynchronisés")
        DefinirEntete("EstEnregistre", "Enregistré")

        ' Propriété calculée, sans intérêt dans la grille : la colonne « Désynchronisés » la dit mieux.
        If dgvListe.Columns.Contains("EstDesynchronise") Then
            dgvListe.Columns("EstDesynchronise").Visible = False
        End If

        If dgvListe.Columns.Contains("Taux") Then
            dgvListe.Columns("Taux").DefaultCellStyle.Format = "P0"
            dgvListe.Columns("Taux").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvListe.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        MettreEnEvidenceDesynchronises()
    End Sub

    ''' <summary>
    ''' Met en évidence les lignes demandant une action : en jaune les groupes hérités de
    ''' l'ancien paramétrage, pas encore enregistrés dans la table ; en rose ceux dont des
    ''' sous-agents ne portent plus les valeurs du groupe.
    ''' </summary>
    Private Sub MettreEnEvidenceDesynchronises()

        For Each ligne As DataGridViewRow In dgvListe.Rows
            Dim groupe As GroupeStatistiqueWU = TryCast(ligne.DataBoundItem, GroupeStatistiqueWU)
            If groupe Is Nothing Then Continue For

            If Not groupe.EstEnregistre Then
                ligne.DefaultCellStyle.BackColor = Drawing.Color.LightGoldenrodYellow
            ElseIf groupe.EstDesynchronise Then
                ligne.DefaultCellStyle.BackColor = Drawing.Color.MistyRose
            End If
        Next
    End Sub

    Private Sub DefinirEntete(nomColonne As String, intitule As String)
        If dgvListe.Columns.Contains(nomColonne) Then
            dgvListe.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

    Private Sub SelectionnerNom(nom As String)

        For Each ligne As DataGridViewRow In dgvListe.Rows
            Dim groupe As GroupeStatistiqueWU = TryCast(ligne.DataBoundItem, GroupeStatistiqueWU)
            If groupe IsNot Nothing AndAlso String.Equals(groupe.Nom, nom, StringComparison.OrdinalIgnoreCase) Then
                ligne.Selected = True
                dgvListe.CurrentCell = ligne.Cells(0)
                Return
            End If
        Next
    End Sub

    Private Function GroupeSelectionne() As GroupeStatistiqueWU
        If dgvListe.CurrentRow Is Nothing Then Return Nothing
        Return TryCast(dgvListe.CurrentRow.DataBoundItem, GroupeStatistiqueWU)
    End Function

    Private Sub dgvListe_SelectionChanged(sender As Object, e As EventArgs) Handles dgvListe.SelectionChanged
        If _chargementEnCours Then Return
        _enCreation = False
        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    Private Sub AfficherFicheSelectionnee()

        If _enCreation Then Return

        Dim groupe As GroupeStatistiqueWU = GroupeSelectionne()

        If groupe Is Nothing Then
            ViderChamps()
            Return
        End If

        txtNom.Text = groupe.Nom
        txtCompteActivite.Text = groupe.CompteActivite
        txtCompteCommission.Text = groupe.CompteCommission
        txtTaux.Text = groupe.Taux.ToString("0.00", Globalization.CultureInfo.CurrentCulture)

        If Not groupe.EstEnregistre Then
            lblStatut.Text = $"Groupe hérité « {groupe.Nom} » : {groupe.NombreSousAgents} sous-agent(s), " &
                             "valeurs lues dans T_Pdv_SA. Cliquez sur « Enregistrer » pour le reprendre " &
                             "dans la table des groupes."
        ElseIf groupe.EstDesynchronise Then
            lblStatut.Text = $"Groupe « {groupe.Nom} » : {groupe.NombreSousAgents} sous-agent(s), dont " &
                             $"{groupe.NombreDesynchronises} ne portant plus les valeurs du groupe."
        Else
            lblStatut.Text = $"Groupe « {groupe.Nom} » : {groupe.NombreSousAgents} sous-agent(s) rattaché(s)."
        End If
    End Sub

    Private Sub ViderChamps()
        txtNom.Text = String.Empty
        txtCompteActivite.Text = String.Empty
        txtCompteCommission.Text = String.Empty
        txtTaux.Text = String.Empty
    End Sub

    ''' <summary>
    ''' Le libellé n'est saisissable qu'à la création : c'est la clé sous laquelle les
    ''' sous-agents se rattachent au groupe, la renommer les détacherait tous d'un coup.
    ''' </summary>
    Private Sub MettreAJourEtatBoutons()

        Dim groupe As GroupeStatistiqueWU = GroupeSelectionne()

        txtNom.ReadOnly = Not _enCreation
        btnSupprimer.Enabled = Not _enCreation AndAlso groupe IsNot Nothing AndAlso groupe.EstEnregistre
        btnSynchroniser.Enabled = Not _enCreation AndAlso groupe IsNot Nothing

        If _enCreation Then
            grpDetail.Text = "Nouveau groupe statistique"
        ElseIf groupe IsNot Nothing AndAlso Not groupe.EstEnregistre Then
            grpDetail.Text = "Groupe hérité — à enregistrer dans la table des groupes"
        Else
            grpDetail.Text = "Fiche du groupe statistique"
        End If
    End Sub

#End Region

#Region "Recherche"

    Private Sub txtRecherche_KeyDown(sender As Object, e As KeyEventArgs) Handles txtRecherche.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            _enCreation = False
            ChargerListe()
        End If
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        _enCreation = False
        ChargerListe()
    End Sub

#End Region

#Region "Création, modification, suppression"

    Private Sub btnNouveau_Click(sender As Object, e As EventArgs) Handles btnNouveau.Click
        PreparerCreation(String.Empty)
    End Sub

    ''' <summary>Bascule le formulaire en saisie d'un nouveau groupe, libellé éventuellement pré-rempli.</summary>
    Private Sub PreparerCreation(nom As String)

        _enCreation = True
        dgvListe.ClearSelection()
        ViderChamps()
        txtNom.Text = nom
        txtTaux.Text = "0,00"
        MettreAJourEtatBoutons()

        lblStatut.Text = "Nouveau groupe : renseignez son compte d'activité, son compte de commission et son taux."

        If nom.Length > 0 Then
            txtCompteActivite.Focus()
        Else
            txtNom.Focus()
        End If
    End Sub

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim taux As Decimal = 0D
        If Not String.IsNullOrWhiteSpace(txtTaux.Text) AndAlso
           Not PointDeVenteSA.EssayerAnalyserTaux(txtTaux.Text, taux) Then

            MessageBox.Show(
                $"Le taux saisi (« {txtTaux.Text} ») n'est pas un nombre." & Environment.NewLine & Environment.NewLine &
                "Saisissez une fraction : 0,70 pour 70 %.",
                "Taux incorrect", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTaux.Focus()
            Return
        End If

        Dim groupe As New GroupeStatistiqueWU() With {
            .Nom = txtNom.Text,
            .CompteActivite = txtCompteActivite.Text,
            .CompteCommission = txtCompteCommission.Text,
            .Taux = taux
        }
        groupe.Normaliser()

        Dim anomalies As List(Of String) = groupe.Anomalies()
        If anomalies.Count > 0 Then
            MessageBox.Show(
                "Saisie incomplète ou incorrecte :" & Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", anomalies),
                "Groupe non enregistrable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Un groupe hérité de l'ancien paramétrage n'a pas encore de ligne dans la table :
        ' l'enregistrer, c'est l'y INSÉRER, pas modifier une ligne inexistante.
        Dim existant As GroupeStatistiqueWU = GroupeSelectionne()
        Dim insertion As Boolean = _enCreation OrElse (existant IsNot Nothing AndAlso Not existant.EstEnregistre)

        ' Modifier un groupe change les comptes et le taux de TOUS ses sous-agents : la portée
        ' réelle de l'opération doit être annoncée avant, pas découverte après.
        If Not insertion AndAlso existant IsNot Nothing AndAlso existant.NombreSousAgents > 0 Then

            Dim confirmation As DialogResult = MessageBox.Show(
                $"Les {existant.NombreSousAgents} sous-agent(s) du groupe « {existant.Nom} » hériteront " &
                "des nouvelles valeurs :" & Environment.NewLine & Environment.NewLine &
                $"    compte d'activité    : {existant.CompteActivite}  ->  {groupe.CompteActivite}" & Environment.NewLine &
                $"    compte de commission : {existant.CompteCommission}  ->  {groupe.CompteCommission}" & Environment.NewLine &
                $"    taux                 : {existant.Taux:0.00}  ->  {groupe.Taux:0.00}" & Environment.NewLine &
                Environment.NewLine &
                "Les pièces comptables générées ensuite utiliseront ces comptes pour tout le groupe." &
                Environment.NewLine & "Confirmer ?",
                "Modifier le groupe", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

            If confirmation <> DialogResult.Yes Then Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            ' Le groupe ne part pas dans T_GroupeStatistique mais dans la file des demandes.
            ' L'alignement des sous-agents qui suivait aussitôt l'enregistrement se fera à
            ' l'autorisation, dans la même transaction : le groupe et son miroir dans T_Pdv_SA
            ' ne peuvent pas diverger.
            Dim demande As DemandeWU = DemandeWU.DepuisGroupe(
                groupe, If(insertion, OperationWU.Creation, OperationWU.Modification))

            If Not DemandeRepository.Soumettre(demande, messageErreur) Then
                MessageBox.Show(messageErreur, "Demande non déposée", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _modifie = True
            _enCreation = False
            lblStatut.Text = $"{demande.Intitule} — demande déposée, en attente d'autorisation."

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe(groupe.Nom)
    End Sub

    Private Sub btnSupprimer_Click(sender As Object, e As EventArgs) Handles btnSupprimer.Click

        Dim groupe As GroupeStatistiqueWU = GroupeSelectionne()
        If groupe Is Nothing Then Return

        If groupe.NombreSousAgents > 0 Then
            MessageBox.Show(
                $"Le groupe « {groupe.Nom} » ne peut pas être supprimé : {groupe.NombreSousAgents} " &
                "sous-agent(s) y sont rattachés." & Environment.NewLine & Environment.NewLine &
                "Rattachez-les d'abord à un autre groupe depuis l'écran des sous-agents.",
                "Suppression impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim reponse As DialogResult = MessageBox.Show(
            $"Supprimer définitivement le groupe « {groupe.Nom} » ?" & Environment.NewLine & Environment.NewLine &
            "Aucun sous-agent n'y est rattaché : la suppression est sans effet sur la comptabilisation." &
            Environment.NewLine & "Cette opération est irréversible.",
            "Confirmer la suppression", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            Dim demande As DemandeWU = DemandeWU.DepuisGroupe(groupe, OperationWU.Suppression)

            If Not DemandeRepository.Soumettre(demande, messageErreur) Then
                MessageBox.Show(messageErreur, "Demande non déposée", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _modifie = True
            lblStatut.Text = $"Groupe « {groupe.Nom} » supprimé."

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe()
    End Sub

    ''' <summary>
    ''' Réaligne les colonnes de T_Pdv_SA sur les valeurs du groupe. Utile après une écriture
    ''' directe en base ou une synchronisation interrompue, que la colonne « Désynchronisés »
    ''' de la grille signale.
    ''' </summary>
    Private Sub btnSynchroniser_Click(sender As Object, e As EventArgs) Handles btnSynchroniser.Click

        Dim groupe As GroupeStatistiqueWU = GroupeSelectionne()
        If groupe Is Nothing Then Return

        Dim reponse As DialogResult = MessageBox.Show(
            $"Appliquer les valeurs du groupe « {groupe.Nom} » à ses {groupe.NombreSousAgents} sous-agent(s) ?" &
            Environment.NewLine & Environment.NewLine &
            $"    compte d'activité    : {groupe.CompteActivite}" & Environment.NewLine &
            $"    compte de commission : {groupe.CompteCommission}" & Environment.NewLine &
            $"    taux                 : {groupe.Taux:0.00}" & Environment.NewLine & Environment.NewLine &
            "Les valeurs actuellement portées par ces sous-agents seront écrasées, " &
            "une fois la demande autorisée.",
            "Synchroniser les sous-agents", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            ' La synchronisation réécrit les comptes et le taux de TOUS les sous-agents du
            ' groupe d'un seul coup : c'est l'écriture la plus lourde de l'application, et donc
            ' celle qu'il serait le moins acceptable de laisser passer sans second regard.
            Dim demande As DemandeWU = DemandeWU.DepuisGroupe(groupe, OperationWU.Synchronisation)

            If Not DemandeRepository.Soumettre(demande, messageErreur) Then
                MessageBox.Show(messageErreur, "Demande non déposée", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _modifie = True
            lblStatut.Text = $"Synchronisation du groupe « {groupe.Nom} » — demande déposée, " &
                             "en attente d'autorisation."

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe(groupe.Nom)
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
