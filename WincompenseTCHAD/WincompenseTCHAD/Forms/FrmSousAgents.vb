Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Gestion des sous-agents Western Union (table T_Pdv_SA) : consultation, création,
''' modification et suppression.
'''
''' Le formulaire n'exécute aucune requête SQL lui-même : tout passe par PdvRepository.
''' Il travaille en deux modes, matérialisés par _enCreation :
'''   - CONSULTATION/MODIFICATION : une fiche existante est sélectionnée dans la liste ;
'''     l'Account est affiché mais verrouillé — c'est la clé sous laquelle les rapports
'''     Western Union désignent le point de vente, la renommer romprait le lien avec
'''     l'historique. Pour changer un Account, il faut supprimer puis recréer la fiche ;
'''   - CRÉATION : les champs sont vides et l'Account est saisissable.
''' </summary>
Public Class FrmSousAgents

    ''' <summary>Liste actuellement affichée, telle que renvoyée par la base.</summary>
    Private _liste As List(Of PointDeVenteSA)

    ''' <summary>Vrai lorsque l'utilisateur saisit une NOUVELLE fiche (non encore enregistrée).</summary>
    Private _enCreation As Boolean = False

    ''' <summary>Empêche le remplissage des champs pendant un rechargement de la liste.</summary>
    Private _chargementEnCours As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Chargement et affichage de la liste"

    Private Sub FrmSousAgents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ChargerListe()
    End Sub

    ''' <summary>Recharge la liste depuis la base en appliquant le filtre de recherche courant.</summary>
    Private Sub ChargerListe(Optional codeASelectionner As String = Nothing)

        Cursor = Cursors.WaitCursor
        _chargementEnCours = True

        Try
            Dim messageErreur As String = String.Empty
            _liste = PdvRepository.ListerSousAgents(txtRecherche.Text, messageErreur)

            dgvListe.DataSource = Nothing
            dgvListe.DataSource = _liste
            RenommerColonnes()

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = messageErreur
                lblNombre.Text = String.Empty
                MessageBox.Show(messageErreur, "Base SQL Server", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                lblNombre.Text = $"{_liste.Count} sous-agent(s)"
                lblStatut.Text = If(_liste.Count = 0,
                                    "Aucun sous-agent ne correspond à la recherche.",
                                    "Sélectionnez une fiche pour la modifier, ou cliquez sur « Nouveau ».")
            End If

        Finally
            _chargementEnCours = False
            Cursor = Cursors.Default
        End Try

        ' Après un enregistrement, on revient sur la fiche concernée plutôt que sur la première :
        ' l'utilisateur garde ainsi sous les yeux ce qu'il vient de saisir.
        If Not String.IsNullOrEmpty(codeASelectionner) Then
            SelectionnerCode(codeASelectionner)
        End If

        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>Donne aux colonnes de la grille des intitulés lisibles plutôt que les noms de propriétés.</summary>
    Private Sub RenommerColonnes()

        DefinirEntete("CodePdv", "Account")
        DefinirEntete("Designation", "Désignation")
        DefinirEntete("GroupeStatistique", "Groupe statistique")
        DefinirEntete("Taux", "Taux")
        DefinirEntete("CompteCompense", "Compte compensation")
        DefinirEntete("CompteCommission", "Compte commission")
        DefinirEntete("CodeAgence", "Code agence")

        If dgvListe.Columns.Contains("Taux") Then
            dgvListe.Columns("Taux").DefaultCellStyle.Format = "P0"
            dgvListe.Columns("Taux").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvListe.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub DefinirEntete(nomColonne As String, intitule As String)
        If dgvListe.Columns.Contains(nomColonne) Then
            dgvListe.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

    ''' <summary>Sélectionne dans la grille la ligne portant l'Account indiqué.</summary>
    Private Sub SelectionnerCode(code As String)

        For Each ligne As DataGridViewRow In dgvListe.Rows
            Dim fiche As PointDeVenteSA = TryCast(ligne.DataBoundItem, PointDeVenteSA)
            If fiche IsNot Nothing AndAlso String.Equals(fiche.CodePdv, code, StringComparison.OrdinalIgnoreCase) Then
                ligne.Selected = True
                dgvListe.CurrentCell = ligne.Cells(0)
                Return
            End If
        Next
    End Sub

    ''' <summary>Fiche actuellement sélectionnée dans la grille, ou Nothing.</summary>
    Private Function FicheSelectionnee() As PointDeVenteSA

        If dgvListe.CurrentRow Is Nothing Then Return Nothing
        Return TryCast(dgvListe.CurrentRow.DataBoundItem, PointDeVenteSA)
    End Function

    Private Sub dgvListe_SelectionChanged(sender As Object, e As EventArgs) Handles dgvListe.SelectionChanged
        If _chargementEnCours Then Return
        _enCreation = False
        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>Recopie la fiche sélectionnée dans les champs de saisie.</summary>
    Private Sub AfficherFicheSelectionnee()

        If _enCreation Then Return

        Dim fiche As PointDeVenteSA = FicheSelectionnee()

        If fiche Is Nothing Then
            ViderChamps()
            Return
        End If

        txtCodePdv.Text = fiche.CodePdv
        txtDesignation.Text = fiche.Designation
        txtGroupeStatistique.Text = fiche.GroupeStatistique
        txtTaux.Text = fiche.Taux.ToString("0.00", Globalization.CultureInfo.CurrentCulture)
        txtCompteCompense.Text = fiche.CompteCompense
        txtCompteCommission.Text = fiche.CompteCommission
        txtCodeAgence.Text = fiche.CodeAgence
    End Sub

    Private Sub ViderChamps()
        txtCodePdv.Text = String.Empty
        txtDesignation.Text = String.Empty
        txtGroupeStatistique.Text = String.Empty
        txtTaux.Text = String.Empty
        txtCompteCompense.Text = String.Empty
        txtCompteCommission.Text = String.Empty
        txtCodeAgence.Text = String.Empty
    End Sub

    ''' <summary>
    ''' L'Account n'est saisissable qu'en création ; la suppression n'a de sens que sur une
    ''' fiche existante.
    ''' </summary>
    Private Sub MettreAJourEtatBoutons()
        txtCodePdv.ReadOnly = Not _enCreation
        btnSupprimer.Enabled = Not _enCreation AndAlso FicheSelectionnee() IsNot Nothing
        grpDetail.Text = If(_enCreation, "Nouveau sous-agent", "Fiche du sous-agent")
    End Sub

#End Region

#Region "Recherche"

    Private Sub txtRecherche_TextChanged(sender As Object, e As EventArgs) Handles txtRecherche.TextChanged
        ' Rien ici : la recherche est lancée par « Actualiser » ou par la touche Entrée, afin de
        ' ne pas interroger la base à chaque caractère frappé.
    End Sub

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

        _enCreation = True
        dgvListe.ClearSelection()
        ViderChamps()
        txtTaux.Text = "0,00"
        MettreAJourEtatBoutons()

        lblStatut.Text = "Saisie d'un nouveau sous-agent : renseignez l'Account puis enregistrez."
        txtCodePdv.Focus()
    End Sub

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim fiche As PointDeVenteSA = LireChamps()
        If fiche Is Nothing Then Return ' Taux illisible : message déjà affiché.

        fiche.Normaliser()

        Dim anomalies As List(Of String) = fiche.Anomalies()
        If anomalies.Count > 0 Then
            MessageBox.Show(
                "Saisie incomplète ou incorrecte :" & Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", anomalies),
                "Fiche non enregistrable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Un même Account dans les deux tables serait silencieusement traité comme un sous-agent :
        ' la recherche interroge T_Pdv_SA en premier. L'utilisateur doit le savoir avant de valider.
        If _enCreation AndAlso PdvRepository.ExisteDansAutreTable(fiche.CodePdv, estSousAgent:=True) Then
            Dim reponse As DialogResult = MessageBox.Show(
                $"L'Account « {fiche.CodePdv} » est déjà enregistré comme AGENCE PROPRE (T_Pdv_EC)." &
                Environment.NewLine & Environment.NewLine &
                "S'il est également créé comme sous-agent, c'est la fiche sous-agent qui sera retenue " &
                "lors de la comptabilisation : la fiche agence deviendrait sans effet." & Environment.NewLine &
                Environment.NewLine & "Créer tout de même ce sous-agent ?",
                "Account déjà utilisé comme agence", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If reponse <> DialogResult.Yes Then Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            Dim reussi As Boolean = If(_enCreation,
                                       PdvRepository.AjouterSousAgent(fiche, messageErreur),
                                       PdvRepository.ModifierSousAgent(fiche, messageErreur))

            If Not reussi Then
                MessageBox.Show(messageErreur, "Enregistrement impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = If(_enCreation,
                                $"Sous-agent « {fiche.CodePdv} » créé.",
                                $"Sous-agent « {fiche.CodePdv} » modifié.")
            _enCreation = False

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe(fiche.CodePdv)
    End Sub

    Private Sub btnSupprimer_Click(sender As Object, e As EventArgs) Handles btnSupprimer.Click

        Dim fiche As PointDeVenteSA = FicheSelectionnee()
        If fiche Is Nothing Then Return

        Dim reponse As DialogResult = MessageBox.Show(
            $"Supprimer définitivement le sous-agent « {fiche.CodePdv} » ({fiche.Designation}) ?" &
            Environment.NewLine & Environment.NewLine &
            "Les rapports Western Union portant cet Account ne seront plus rattachés à aucun point " &
            "de vente : ils apparaîtront en INCONNU et leur pièce comptable utilisera le compte " &
            "courant WU au lieu du compte de compensation." & Environment.NewLine & Environment.NewLine &
            "Cette suppression est irréversible.",
            "Confirmer la suppression", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            If Not PdvRepository.SupprimerSousAgent(fiche.CodePdv, messageErreur) Then
                MessageBox.Show(messageErreur, "Suppression impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = $"Sous-agent « {fiche.CodePdv} » supprimé."

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe()
    End Sub

    ''' <summary>
    ''' Construit une fiche à partir des champs saisis. Retourne Nothing si le taux n'est pas
    ''' un nombre exploitable — seule vérification faite ici, les autres relevant de
    ''' PointDeVenteSA.Anomalies.
    ''' </summary>
    Private Function LireChamps() As PointDeVenteSA

        Dim taux As Decimal = 0D

        If Not String.IsNullOrWhiteSpace(txtTaux.Text) AndAlso
           Not PointDeVenteSA.EssayerAnalyserTaux(txtTaux.Text, taux) Then

            MessageBox.Show(
                $"Le taux saisi (« {txtTaux.Text} ») n'est pas un nombre." & Environment.NewLine & Environment.NewLine &
                "Saisissez une fraction : 0,70 pour 70 %.",
                "Taux incorrect", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTaux.Focus()
            Return Nothing
        End If

        Return New PointDeVenteSA() With {
            .CodePdv = txtCodePdv.Text,
            .Designation = txtDesignation.Text,
            .GroupeStatistique = txtGroupeStatistique.Text,
            .Taux = taux,
            .CompteCompense = txtCompteCompense.Text,
            .CompteCommission = txtCompteCommission.Text,
            .CodeAgence = txtCodeAgence.Text
        }
    End Function

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
