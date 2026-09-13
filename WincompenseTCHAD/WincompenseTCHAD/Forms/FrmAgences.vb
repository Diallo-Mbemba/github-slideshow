Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Gestion des agences propres Ecobank (table T_Pdv_EC) : consultation, création,
''' modification et suppression.
'''
''' Même organisation que FrmSousAgents : aucune requête SQL ici, tout passe par
''' PdvRepository ; l'Account (Codesite) n'est saisissable qu'à la création, car c'est la
''' clé sous laquelle les rapports Western Union désignent l'agence.
'''
''' Une agence propre ne rétrocède aucune commission — la banque en conserve 100 % — d'où
''' l'absence de taux et de comptes de compensation/commission, contrairement aux sous-agents.
''' </summary>
Public Class FrmAgences

    Private _liste As List(Of PointDeVenteEC)
    Private _enCreation As Boolean = False
    Private _chargementEnCours As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Chargement et affichage de la liste"

    Private Sub FrmAgences_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Habillage et adaptation à l'écran du poste, avant tout autre traitement : la fenêtre
        ' prend sa taille définitive une fois pour toutes, et les contrôles ancrés suivent.
        ThemeWU.Appliquer(Me, "Agences propres", btnEnregistrer)
        DimensionsWU.Adapter(Me, True)

        If Not SessionWU.PeutGererLesPointsDeVente Then
            MessageBox.Show("Le paramétrage des agences propres est réservé aux commerciaux et aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        ChargerListe()
    End Sub

    Private Sub ChargerListe(Optional codeASelectionner As String = Nothing)

        Cursor = Cursors.WaitCursor
        _chargementEnCours = True

        Try
            Dim messageErreur As String = String.Empty
            _liste = PdvRepository.ListerAgences(txtRecherche.Text, messageErreur)

            dgvListe.DataSource = Nothing
            dgvListe.DataSource = _liste
            RenommerColonnes()

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = messageErreur
                lblNombre.Text = String.Empty
                MessageBox.Show(messageErreur, "Base SQL Server", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                lblNombre.Text = $"{_liste.Count} agence(s)"
                lblStatut.Text = If(_liste.Count = 0,
                                    "Aucune agence ne correspond à la recherche.",
                                    "Sélectionnez une fiche pour la modifier, ou cliquez sur « Nouveau ».")
            End If

        Finally
            _chargementEnCours = False
            Cursor = Cursors.Default
        End Try

        If Not String.IsNullOrEmpty(codeASelectionner) Then
            SelectionnerCode(codeASelectionner)
        End If

        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    Private Sub RenommerColonnes()

        DefinirEntete("CodeSite", "Account")
        DefinirEntete("Designation", "Désignation")
        DefinirEntete("CodeAgenceVoyager", "Code agence Voyager")

        dgvListe.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub DefinirEntete(nomColonne As String, intitule As String)
        If dgvListe.Columns.Contains(nomColonne) Then
            dgvListe.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

    Private Sub SelectionnerCode(code As String)

        For Each ligne As DataGridViewRow In dgvListe.Rows
            Dim fiche As PointDeVenteEC = TryCast(ligne.DataBoundItem, PointDeVenteEC)
            If fiche IsNot Nothing AndAlso String.Equals(fiche.CodeSite, code, StringComparison.OrdinalIgnoreCase) Then
                ligne.Selected = True
                dgvListe.CurrentCell = ligne.Cells(0)
                Return
            End If
        Next
    End Sub

    Private Function FicheSelectionnee() As PointDeVenteEC
        If dgvListe.CurrentRow Is Nothing Then Return Nothing
        Return TryCast(dgvListe.CurrentRow.DataBoundItem, PointDeVenteEC)
    End Function

    Private Sub dgvListe_SelectionChanged(sender As Object, e As EventArgs) Handles dgvListe.SelectionChanged
        If _chargementEnCours Then Return
        _enCreation = False
        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    Private Sub AfficherFicheSelectionnee()

        If _enCreation Then Return

        Dim fiche As PointDeVenteEC = FicheSelectionnee()

        If fiche Is Nothing Then
            ViderChamps()
            Return
        End If

        txtCodeSite.Text = fiche.CodeSite
        txtDesignation.Text = fiche.Designation
        txtCodeAgenceVoyager.Text = fiche.CodeAgenceVoyager
    End Sub

    Private Sub ViderChamps()
        txtCodeSite.Text = String.Empty
        txtDesignation.Text = String.Empty
        txtCodeAgenceVoyager.Text = String.Empty
    End Sub

    Private Sub MettreAJourEtatBoutons()
        txtCodeSite.ReadOnly = Not _enCreation
        btnSupprimer.Enabled = Not _enCreation AndAlso FicheSelectionnee() IsNot Nothing
        grpDetail.Text = If(_enCreation, "Nouvelle agence", "Fiche de l'agence")
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

        _enCreation = True
        dgvListe.ClearSelection()
        ViderChamps()
        MettreAJourEtatBoutons()

        lblStatut.Text = "Saisie d'une nouvelle agence : renseignez l'Account puis enregistrez."
        txtCodeSite.Focus()
    End Sub

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim fiche As New PointDeVenteEC() With {
            .CodeSite = txtCodeSite.Text,
            .Designation = txtDesignation.Text,
            .CodeAgenceVoyager = txtCodeAgenceVoyager.Text
        }
        fiche.Normaliser()

        Dim anomalies As List(Of String) = fiche.Anomalies()
        If anomalies.Count > 0 Then
            MessageBox.Show(
                "Saisie incomplète :" & Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", anomalies),
                "Fiche non enregistrable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' La désignation est facultative en base mais alimente le libellé des écritures :
        ' on avertit sans bloquer.
        Dim avertissements As List(Of String) = fiche.Avertissements()
        If avertissements.Count > 0 Then
            Dim suite As DialogResult = MessageBox.Show(
                String.Join(Environment.NewLine, avertissements) & Environment.NewLine & Environment.NewLine &
                "Enregistrer tout de même ?",
                "Fiche incomplète", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If suite <> DialogResult.Yes Then Return
        End If

        If _enCreation AndAlso PdvRepository.ExisteDansAutreTable(fiche.CodeSite, estSousAgent:=False) Then
            Dim reponse As DialogResult = MessageBox.Show(
                $"L'Account « {fiche.CodeSite} » est déjà enregistré comme SOUS-AGENT (T_Pdv_SA)." &
                Environment.NewLine & Environment.NewLine &
                "La comptabilisation interrogeant T_Pdv_SA en premier, c'est la fiche sous-agent qui " &
                "sera retenue : cette fiche agence resterait sans effet." & Environment.NewLine &
                Environment.NewLine & "Créer tout de même cette agence ?",
                "Account déjà utilisé comme sous-agent", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If reponse <> DialogResult.Yes Then Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            Dim reussi As Boolean = If(_enCreation,
                                       PdvRepository.AjouterAgence(fiche, messageErreur),
                                       PdvRepository.ModifierAgence(fiche, messageErreur))

            If Not reussi Then
                MessageBox.Show(messageErreur, "Enregistrement impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = If(_enCreation,
                                $"Agence « {fiche.CodeSite} » créée.",
                                $"Agence « {fiche.CodeSite} » modifiée.")
            _enCreation = False

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe(fiche.CodeSite)
    End Sub

    Private Sub btnSupprimer_Click(sender As Object, e As EventArgs) Handles btnSupprimer.Click

        Dim fiche As PointDeVenteEC = FicheSelectionnee()
        If fiche Is Nothing Then Return

        Dim reponse As DialogResult = MessageBox.Show(
            $"Supprimer définitivement l'agence « {fiche.CodeSite} » ({fiche.Designation}) ?" &
            Environment.NewLine & Environment.NewLine &
            "Les rapports Western Union portant cet Account ne seront plus rattachés à aucun point " &
            "de vente : ils apparaîtront en INCONNU dans la grille de contrôle." &
            Environment.NewLine & Environment.NewLine & "Cette suppression est irréversible.",
            "Confirmer la suppression", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            If Not PdvRepository.SupprimerAgence(fiche.CodeSite, messageErreur) Then
                MessageBox.Show(messageErreur, "Suppression impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = $"Agence « {fiche.CodeSite} » supprimée."

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe()
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
