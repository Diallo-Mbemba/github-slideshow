Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Liste des sous-agents par groupe statistique.
'''
''' Écran de consultation uniquement : aucune modification n'y est possible, c'est un état de
''' restitution du paramétrage. La grille du haut récapitule les groupes, celle du bas détaille
''' les sous-agents — tous, ou ceux du seul groupe sélectionné.
'''
''' Les sous-agents sans groupe sont regroupés sous un libellé explicite plutôt qu'omis : ce
''' sont précisément ceux qui n'héritent d'aucun compte ni d'aucun taux, et qu'il faut voir.
''' </summary>
Public Class FrmSousAgentsParGroupe

    ''' <summary>Libellé sous lequel sont présentés les sous-agents sans groupe statistique.</summary>
    Private Const SANS_GROUPE As String = "(sans groupe statistique)"

    ''' <summary>Entrée de la liste déroulante affichant tous les groupes à la fois.</summary>
    Private Const TOUS_LES_GROUPES As String = "(tous les groupes)"

    ''' <summary>Tous les sous-agents, lus une seule fois puis filtrés en mémoire.</summary>
    Private _sousAgents As New List(Of PointDeVenteSA)

    ''' <summary>Groupes statistiques, enregistrés comme hérités.</summary>
    Private _groupes As New List(Of GroupeStatistiqueWU)

    Private _chargementEnCours As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Chargement"

    Private Sub FrmSousAgentsParGroupe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ChargerDonnees()
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        ChargerDonnees()
    End Sub

    ''' <summary>
    ''' Lit les groupes et la totalité des sous-agents. Le détail est ensuite filtré en mémoire :
    ''' un point de vente est une donnée de paramétrage, peu volumineuse, qu'il est inutile de
    ''' relire à chaque changement de groupe affiché.
    ''' </summary>
    Private Sub ChargerDonnees()

        Cursor = Cursors.WaitCursor
        _chargementEnCours = True

        Try
            Dim erreurGroupes As String = String.Empty
            Dim erreurPdv As String = String.Empty

            _groupes = PdvRepository.ListerGroupes(String.Empty, erreurGroupes)
            _sousAgents = PdvRepository.ListerSousAgents(String.Empty, erreurPdv)

            Dim erreur As String = If(Not String.IsNullOrEmpty(erreurGroupes), erreurGroupes, erreurPdv)
            If Not String.IsNullOrEmpty(erreur) Then
                lblStatut.Text = "Lecture incomplète : voir le message affiché."
                MessageBox.Show(erreur, "Base SQL Server", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            AfficherGroupes()
            RemplirListeDeroulante()

        Finally
            _chargementEnCours = False
            Cursor = Cursors.Default
        End Try

        AfficherSousAgents()
    End Sub

    ''' <summary>
    ''' Récapitulatif par groupe. Les sous-agents sans groupe forment une ligne supplémentaire,
    ''' sans compte ni taux : elle rend visible d'un coup d'oeil un paramétrage incomplet.
    ''' </summary>
    Private Sub AfficherGroupes()

        Dim recapitulatif As New DataTable("Groupes")
        recapitulatif.Columns.Add("Groupe", GetType(String))
        recapitulatif.Columns.Add("SousAgents", GetType(Integer))
        recapitulatif.Columns.Add("CompteActivite", GetType(String))
        recapitulatif.Columns.Add("CompteCommission", GetType(String))
        recapitulatif.Columns.Add("Taux", GetType(Decimal))
        recapitulatif.Columns.Add("Etat", GetType(String))

        For Each groupe As GroupeStatistiqueWU In _groupes
            recapitulatif.Rows.Add(groupe.Nom,
                                   CompterSousAgents(groupe.Nom),
                                   groupe.CompteActivite,
                                   groupe.CompteCommission,
                                   groupe.Taux,
                                   EtatDuGroupe(groupe))
        Next

        Dim orphelins As Integer = CompterSousAgents(String.Empty)
        If orphelins > 0 Then
            recapitulatif.Rows.Add(SANS_GROUPE, orphelins, String.Empty, String.Empty, 0D,
                                   "Aucun compte ni taux hérité")
        End If

        dgvGroupes.DataSource = recapitulatif

        DefinirEntete(dgvGroupes, "SousAgents", "Sous-agents")
        DefinirEntete(dgvGroupes, "CompteActivite", "Compte d'activité")
        DefinirEntete(dgvGroupes, "CompteCommission", "Compte de commission")
        DefinirEntete(dgvGroupes, "Etat", "État")

        If dgvGroupes.Columns.Contains("Taux") Then
            dgvGroupes.Columns("Taux").DefaultCellStyle.Format = "P0"
            dgvGroupes.Columns("Taux").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
        If dgvGroupes.Columns.Contains("SousAgents") Then
            dgvGroupes.Columns("SousAgents").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvGroupes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvGroupes.ClearSelection()

        lblNombre.Text = $"{_groupes.Count} groupe(s) — {_sousAgents.Count} sous-agent(s)"
    End Sub

    ''' <summary>État du groupe, dans les mêmes termes que l'écran des groupes statistiques.</summary>
    Private Shared Function EtatDuGroupe(groupe As GroupeStatistiqueWU) As String

        If Not groupe.EstEnregistre Then
            Return "Hérité — pas encore enregistré"
        End If

        If groupe.EstDesynchronise Then
            Return $"{groupe.NombreDesynchronises} sous-agent(s) désynchronisé(s)"
        End If

        Return String.Empty
    End Function

    Private Sub RemplirListeDeroulante()

        cboGroupe.BeginUpdate()
        Try
            cboGroupe.Items.Clear()
            cboGroupe.Items.Add(TOUS_LES_GROUPES)

            For Each groupe As GroupeStatistiqueWU In _groupes
                cboGroupe.Items.Add(groupe.Nom)
            Next

            If CompterSousAgents(String.Empty) > 0 Then
                cboGroupe.Items.Add(SANS_GROUPE)
            End If
        Finally
            cboGroupe.EndUpdate()
        End Try

        cboGroupe.SelectedIndex = 0
    End Sub

#End Region

#Region "Affichage du détail"

    Private Sub cboGroupe_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGroupe.SelectedIndexChanged
        If _chargementEnCours Then Return
        AfficherSousAgents()
    End Sub

    ''' <summary>Cliquer un groupe du récapitulatif restreint le détail à ce seul groupe.</summary>
    Private Sub dgvGroupes_SelectionChanged(sender As Object, e As EventArgs) Handles dgvGroupes.SelectionChanged

        If _chargementEnCours OrElse dgvGroupes.CurrentRow Is Nothing Then Return

        Dim nom As String = Convert.ToString(dgvGroupes.CurrentRow.Cells("Groupe").Value)
        If String.IsNullOrEmpty(nom) Then Return

        Dim index As Integer = cboGroupe.Items.IndexOf(nom)
        If index >= 0 AndAlso cboGroupe.SelectedIndex <> index Then
            cboGroupe.SelectedIndex = index   ' déclenche AfficherSousAgents
        End If
    End Sub

    ''' <summary>
    ''' Remplit la grille de détail avec les sous-agents du groupe choisi, ou avec tous, triés
    ''' par groupe puis par Account : la liste se lit alors groupe par groupe.
    ''' </summary>
    Private Sub AfficherSousAgents()

        Dim groupeChoisi As String = Convert.ToString(cboGroupe.SelectedItem)
        Dim tous As Boolean = String.IsNullOrEmpty(groupeChoisi) OrElse groupeChoisi = TOUS_LES_GROUPES
        Dim sansGroupe As Boolean = (groupeChoisi = SANS_GROUPE)

        Dim retenus As New List(Of PointDeVenteSA)

        For Each pdv As PointDeVenteSA In _sousAgents

            Dim groupePdv As String = If(pdv.GroupeStatistique, String.Empty).Trim()

            If tous OrElse
               (sansGroupe AndAlso groupePdv.Length = 0) OrElse
               (Not sansGroupe AndAlso String.Equals(groupePdv, groupeChoisi, StringComparison.OrdinalIgnoreCase)) Then
                retenus.Add(pdv)
            End If
        Next

        ' Tri par groupe puis par Account : c'est ce qui rend la liste lisible « par groupe ».
        retenus.Sort(AddressOf ComparerParGroupePuisAccount)

        Dim detail As New DataTable("SousAgents")
        detail.Columns.Add("Groupe", GetType(String))
        detail.Columns.Add("Account", GetType(String))
        detail.Columns.Add("Designation", GetType(String))
        detail.Columns.Add("CompteActivite", GetType(String))
        detail.Columns.Add("CompteCommission", GetType(String))
        detail.Columns.Add("Taux", GetType(Decimal))
        detail.Columns.Add("CodeAgence", GetType(String))

        For Each pdv As PointDeVenteSA In retenus
            Dim groupePdv As String = If(pdv.GroupeStatistique, String.Empty).Trim()
            detail.Rows.Add(If(groupePdv.Length = 0, SANS_GROUPE, groupePdv),
                            pdv.CodePdv, pdv.Designation,
                            pdv.CompteCompense, pdv.CompteCommission,
                            pdv.Taux, pdv.CodeAgence)
        Next

        dgvSousAgents.DataSource = detail

        DefinirEntete(dgvSousAgents, "Designation", "Désignation")
        DefinirEntete(dgvSousAgents, "CompteActivite", "Compte d'activité")
        DefinirEntete(dgvSousAgents, "CompteCommission", "Compte de commission")
        DefinirEntete(dgvSousAgents, "CodeAgence", "Code agence")

        If dgvSousAgents.Columns.Contains("Taux") Then
            dgvSousAgents.Columns("Taux").DefaultCellStyle.Format = "P0"
            dgvSousAgents.Columns("Taux").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvSousAgents.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        ' En affichage complet, une ligne sur deux change de teinte À CHAQUE CHANGEMENT DE GROUPE
        ' (et non une ligne sur deux) : les blocs se distinguent alors sans ligne de séparation.
        If tous Then AlternerLesTeintesParGroupe()

        lblDetail.Text = If(tous,
                            $"Sous-agents, tous groupes confondus ({retenus.Count})",
                            $"Sous-agents du groupe « {groupeChoisi} » ({retenus.Count})")

        lblStatut.Text = If(retenus.Count = 0,
                            "Aucun sous-agent dans ce groupe.",
                            "Cliquez un groupe du récapitulatif pour n'afficher que ses sous-agents.")
    End Sub

    ''' <summary>Tri par groupe (sans groupe en dernier) puis par Account, insensible à la casse.</summary>
    Private Shared Function ComparerParGroupePuisAccount(x As PointDeVenteSA, y As PointDeVenteSA) As Integer

        Dim groupeX As String = If(x.GroupeStatistique, String.Empty).Trim()
        Dim groupeY As String = If(y.GroupeStatistique, String.Empty).Trim()

        ' Les sous-agents sans groupe ferment la liste : ce sont des cas à traiter, pas un groupe.
        If groupeX.Length = 0 AndAlso groupeY.Length > 0 Then Return 1
        If groupeY.Length = 0 AndAlso groupeX.Length > 0 Then Return -1

        Dim comparaison As Integer = String.Compare(groupeX, groupeY, StringComparison.OrdinalIgnoreCase)
        If comparaison <> 0 Then Return comparaison

        Return String.Compare(If(x.CodePdv, String.Empty), If(y.CodePdv, String.Empty),
                              StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Alterne la teinte de fond d'un bloc de groupe à l'autre, pour les séparer visuellement.</summary>
    Private Sub AlternerLesTeintesParGroupe()

        Dim groupePrecedent As String = Nothing
        Dim teinte As Boolean = False

        For Each ligne As DataGridViewRow In dgvSousAgents.Rows

            Dim groupe As String = Convert.ToString(ligne.Cells("Groupe").Value)

            If groupePrecedent IsNot Nothing AndAlso
               Not String.Equals(groupe, groupePrecedent, StringComparison.OrdinalIgnoreCase) Then
                teinte = Not teinte
            End If

            If teinte Then
                ligne.DefaultCellStyle.BackColor = Drawing.Color.WhiteSmoke
            End If

            groupePrecedent = groupe
        Next
    End Sub

#End Region

#Region "Utilitaires"

    ''' <summary>Nombre de sous-agents d'un groupe ; chaîne vide pour ceux qui n'en ont aucun.</summary>
    Private Function CompterSousAgents(nomGroupe As String) As Integer

        Dim total As Integer = 0

        For Each pdv As PointDeVenteSA In _sousAgents
            Dim groupePdv As String = If(pdv.GroupeStatistique, String.Empty).Trim()

            If nomGroupe.Length = 0 Then
                If groupePdv.Length = 0 Then total += 1
            ElseIf String.Equals(groupePdv, nomGroupe, StringComparison.OrdinalIgnoreCase) Then
                total += 1
            End If
        Next

        Return total
    End Function

    Private Shared Sub DefinirEntete(grille As DataGridView, nomColonne As String, intitule As String)
        If grille.Columns.Contains(nomColonne) Then
            grille.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
