Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' Gestion des comptes utilisateurs et consultation du journal des connexions.
'''
''' Écran réservé aux administrateurs. Le menu qui y conduit est masqué aux autres rôles, mais
''' le contrôle est refait ici : un écran qui ne compte que sur le menu pour être protégé n'est
''' pas protégé.
'''
''' Les mots de passe ne sont jamais saisis par l'administrateur pour le compte d'un tiers :
''' une réinitialisation tire un mot de passe provisoire, affiché une seule fois, que
''' l'utilisateur devra remplacer à sa première connexion.
''' </summary>
Public Class FrmUtilisateurs

    ''' <summary>Comptes affichés, dans l'ordre de la grille.</summary>
    Private _utilisateurs As New List(Of UtilisateurWU)

    ''' <summary>Nombres de lignes de journal proposés.</summary>
    Private Shared ReadOnly NOMBRES_LIGNES As Integer() = New Integer() {100, 500, 2000}

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture"

    Private Sub FrmUtilisateurs_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesUtilisateurs Then
            MessageBox.Show("La gestion des utilisateurs est réservée aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        For Each nombre As Integer In NOMBRES_LIGNES
            cboLignes.Items.Add(nombre.ToString(Globalization.CultureInfo.CurrentCulture))
        Next

        ' Fixer la sélection déclenche l'événement qui charge le journal : le demander ici en
        ' plus le lirait deux fois.
        cboLignes.SelectedIndex = 0

        ChargerUtilisateurs()
    End Sub

#End Region

#Region "Liste des comptes"

    Private Sub ChargerUtilisateurs()

        lblMessage.Text = String.Empty

        Dim messageErreur As String = String.Empty
        _utilisateurs = UtilisateurRepository.Lister(messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then lblMessage.Text = messageErreur

        dgvUtilisateurs.DataSource = ConstruireTable(_utilisateurs)
        AjusterColonnes()
    End Sub

    ''' <summary>
    ''' Met les comptes en table pour l'affichage. L'empreinte, le sel et le nombre d'itérations
    ''' n'y figurent pas : rien de ce qui touche au secret n'a à s'afficher à l'écran.
    ''' </summary>
    Private Shared Function ConstruireTable(utilisateurs As List(Of UtilisateurWU)) As DataTable

        Dim table As New DataTable("Utilisateurs")

        table.Columns.Add("Identifiant", GetType(String))
        table.Columns.Add("Nom et prénom", GetType(String))
        table.Columns.Add("Rôle", GetType(String))
        table.Columns.Add("Fonction", GetType(String))
        table.Columns.Add("État", GetType(String))
        table.Columns.Add("Échecs", GetType(Integer))
        table.Columns.Add("Dernière connexion", GetType(String))
        table.Columns.Add("Créé le", GetType(String))
        table.Columns.Add("Créé par", GetType(String))
        table.Columns.Add("Modifié le", GetType(String))
        table.Columns.Add("Modifié par", GetType(String))

        For Each utilisateur As UtilisateurWU In utilisateurs

            table.Rows.Add(utilisateur.Identifiant,
                           utilisateur.NomComplet,
                           utilisateur.LibelleRole,
                           utilisateur.LibelleFonction,
                           LibelleEtat(utilisateur),
                           utilisateur.EchecsConsecutifs,
                           LibelleDate(utilisateur.DerniereConnexion),
                           LibelleDate(utilisateur.DateCreation),
                           utilisateur.CreePar,
                           LibelleDate(utilisateur.DateModification),
                           utilisateur.ModifiePar)
        Next

        Return table
    End Function

    ''' <summary>État du compte, du plus bloquant au plus anodin.</summary>
    Private Shared Function LibelleEtat(utilisateur As UtilisateurWU) As String

        If Not utilisateur.Actif Then Return "Désactivé"
        If EstVerrouille(utilisateur) Then Return "Verrouillé"
        If utilisateur.DoitChangerMotDePasse Then Return "Mot de passe à changer"
        Return "Actif"
    End Function

    Private Shared Function EstVerrouille(utilisateur As UtilisateurWU) As Boolean
        Return utilisateur.EchecsConsecutifs >= UtilisateurRepository.ECHECS_AVANT_VERROUILLAGE
    End Function

    Private Shared Function LibelleDate(valeur As Date?) As String
        If Not valeur.HasValue Then Return String.Empty
        Return valeur.Value.ToString("dd/MM/yyyy HH:mm", Globalization.CultureInfo.CurrentCulture)
    End Function

    Private Sub AjusterColonnes()

        For Each colonne As DataGridViewColumn In dgvUtilisateurs.Columns
            colonne.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    ''' <summary>
    ''' Compte sélectionné dans la grille, ou Nothing.
    '''
    ''' Le compte est retrouvé par son identifiant et non par le rang de la ligne : l'utilisateur
    ''' peut trier la grille en cliquant un en-tête, et le rang ne correspondrait alors plus à
    ''' celui de la liste — on réinitialiserait le mot de passe de quelqu'un d'autre.
    ''' </summary>
    Private Function Selectionne() As UtilisateurWU

        If dgvUtilisateurs.CurrentRow Is Nothing Then Return Nothing

        Dim cellule As DataGridViewCell = dgvUtilisateurs.CurrentRow.Cells("Identifiant")
        If cellule Is Nothing OrElse cellule.Value Is Nothing Then Return Nothing

        Dim identifiant As String = Convert.ToString(cellule.Value)

        Return _utilisateurs.FirstOrDefault(
            Function(candidat) String.Equals(candidat.Identifiant, identifiant,
                                             StringComparison.OrdinalIgnoreCase))
    End Function

    ''' <summary>Compte sélectionné, ou message explicatif si aucun ne l'est.</summary>
    Private Function SelectionneOuMessage() As UtilisateurWU

        Dim utilisateur As UtilisateurWU = Selectionne()

        If utilisateur Is Nothing Then
            lblMessage.Text = "Sélectionnez d'abord un compte dans la liste."
        End If

        Return utilisateur
    End Function

#End Region

#Region "Actions sur les comptes"

    Private Sub btnNouveau_Click(sender As Object, e As EventArgs) Handles btnNouveau.Click

        Using edition As New FrmUtilisateurEdition(Nothing, False)

            If edition.ShowDialog(Me) <> DialogResult.OK Then Return

            ChargerUtilisateurs()
            ChargerJournal()
            lblMessage.Text = $"Compte « {edition.UtilisateurEnregistre.Identifiant} » créé. " &
                              "Le mot de passe initial devra être changé à la première connexion."
        End Using
    End Sub

    Private Sub btnModifier_Click(sender As Object, e As EventArgs) Handles btnModifier.Click

        lblMessage.Text = String.Empty
        Dim utilisateur As UtilisateurWU = SelectionneOuMessage()
        If utilisateur Is Nothing Then Return

        Using edition As New FrmUtilisateurEdition(utilisateur, False)

            If edition.ShowDialog(Me) <> DialogResult.OK Then Return

            ChargerUtilisateurs()
            lblMessage.Text = $"Compte « {utilisateur.Identifiant} » modifié."
        End Using
    End Sub

    Private Sub dgvUtilisateurs_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) _
        Handles dgvUtilisateurs.CellDoubleClick

        If e.RowIndex < 0 Then Return
        btnModifier.PerformClick()
    End Sub

    Private Sub btnReinitialiser_Click(sender As Object, e As EventArgs) Handles btnReinitialiser.Click

        lblMessage.Text = String.Empty
        Dim utilisateur As UtilisateurWU = SelectionneOuMessage()
        If utilisateur Is Nothing Then Return

        Dim reponse As DialogResult = MessageBox.Show(
            $"Réinitialiser le mot de passe du compte « {utilisateur.Identifiant} » ?" & Environment.NewLine & Environment.NewLine &
            "Un mot de passe provisoire sera tiré au hasard et affiché une seule fois." & Environment.NewLine &
            "L'utilisateur devra en choisir un autre dès sa première connexion.",
            "Réinitialisation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then Return

        Dim provisoire As String = MotDePasseService.GenererProvisoire()
        Dim messageErreur As String = String.Empty

        ' changementExige à vrai : l'administrateur connaît ce mot de passe, il ne doit pas
        ' rester en service au-delà de la première connexion de son titulaire.
        If Not UtilisateurRepository.ChangerMotDePasse(utilisateur.Identifiant, provisoire, True, messageErreur) Then
            lblMessage.Text = messageErreur
            Return
        End If

        UtilisateurRepository.Journaliser(utilisateur.Identifiant, True,
                                          $"Mot de passe réinitialisé par {SessionWU.Identifiant}")

        MessageBox.Show(
            $"Mot de passe provisoire du compte « {utilisateur.Identifiant} » :" & Environment.NewLine & Environment.NewLine &
            provisoire & Environment.NewLine & Environment.NewLine &
            "Communiquez-le à son titulaire. Il ne sera plus affiché.",
            "Mot de passe provisoire", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ChargerUtilisateurs()
        ChargerJournal()
    End Sub

    Private Sub btnDeverrouiller_Click(sender As Object, e As EventArgs) Handles btnDeverrouiller.Click

        lblMessage.Text = String.Empty
        Dim utilisateur As UtilisateurWU = SelectionneOuMessage()
        If utilisateur Is Nothing Then Return

        If Not EstVerrouille(utilisateur) Then
            lblMessage.Text = $"Le compte « {utilisateur.Identifiant} » n'est pas verrouillé."
            Return
        End If

        Dim messageErreur As String = String.Empty

        If Not UtilisateurRepository.Deverrouiller(utilisateur.Identifiant, messageErreur) Then
            lblMessage.Text = messageErreur
            Return
        End If

        UtilisateurRepository.Journaliser(utilisateur.Identifiant, True,
                                          $"Compte déverrouillé par {SessionWU.Identifiant}")

        ChargerUtilisateurs()
        ChargerJournal()
        lblMessage.Text = $"Compte « {utilisateur.Identifiant} » déverrouillé."
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        ChargerUtilisateurs()
    End Sub

#End Region

#Region "Journal des connexions"

    Private Sub ChargerJournal()

        lblMessageJournal.Text = String.Empty

        Dim messageErreur As String = String.Empty
        Dim journal As DataTable = UtilisateurRepository.ListerConnexions(NombreLignesChoisi(), messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then lblMessageJournal.Text = messageErreur

        dgvConnexions.DataSource = journal

        If dgvConnexions.Columns.Count = 0 Then Return

        dgvConnexions.Columns("DateConnexion").HeaderText = "Date et heure"
        dgvConnexions.Columns("DateConnexion").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss"
        dgvConnexions.Columns("CompteWindows").HeaderText = "Compte Windows"
        dgvConnexions.Columns("Succes").HeaderText = "Réussie"
        dgvConnexions.Columns("Motif").HeaderText = "Motif du refus"

        For Each colonne As DataGridViewColumn In dgvConnexions.Columns
            colonne.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Function NombreLignesChoisi() As Integer

        Dim index As Integer = cboLignes.SelectedIndex
        If index < 0 OrElse index >= NOMBRES_LIGNES.Length Then Return NOMBRES_LIGNES(0)
        Return NOMBRES_LIGNES(index)
    End Function

    Private Sub btnActualiserJournal_Click(sender As Object, e As EventArgs) Handles btnActualiserJournal.Click
        ChargerJournal()
    End Sub

    Private Sub cboLignes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLignes.SelectedIndexChanged
        ChargerJournal()
    End Sub

#End Region

End Class
