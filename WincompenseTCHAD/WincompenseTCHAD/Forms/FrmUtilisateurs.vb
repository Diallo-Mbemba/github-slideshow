Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' Gestion des comptes utilisateurs et consultation du journal des connexions.
'''
''' DEUX RÔLES L'OUVRENT, ET N'Y PEUVENT PAS LA MÊME CHOSE. L'administrateur dispose de tout ;
''' le contrôleur crée des agents et des commerciaux, active, désactive et déverrouille — et
''' sur aucun administrateur, ni sur son propre compte. Le menu qui conduit ici est masqué aux
''' autres rôles, mais le contrôle est refait ici : un écran qui ne compte que sur le menu pour
''' être protégé n'est pas protégé.
'''
''' LES BOUTONS SUIVENT LA LIGNE SÉLECTIONNÉE, et pas seulement le rôle de celui qui regarde :
''' le même contrôleur peut désactiver un agent et pas l'administrateur de la ligne suivante.
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
        IconesWU.Habiller(Me)
    End Sub

#Region "Ouverture"

    Private Sub FrmUtilisateurs_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesUtilisateurs Then
            MessageBox.Show("La gestion des utilisateurs est réservée à l'administrateur et au contrôleur.",
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
        AppliquerLesDroits()
    End Sub

    ''' <summary>
    ''' Éteint les boutons que le rôle de l'utilisateur connecté, OU la ligne sélectionnée, ne
    ''' permettent pas.
    '''
    ''' CE N'EST PAS LA SERRURE, C'EST LA POLITESSE. Le dépôt refuse de lui-même ce qui n'est
    ''' pas permis ; ces boutons évitent seulement qu'on saisisse une action pour s'entendre
    ''' dire non. L'un sans l'autre serait une faute : des boutons seuls se contournent, un
    ''' refus seul se découvre trop tard.
    ''' </summary>
    Private Sub AppliquerLesDroits()

        ' LA VARIABLE NE S'APPELLE PAS « selectionne » : VB ne distingue pas la casse, et elle
        ' masquerait la fonction Selectionne() dont elle reçoit le résultat. L'appel se lirait
        ' alors comme une INDEXATION de la variable, et le compilateur refuse (BC30367).
        Dim compte As UtilisateurWU = Selectionne()

        btnModifier.Enabled = SessionWU.PeutModifierLesComptes
        btnReinitialiser.Enabled = SessionWU.PeutReinitialiserLesMotsDePasse

        ' Sans ligne sélectionnée, le bouton reste actif : il affichera « sélectionnez d'abord
        ' un compte », ce qui est plus parlant qu'un bouton éteint dont on ne sait pas pourquoi.
        Dim agirPossible As Boolean = compte Is Nothing OrElse
                                      SessionWU.PeutAgirSurLActivite(compte)

        btnActiver.Enabled = agirPossible
        btnDeverrouiller.Enabled = agirPossible

        btnActiver.Text = LibelleDuBoutonActiver(compte)
    End Sub

    ''' <summary>
    ''' Le bouton dit ce qu'il va faire au compte sélectionné, et non ce qu'il sait faire en
    ''' général : un bouton « Activer / désactiver » laisse chercher dans quel sens il agit.
    ''' </summary>
    Private Shared Function LibelleDuBoutonActiver(compte As UtilisateurWU) As String

        If compte Is Nothing Then Return "Activer / &désactiver"
        If compte.Actif Then Return "&Désactiver le compte"
        Return "Activer le &compte"
    End Function

    Private Sub dgvUtilisateurs_SelectionChanged(sender As Object, e As EventArgs) _
        Handles dgvUtilisateurs.SelectionChanged

        AppliquerLesDroits()
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

        ' Le double-clic sur une ligne passe par ici : le bouton éteint ne suffirait pas à
        ' fermer cette porte-là.
        If Not SessionWU.PeutModifierLesComptes Then
            lblMessage.Text = "Votre rôle ne permet pas de modifier un compte."
            Return
        End If

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

    ''' <summary>
    ''' Active ou désactive le compte sélectionné, et rien d'autre.
    '''
    ''' LA CONFIRMATION NOMME LE COMPTE ET LE SENS. Désactiver quelqu'un l'empêche de se
    ''' connecter le lendemain matin : ce n'est pas un geste à faire en passant sur la mauvaise
    ''' ligne d'une grille triée autrement qu'on ne croit.
    ''' </summary>
    Private Sub btnActiver_Click(sender As Object, e As EventArgs) Handles btnActiver.Click

        lblMessage.Text = String.Empty
        Dim utilisateur As UtilisateurWU = SelectionneOuMessage()
        If utilisateur Is Nothing Then Return

        Dim activer As Boolean = Not utilisateur.Actif
        Dim verbe As String = If(activer, "Activer", "Désactiver")

        Dim reponse As DialogResult = MessageBox.Show(
            $"{verbe} le compte « {utilisateur.Identifiant} » ({utilisateur.NomComplet}) ?" &
            Environment.NewLine & Environment.NewLine &
            If(activer,
               "Son titulaire pourra de nouveau se connecter avec son mot de passe actuel.",
               "Son titulaire ne pourra plus se connecter. Son historique et ses écritures sont conservés."),
            $"{verbe} un compte", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then Return

        Dim messageErreur As String = String.Empty

        If Not UtilisateurRepository.DefinirActivite(utilisateur, activer, messageErreur) Then
            lblMessage.Text = messageErreur
            Return
        End If

        UtilisateurRepository.Journaliser(utilisateur.Identifiant, True,
                                          $"Compte {If(activer, "activé", "désactivé")} par {SessionWU.Identifiant}")

        ChargerUtilisateurs()
        ChargerJournal()
        lblMessage.Text = $"Compte « {utilisateur.Identifiant} » {If(activer, "activé", "désactivé")}."
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

        If Not UtilisateurRepository.Deverrouiller(utilisateur, messageErreur) Then
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
