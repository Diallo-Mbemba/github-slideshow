Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' Écran de l'authorizer : les demandes déposées sur le référentiel, et la décision.
'''
''' La grille du haut liste les demandes en attente ; celle du bas montre, champ par champ, ce
''' que la demande changerait — valeur actuelle à gauche, valeur demandée à droite. Autoriser
''' sans voir ce qu'on autorise ne serait qu'un clic de plus, pas un contrôle.
'''
''' L'écran refuse de décider d'une demande que l'utilisateur a lui-même saisie. La règle est
''' portée par DemandeRepository et redoublée par une contrainte de la base ; elle est
''' simplement rendue visible ici, pour que le refus ne surprenne pas au dernier moment.
''' </summary>
Public Class FrmDemandes

    ''' <summary>Demandes en attente, dans l'ordre de la grille.</summary>
    Private _attente As New List(Of DemandeWU)

    ''' <summary>Nombre de lignes d'historique présentées.</summary>
    Private Const LIGNES_HISTORIQUE As Integer = 300

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture"

    Private Sub FrmDemandes_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesPointsDeVente Then
            MessageBox.Show("Les autorisations du référentiel sont réservées aux commerciaux et aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        ' Un inputer peut consulter la file — voir où en est sa saisie — mais rien y décider.
        Dim decide As Boolean = SessionWU.PeutAutoriserLesPointsDeVente
        btnAutoriser.Enabled = decide
        btnRejeter.Enabled = decide

        ChargerAttente()
        ChargerHistorique()

        If Not decide Then
            lblMessage.Text = "Consultation seule : la décision est réservée à la fonction « authorizer »."
        End If
    End Sub

#End Region

#Region "Demandes en attente"

    Private Sub ChargerAttente()

        lblMessage.Text = String.Empty

        Dim messageErreur As String = String.Empty
        _attente = DemandeRepository.ListerEnAttente(messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then lblMessage.Text = messageErreur

        dgvAttente.DataSource = ConstruireTableAttente(_attente)
        AjusterColonnes(dgvAttente)

        AfficherDetail()
    End Sub

    Private Shared Function ConstruireTableAttente(demandes As List(Of DemandeWU)) As DataTable

        Dim table As New DataTable("Attente")

        table.Columns.Add("Numero", GetType(Long))
        table.Columns.Add("Objet", GetType(String))
        table.Columns.Add("Opération", GetType(String))
        table.Columns.Add("Clé", GetType(String))
        table.Columns.Add("Désignation", GetType(String))
        table.Columns.Add("Saisi par", GetType(String))
        table.Columns.Add("Saisi le", GetType(String))

        For Each demande As DemandeWU In demandes

            table.Rows.Add(demande.IdDemande,
                           demande.LibelleObjet,
                           demande.LibelleOperation,
                           demande.Cle,
                           demande.Designation,
                           demande.SaisiPar,
                           LibelleDate(demande.DateSaisie))
        Next

        Return table
    End Function

    ''' <summary>Demande sélectionnée dans la grille du haut, ou Nothing.</summary>
    Private Function Selectionnee() As DemandeWU

        If dgvAttente.CurrentRow Is Nothing Then Return Nothing

        ' L'événement de sélection se déclenche pendant la liaison des données, alors que les
        ' colonnes n'existent pas encore : y accéder sans vérifier lèverait une exception.
        If Not dgvAttente.Columns.Contains("Numero") Then Return Nothing

        Dim cellule As DataGridViewCell = dgvAttente.CurrentRow.Cells("Numero")
        If cellule Is Nothing OrElse cellule.Value Is Nothing Then Return Nothing

        Dim identifiant As Long = Convert.ToInt64(cellule.Value, Globalization.CultureInfo.InvariantCulture)

        ' La demande est retrouvée par son numéro et non par le rang de la ligne : la grille
        ' peut être triée d'un clic sur un en-tête, et l'on autoriserait alors autre chose que
        ' ce qui est affiché.
        Return _attente.FirstOrDefault(Function(candidate) candidate.IdDemande = identifiant)
    End Function

    Private Sub dgvAttente_SelectionChanged(sender As Object, e As EventArgs) Handles dgvAttente.SelectionChanged
        AfficherDetail()
    End Sub

#End Region

#Region "Ce que la demande changerait"

    ''' <summary>
    ''' Confronte, champ par champ, ce qui est en base et ce qui est demandé.
    ''' </summary>
    Private Sub AfficherDetail()

        Dim demande As DemandeWU = Selectionnee()

        If demande Is Nothing Then
            dgvDetail.DataSource = Nothing
            Return
        End If

        Dim table As New DataTable("Detail")
        table.Columns.Add("Champ", GetType(String))
        table.Columns.Add("Valeur actuelle", GetType(String))
        table.Columns.Add("Valeur demandée", GetType(String))
        table.Columns.Add("Changement", GetType(String))

        For Each ligne As String() In Comparaison(demande)
            table.Rows.Add(ligne(0), ligne(1), ligne(2),
                           If(String.Equals(ligne(1), ligne(2), StringComparison.Ordinal), String.Empty, "modifié"))
        Next

        dgvDetail.DataSource = table
        AjusterColonnes(dgvDetail)
        ColorerLesChangements()
    End Sub

    ''' <summary>Lignes de comparaison, propres à chaque type d'objet.</summary>
    Private Shared Function Comparaison(demande As DemandeWU) As List(Of String())

        Dim lignes As New List(Of String())

        Select Case demande.TypeObjet

            Case TypeObjetWU.SousAgent
                Dim actuel As PointDeVenteSA = SousAgentActuel(demande.Cle)
                lignes.Add({"Account", Si(actuel, actuel?.CodePdv), demande.Cle})
                lignes.Add({"Désignation", Si(actuel, actuel?.Designation), demande.Designation})
                lignes.Add({"Groupe statistique", Si(actuel, actuel?.GroupeStatistique), demande.GroupeStatistique})
                lignes.Add({"Compte de compensation", Si(actuel, actuel?.CompteCompense), demande.CompteActivite})
                lignes.Add({"Compte de commission", Si(actuel, actuel?.CompteCommission), demande.CompteCommission})
                lignes.Add({"Taux", SiTaux(actuel Is Nothing, If(actuel Is Nothing, 0D, actuel.Taux)), Pourcentage(demande.Taux)})
                lignes.Add({"Code agence", Si(actuel, actuel?.CodeAgence), demande.CodeRattachement})

            Case TypeObjetWU.Agence
                Dim actuel As PointDeVenteEC = AgenceActuelle(demande.Cle)
                lignes.Add({"Account", Si(actuel, actuel?.CodeSite), demande.Cle})
                lignes.Add({"Désignation", Si(actuel, actuel?.Designation), demande.Designation})
                lignes.Add({"Code agence Voyager", Si(actuel, actuel?.CodeAgenceVoyager), demande.CodeRattachement})

            Case TypeObjetWU.Groupe
                Dim actuel As GroupeStatistiqueWU = GroupeActuel(demande.Cle)
                lignes.Add({"Groupe", Si(actuel, actuel?.Nom), demande.Cle})
                lignes.Add({"Compte d'activité", Si(actuel, actuel?.CompteActivite), demande.CompteActivite})
                lignes.Add({"Compte de commission", Si(actuel, actuel?.CompteCommission), demande.CompteCommission})
                lignes.Add({"Taux", SiTaux(actuel Is Nothing, If(actuel Is Nothing, 0D, actuel.Taux)), Pourcentage(demande.Taux)})
        End Select

        ' Une suppression ne propose aucune valeur : la colonne de droite se vide, pour qu'on
        ' voie exactement ce qui disparaîtrait.
        If demande.Operation = OperationWU.Suppression Then
            For Each ligne As String() In lignes
                ligne(2) = String.Empty
            Next
        End If

        Return lignes
    End Function

    ''' <summary>Valeur actuelle, ou la mention « (inexistant) » si l'objet n'est pas en base.</summary>
    Private Shared Function Si(objet As Object, valeur As String) As String

        If objet Is Nothing Then Return "(inexistant)"
        Return If(valeur, String.Empty)
    End Function

    Private Shared Function SiTaux(absent As Boolean, valeur As Decimal) As String

        If absent Then Return "(inexistant)"
        Return Pourcentage(valeur)
    End Function

    Private Shared Function Pourcentage(valeur As Decimal) As String
        Return valeur.ToString("0.00", Globalization.CultureInfo.CurrentCulture) & " %"
    End Function

    Private Shared Function SousAgentActuel(cle As String) As PointDeVenteSA

        Dim messageErreur As String = String.Empty

        Return PdvRepository.ListerSousAgents(cle, messageErreur).
            FirstOrDefault(Function(p) String.Equals(p.CodePdv, cle, StringComparison.OrdinalIgnoreCase))
    End Function

    Private Shared Function AgenceActuelle(cle As String) As PointDeVenteEC

        Dim messageErreur As String = String.Empty

        Return PdvRepository.ListerAgences(cle, messageErreur).
            FirstOrDefault(Function(p) String.Equals(p.CodeSite, cle, StringComparison.OrdinalIgnoreCase))
    End Function

    Private Shared Function GroupeActuel(cle As String) As GroupeStatistiqueWU

        Dim messageErreur As String = String.Empty

        Return PdvRepository.ListerGroupes(cle, messageErreur).
            FirstOrDefault(Function(g) String.Equals(g.Nom, cle, StringComparison.OrdinalIgnoreCase) AndAlso g.EstEnregistre)
    End Function

    ''' <summary>Met en évidence les seules lignes qui changent.</summary>
    Private Sub ColorerLesChangements()

        For Each ligne As DataGridViewRow In dgvDetail.Rows

            Dim cellule As DataGridViewCell = ligne.Cells("Changement")
            If cellule Is Nothing OrElse cellule.Value Is Nothing Then Continue For

            If Convert.ToString(cellule.Value).Length = 0 Then Continue For

            ligne.DefaultCellStyle.BackColor = Drawing.Color.LightGoldenrodYellow
            ligne.DefaultCellStyle.Font = New Drawing.Font(dgvDetail.Font, Drawing.FontStyle.Bold)
        Next
    End Sub

#End Region

#Region "Décision"

    Private Sub btnAutoriser_Click(sender As Object, e As EventArgs) Handles btnAutoriser.Click

        lblMessage.Text = String.Empty

        Dim demande As DemandeWU = Selectionnee()
        If demande Is Nothing Then
            lblMessage.Text = "Sélectionnez d'abord une demande."
            Return
        End If

        Dim reponse As DialogResult = MessageBox.Show(
            demande.Intitule & Environment.NewLine & Environment.NewLine &
            $"Saisie par {demande.SaisiPar} le {LibelleDate(demande.DateSaisie)}." &
            Environment.NewLine & Environment.NewLine &
            "Autoriser cette demande ? Elle sera portée immédiatement dans la base, et la " &
            "comptabilisation du jour en tiendra compte.",
            "Autoriser", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            If Not DemandeRepository.Autoriser(demande, messageErreur) Then
                MessageBox.Show(messageErreur, "Autorisation impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                              $"Demande {demande.IdDemande} autorisée — {demande.Intitule}")

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerAttente()
        ChargerHistorique()
        lblMessage.Text = $"{demande.Intitule} — autorisée et appliquée."
    End Sub

    Private Sub btnRejeter_Click(sender As Object, e As EventArgs) Handles btnRejeter.Click

        lblMessage.Text = String.Empty

        Dim demande As DemandeWU = Selectionnee()
        If demande Is Nothing Then
            lblMessage.Text = "Sélectionnez d'abord une demande."
            Return
        End If

        Dim motif As String = InputBox(
            demande.Intitule & Environment.NewLine & Environment.NewLine &
            "Motif du rejet — il sera lu par " & demande.SaisiPar & ", qui doit savoir quoi corriger :",
            "Rejeter la demande")

        If String.IsNullOrWhiteSpace(motif) Then
            lblMessage.Text = "Rejet abandonné : aucun motif n'a été saisi."
            Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            If Not DemandeRepository.Rejeter(demande, motif, messageErreur) Then
                MessageBox.Show(messageErreur, "Rejet impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                              $"Demande {demande.IdDemande} rejetée — {demande.Intitule}")

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerAttente()
        ChargerHistorique()
        lblMessage.Text = $"{demande.Intitule} — rejetée."
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        ChargerAttente()
    End Sub

#End Region

#Region "Demandes décidées"

    Private Sub ChargerHistorique()

        lblMessageHistorique.Text = String.Empty

        Dim messageErreur As String = String.Empty
        Dim decidees As List(Of DemandeWU) = DemandeRepository.ListerHistorique(LIGNES_HISTORIQUE, messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then lblMessageHistorique.Text = messageErreur

        Dim table As New DataTable("Historique")

        table.Columns.Add("Numero", GetType(Long))
        table.Columns.Add("Objet", GetType(String))
        table.Columns.Add("Opération", GetType(String))
        table.Columns.Add("Clé", GetType(String))
        table.Columns.Add("Décision", GetType(String))
        table.Columns.Add("Saisi par", GetType(String))
        table.Columns.Add("Décidé par", GetType(String))
        table.Columns.Add("Décidé le", GetType(String))
        table.Columns.Add("Motif du rejet", GetType(String))

        For Each demande As DemandeWU In decidees

            table.Rows.Add(demande.IdDemande,
                           demande.LibelleObjet,
                           demande.LibelleOperation,
                           demande.Cle,
                           demande.LibelleStatut,
                           demande.SaisiPar,
                           demande.DecidePar,
                           LibelleDate(demande.DateDecision),
                           demande.MotifRejet)
        Next

        dgvHistorique.DataSource = table
        AjusterColonnes(dgvHistorique)
    End Sub

    Private Sub btnActualiserHistorique_Click(sender As Object, e As EventArgs) Handles btnActualiserHistorique.Click
        ChargerHistorique()
    End Sub

#End Region

#Region "Utilitaires d'affichage"

    Private Shared Sub AjusterColonnes(grille As DataGridView)

        If grille.Columns.Contains("Numero") Then grille.Columns("Numero").HeaderText = "N°"

        For Each colonne As DataGridViewColumn In grille.Columns
            colonne.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Shared Function LibelleDate(valeur As Date?) As String

        If Not valeur.HasValue Then Return String.Empty
        Return valeur.Value.ToString("dd/MM/yyyy HH:mm", Globalization.CultureInfo.CurrentCulture)
    End Function

#End Region

End Class
