Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Demande d'annulation d'une journée comptabilisée.
'''
''' CET ÉCRAN N'ANNULE RIEN. Il dépose une demande, que seul un authorizer pourra autoriser —
''' et qui ne peut pas être autorisée par celui qui l'a déposée. Retirer une journée de la
''' comptabilité est plus grave que modifier un taux de sous-agent, et le taux exige déjà
''' deux personnes.
'''
''' IL DIT AUSSI CE QU'IL NE FAIT PAS. Si le fichier destiné au core banking a été injecté,
''' les écritures sont dans les livres de la banque : annuler la journée dans Wincompense ne
''' les en retire pas. L'écran pose la question, affiche ce qu'il sait de la production du
''' fichier, et conserve la réponse dans l'archive — mais l'extourne se demande en
''' comptabilité, et personne d'autre ne le fera à la place de l'agent.
''' </summary>
Public Class FrmAnnulerComptabilisation

    Private ReadOnly _jour As Date
    Private ReadOnly _contenu As AnnulationRepository.ContenuJournee

    ''' <summary>Vrai si la demande a effectivement été déposée.</summary>
    Public ReadOnly Property DemandeDeposee As Boolean
        Get
            Return _deposee
        End Get
    End Property
    Private _deposee As Boolean = False

    ''' <summary>Constructeur requis par le Concepteur Windows Forms.</summary>
    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
        _jour = Date.Today
        _contenu = New AnnulationRepository.ContenuJournee()
    End Sub

    Public Sub New(journee As Date, contenuJournee As AnnulationRepository.ContenuJournee)

        InitializeComponent()
        IconesWU.Habiller(Me)

        _jour = journee.Date
        _contenu = If(contenuJournee, New AnnulationRepository.ContenuJournee())
    End Sub

#Region "Ouverture"

    Private Sub FrmAnnulerComptabilisation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblJournee.Text = $"Journée du {_jour:dd/MM/yyyy}"
        lblContenu.Text = "Seront retirés des tables vivantes et conservés en archive : " &
                          Environment.NewLine & _contenu.Intitule

        RemplirLesMotifs()
        AnnoncerLeFichierCoreBanking()
        AfficherLaConsequence()
    End Sub

    Private Sub RemplirLesMotifs()

        cboMotif.DataSource = AnnulationWU.MotifsProposes()

        ' Aucun motif présélectionné : le premier de la liste serait choisi par défaut, et
        ' un motif choisi par défaut n'explique rien.
        cboMotif.SelectedIndex = -1
    End Sub

    ''' <summary>
    ''' Dit ce que l'application SAIT de la production du fichier core banking — pas ce
    ''' qu'elle suppose.
    '''
    ''' Une production connue ne prouve pas l'injection, et l'absence de production connue ne
    ''' prouve pas le contraire : les fichiers sortis avant la mise en service de la trace ne
    ''' sont pas enregistrés. L'écran le dit dans ces termes, et laisse l'agent répondre.
    ''' </summary>
    Private Sub AnnoncerLeFichierCoreBanking()

        Dim production As CoreBankingRepository.Production = CoreBankingRepository.DerniereProduction(_jour)

        If production Is Nothing Then
            lblQuestion.Text =
                "Aucune production de fichier n'est enregistrée pour cette journée — ce qui " &
                "ne veut pas dire qu'aucun fichier n'est parti : les fichiers produits avant " &
                "la mise en service de cette trace ne sont pas enregistrés." &
                Environment.NewLine &
                "Ce fichier a-t-il déjà été injecté dans le core banking ?"
            Return
        End If

        lblQuestion.Text = production.Avertissement & Environment.NewLine &
                           "Ce fichier a-t-il déjà été injecté dans le core banking ?"

        ' Un fichier dont on sait qu'il est sorti mérite que la réponse la plus prudente soit
        ' proposée : c'est celle qui oblige à regarder avant de continuer.
        rdoInjecte.Checked = True
    End Sub

#End Region

#Region "Conséquence de la réponse"

    Private Sub rdoInjecte_CheckedChanged(sender As Object, e As EventArgs) Handles rdoInjecte.CheckedChanged
        AfficherLaConsequence()
    End Sub

    Private Sub rdoNonInjecte_CheckedChanged(sender As Object, e As EventArgs) Handles rdoNonInjecte.CheckedChanged
        AfficherLaConsequence()
    End Sub

    ''' <summary>
    ''' Écrit noir sur blanc ce que l'annulation fera, et ce qu'elle ne fera pas.
    ''' </summary>
    Private Sub AfficherLaConsequence()

        If rdoInjecte.Checked Then
            lblExtourne.ForeColor = Drawing.Color.Firebrick
            lblExtourne.Text =
                "LES ÉCRITURES SONT DANS LES LIVRES DE LA BANQUE." & Environment.NewLine &
                "L'annulation retire la journée des rapports de Wincompense ; elle n'extourne " &
                "rien. L'extourne est à demander en comptabilité, séparément."
            Return
        End If

        lblExtourne.ForeColor = Drawing.SystemColors.GrayText
        lblExtourne.Text =
            "La journée sortira des rapports d'activité et de la liste des pièces en vigueur." &
            Environment.NewLine &
            "Elle restera consultable, marquée ANNULÉE, avec son motif et les deux signatures."
    End Sub

#End Region

#Region "Dépôt"

    Private Sub btnDeposer_Click(sender As Object, e As EventArgs) Handles btnDeposer.Click

        Dim choix As AnnulationWU.ChoixMotif = TryCast(cboMotif.SelectedItem, AnnulationWU.ChoixMotif)

        If choix Is Nothing Then
            MessageBox.Show(Me,
                "Choisissez le motif de l'annulation." & Environment.NewLine & Environment.NewLine &
                "Sans lui, l'archive ne dira pas pourquoi la journée a été retirée — et c'est " &
                "précisément ce qu'on lui demandera.",
                "Motif obligatoire", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            cboMotif.Focus()
            Return
        End If

        Dim annulation As New AnnulationWU() With {
            .DateActivite = _jour,
            .Motif = choix.Valeur,
            .Commentaire = txtCommentaire.Text.Trim(),
            .CoreBankingInjecte = rdoInjecte.Checked
        }

        If Not ConfirmerLeDepot(annulation) Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim depose As Boolean = AnnulationRepository.Proposer(annulation, messageErreur)
        Cursor = Cursors.Default

        If Not depose Then
            MessageBox.Show(Me, messageErreur, "Demande non déposée",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' Le motif « Autre » sans explication est le refus le plus fréquent : autant
            ' remettre le curseur là où il faut écrire.
            If annulation.Motif = MotifAnnulationWU.Autre AndAlso
               txtCommentaire.Text.Trim().Length = 0 Then txtCommentaire.Focus()

            Return
        End If

        _deposee = True

        MessageBox.Show(Me,
            $"La demande d'annulation de la journée du {_jour:dd/MM/yyyy} est déposée." &
            Environment.NewLine & Environment.NewLine &
            "La journée reste comptabilisée et continue de figurer dans les rapports tant " &
            "qu'un authorizer ne l'a pas autorisée. Ce ne peut pas être vous.",
            "Demande déposée", MessageBoxButtons.OK, MessageBoxIcon.Information)

        DialogResult = DialogResult.OK
        Close()
    End Sub

    ''' <summary>
    ''' Fait relire à l'agent ce qu'il demande. Une confirmation qui répète les chiffres vaut
    ''' mieux qu'une qui demande « êtes-vous sûr ? » : on est toujours sûr.
    ''' </summary>
    Private Function ConfirmerLeDepot(annulation As AnnulationWU) As Boolean

        Dim texte As String =
            $"Demander l'annulation de la journée du {_jour:dd/MM/yyyy} ?" &
            Environment.NewLine & Environment.NewLine &
            _contenu.Intitule & Environment.NewLine & Environment.NewLine &
            $"Motif : {annulation.LibelleMotif}"

        If annulation.CoreBankingInjecte Then
            texte &= Environment.NewLine & Environment.NewLine &
                     "Vous avez indiqué que le fichier core banking est DÉJÀ INJECTÉ : " &
                     "l'extourne restera à demander en comptabilité."
        End If

        Return MessageBox.Show(Me, texte, "Confirmer la demande",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

End Class
