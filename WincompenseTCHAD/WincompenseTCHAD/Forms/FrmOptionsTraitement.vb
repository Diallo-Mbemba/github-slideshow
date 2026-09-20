Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Les options de traitement de la banque, réservées à l'administrateur.
'''
''' POURQUOI CET ÉCRAN EXISTE
'''
''' Le visa d'une journée constate, ou bloque. La banque choisit. Imposer l'un ou l'autre
''' serait décider à sa place : un contrôle bloquant arrête la compense le matin où le chef de
''' service est absent, et un contrôle qui n'arrête rien ne contrôle pas grand-chose. Le choix
''' se fait ici, une fois, et vaut pour tous les postes.
'''
''' L'ÉCRAN DIT CE QUE CHAQUE CHOIX PRODUIT
'''
''' Une case à cocher dont on ne voit pas la conséquence se coche au hasard. Sous les deux
''' options, une phrase décrit ce qui se passera au moment de produire le fichier core banking.
''' </summary>
Public Class FrmOptionsTraitement

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Ouverture"

    Private Sub FrmOptionsTraitement_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesComptesSystemes Then
            MessageBox.Show("Les options de traitement sont réservées à l'administrateur." &
                            Environment.NewLine & Environment.NewLine &
                            "Elles décrivent la façon de travailler de la banque, et valent " &
                            "pour tous les postes.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        Charger()
    End Sub

    Private Sub Charger()

        ' La valeur est relue en base, et non prise dans le cache : on vient précisément la
        ' modifier, et quelqu'un d'autre a pu le faire avant.
        OptionsWU.Oublier()

        Dim actif As Boolean = OptionsWU.VisaAvantCoreBanking

        rdoVisaOui.Checked = actif
        rdoVisaNon.Checked = Not actif

        AfficherLaConsequence()

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = String.Empty
    End Sub

#End Region

#Region "Conséquence du choix"

    Private Sub rdoVisaOui_CheckedChanged(sender As Object, e As EventArgs) Handles rdoVisaOui.CheckedChanged
        AfficherLaConsequence()
    End Sub

    Private Sub rdoVisaNon_CheckedChanged(sender As Object, e As EventArgs) Handles rdoVisaNon.CheckedChanged
        AfficherLaConsequence()
    End Sub

    ''' <summary>
    ''' Décrit ce que le choix produira. Une option dont on ne voit pas la conséquence se
    ''' coche au hasard.
    ''' </summary>
    Private Sub AfficherLaConsequence()

        If rdoVisaOui.Checked Then
            lblConsequence.ForeColor = Drawing.Color.Firebrick
            lblConsequence.Text =
                "L'agent de la compense ne pourra pas produire le fichier tant qu'un " &
                "authorizer n'aura pas visé la journée, sur le bordereau." & Environment.NewLine &
                "Prévoyez PLUSIEURS authorizers : avec un seul, une journée d'absence bloque " &
                "la compense."
            Return
        End If

        lblConsequence.ForeColor = Drawing.SystemColors.GrayText
        lblConsequence.Text =
            "L'agent sera averti que la journée n'est pas visée — et que des Accounts n'ont " &
            "pas été comptabilisés, s'il y en a — puis libre de continuer." & Environment.NewLine &
            "Le visa reste attendu : il constate, il n'empêche pas."
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim actif As Boolean = rdoVisaOui.Checked

        If actif AndAlso Not Confirmer() Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim enregistre As Boolean = OptionsWU.Enregistrer(
            OptionsWU.CLE_VISA_AVANT_CORE_BANKING, OptionsWU.Texte(actif),
            OptionsWU.LIBELLE_VISA_AVANT_CORE_BANKING, messageErreur)
        Cursor = Cursors.Default

        If Not enregistre Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Option non enregistrée."
            FrmDiagnostic.Afficher(Me, "Options de traitement", messageErreur)
            Return
        End If

        UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                          "Visa avant core banking : " & OptionsWU.Texte(actif))

        lblStatut.ForeColor = Drawing.Color.DarkGreen
        lblStatut.Text = "Option enregistrée : visa obligatoire " & OptionsWU.Texte(actif).ToLowerInvariant() & "."

        AfficherLaDerniereModification()
    End Sub

    ''' <summary>
    ''' Activer un contrôle bloquant mérite une confirmation : c'est le jour où le chef de
    ''' service est absent qu'on s'aperçoit de ce qu'on a coché.
    ''' </summary>
    Private Function Confirmer() As Boolean

        Return MessageBox.Show(Me,
            "Rendre le visa OBLIGATOIRE avant la production du fichier core banking ?" &
            Environment.NewLine & Environment.NewLine &
            "À partir de maintenant, aucune journée ne pourra être chargée dans le core " &
            "banking sans qu'un authorizer l'ait visée." & Environment.NewLine & Environment.NewLine &
            "Assurez-vous qu'il y a PLUSIEURS authorizers : avec un seul, une journée " &
            "d'absence bloque la compense.",
            "Confirmer l'option", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    Private Sub AfficherLaDerniereModification()

        lblDerniereModification.Text =
            $"Modifiée le {Date.Now:dd/MM/yyyy à HH:mm} par {SessionWU.Auteur}."
    End Sub

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
