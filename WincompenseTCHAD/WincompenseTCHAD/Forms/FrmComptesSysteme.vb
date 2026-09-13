Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Paramétrage des comptes comptables utilisés par la pièce comptable Western Union.
'''
''' Les comptes sont lus dans la table SQL Server SystemeWU et y sont réécrits. Ce formulaire
''' ne contient que de l'orchestration : la lecture, l'écriture et les règles de validation
''' sont portées respectivement par WURepository et ComptesSystemeWU.
'''
''' Tant que l'utilisateur n'a pas enregistré, il travaille sur une COPIE : les comptes en
''' service (ComptesSystemeWU.Actuels) ne sont remplacés qu'une fois l'écriture en base réussie.
''' Une modification abandonnée ne peut donc pas fausser une pièce comptable.
''' </summary>
Public Class FrmComptesSysteme

    ''' <summary>Comptes en cours d'édition (copie de travail).</summary>
    Private _comptes As ComptesSystemeWU

    ''' <summary>Vrai si l'utilisateur a effectivement enregistré une modification.</summary>
    Public ReadOnly Property ModificationEnregistree As Boolean
        Get
            Return _enregistre
        End Get
    End Property
    Private _enregistre As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Chargement"

    Private Sub FrmComptesSysteme_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Relecture de la base à l'ouverture : les comptes ont pu être modifiés par un autre
        ' poste depuis le démarrage de l'application.
        Dim messageErreur As String = String.Empty
        Dim comptesBase As ComptesSystemeWU = WURepository.ChargerComptesSysteme(messageErreur)

        If String.IsNullOrEmpty(messageErreur) Then
            _comptes = comptesBase
        Else
            ' Base inaccessible : on présente les comptes en service plutôt qu'un formulaire vide,
            ' mais l'enregistrement échouera tant que la base ne répond pas — d'où l'avertissement.
            _comptes = ComptesSystemeWU.Actuels.Copier()
            MessageBox.Show(
                messageErreur & Environment.NewLine & Environment.NewLine &
                "Les comptes affichés sont ceux actuellement utilisés par l'application." & Environment.NewLine &
                "Toute modification restera sans effet tant que la base ne sera pas accessible.",
                "Paramétrage en lecture seule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        RemplirChamps()
        AfficherOrigine()
    End Sub

    ''' <summary>
    ''' Renseigne chaque liste déroulante avec la valeur paramétrée et, lorsqu'elle en diffère,
    ''' la valeur par défaut de l'application : l'utilisateur peut ainsi revenir au compte
    ''' d'origine sans avoir à le connaître par cœur. La saisie libre reste possible.
    ''' </summary>
    Private Sub RemplirChamps()
        RemplirListe(cboCompteCourant, _comptes.CompteCourant, ConstantesWU.CPT_COMPTE_COURANT)
        RemplirListe(cboInterBancaire, _comptes.CompteInterBancaire, ConstantesWU.CPT_ATTENTE)
        RemplirListe(cboCommissionTransfert, _comptes.CommissionTransfertBanque, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE)
        RemplirListe(cboCommissionEnvoi, _comptes.CommissionEnvoiBanque, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE)
        RemplirListe(cboCommissionPaiement, _comptes.CommissionPaiementBanque, ConstantesWU.CPT_COMMISSION_PAIEMENT_BANQUE)
        RemplirListe(cboImpotsTaxeEnvoi, _comptes.ImpotsTaxeEnvoi, ConstantesWU.CPT_IMPOTS_TAXE_ENVOI)
        RemplirListe(cboTVA, _comptes.TVACollectee, ConstantesWU.CPT_TVA_COLLECTEE)
        RemplirListe(cboTTAEnvoi, _comptes.TTAEnvoi, ConstantesWU.CPT_TTA_ENVOI)
        RemplirListe(cboTTAReception, _comptes.TTAReception, ConstantesWU.CPT_TTA_RECEPTION)
    End Sub

    Private Shared Sub RemplirListe(liste As ComboBox, valeurParametree As String, valeurParDefaut As String)

        liste.Items.Clear()

        Dim valeur As String = If(valeurParametree, String.Empty).Trim()

        If valeur.Length > 0 Then
            liste.Items.Add(valeur)
        End If

        If Not String.IsNullOrWhiteSpace(valeurParDefaut) AndAlso
           Not String.Equals(valeurParDefaut, valeur, StringComparison.Ordinal) Then
            liste.Items.Add(valeurParDefaut)
        End If

        liste.Text = valeur
    End Sub

    ''' <summary>Indique d'où viennent les comptes affichés : sans cela, impossible de savoir si l'on édite la base ou un repli.</summary>
    Private Sub AfficherOrigine()

        If _comptes.ChargeDepuisBase Then
            Dim code As String = If(String.IsNullOrWhiteSpace(_comptes.CodeParametrage),
                                    "(sans code)", $"code « {_comptes.CodeParametrage} »")
            lblInfo.Text = $"Comptes lus dans la table SystemeWU, ligne {code}."
        Else
            lblInfo.Text = "Comptes par défaut de l'application : aucune ligne de paramétrage n'a pu être lue " &
                           "dans la table SystemeWU."
        End If
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim saisie As New ComptesSystemeWU() With {
            .CompteCourant = cboCompteCourant.Text,
            .CompteInterBancaire = cboInterBancaire.Text,
            .CommissionTransfertBanque = cboCommissionTransfert.Text,
            .CommissionEnvoiBanque = cboCommissionEnvoi.Text,
            .CommissionPaiementBanque = cboCommissionPaiement.Text,
            .ImpotsTaxeEnvoi = cboImpotsTaxeEnvoi.Text,
            .TVACollectee = cboTVA.Text,
            .TTAEnvoi = cboTTAEnvoi.Text,
            .TTAReception = cboTTAReception.Text,
            .CodeParametrage = _comptes.CodeParametrage,
            .ChargeDepuisBase = _comptes.ChargeDepuisBase
        }

        saisie.Normaliser()

        ' --- Un compte vide produirait des écritures sans numéro de compte : blocage. ---
        Dim manquants As List(Of String) = saisie.ComptesManquants()
        If manquants.Count > 0 Then
            MessageBox.Show(
                "Les comptes suivants ne sont pas renseignés :" & Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", manquants) & Environment.NewLine & Environment.NewLine &
                "Tous les comptes doivent être renseignés pour que la pièce comptable soit exploitable.",
                "Paramétrage incomplet", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' --- Un compte non numérique est probablement une faute de frappe : avertissement, pas blocage. ---
        Dim suspects As List(Of String) = saisie.ComptesNonNumeriques()
        If suspects.Count > 0 Then
            Dim reponse As DialogResult = MessageBox.Show(
                "Les comptes suivants ne sont pas composés uniquement de chiffres :" &
                Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", suspects) & Environment.NewLine & Environment.NewLine &
                "Il s'agit le plus souvent d'une faute de frappe (espace ou lettre)." & Environment.NewLine &
                "Voulez-vous tout de même enregistrer ces valeurs ?",
                "Comptes inhabituels", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If reponse <> DialogResult.Yes Then Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            If Not WURepository.EnregistrerComptesSysteme(saisie, messageErreur) Then
                MessageBox.Show(messageErreur, "Enregistrement impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' L'écriture en base a réussi : les comptes deviennent ceux en service.
            _comptes = saisie
            ComptesSystemeWU.Actuels = saisie
            _enregistre = True

            MessageBox.Show(
                "Les comptes ont été enregistrés dans la table SystemeWU." & Environment.NewLine & Environment.NewLine &
                "Ils sont pris en compte immédiatement : toute pièce comptable générée à partir " &
                "de maintenant utilisera ces comptes.",
                "Enregistrement effectué", MessageBoxButtons.OK, MessageBoxIcon.Information)

            RemplirChamps()
            AfficherOrigine()

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
