Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>
''' Ce que la banque a réellement encaissé sur les commissions Western Union.
'''
''' POURQUOI CET ÉCRAN N'EXISTAIT PAS
'''
''' L'onglet « Évolution des commissions » des rapports d'activité montre la commission
''' GÉNÉRÉE par l'activité — celle que Western Union verse. Sur un sous-agent à 70 %, les sept
''' dixièmes affichés ne sont pas à la banque : ils lui sont rétrocédés. Aucun écran ne disait
''' donc ce que la banque gagne.
'''
''' DEUX SOURCES, DEUX RÈGLES
'''
''' Sur un sous-agent, la banque garde ce qui n'est pas rétrocédé ; sur une agence propre, elle
''' garde tout, le taux valant zéro. La somme des deux est ce qu'elle a encaissé.
'''
''' RIEN N'EST RECALCULÉ
'''
''' La part de la banque est celle qui a été calculée le jour même et conservée avec
''' l'historique. La recalculer avec les taux d'aujourd'hui ferait varier rétroactivement ce
''' que la banque a gagné le mois dernier — c'est ce que l'application refuse déjà de faire
''' pour les pièces comptables.
'''
''' ET LE CHIFFRE EST CONFRONTÉ À LA COMPTABILITÉ
'''
''' Le total est comparé à ce que les pièces conservées ont effectivement porté sur les comptes
''' de commission de la banque. Les deux doivent coïncider ; l'écran le dit, plutôt que de le
''' laisser croire.
''' </summary>
Public Class FrmCommissionsBanque

    Private _lignes As List(Of LigneHistoriqueWU)
    Private _synthese As DataTable
    Private _parJour As DataTable
    Private _comptes As DataTable

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
        _lignes = New List(Of LigneHistoriqueWU)()
        _synthese = New DataTable()
        _parJour = New DataTable()
        _comptes = New DataTable()
    End Sub

#Region "Ouverture"

    Private Sub FrmCommissionsBanque_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutVoirLesRapports Then
            MessageBox.Show("La consultation des états est réservée aux utilisateurs identifiés.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        btnExporter.Enabled = False

        Dim premiere As Date
        Dim derniere As Date
        Dim messageErreur As String = String.Empty

        If Not HistoriqueRepository.ObtenirBornes(premiere, derniere, messageErreur) Then

            If messageErreur.Length > 0 Then
                lblDisponible.Text = "Historique illisible."
                FrmDiagnostic.Afficher(Me, "Historique", messageErreur)
            Else
                lblDisponible.Text = "Aucune journée comptabilisée à ce jour."
                lblStatut.Text = "L'historique se remplit à chaque génération de pièce comptable."
            End If

            Return
        End If

        dtpDebut.Value = premiere
        dtpFin.Value = derniere
        lblDisponible.Text = $"Journées disponibles : du {premiere:dd/MM/yyyy} au {derniere:dd/MM/yyyy}"

        Afficher()
    End Sub

    ''' <summary>
    ''' Cale la période sur le mois entier de la date de fin.
    '''
    ''' L'édition destinée à la signature est mensuelle : un état signé qui couvrirait
    ''' vingt-trois jours parce que l'agent a mal cliqué serait un état faux, et rien ne le
    ''' dirait.
    ''' </summary>
    Private Sub btnMois_Click(sender As Object, e As EventArgs) Handles btnMois.Click

        Dim reference As Date = dtpFin.Value.Date
        Dim premier As New Date(reference.Year, reference.Month, 1)

        dtpDebut.Value = premier
        dtpFin.Value = premier.AddMonths(1).AddDays(-1)

        Afficher()
    End Sub

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click
        Afficher()
    End Sub

#End Region

#Region "Construction de l'état"

    Private Sub Afficher()

        Dim debut As Date = dtpDebut.Value.Date
        Dim fin As Date = dtpFin.Value.Date

        If debut > fin Then
            MessageBox.Show(Me, "La date de début est postérieure à la date de fin.",
                            "Période invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        _lignes = HistoriqueRepository.ListerPeriode(debut, fin, messageErreur)
        Cursor = Cursors.Default

        If messageErreur.Length > 0 Then
            btnExporter.Enabled = False
            FrmDiagnostic.Afficher(Me, "Historique", messageErreur)
            Return
        End If

        _synthese = RapportActiviteService.ConstruireCommissionsBanque(_lignes)
        _parJour = RapportActiviteService.ConstruireCommissionsBanqueParJour(_lignes)
        _comptes = LireLesComptes(debut, fin)

        dgvSynthese.DataSource = _synthese
        dgvParJour.DataSource = _parJour
        dgvComptes.DataSource = _comptes

        FormaterLaSynthese()
        FormaterLeParJour()
        FormaterLesComptes()

        AvertirDesJourneesNonDocumentees()
        AfficherLeControle()

        btnExporter.Enabled = _lignes.Count > 0
        lblStatut.Text = $"{_lignes.Count} ligne(s) d'historique sur la période."
    End Sub

    ''' <summary>
    ''' Ce que les pièces conservées ont porté sur les comptes de commission de la banque.
    '''
    ''' Les comptes sont ceux du paramétrage COURANT : si la banque en a changé au milieu de
    ''' la période, les écritures passées sur l'ancien compte n'y figurent pas. Le contrôle
    ''' d'écart le signalera, et c'est bien ainsi — un écart inexpliqué vaut mieux qu'un total
    ''' bricolé.
    ''' </summary>
    Private Function LireLesComptes(debut As Date, fin As Date) As DataTable

        Dim comptes As ComptesSystemeWU = ComptesSystemeWU.Actuels

        If comptes Is Nothing Then
            Dim motif As String = String.Empty
            comptes = WURepository.ChargerComptesSysteme(motif)
        End If

        Dim recherches As New List(Of String) From {
            comptes.CommissionTransfertBanque,
            comptes.CommissionEnvoiBanque,
            comptes.CommissionPaiementBanque
        }

        Dim messageErreur As String = String.Empty
        Dim totaux As DataTable = PieceRepository.TotauxParCompte(debut, fin, recherches, messageErreur)

        If messageErreur.Length > 0 Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = PremiereLigne(messageErreur)
        Else
            lblStatut.ForeColor = Drawing.SystemColors.GrayText
        End If

        Return totaux
    End Function

#End Region

#Region "Mise en forme"

    Private Sub FormaterLaSynthese()

        dgvSynthese.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvSynthese, "SousAgents", "Sous-agents", 100)
        Entete(dgvSynthese, "AgencesPropres", "Agences propres", 100)
        Entete(dgvSynthese, "TotalBanque", "TOTAL BANQUE", 110)
        Entete(dgvSynthese, "Retrocede", "Rétrocédé aux SA", 110)
        Entete(dgvSynthese, "CommissionTotale", "Commission totale", 110)

        If dgvSynthese.Columns.Contains("Nature") Then
            dgvSynthese.Columns("Nature").HeaderText = "Nature"
            dgvSynthese.Columns("Nature").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        Montants(dgvSynthese, "SousAgents", "AgencesPropres", "TotalBanque",
                 "Retrocede", "CommissionTotale")

        ' La ligne TOTAL est la seule qu'on lise vraiment : elle se voit.
        For rang As Integer = 0 To dgvSynthese.Rows.Count - 1
            If Not String.Equals(Convert.ToString(dgvSynthese.Rows(rang).Cells("Nature").Value),
                                 "TOTAL", StringComparison.Ordinal) Then Continue For

            dgvSynthese.Rows(rang).DefaultCellStyle.Font =
                New Drawing.Font(dgvSynthese.Font, Drawing.FontStyle.Bold)
        Next
    End Sub

    Private Sub FormaterLeParJour()

        dgvParJour.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvParJour, "Date", "Journée", 90)
        Entete(dgvParJour, "SousAgents", "Sous-agents", 110)
        Entete(dgvParJour, "AgencesPropres", "Agences propres", 110)
        Entete(dgvParJour, "TotalBanque", "TOTAL BANQUE", 120)
        Entete(dgvParJour, "Retrocede", "Rétrocédé aux SA", 120)
        Entete(dgvParJour, "Documentee", "Répartition", 90)

        If dgvParJour.Columns.Contains("CommissionTotale") Then
            dgvParJour.Columns("CommissionTotale").HeaderText = "Commission totale"
            dgvParJour.Columns("CommissionTotale").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvParJour.Columns("CommissionTotale").DefaultCellStyle.Format = "N0"
            dgvParJour.Columns("CommissionTotale").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight
        End If

        Montants(dgvParJour, "SousAgents", "AgencesPropres", "TotalBanque", "Retrocede")
        MarquerLesJourneesIncompletes()
    End Sub

    ''' <summary>
    ''' Une journée dont la répartition n'est pas conservée porte un zéro côté sous-agents qui
    ''' n'est pas un chiffre mais une absence. Elle est grisée, et sa colonne le dit.
    ''' </summary>
    Private Sub MarquerLesJourneesIncompletes()

        If Not dgvParJour.Columns.Contains("Documentee") Then Return

        For rang As Integer = 0 To dgvParJour.Rows.Count - 1

            Dim valeur As String = Convert.ToString(dgvParJour.Rows(rang).Cells("Documentee").Value)
            If Not String.Equals(valeur, "NON", StringComparison.Ordinal) Then Continue For

            dgvParJour.Rows(rang).DefaultCellStyle.ForeColor = Drawing.Color.Firebrick
        Next
    End Sub

    Private Sub FormaterLesComptes()

        dgvComptes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvComptes, "Ecritures", "Écr.", 55)
        Entete(dgvComptes, "Debit", "Débit", 100)
        Entete(dgvComptes, "Credit", "Crédit", 100)
        Entete(dgvComptes, "Net", "Net crédité", 110)

        If dgvComptes.Columns.Contains("Compte") Then
            dgvComptes.Columns("Compte").HeaderText = "Compte"
            dgvComptes.Columns("Compte").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        Montants(dgvComptes, "Debit", "Credit", "Net")

        If dgvComptes.Columns.Contains("Ecritures") Then
            dgvComptes.Columns("Ecritures").DefaultCellStyle.Format = "N0"
            dgvComptes.Columns("Ecritures").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    Private Shared Sub Entete(grille As DataGridView, colonne As String, libelle As String,
                              largeur As Integer)

        If Not grille.Columns.Contains(colonne) Then Return
        grille.Columns(colonne).HeaderText = libelle
        GrilleWU.LargeurFixe(grille, colonne, largeur)
    End Sub

    ''' <summary>Aligne à droite et groupe par milliers, sans décimale : le FCFA n'en a pas.</summary>
    Private Shared Sub Montants(grille As DataGridView, ParamArray colonnes() As String)

        For Each colonne As String In colonnes
            If Not grille.Columns.Contains(colonne) Then Continue For
            grille.Columns(colonne).DefaultCellStyle.Format = "N0"
            grille.Columns(colonne).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next
    End Sub

#End Region

#Region "Avertissement et contrôle"

    ''' <summary>
    ''' Dit franchement ce que l'état ne sait pas.
    '''
    ''' La répartition n'est conservée que depuis la mise en service des colonnes de part
    ''' bancaire. Avant, elle est connue pour les agences propres — leur taux vaut zéro par
    ''' construction — et inconnue pour les sous-agents. Afficher zéro sans le dire ferait
    ''' croire à un chiffre.
    ''' </summary>
    Private Sub AvertirDesJourneesNonDocumentees()

        Dim incompletes As List(Of Date) = RapportActiviteService.JourneesNonDocumentees(_lignes)

        If incompletes.Count = 0 Then
            lblAvertissement.Text = String.Empty
            Return
        End If

        Dim premier As Date? = RapportActiviteService.PremierJourDocumente(_lignes)

        Dim phrase As String =
            $"{incompletes.Count} journée(s) de cette période n'ont pas de répartition conservée : " &
            $"la part des SOUS-AGENTS y est inconnue, et non nulle " &
            $"(de {incompletes(0):dd/MM/yyyy} à {incompletes(incompletes.Count - 1):dd/MM/yyyy})."

        If premier.HasValue Then
            phrase &= Environment.NewLine &
                      $"La répartition est disponible à partir du {premier.Value:dd/MM/yyyy}."
        Else
            phrase &= Environment.NewLine &
                      "Aucune journée de cette période n'est répartie : exécutez " &
                      "Scripts\16_CommissionsBanque.sql, puis comptabilisez normalement."
        End If

        lblAvertissement.Text = phrase
    End Sub

    ''' <summary>
    ''' Confronte le total de l'état à ce que les pièces ont réellement porté en comptabilité.
    '''
    ''' Les deux doivent coïncider. Le constater vaut mieux que le croire — c'est la même
    ''' logique que l'équilibre débit/crédit affiché sur chaque pièce.
    ''' </summary>
    Private Sub AfficherLeControle()

        Dim selonHistorique As Decimal = RapportActiviteService.TotalPartBanque(_lignes)
        Dim selonPieces As Long = TotalDesComptes()

        Dim ecart As Decimal = selonHistorique - selonPieces

        Dim texte As String =
            $"Total selon l'historique : {selonHistorique:N0} FCFA     " &
            $"Total porté sur les comptes de commission : {selonPieces:N0} FCFA"

        ' Les montants de l'historique portent deux décimales, la pièce est arrondie à
        ' l'unité : quelques francs d'écart sont l'arrondi, pas une anomalie.
        If Math.Abs(ecart) <= _lignes.Count Then
            lblControle.ForeColor = Drawing.Color.DarkGreen
            lblControle.Text = texte & Environment.NewLine & "Les deux sources concordent."
            Return
        End If

        lblControle.ForeColor = Drawing.Color.Firebrick
        lblControle.Text = texte & Environment.NewLine &
                           $"ÉCART DE {ecart:N0} FCFA — les comptes de commission ont peut-être " &
                           "changé pendant la période, ou une journée n'a pas de pièce conservée."
    End Sub

    Private Function TotalDesComptes() As Long

        Dim total As Long = 0L
        If Not _comptes.Columns.Contains("Net") Then Return total

        For Each ligne As DataRow In _comptes.Rows
            total += Convert.ToInt64(ligne("Net"), CultureInfo.InvariantCulture)
        Next

        Return total
    End Function

#End Region

#Region "Export en PDF"

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        sfdEtat.FileName = If(chkASigner.Checked,
                              $"Commissions-banque-a-signer-{dtpDebut.Value:yyyyMM}.pdf",
                              $"Commissions-banque-{dtpDebut.Value:yyyyMMdd}-{dtpFin.Value:yyyyMMdd}.pdf")
        sfdEtat.OverwritePrompt = True

        If sfdEtat.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim avancement As FrmProgression = Nothing

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, "Export PDF des commissions de la banque")

            ' En PDF et non en classeur : cet état se signe, se classe et se transmet. Un
            ' tableur invite à retoucher les chiffres, et un chiffre retouché dans le fichier
            ' qu'on présente n'est plus le chiffre de la banque.
            ExcelExportService.ExporterEnPdf(
                "ECOBANK TCHAD — COMMISSIONS ENCAISSÉES SUR L'ACTIVITÉ WESTERN UNION",
                SousTitres(), Blocs(), "Commissions", sfdEtat.FileName,
                progression:=avancement.Progression)

            avancement.Fermer()

            lblStatut.ForeColor = Drawing.SystemColors.GrayText
            lblStatut.Text = "PDF produit : " & IO.Path.GetFileName(sfdEtat.FileName) &
                             " — il s'ouvre aussitôt."

        Catch ex As Exception
            If avancement IsNot Nothing Then avancement.Fermer()

            MessageBox.Show(Me, "Export PDF impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Function SousTitres() As List(Of SousTitreExcel)

        ' SousTitreExcel n'a pas de constructeur sans paramètre : le texte se passe à la
        ' construction, et non par un initialiseur d'objet.
        '
        ' La variable ne s'appelle pas « sousTitres » : en Visual Basic, le nom d'une
        ' fonction EST sa variable de retour, et une locale homonyme est refusée (BC30290).
        Dim lignes As New List(Of SousTitreExcel) From {
            New SousTitreExcel($"Période du {dtpDebut.Value:dd/MM/yyyy} au {dtpFin.Value:dd/MM/yyyy}"),
            New SousTitreExcel("Part revenant à la banque, telle qu'elle a été calculée chaque " &
                               "jour. Aucun montant n'est recalculé.")
        }

        ' L'avertissement suit l'état jusque dans le classeur : un chiffre incomplet exporté
        ' sans sa réserve deviendrait un chiffre tout court.
        If lblAvertissement.Text.Length > 0 Then
            lignes.Add(New SousTitreExcel(SurUneLigne(lblAvertissement.Text), True))
        End If

        lignes.Add(New SousTitreExcel(SurUneLigne(lblControle.Text)))

        If chkASigner.Checked Then
            lignes.Add(New SousTitreExcel("ÉDITION DESTINÉE À LA SIGNATURE", True))
        End If

        ' Un état qui se classe doit dire quand il a été tiré : deux éditions d'une même
        ' période peuvent différer si une journée a été annulée entre-temps.
        lignes.Add(New SousTitreExcel($"Édité le {Date.Now:dd/MM/yyyy à HH:mm} par {SessionWU.Auteur}"))

        Return lignes
    End Function

    Private Function Blocs() As List(Of BlocExcel)

        Dim formats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"SousAgents", "# ##0"}, {"AgencesPropres", "# ##0"}, {"TotalBanque", "# ##0"},
            {"Retrocede", "# ##0"}, {"CommissionTotale", "# ##0"},
            {"Debit", "# ##0"}, {"Credit", "# ##0"}, {"Net", "# ##0"}
        }

        Dim entetes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"SousAgents", "Sous-agents"}, {"AgencesPropres", "Agences propres"},
            {"TotalBanque", "TOTAL BANQUE"}, {"Retrocede", "Rétrocédé aux sous-agents"},
            {"CommissionTotale", "Commission totale"}, {"Documentee", "Répartition connue"},
            {"Ecritures", "Écritures"}, {"Net", "Net crédité"}
        }

        Dim liste As New List(Of BlocExcel) From {
            New BlocExcel() With {
                .Titre = "Répartition sur la période",
                .Donnees = _synthese, .Formats = formats, .Entetes = entetes
            },
            New BlocExcel() With {
                .Titre = "Comptes bancaires crédités, d'après les pièces conservées",
                .Donnees = _comptes, .Formats = formats, .Entetes = entetes
            },
            New BlocExcel() With {
                .Titre = "Jour par jour",
                .Donnees = _parJour, .Formats = formats, .Entetes = entetes
            }
        }

        If chkASigner.Checked Then liste.Add(BlocDesSignatures())

        Return liste
    End Function

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

    ''' <summary>
    ''' Les cartouches de signature, pour l'édition mensuelle qui se classe.
    '''
    ''' Trois lignes hautes : une case où l'on doit écrire un nom et signer à la main a
    ''' besoin de place, et la hauteur d'une ligne de tableau ne suffit pas.
    ''' </summary>
    Private Shared Function BlocDesSignatures() As BlocExcel

        Dim table As New DataTable("Signatures")
        table.Columns.Add("Fonction", GetType(String))
        table.Columns.Add("Nom", GetType(String))
        table.Columns.Add("Date", GetType(String))
        table.Columns.Add("Signature", GetType(String))

        table.Rows.Add("Établi par — agent de la compense", String.Empty, String.Empty, String.Empty)
        table.Rows.Add("Vérifié par — chef de service", String.Empty, String.Empty, String.Empty)
        table.Rows.Add("Approuvé par", String.Empty, String.Empty, String.Empty)

        Return New BlocExcel() With {
            .Titre = "Visa et signatures",
            .Donnees = table,
            .HauteurLignes = 42R
        }
    End Function

    ''' <summary>Un libellé de plusieurs lignes, ramené à une seule pour un sous-titre.</summary>
    Private Shared Function SurUneLigne(texte As String) As String
        Return If(texte, String.Empty).Replace(Environment.NewLine, " ").Replace(vbLf, " ")
    End Function

    ''' <summary>La première ligne d'un message, pour un libellé qui n'en tient qu'une.</summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim fin As Integer = texte.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If fin < 0 Then Return texte
        Return texte.Substring(0, fin)
    End Function

#End Region

End Class
