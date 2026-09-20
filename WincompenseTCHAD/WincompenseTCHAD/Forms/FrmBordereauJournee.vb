Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>
''' Le bordereau de fin de journée : ce qui s'imprime, se signe et se classe.
'''
''' CE N'EST PAS UNE PIÈCE COMPTABLE. La pièce porte déjà les quatre cartouches de la banque,
''' et les écritures sont couvertes par une signature. Le bordereau atteste de la FAÇON dont
''' la journée a été faite — et surtout de ce qui n'a pas pu l'être.
'''
''' LE VISA
'''
''' Modifier un taux de sous-agent exige deux personnes ; comptabiliser une journée entière
''' n'en exigeait qu'une. Le visa corrige ce déséquilibre. Il est réservé à la fonction
''' « authorizer », et celui qui a comptabilisé la journée ne peut pas la viser — la base le
''' refuse autant que l'application.
'''
''' Recomptabiliser une journée EFFACE son visa : il atteste d'un traitement précis, et refaire
''' la journée en produit un autre.
''' </summary>
Public Class FrmBordereauJournee

    Private ReadOnly _jour As Date
    Private _contenu As BordereauService.Contenu

    ''' <summary>Constructeur requis par le Concepteur Windows Forms.</summary>
    Public Sub New()
        InitializeComponent()
        _jour = Date.Today
    End Sub

    Public Sub New(journee As Date)
        InitializeComponent()
        _jour = journee.Date
    End Sub

#Region "Ouverture"

    Private Sub FrmBordereauJournee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Charger()
    End Sub

    Private Sub Charger()

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        _contenu = BordereauService.Construire(_jour, messageErreur)
        Cursor = Cursors.Default

        If _contenu Is Nothing Then
            btnExporter.Enabled = False
            btnViser.Enabled = False
            lblJournee.Text = $"Journée du {_jour:dd/MM/yyyy}"
            FrmDiagnostic.Afficher(Me, "Bordereau", messageErreur)
            Return
        End If

        lblJournee.Text = _contenu.Traitement.Intitule

        dgvRecapitulatif.DataSource = TableRecapitulative()
        dgvEcartes.DataSource = _contenu.Ecartes

        FormaterLeRecapitulatif()
        FormaterLesEcartes()

        AfficherLeVisa()
        AfficherLesAlertes()

        btnExporter.Enabled = Not _contenu.EstVide
        ReglerLeBoutonViser()
    End Sub

    ''' <summary>
    ''' Les six blocs du bordereau réunis en une seule grille, avec leur section.
    '''
    ''' Six grilles empilées à l'écran seraient illisibles ; le PDF, lui, les sépare vraiment,
    ''' parce qu'une page imprimée se lit autrement qu'un écran.
    ''' </summary>
    Private Function TableRecapitulative() As DataTable

        Dim table As New DataTable("Recapitulatif")
        table.Columns.Add("Section", GetType(String))
        table.Columns.Add("Rubrique", GetType(String))
        table.Columns.Add("Valeur", GetType(String))

        Ajouter(table, "1. Identification", _contenu.Identification)
        AjouterLesSources(table)
        Ajouter(table, "3. Ce qui a été traité", _contenu.Volumes)
        Ajouter(table, "4. Montants", _contenu.Montants)
        Ajouter(table, "5. Pièce comptable", _contenu.Piece)
        Ajouter(table, "7. Core banking", _contenu.CoreBanking)

        Return table
    End Function

    ''' <summary>Recopie un bloc à deux colonnes sous son intitulé de section.</summary>
    Private Shared Sub Ajouter(cible As DataTable, section As String, source As DataTable)

        If source Is Nothing Then Return

        For Each ligne As DataRow In source.Rows
            cible.Rows.Add(section, Convert.ToString(ligne(0)), Convert.ToString(ligne(1)))
        Next
    End Sub

    ''' <summary>
    ''' Les rapports Western Union ont trois colonnes — nom et empreinte — là où les autres
    ''' blocs en ont deux. L'empreinte suit le nom, sur la même ligne.
    ''' </summary>
    Private Sub AjouterLesSources(cible As DataTable)

        If _contenu.Sources Is Nothing Then Return

        For Each ligne As DataRow In _contenu.Sources.Rows

            Dim empreinte As String = Convert.ToString(ligne(2))
            Dim valeur As String = Convert.ToString(ligne(1))

            If empreinte.Length > 0 Then valeur &= "     empreinte " & empreinte

            cible.Rows.Add("2. Rapports traités", Convert.ToString(ligne(0)), valeur)
        Next
    End Sub

#End Region

#Region "Mise en forme"

    Private Sub FormaterLeRecapitulatif()

        dgvRecapitulatif.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        GrilleWU.LargeurFixe(dgvRecapitulatif, "Section", 150)
        GrilleWU.LargeurFixe(dgvRecapitulatif, "Rubrique", 250)

        If dgvRecapitulatif.Columns.Contains("Valeur") Then
            dgvRecapitulatif.Columns("Valeur").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        MettreEnEvidence()
    End Sub

    ''' <summary>
    ''' Deux lignes doivent sauter aux yeux : le nombre d'Accounts écartés et l'équilibre de
    ''' la pièce. Ce sont celles qu'un supérieur cherche avant de signer.
    ''' </summary>
    Private Sub MettreEnEvidence()

        For rang As Integer = 0 To dgvRecapitulatif.Rows.Count - 1

            Dim rubrique As String = Convert.ToString(dgvRecapitulatif.Rows(rang).Cells("Rubrique").Value)
            Dim valeur As String = Convert.ToString(dgvRecapitulatif.Rows(rang).Cells("Valeur").Value)

            If rubrique.StartsWith("ACCOUNTS ÉCARTÉS", StringComparison.Ordinal) Then
                dgvRecapitulatif.Rows(rang).DefaultCellStyle.Font =
                    New Drawing.Font(dgvRecapitulatif.Font, Drawing.FontStyle.Bold)

                If _contenu.Traitement.NombreEcartes > 0 Then
                    dgvRecapitulatif.Rows(rang).DefaultCellStyle.ForeColor = Drawing.Color.Firebrick
                End If
            End If

            If String.Equals(rubrique, "Équilibre", StringComparison.Ordinal) Then
                dgvRecapitulatif.Rows(rang).DefaultCellStyle.ForeColor =
                    If(String.Equals(valeur, "ÉQUILIBRÉE", StringComparison.Ordinal),
                       Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
            End If
        Next
    End Sub

    Private Sub FormaterLesEcartes()

        dgvEcartes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvEcartes, "Account", "Account", 130)
        Entete(dgvEcartes, "Envois", "Envois", 70)
        Entete(dgvEcartes, "Paiements", "Paiements", 80)
        Entete(dgvEcartes, "PrincipalEnvoi", "Principal envoyé", 130)
        Entete(dgvEcartes, "PrincipalPaye", "Principal payé", 130)
        Entete(dgvEcartes, "Commission", "Commission générée", 140)

        If dgvEcartes.Columns.Contains("Designation") Then
            dgvEcartes.Columns("Designation").HeaderText = "Désignation"
            dgvEcartes.Columns("Designation").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        For Each nom As String In New String() {"Envois", "Paiements", "PrincipalEnvoi",
                                                "PrincipalPaye", "Commission"}
            If Not dgvEcartes.Columns.Contains(nom) Then Continue For
            dgvEcartes.Columns(nom).DefaultCellStyle.Format = "N0"
            dgvEcartes.Columns(nom).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next
    End Sub

    Private Shared Sub Entete(grille As DataGridView, colonne As String, libelle As String,
                              largeur As Integer)

        If Not grille.Columns.Contains(colonne) Then Return
        grille.Columns(colonne).HeaderText = libelle
        GrilleWU.LargeurFixe(grille, colonne, largeur)
    End Sub

#End Region

#Region "Visa et alertes"

    Private Sub AfficherLeVisa()

        If _contenu.Traitement.EstVisee Then
            lblVisa.ForeColor = Drawing.Color.DarkGreen
            lblVisa.Text = _contenu.Traitement.LibelleVisa &
                           If(_contenu.Traitement.CommentaireVisa.Length > 0,
                              "     « " & _contenu.Traitement.CommentaireVisa & " »", String.Empty)
            Return
        End If

        lblVisa.ForeColor = Drawing.Color.Firebrick
        lblVisa.Text = "EN ATTENTE DE VISA — cette journée n'a pas encore été relue par un supérieur."
    End Sub

    ''' <summary>
    ''' Les deux choses qu'un supérieur doit voir avant de signer, écrites en toutes lettres :
    ''' ce qui n'a pas été comptabilisé, et une pièce qui ne s'équilibre pas.
    ''' </summary>
    Private Sub AfficherLesAlertes()

        Dim alertes As New List(Of String)()
        Dim traitement As TraitementJourneeWU = _contenu.Traitement

        If traitement.NombreEcartes > 0 Then
            alertes.Add($"{traitement.NombreEcartes} Account(s) ont travaillé ce jour-là sans être " &
                        "comptabilisés : ils ne figurent nulle part sur la pièce.")
        End If

        If Not traitement.Equilibree Then
            alertes.Add("La pièce de cette journée ne s'équilibre pas.")
        End If

        If Not traitement.Enregistre Then
            alertes.Add("Journée comptabilisée avant la mise en service du bordereau : le nom " &
                        "des rapports Western Union n'a pas été conservé, et elle ne peut pas " &
                        "être visée.")
        End If

        lblEcartes.Text = If(traitement.NombreEcartes > 0,
                             $"Accounts NON COMPTABILISÉS ({traitement.NombreEcartes})",
                             "Accounts non comptabilisés — aucun")

        lblAlerte.Text = String.Join(Environment.NewLine, alertes)
    End Sub

    Private Sub ReglerLeBoutonViser()

        Dim traitement As TraitementJourneeWU = _contenu.Traitement

        btnViser.Enabled = traitement.Enregistre AndAlso
                           Not traitement.EstVisee AndAlso
                           SessionWU.PeutAutoriserLesPointsDeVente AndAlso
                           Not String.Equals(traitement.ComptabilisePar, SessionWU.Auteur,
                                             StringComparison.OrdinalIgnoreCase)
    End Sub

#End Region

#Region "Le visa"

    Private Sub btnViser_Click(sender As Object, e As EventArgs) Handles btnViser.Click

        If _contenu Is Nothing Then Return

        If Not Confirmer() Then Return

        Dim commentaire As String = InputBox(
            "Observation éventuelle (facultatif) :" & Environment.NewLine & Environment.NewLine &
            "Elle sera conservée avec le visa et imprimée sur le bordereau.",
            "Viser la journée", String.Empty)

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim vise As Boolean = TraitementRepository.Viser(_jour, commentaire, messageErreur)
        Cursor = Cursors.Default

        If Not vise Then
            FrmDiagnostic.Afficher(Me, "Visa refusé", messageErreur)
            Return
        End If

        UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                          $"Journée du {_jour:dd/MM/yyyy} visée")

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = "Journée visée."

        Charger()
    End Sub

    ''' <summary>
    ''' Fait relire au supérieur ce qu'il vise. Une confirmation qui répète les chiffres vaut
    ''' mieux qu'une qui demande « êtes-vous sûr ? » : on est toujours sûr.
    ''' </summary>
    Private Function Confirmer() As Boolean

        Dim traitement As TraitementJourneeWU = _contenu.Traitement

        Dim texte As New System.Text.StringBuilder()

        texte.AppendLine($"Viser la journée du {traitement.DateActivite:dd/MM/yyyy} ?")
        texte.AppendLine()
        texte.AppendLine($"{traitement.NombrePdv} point(s) de vente, " &
                         $"{traitement.NombreEnvois} envoi(s), {traitement.NombrePaiements} paiement(s).")
        texte.AppendLine($"Pièce : {traitement.TotalDebit:N0} au débit, " &
                         $"{traitement.TotalCredit:N0} au crédit — " &
                         If(traitement.Equilibree, "équilibrée.", "DÉSÉQUILIBRÉE."))

        If traitement.NombreEcartes > 0 Then
            texte.AppendLine()
            texte.AppendLine($"{traitement.NombreEcartes} ACCOUNT(S) N'ONT PAS ÉTÉ COMPTABILISÉS. " &
                             "Ils figurent au bas de cet écran : relisez-les avant de viser.")
        End If

        texte.AppendLine()
        texte.Append("Votre nom et l'heure seront conservés avec la journée.")

        Return MessageBox.Show(Me, texte.ToString(), "Confirmer le visa",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

#Region "Impression"

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        If _contenu Is Nothing Then Return

        sfdBordereau.FileName = _contenu.NomDeFichier
        sfdBordereau.OverwritePrompt = True

        If sfdBordereau.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim avancement As FrmProgression = Nothing

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, "Impression du bordereau")

            BordereauService.ExporterEnPdf(_contenu, sfdBordereau.FileName, avancement.Progression)

            avancement.Fermer()

            lblStatut.ForeColor = Drawing.SystemColors.GrayText
            lblStatut.Text = "Bordereau produit : " & IO.Path.GetFileName(sfdBordereau.FileName)

        Catch ex As Exception
            If avancement IsNot Nothing Then avancement.Fermer()

            MessageBox.Show(Me, "Impression impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
