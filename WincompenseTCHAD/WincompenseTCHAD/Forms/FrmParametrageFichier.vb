Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>
''' Sortie et rechargement du paramétrage, par fichier.
'''
''' CE QUE CET ÉCRAN N'EST PAS : une sauvegarde de la base. `BACKUP DATABASE` demande un droit
''' de niveau serveur que l'application n'a pas et ne doit pas avoir, écrit sur le disque du
''' serveur et non sur celui du poste, et une sauvegarde qui ne se déclenche que si quelqu'un
''' ouvre l'application n'est pas une sauvegarde. Celle-là se fait sur le serveur, par l'équipe
''' qui le tient.
'''
''' CE QU'IL EST : de quoi remonter une installation NEUVE sans resaisir à la main des centaines
''' de sous-agents. C'est peu, et c'est exactement ce qu'une application de poste peut promettre
''' sans mentir.
'''
''' RÉSERVÉ À L'ADMINISTRATEUR, des deux côtés. Sortir le paramétrage, c'est emporter le plan
''' comptable de la banque ; le charger, c'est écrire dans les tables qui déterminent toutes les
''' écritures.
''' </summary>
Public Class FrmParametrageFichier

    ''' <summary>Le lot lu dans le fichier choisi, ou Nothing tant qu'aucun n'est choisi.</summary>
    Private _lot As LotParametrageWU

    ''' <summary>Ce que le chargement ferait, tel qu'il est affiché.</summary>
    Private _rapport As ParametrageImportService.Rapport

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Ouverture"

    Private Sub FrmParametrageFichier_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesComptesSystemes Then
            MessageBox.Show("Le fichier de secours du paramétrage est réservé à l'administrateur." &
                            Environment.NewLine & Environment.NewLine &
                            "Il porte le plan comptable de la banque et le référentiel complet " &
                            "des points de vente.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        ReglerLesBoutons()
        AfficherLeContenuDeLaBase()
    End Sub

    ''' <summary>
    ''' Montre ce que contient la base, avant d'en faire un fichier. Un export dont on ne sait
    ''' pas ce qu'il emporte ne rassure personne.
    ''' </summary>
    Private Sub AfficherLeContenuDeLaBase()

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim lot As LotParametrageWU = ParametrageExportService.Construire(Nothing, messageErreur)
        Cursor = Cursors.Default

        dgvContenu.DataSource = TableDuContenu(lot)
        FormaterLeContenu()

        If messageErreur.Length > 0 Then
            lblResultatExport.ForeColor = Drawing.Color.Firebrick
            lblResultatExport.Text = PremiereLigne(messageErreur)
            Return
        End If

        lblResultatExport.ForeColor = Drawing.SystemColors.GrayText
        lblResultatExport.Text = lot.Provenance
    End Sub

    Private Shared Function TableDuContenu(lot As LotParametrageWU) As DataTable

        Dim table As New DataTable("Contenu")
        table.Columns.Add("Objet", GetType(String))
        table.Columns.Add("Nombre", GetType(Integer))
        table.Columns.Add("Rechargeable", GetType(String))

        If lot Is Nothing Then Return table

        table.Rows.Add("Comptes comptables", If(lot.Comptes Is Nothing, 0, 9),
                       "oui, sur demande explicite")
        table.Rows.Add("Groupes statistiques", lot.Groupes.Count, "oui")
        table.Rows.Add("Sous-agents", lot.SousAgents.Count, "oui")
        table.Rows.Add("Agences propres", lot.Agences.Count, "oui")
        table.Rows.Add("Utilisateurs", lot.Utilisateurs.Count,
                       "NON — pour mémoire seulement, aucun mot de passe")

        Return table
    End Function

    Private Sub FormaterLeContenu()

        dgvContenu.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvContenu, "Objet", "Objet", 220)
        Entete(dgvContenu, "Nombre", "Nombre", 80)

        If dgvContenu.Columns.Contains("Nombre") Then
            dgvContenu.Columns("Nombre").DefaultCellStyle.Format = "N0"
            dgvContenu.Columns("Nombre").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        If dgvContenu.Columns.Contains("Rechargeable") Then
            dgvContenu.Columns("Rechargeable").HeaderText = "Rechargeable"
            dgvContenu.Columns("Rechargeable").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If
    End Sub

    Private Shared Sub Entete(grille As DataGridView, colonne As String, libelle As String, largeur As Integer)
        If Not grille.Columns.Contains(colonne) Then Return
        grille.Columns(colonne).HeaderText = libelle
        GrilleWU.LargeurFixe(grille, colonne, largeur)
    End Sub

#End Region

#Region "Export"

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        Dim messageErreur As String = String.Empty

        sfdParametrage.FileName = ParametrageFichierWU.NomDeFichier(NomDeLaBase())
        sfdParametrage.OverwritePrompt = True

        If sfdParametrage.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim avancement As FrmProgression = Nothing

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, "Export du paramétrage")
            avancement.Progression.Commencer(ParametrageExportService.ETAPES)

            Dim avertissement As String = String.Empty

            Dim lot As LotParametrageWU = ParametrageExportService.Exporter(
                sfdParametrage.FileName, avancement.Progression, avertissement, messageErreur)

            avancement.Fermer()

            If lot Is Nothing Then
                lblResultatExport.ForeColor = Drawing.Color.Firebrick
                lblResultatExport.Text = PremiereLigne(messageErreur)
                FrmDiagnostic.Afficher(Me, "Export impossible", messageErreur)
                Return
            End If

            AnnoncerLExport(lot, avertissement)

        Catch ex As Exception
            If avancement IsNot Nothing Then avancement.Fermer()

            MessageBox.Show(Me, "Export impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub AnnoncerLExport(lot As LotParametrageWU, avertissement As String)

        Dim nom As String = IO.Path.GetFileName(sfdParametrage.FileName)

        If avertissement.Length > 0 Then
            lblResultatExport.ForeColor = Drawing.Color.Firebrick
            lblResultatExport.Text = $"{nom} — INCOMPLET. {PremiereLigne(avertissement)}"
            FrmDiagnostic.Afficher(Me, "Export incomplet", avertissement)
            Return
        End If

        lblResultatExport.ForeColor = Drawing.Color.DarkGreen
        lblResultatExport.Text = $"{nom} — {lot.Intitule}." & Environment.NewLine &
                                 "Rangez-le hors de ce poste : un fichier de secours qui brûle " &
                                 "avec le serveur ne sert à rien."
    End Sub

    ''' <summary>Le nom de la base, pour nommer le fichier. Vide si la chaîne est inhabituelle.</summary>
    Private Shared Function NomDeLaBase() As String

        Try
            Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder(
                WURepository.ObtenirChaineConnexion())
            Return constructeur.InitialCatalog

        Catch ex As ArgumentException
            Return String.Empty
        Catch ex As FormatException
            Return String.Empty
        End Try
    End Function

#End Region

#Region "Choix du fichier à charger"

    Private Sub btnChoisir_Click(sender As Object, e As EventArgs) Handles btnChoisir.Click

        If ofdParametrage.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        _lot = ParametrageFichierWU.Lire(ofdParametrage.FileName, messageErreur)
        Cursor = Cursors.Default

        If _lot Is Nothing Then
            OublierLeFichier()
            FrmDiagnostic.Afficher(Me, "Archive non exploitable", messageErreur)
            Return
        End If

        lblFichier.Text = IO.Path.GetFileName(ofdParametrage.FileName)
        AfficherLaProvenance()

        If Not ControlerLaBase() Then Return

        Analyser()
    End Sub

    Private Sub AfficherLaProvenance()

        If _lot.Modifie Then
            lblProvenance.ForeColor = Drawing.Color.Firebrick
            lblProvenance.Text = _lot.Provenance & "  —  CONTENU MODIFIÉ depuis l'export."
            Return
        End If

        lblProvenance.ForeColor = Drawing.SystemColors.GrayText
        lblProvenance.Text = _lot.Provenance
    End Sub

    ''' <summary>
    ''' Avertit si le fichier vient d'une AUTRE base que celle à laquelle on est connecté.
    '''
    ''' Ce n'est pas interdit — c'est même le cas normal, puisqu'on charge dans une base neuve.
    ''' Mais charger le paramétrage du Tchad dans la base d'un autre pays produirait un
    ''' référentiel plausible et faux, et rien ne le signalerait ensuite.
    ''' </summary>
    Private Function ControlerLaBase() As Boolean

        Dim base As String = NomDeLaBase()

        If base.Length = 0 OrElse _lot.Base.Length = 0 Then Return True
        If String.Equals(base, _lot.Base, StringComparison.OrdinalIgnoreCase) Then Return True

        Dim reponse As DialogResult = MessageBox.Show(Me,
            $"Ce fichier a été exporté depuis la base « {_lot.Base} »." & Environment.NewLine &
            $"Vous êtes connecté à « {base} »." & Environment.NewLine & Environment.NewLine &
            "C'est normal pour une installation neuve. Ce l'est beaucoup moins si les deux " &
            "bases appartiennent à des pays ou à des entités différentes : le référentiel " &
            "chargé serait plausible et faux." & Environment.NewLine & Environment.NewLine &
            "Continuer l'analyse ?",
            "Base différente", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2)

        If reponse = DialogResult.Yes Then Return True

        OublierLeFichier()
        Return False
    End Function

    Private Sub OublierLeFichier()

        _lot = Nothing
        _rapport = Nothing

        lblFichier.Text = "Aucune archive choisie."
        lblProvenance.Text = String.Empty

        dgvAnalyse.DataSource = Nothing
        dgvComptes.DataSource = Nothing

        ReglerLesBoutons()
    End Sub

#End Region

#Region "Analyse"

    ''' <summary>Confronte le fichier à la base, sans rien écrire.</summary>
    Private Sub Analyser()

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        _rapport = ParametrageImportService.Analyser(_lot, messageErreur)
        Cursor = Cursors.Default

        If messageErreur.Length > 0 Then
            _rapport = Nothing
            ReglerLesBoutons()
            FrmDiagnostic.Afficher(Me, "Analyse impossible", messageErreur)
            Return
        End If

        dgvAnalyse.DataSource = TableDeLAnalyse(_rapport)
        FormaterLAnalyse()

        dgvComptes.DataSource = TableDesComptes(_rapport)
        FormaterLesComptes()

        lblResultatImport.ForeColor = Drawing.SystemColors.ControlText
        lblResultatImport.Text = _rapport.Intitule

        ReglerLesBoutons()
    End Sub

    Private Shared Function TableDeLAnalyse(rapport As ParametrageImportService.Rapport) As DataTable

        Dim table As New DataTable("Analyse")
        table.Columns.Add("Objet", GetType(String))
        table.Columns.Add("Cle", GetType(String))
        table.Columns.Add("Etat", GetType(String))
        table.Columns.Add("Motif", GetType(String))

        For Each ligne As ParametrageImportService.LigneRapport In rapport.Lignes
            table.Rows.Add(ligne.Objet, ligne.Cle, ligne.LibelleEtat, ligne.Motif)
        Next

        Return table
    End Function

    Private Sub FormaterLAnalyse()

        dgvAnalyse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvAnalyse, "Objet", "Objet", 100)
        Entete(dgvAnalyse, "Cle", "Clé", 110)
        Entete(dgvAnalyse, "Etat", "État", 95)

        If dgvAnalyse.Columns.Contains("Motif") Then
            dgvAnalyse.Columns("Motif").HeaderText = "Motif / remarque"
            dgvAnalyse.Columns("Motif").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        ColorerLAnalyse()
    End Sub

    ''' <summary>
    ''' Le refus se voit au premier coup d'œil, et le « déjà présent » s'efface : sur trois
    ''' cents lignes, ce sont les deux seules informations que l'on cherche.
    ''' </summary>
    Private Sub ColorerLAnalyse()

        If _rapport Is Nothing Then Return

        For rang As Integer = 0 To dgvAnalyse.Rows.Count - 1

            If rang >= _rapport.Lignes.Count Then Exit For

            Select Case _rapport.Lignes(rang).Etat

                Case ParametrageImportService.EtatLigneWU.Refusee
                    dgvAnalyse.Rows(rang).DefaultCellStyle.ForeColor = Drawing.Color.Firebrick

                Case ParametrageImportService.EtatLigneWU.DejaPresente
                    dgvAnalyse.Rows(rang).DefaultCellStyle.ForeColor = Drawing.SystemColors.GrayText
            End Select
        Next
    End Sub

    Private Shared Function TableDesComptes(rapport As ParametrageImportService.Rapport) As DataTable

        Dim table As New DataTable("Comptes")
        table.Columns.Add("Compte", GetType(String))
        table.Columns.Add("Actuel", GetType(String))
        table.Columns.Add("Propose", GetType(String))

        For Each ecart As ParametrageColonnes.Difference In rapport.Differences
            table.Rows.Add(ecart.Signification, ecart.Actuel, ecart.Propose)
        Next

        Return table
    End Function

    Private Sub FormaterLesComptes()

        dgvComptes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvComptes, "Actuel", "Dans la base", 90)
        Entete(dgvComptes, "Propose", "Dans le fichier", 90)

        If dgvComptes.Columns.Contains("Compte") Then
            dgvComptes.Columns("Compte").HeaderText = "Compte"
            dgvComptes.Columns("Compte").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

        If _rapport Is Nothing Then Return

        For rang As Integer = 0 To dgvComptes.Rows.Count - 1
            If rang >= _rapport.Differences.Count Then Exit For
            If Not _rapport.Differences(rang).Change Then Continue For

            dgvComptes.Rows(rang).DefaultCellStyle.ForeColor = Drawing.Color.Firebrick
        Next
    End Sub

#End Region

#Region "Chargement"

    Private Sub btnCharger_Click(sender As Object, e As EventArgs) Handles btnCharger.Click

        If _lot Is Nothing OrElse _rapport Is Nothing Then Return

        ' L'analyse est refaite maintenant : entre son affichage et ce clic, quelqu'un d'autre
        ' a pu créer un sous-agent, et une clé dupliquée en pleine transaction ferait tout
        ' échouer au lieu d'être simplement ignorée.
        Analyser()
        If _rapport Is Nothing Then Return

        If _rapport.RienAFaire Then
            MessageBox.Show(Me,
                "Il n'y a rien à charger : tout ce que porte ce fichier est déjà dans la base, " &
                "ou a été refusé." & Environment.NewLine & Environment.NewLine & _rapport.Intitule,
                "Rien à faire", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If Not Confirmer() Then Return

        Dim avancement As FrmProgression = Nothing
        Dim messageErreur As String = String.Empty
        Dim nombreCrees As Integer = 0

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, "Chargement du paramétrage")
            ' Le budget d'étapes doit être EXACT : promettre une étape de plus que ce qui sera
            ' annoncé laisserait la barre s'arrêter juste avant la fin, et douter du reste.
            avancement.Progression.Commencer(
                _rapport.ACreer + If(chkRemplacerComptes.Checked, 1, 0))

            Dim charge As Boolean = ParametrageImportService.Appliquer(
                _lot, _rapport, chkRemplacerComptes.Checked, avancement.Progression,
                nombreCrees, messageErreur)

            avancement.Fermer()

            If Not charge Then
                lblResultatImport.ForeColor = Drawing.Color.Firebrick
                lblResultatImport.Text = "Chargement interrompu — rien n'a été écrit."
                FrmDiagnostic.Afficher(Me, "Chargement interrompu", messageErreur)
                Return
            End If

            AnnoncerLeChargement(nombreCrees, messageErreur)

        Catch ex As Exception
            If avancement IsNot Nothing Then avancement.Fermer()

            MessageBox.Show(Me, "Chargement impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try

        ' La base a changé : l'analyse affichée ne vaut plus, et le contenu exportable non plus.
        Analyser()
        AfficherLeContenuDeLaBase()
    End Sub

    Private Function Confirmer() As Boolean

        Dim texte As New System.Text.StringBuilder()

        texte.AppendLine($"Charger {_rapport.ACreer} ligne(s) dans la base ?")
        texte.AppendLine()
        texte.AppendLine(_rapport.Intitule)

        If chkRemplacerComptes.Checked AndAlso _rapport.ComptesDifferents Then
            texte.AppendLine()
            texte.AppendLine("LES COMPTES COMPTABLES SERONT REMPLACÉS par ceux du fichier.")
            texte.AppendLine("Ce sont eux qui déterminent toute la pièce comptable : relisez la " &
                             "colonne de droite avant de valider.")
        End If

        If _lot.Modifie Then
            texte.AppendLine()
            texte.AppendLine("Le contenu de ce fichier a été MODIFIÉ depuis son export : vous ne " &
                             "chargez plus ce que l'application avait produit.")
        End If

        texte.AppendLine()
        texte.Append("Rien de ce qui est déjà en base ne sera remplacé.")

        Return MessageBox.Show(Me, texte.ToString(), "Confirmer le chargement",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    Private Sub AnnoncerLeChargement(nombreCrees As Integer, messageComptes As String)

        If messageComptes.Length > 0 Then
            lblResultatImport.ForeColor = Drawing.Color.Firebrick
            lblResultatImport.Text = $"{nombreCrees} ligne(s) créée(s), comptes NON chargés."
            FrmDiagnostic.Afficher(Me, "Comptes comptables non chargés", messageComptes)
            Return
        End If

        lblResultatImport.ForeColor = Drawing.Color.DarkGreen
        lblResultatImport.Text = $"{nombreCrees} ligne(s) créée(s)." &
                                 If(chkRemplacerComptes.Checked AndAlso _rapport.ComptesDifferents,
                                    " Comptes comptables remplacés.", String.Empty)
    End Sub

#End Region

#Region "État des boutons"

    Private Sub ReglerLesBoutons()

        Dim analyse As Boolean = _rapport IsNot Nothing

        chkRemplacerComptes.Enabled = analyse AndAlso _rapport.ComptesDifferents
        If Not chkRemplacerComptes.Enabled Then chkRemplacerComptes.Checked = False

        btnCharger.Enabled = analyse AndAlso Not _rapport.RienAFaire

        lblComptes.Text = If(analyse AndAlso _rapport.ComptesDifferents,
                             "Comptes comptables — DIFFÉRENTS du fichier",
                             "Comptes comptables")
    End Sub

    Private Sub chkRemplacerComptes_CheckedChanged(sender As Object, e As EventArgs) Handles chkRemplacerComptes.CheckedChanged

        If _rapport Is Nothing Then Return
        btnCharger.Enabled = Not _rapport.RienAFaire OrElse chkRemplacerComptes.Checked
    End Sub

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

    ''' <summary>La première ligne d'un message, pour un libellé qui n'en tient qu'une.</summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim fin As Integer = texte.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If fin < 0 Then Return texte
        Return texte.Substring(0, fin)
    End Function

#End Region

End Class
