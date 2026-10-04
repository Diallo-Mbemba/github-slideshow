Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' Rejoue le calcul des écarts de change sur un rapport de règlement, et montre tout : les
''' transactions retenues une à une, les lignes écartées avec leur motif, et les totaux.
'''
''' POURQUOI CET ÉCRAN EXISTE, ET CE QU'IL REMPLACE. La banque a écarté un projet de tests
''' automatisés. Un calcul comptable qui porte un demi-million de francs par semaine ne peut
''' pas pour autant être livré sur parole : il faut pouvoir le CONFRONTER à un rapport dont
''' on connaît déjà le résultat. Cet écran est ce moyen de confrontation, et il a sur un
''' projet de tests un avantage que la banque appréciera : il reste disponible APRÈS la
''' livraison, sur le poste de l'agent, le jour où un comptable contestera un total.
'''
''' IL N'ENREGISTRE RIEN, ET C'EST SA DÉFINITION. Aucune écriture en base, aucune pièce
''' produite, aucun fichier écrit. On peut l'ouvrir sur la production en pleine journée de
''' compense sans prendre le moindre risque : il lit un rapport et calcule en mémoire.
'''
''' IL EXERCE LE CODE LIVRÉ, PAS UNE COPIE. Tout le calcul est dans ChangeService, qui est
''' pur ; cet écran ne fait que lui passer un DataTable et des réglages. Ce qu'il affiche est
''' donc, au franc près, ce que la pièce comptable portera.
''' </summary>
Public Class FrmControleChange

    ''' <summary>Noms des colonnes de la grille. Écrits une fois, employés partout.</summary>
    Private Const COL_MTCN As String = "MTCN"
    Private Const COL_DATE As String = "Règlement"
    Private Const COL_SENS As String = "Sens"
    Private Const COL_PRODUIT As String = "Produit"
    Private Const COL_ACCOUNT As String = "Account"
    Private Const COL_STATUT As String = "Statut"
    Private Const COL_LOCAL As String = "Montant local"
    Private Const COL_DEVISE As String = "Montant devise"
    Private Const COL_CONTREVALEUR As String = "Contre-valeur"
    Private Const COL_ECART As String = "Écart"
    Private Const COL_NATURE As String = "Nature"

    ''' <summary>Le rapport chargé. Nothing tant qu'aucun fichier n'a été choisi.</summary>
    Private _rapport As DataTable

    ''' <summary>Nom du fichier chargé, tel qu'il s'affiche et se recopie dans la synthèse.</summary>
    Private _nomDuRapport As String = String.Empty

    ''' <summary>
    ''' La vue alimentant la grille. Conservée pour que le filtre par MTCN agisse sur elle
    ''' plutôt que de relancer le calcul : filtrer n'est pas recalculer.
    ''' </summary>
    Private _vue As DataView

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Chargement de l'écran"

    Private Sub FrmControleChange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim defauts As OptionsChangeWU = OptionsChangeWU.Actuelles()

        txtParite.Text = defauts.Parite.ToString("0.###", Globalization.CultureInfo.CurrentCulture)
        chkEnAttente.Checked = defauts.InclureEnvoisEnAttente

        lblIntro.Text =
            "Rejoue le calcul des gains et pertes de change sur un rapport de règlement — sans rien " &
            "enregistrer, et avec le code même qui produira la pièce comptable." & Environment.NewLine &
            "Les réglages proposés sont ceux en service : les changer ici ne vaut que pour ce contrôle."

        MettreAJourEtatBoutons()
    End Sub

    Private Sub MettreAJourEtatBoutons()

        btnAfficher.Enabled = _rapport IsNot Nothing
        btnCopier.Enabled = txtSynthese.Text.Length > 0
        txtMtcn.Enabled = _vue IsNot Nothing
    End Sub

#End Region

#Region "Chargement du rapport"

    Private Sub btnCharger_Click(sender As Object, e As EventArgs) Handles btnCharger.Click

        If ofdRapport.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim table As DataTable

        Try
            table = WUReportService.LireRapportWU(ofdRapport.FileName)

        Catch ex As RapportInvalideException
            MessageBox.Show(Me, ex.Message, "Rapport illisible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        ' UNE COLONNE MANQUANTE ARRÊTE TOUT, et ne se contente pas d'un avertissement. Lue
        ' absente, elle vaudrait zéro : un rapport privé de ClearPrincipalLOC afficherait une
        ' contre-valeur nulle et un « gain » égal à la totalité des montants encaissés. Mieux
        ' vaut refuser le fichier que présenter des centaines de millions de francs de gain.
        Dim absentes As List(Of String) = ChangeService.ColonnesManquantes(table)

        If absentes.Count > 0 Then
            MessageBox.Show(Me,
                            "Ce rapport ne porte pas les colonnes nécessaires au calcul des écarts " &
                            "de change :" & Environment.NewLine & Environment.NewLine &
                            "    " & String.Join(", ", absentes) & Environment.NewLine & Environment.NewLine &
                            "Il s'agit probablement d'un rapport d'ACTIVITÉ et non d'un rapport de " &
                            "RÈGLEMENT, ou d'un format que l'application ne connaît pas encore.",
                            "Colonnes manquantes", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        _rapport = table
        _nomDuRapport = IO.Path.GetFileName(ofdRapport.FileName)

        lblFichier.Text = $"{_nomDuRapport} — {table.Rows.Count:N0} lignes, {table.Columns.Count:N0} colonnes."
        lblFichier.ForeColor = Drawing.SystemColors.ControlText

        Vider()
        lblStatut.Text = "Rapport chargé. Cliquez sur Calculer."
        lblStatut.ForeColor = Drawing.SystemColors.ControlText

        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>Efface le résultat précédent : un rapport neuf ne doit pas garder les totaux de l'ancien.</summary>
    Private Sub Vider()

        dgvEcarts.DataSource = Nothing
        _vue = Nothing
        txtSynthese.Text = String.Empty
        txtMtcn.Text = String.Empty
    End Sub

#End Region

#Region "Calcul"

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click

        If _rapport Is Nothing Then Return

        Dim parite As Decimal
        If Not LireLaParite(parite) Then Return

        Dim options As New OptionsChangeWU()
        options.Parite = parite
        options.InclureEnvoisEnAttente = chkEnAttente.Checked

        Dim resultat As ResultatChangeWU

        ' Le sablier est posé sur LA FENÊTRE, et non sur Cursor.Current : dans un formulaire,
        ' le membre hérité Me.Cursor masque le type Windows Forms du même nom, et
        ' « Cursor.Current » se lirait comme un membre du curseur de cette fenêtre. C'est le
        ' même piège que Me.Text face à System.Text, et c'est l'idiome de tout le projet.
        Cursor = Cursors.WaitCursor
        Try
            resultat = ChangeService.Calculer(_rapport, options, _nomDuRapport)
        Finally
            Cursor = Cursors.Default
        End Try

        RemplirLaGrille(resultat)

        txtSynthese.Text = Normaliser(resultat.Synthese())
        txtSynthese.Select(0, 0)

        AfficherLeStatut(resultat)
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>
    ''' Lit la parité saisie. Une parité nulle ou négative est refusée : elle rendrait toutes
    ''' les contre-valeurs nulles, et donc le gain de change égal au chiffre d'affaires.
    '''
    ''' La lecture emploie WUReportService.ToDecimalSafe, qui accepte le point comme la
    ''' virgule : l'agent ne doit pas avoir à devenir le séparateur décimal de son poste.
    ''' </summary>
    Private Function LireLaParite(ByRef parite As Decimal) As Boolean

        parite = WUReportService.ToDecimalSafe(txtParite.Text)

        If parite > 0D Then Return True

        MessageBox.Show(Me,
                        "La parité doit être un nombre strictement positif." & Environment.NewLine &
                        Environment.NewLine &
                        $"Parité de référence : {ConstantesWU.TAUX_CONVERSION.ToString("0.###", Globalization.CultureInfo.CurrentCulture)} " &
                        $"({ConstantesWU.DEVISE_EURO} vers {ConstantesWU.DEVISE_FCFA}).",
                        "Parité invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        txtParite.Focus()
        txtParite.SelectAll()
        Return False
    End Function

    ''' <summary>
    ''' Le statut en bas d'écran : les chiffres qu'on cherche d'abord, et le contrôle.
    '''
    ''' LE CONTRÔLE ARITHMÉTIQUE EST ANNONCÉ MÊME QUAND IL PASSE. Un contrôle qui ne parle
    ''' que pour se plaindre n'est pas lu le jour où il se plaint.
    ''' </summary>
    Private Sub AfficherLeStatut(resultat As ResultatChangeWU)

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")

        lblStatut.Text =
            $"{resultat.Ecarts.Count:N0} transactions retenues, {resultat.Exclusions.Count:N0} écartées — " &
            $"gains {resultat.Gains.ToString("N0", fr)}, pertes {resultat.Pertes.ToString("N0", fr)}, " &
            $"net {resultat.Net.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}." & Environment.NewLine &
            If(resultat.EstCoherent,
               "Contrôle arithmétique : OK (la différence des sommes est égale au net).",
               $"CONTRÔLE EN ÉCHEC : écart de {resultat.EcartDeControle.ToString("N0", fr)} entre les deux " &
               "façons de calculer le net. Ne comptabilisez rien et signalez-le.")

        lblStatut.ForeColor = If(resultat.EstCoherent,
                                 Drawing.SystemColors.ControlText,
                                 Drawing.Color.FromArgb(183, 28, 28))
    End Sub

#End Region

#Region "Grille"

    ''' <summary>
    ''' Porte le détail dans la grille, via un DataTable typé.
    '''
    ''' POURQUOI UN DataTable ET NON LA LISTE D'OBJETS. Deux raisons, et aucune n'est
    ''' esthétique : les colonnes de montants sont alors de VRAIS Decimal, donc le tri par
    ''' écart décroissant range 7 291 avant 998 au lieu de les ranger alphabétiquement ; et le
    ''' filtre par MTCN passe par RowFilter, qui n'exige pas de reconstruire la grille.
    ''' </summary>
    Private Sub RemplirLaGrille(resultat As ResultatChangeWU)

        Dim table As New DataTable("Ecarts")

        table.Columns.Add(COL_MTCN, GetType(String))
        table.Columns.Add(COL_DATE, GetType(Date))
        table.Columns.Add(COL_SENS, GetType(String))
        table.Columns.Add(COL_PRODUIT, GetType(String))
        table.Columns.Add(COL_ACCOUNT, GetType(String))
        table.Columns.Add(COL_STATUT, GetType(String))
        table.Columns.Add(COL_LOCAL, GetType(Decimal))
        table.Columns.Add(COL_DEVISE, GetType(Decimal))
        table.Columns.Add(COL_CONTREVALEUR, GetType(Decimal))
        table.Columns.Add(COL_ECART, GetType(Decimal))
        table.Columns.Add(COL_NATURE, GetType(String))

        For Each ligne As EcartChangeWU In resultat.Ecarts

            Dim row As DataRow = table.NewRow()

            row(COL_MTCN) = ligne.Mtcn
            row(COL_DATE) = If(ligne.DateReglement.HasValue, CType(ligne.DateReglement.Value, Object), DBNull.Value)
            row(COL_SENS) = ligne.SensLisible
            row(COL_PRODUIT) = ligne.CodeProduit
            row(COL_ACCOUNT) = ligne.Account
            row(COL_STATUT) = ligne.Statut
            row(COL_LOCAL) = ligne.MontantLocal
            row(COL_DEVISE) = ligne.MontantEnDevise
            row(COL_CONTREVALEUR) = ligne.ContreValeur
            row(COL_ECART) = ligne.Ecart
            row(COL_NATURE) = ligne.NatureLisible

            table.Rows.Add(row)
        Next

        _vue = New DataView(table)
        dgvEcarts.DataSource = _vue

        FormaterLesColonnes()
    End Sub

    Private Sub FormaterLesColonnes()

        If dgvEcarts.Columns.Count = 0 Then Return

        dgvEcarts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        AlignerUnMontant(COL_LOCAL, "N0")
        AlignerUnMontant(COL_DEVISE, "N2")
        AlignerUnMontant(COL_CONTREVALEUR, "N0")
        AlignerUnMontant(COL_ECART, "N0")

        If dgvEcarts.Columns.Contains(COL_DATE) Then
            dgvEcarts.Columns(COL_DATE).DefaultCellStyle.Format = "dd/MM/yyyy"
        End If

        ' Les colonnes courtes n'ont pas besoin du partage de largeur : le laisser leur donner
        ' autant de place qu'aux montants rejetterait l'écart hors de l'écran.
        GrilleWU.LargeurFixe(dgvEcarts, COL_SENS, 80)
        GrilleWU.LargeurFixe(dgvEcarts, COL_PRODUIT, 70)
        GrilleWU.LargeurFixe(dgvEcarts, COL_ACCOUNT, 90)
        GrilleWU.LargeurFixe(dgvEcarts, COL_STATUT, 60)
        GrilleWU.LargeurFixe(dgvEcarts, COL_NATURE, 70)
        GrilleWU.LargeurFixe(dgvEcarts, COL_DATE, 90)
    End Sub

    Private Sub AlignerUnMontant(nomColonne As String, format As String)

        If Not dgvEcarts.Columns.Contains(nomColonne) Then Return

        Dim colonne As DataGridViewColumn = dgvEcarts.Columns(nomColonne)
        colonne.DefaultCellStyle.Format = format
        colonne.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    ''' <summary>
    ''' Colore l'écart : vert pour un gain, rouge pour une perte, gris pour une conversion exacte.
    '''
    ''' LA COULEUR EST POSÉE À L'AFFICHAGE, et non ligne par ligne après le remplissage. Un
    ''' parcours des 2 400 lignes pour y poser un style serait à refaire à chaque filtre, et
    ''' l'oubli ne se verrait qu'à l'écran. CellFormatting ne se déclenche que sur les cellules
    ''' RÉELLEMENT VISIBLES — une vingtaine — et ne peut pas se désynchroniser des données.
    ''' </summary>
    Private Sub dgvEcarts_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) _
        Handles dgvEcarts.CellFormatting

        If e.ColumnIndex < 0 OrElse e.ColumnIndex >= dgvEcarts.Columns.Count Then Return
        If dgvEcarts.Columns(e.ColumnIndex).Name <> COL_ECART Then Return
        If e.Value Is Nothing OrElse e.Value Is DBNull.Value Then Return

        Dim ecart As Decimal = Convert.ToDecimal(e.Value)

        If ecart > 0D Then
            e.CellStyle.ForeColor = Drawing.Color.FromArgb(27, 94, 32)
        ElseIf ecart < 0D Then
            e.CellStyle.ForeColor = Drawing.Color.FromArgb(183, 28, 28)
        Else
            e.CellStyle.ForeColor = Drawing.SystemColors.GrayText
        End If
    End Sub

    ''' <summary>
    ''' Ne garde que les transactions dont le MTCN contient le texte saisi.
    '''
    ''' C'est la question que la banque pose toujours : « et le transfert 0050908785, vous
    ''' trouvez quoi ? ». Sur deux mille quatre cents lignes, la chercher à la molette n'est
    ''' pas une réponse.
    ''' </summary>
    Private Sub txtMtcn_TextChanged(sender As Object, e As EventArgs) Handles txtMtcn.TextChanged

        If _vue Is Nothing Then Return

        Dim cherche As String = txtMtcn.Text.Trim()

        ' L'apostrophe doublée : sans elle, une saisie malheureuse casserait l'expression de
        ' filtre et lèverait une exception au beau milieu d'une frappe.
        If cherche.Length = 0 Then
            _vue.RowFilter = String.Empty
        Else
            _vue.RowFilter = $"[{COL_MTCN}] LIKE '%{cherche.Replace("'", "''")}%'"
        End If

        lblFiltre.Text = If(cherche.Length = 0,
                            "Saisissez un MTCN pour ne garder que lui ; effacez pour revoir toutes les transactions.",
                            $"{_vue.Count:N0} transaction(s) dont le MTCN contient « {cherche} ».")
    End Sub

#End Region

#Region "Restitution"

    Private Sub btnCopier_Click(sender As Object, e As EventArgs) Handles btnCopier.Click

        If txtSynthese.Text.Length = 0 Then Return

        Try
            Clipboard.SetText(txtSynthese.Text)
            lblStatut.Text = "Synthèse copiée dans le presse-papiers."
            lblStatut.ForeColor = Drawing.SystemColors.ControlText

        Catch ex As System.Runtime.InteropServices.ExternalException
            ' Presse-papiers verrouillé par une autre application : le texte reste lisible et
            ' sélectionnable à l'écran, rien n'est perdu.
            MessageBox.Show(Me,
                            "Le presse-papiers est occupé par une autre application." & Environment.NewLine &
                            Environment.NewLine &
                            "Ouvrez l'onglet « Synthèse », sélectionnez le texte à la souris, puis Ctrl+C.",
                            "Copie impossible", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    ''' <summary>Normalise les sauts de ligne : une zone de texte Windows attend vbCrLf.</summary>
    Private Shared Function Normaliser(texte As String) As String
        Return texte.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
    End Function

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
