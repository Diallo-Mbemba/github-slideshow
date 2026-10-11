Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' LE FORMAT DU FICHIER CORE BANKING : l'écran où la banque écrit elle-même les codes que
''' porteront les colonnes de son fichier d'interface.
'''
''' POURQUOI CET ÉCRAN EXISTE
'''
''' La banque a déjà changé quatre de ces codes d'un coup — BRN de N01 à S51, TXNCD débit de
''' U24 à F03, TXNCD crédit de F15 à F57, et COST_CENTER retirée. Chacun a coûté un cycle
''' complet : modification du code, compilation, commit, pull, redéploiement sur tous les
''' postes. Ces codes sont les siens ; elle doit pouvoir les corriger sans nous.
'''
''' CE N'EST PAS LE MÊME ENJEU QUE LA NARRATIVE, ET L'ÉCRAN LE TRAITE AUTREMENT
'''
''' Un libellé mal écrit fait une ligne laide dans le grand livre. Un TXNCD mal écrit fait une
''' écriture du MAUVAIS TYPE, ou un lot rejeté le matin de la compense. D'où trois garde-fous
''' plutôt qu'un :
'''
'''   — un CONTRÔLE à la saisie, qui refuse ce qui est faux dans tous les cas ;
'''   — un JOURNAL, qui dit qui a changé quoi, quand, et ce qu'il y avait avant ;
'''   — un APERÇU d'une ligne de fichier, parce que personne ne relit un code de trois
'''     caractères dans une grille, et que tout le monde relit une ligne.
'''
''' CE QUI N'EST PAS ICI, ET N'Y SERA PAS
'''
''' DRCR (« D » / « C ») n'est pas un code de la banque : c'est la sémantique du débit et du
''' crédit. AMOUNT, ACNO, VALDT et ADDLTEXT sont des données. INSTR_NO est toujours vide.
''' BATCHNO est une règle. Voir FormatCoreBankingWU, qui porte ce raisonnement.
'''
''' ADMINISTRATEUR SEUL. Le contrôle est posé par OptionsWU.Enregistrer, et non par cet
''' écran : un droit vérifié à l'affichage se contourne, un droit vérifié à l'écriture non.
''' </summary>
Public Class FrmFormatCoreBanking

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "État de l'écran"

    ''' <summary>Nom de la colonne de la grille des comptes. Une seule, sans en-tête bavard.</summary>
    Private Const COL_COMPTE As String = "Compte"

    ''' <summary>
    ''' Combien de modifications la grille du journal charge. Un journal se consulte et ne se
    ''' télécharge pas.
    ''' </summary>
    Private Const LIMITE_JOURNAL As Integer = 200

    ''' <summary>Vrai si le journal est tenu : son absence n'empêche rien, elle se dit.</summary>
    Private _journalDisponible As Boolean = False

    ''' <summary>
    ''' Vrai pendant le chargement, pour que les événements de saisie ne recalculent pas
    ''' l'aperçu sur un écran encore à moitié rempli.
    ''' </summary>
    Private _enChargement As Boolean = False

#End Region

#Region "Ouverture"

    Private Sub FrmFormatCoreBanking_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Charger()
    End Sub

    Private Sub Charger()

        _enChargement = True
        Try
            ' Les valeurs sont relues en base, et non prises dans le cache : on vient
            ' précisément les modifier, et quelqu'un d'autre a pu le faire avant.
            OptionsWU.Oublier()
            JournalParametreWU.Oublier()

            _journalDisponible = JournalParametreWU.Disponible()

            Dim format As FormatCoreBankingWU = OptionsWU.FormatCoreBanking

            txtDetbsjrnl.Text = format.Detbsjrnl
            txtBrn.Text = format.Brn
            txtSrccode.Text = format.Srccode
            txtTxncdDebit.Text = format.TxncdDebit
            txtTxncdCredit.Text = format.TxncdCredit
            txtAgenceSiege.Text = format.AgenceSiege

            For Each zone As TextBox In Zones()
                zone.MaxLength = ConstantesWU.CB_CODE_LONGUEUR_MAX
            Next

            ChargerLesComptes(format)
            ChargerLeJournal()

            lblDerniereModification.Text = Historique()
            lblStatut.ForeColor = Drawing.SystemColors.GrayText
            lblStatut.Text = String.Empty

        Finally
            _enChargement = False
        End Try

        AfficherLApercu()
    End Sub

    ''' <summary>Les six zones de code, dans l'ordre où l'écran les présente.</summary>
    Private Function Zones() As List(Of TextBox)

        Dim liste As New List(Of TextBox)

        liste.Add(txtDetbsjrnl)
        liste.Add(txtBrn)
        liste.Add(txtSrccode)
        liste.Add(txtTxncdDebit)
        liste.Add(txtTxncdCredit)
        liste.Add(txtAgenceSiege)

        Return liste
    End Function

    ''' <summary>L'intitulé d'une zone, pour les messages. Celui de l'étiquette, sans ses deux points.</summary>
    Private Function IntituleDe(zone As TextBox) As String

        If zone Is txtDetbsjrnl Then Return "DETBSJRNL"
        If zone Is txtBrn Then Return "BRN (agence du lot)"
        If zone Is txtSrccode Then Return "SRCCODE"
        If zone Is txtTxncdDebit Then Return "TXNCD au débit"
        If zone Is txtTxncdCredit Then Return "TXNCD au crédit"
        If zone Is txtAgenceSiege Then Return "ACBRN (agence du siège)"

        Return zone.Name
    End Function

    ''' <summary>La dernière trace connue, ou la raison pour laquelle il n'y en a pas.</summary>
    Private Function Historique() As String

        Dim trace As String = OptionsWU.DerniereModification(OptionsWU.CLE_CB_BRN)

        If trace.Length = 0 Then
            Return "Aucun code n'est enregistré dans la base : ceux qui s'affichent sont ceux " &
                   "de l'application, et le fichier sort comme aujourd'hui."
        End If

        Return trace
    End Function

#End Region

#Region "Les comptes du siège"

    ''' <summary>
    ''' Remplit la grille des comptes rattachés d'office au siège.
    '''
    ''' UNE SEULE COLONNE, ET PAS D'EN-TÊTE BAVARD : la phrase d'aide au-dessus dit tout. Une
    ''' grille à une colonne large se saisit mieux qu'un champ texte à virgules, et c'est elle
    ''' qui cache le séparateur — la banque voit des comptes, pas un format de stockage.
    ''' </summary>
    Private Sub ChargerLesComptes(format As FormatCoreBankingWU)

        Dim table As New DataTable("comptesDuSiege")
        table.Columns.Add(COL_COMPTE, GetType(String))

        For Each compte As String In format.ComptesDuSiege
            table.Rows.Add(compte)
        Next

        dgvComptes.DataSource = table

        If Not dgvComptes.Columns.Contains(COL_COMPTE) Then Return

        dgvComptes.Columns(COL_COMPTE).HeaderText = "Numéro de compte"
        dgvComptes.Columns(COL_COMPTE).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
    End Sub

    ''' <summary>
    ''' Les comptes actuellement saisis dans la grille, vides et doublons compris.
    '''
    ''' LES DOUBLONS NE SONT PAS ABSORBÉS ICI, et c'est voulu : c'est le contrôle qui les
    ''' signale. Celui qui a saisi deux fois le même compte doit l'apprendre, et non le voir
    ''' disparaître sans explication.
    ''' </summary>
    Private Function ComptesSaisis() As List(Of String)

        Dim comptes As New List(Of String)

        For Each ligne As DataGridViewRow In dgvComptes.Rows

            If ligne.IsNewRow Then Continue For

            Dim cellule As DataGridViewCell = ligne.Cells(COL_COMPTE)
            If cellule Is Nothing OrElse cellule.Value Is Nothing Then Continue For

            Dim compte As String = Convert.ToString(cellule.Value).Trim()
            If compte.Length = 0 Then Continue For

            comptes.Add(compte)
        Next

        Return comptes
    End Function

    Private Sub dgvComptes_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvComptes.CellEndEdit
        AfficherLApercu()
    End Sub

    Private Sub dgvComptes_UserDeletedRow(sender As Object, e As DataGridViewRowEventArgs) Handles dgvComptes.UserDeletedRow
        AfficherLApercu()
    End Sub

#End Region

#Region "L'aperçu"

    Private Sub Code_TextChanged(sender As Object, e As EventArgs) _
        Handles txtDetbsjrnl.TextChanged, txtBrn.TextChanged, txtSrccode.TextChanged,
                txtTxncdDebit.TextChanged, txtTxncdCredit.TextChanged, txtAgenceSiege.TextChanged

        If _enChargement Then Return
        AfficherLApercu()
    End Sub

    ''' <summary>
    ''' Montre une ligne du fichier telle qu'elle sortira, avec les codes saisis.
    '''
    ''' L'APERÇU EST CALCULÉ PAR LE MÊME CODE QUE LE FICHIER — LigneCoreBankingWU.Construire,
    ''' et non une phrase reconstituée ici. Un aperçu qui s'écrirait à côté finirait par mentir
    ''' le jour où l'un des deux changerait.
    '''
    ''' LA LIGNE MONTRÉE EST UN DÉBIT SUR UN COMPTE DU SIÈGE, et ce choix n'est pas neutre :
    ''' c'est celle qui emploie le PLUS de codes à la fois — DETBSJRNL, BRN, SRCCODE, le TXNCD
    ''' du débit, et l'agence du siège par la règle d'aiguillage. Une ligne de point de vente
    ''' n'en montrerait pas autant.
    ''' </summary>
    Private Sub AfficherLApercu()

        Dim format As New FormatCoreBankingWU(txtDetbsjrnl.Text, txtBrn.Text, txtSrccode.Text,
                                              txtTxncdDebit.Text, txtTxncdCredit.Text,
                                              txtAgenceSiege.Text, ComptesSaisis())

        ' Le compte du siège s'il y en a un : c'est lui qui déclenche la règle d'aiguillage.
        ' À défaut, un compte quelconque — l'aperçu montre alors l'agence du point de vente.
        Dim compte As String = "123456789"
        If format.ComptesDuSiege.Count > 0 Then compte = format.ComptesDuSiege(0)

        Dim ligne As LigneCoreBankingWU = LigneCoreBankingWU.Construire(
            compte, "LIGNE D'EXEMPLE", 1250000L, True, "N07", Date.Today, "07Q4", format)

        lblApercu.Text =
            $"{ligne.DETBSJRNL}  {ligne.BRN}  {ligne.BATCHNO}  {ligne.SRCCODE}  " &
            $"{ligne.AMOUNT}  {ligne.ACNO}  {ligne.DRCR}  {ligne.ACBRN}  {ligne.TXNCD}  " &
            $"{ligne.VALDT:dd/MM/yyyy}"
    End Sub

#End Region

#Region "Le journal"

    Private Sub ChargerLeJournal()

        If Not _journalDisponible Then
            dgvJournal.DataSource = Nothing
            lblJournal.ForeColor = Drawing.Color.Firebrick
            lblJournal.Text = JournalParametreWU.MESSAGE_TABLE_ABSENTE.Replace(Environment.NewLine, " ")
            Return
        End If

        Dim messageErreur As String = String.Empty
        Dim lignes As List(Of JournalParametreWU.Modification) =
            JournalParametreWU.Lister(LIMITE_JOURNAL, messageErreur)

        dgvJournal.DataSource = lignes
        HabillerLaGrilleDuJournal()

        If messageErreur.Length > 0 Then
            lblJournal.ForeColor = Drawing.Color.Firebrick
            lblJournal.Text = messageErreur.Replace(Environment.NewLine, " ")
            Return
        End If

        lblJournal.ForeColor = Drawing.SystemColors.GrayText

        If lignes.Count = 0 Then
            lblJournal.Text = "Aucune modification enregistrée : le paramétrage est celui " &
                              "d'origine, ou il a été changé avant la création du journal."
            Return
        End If

        lblJournal.Text =
            $"{lignes.Count} modification(s), de la plus récente à la plus ancienne " &
            $"(les {LIMITE_JOURNAL} dernières). Ce journal est celui de TOUT le paramétrage, " &
            "narrative comprise, et il ne peut être ni modifié ni effacé depuis l'application."
    End Sub

    Private Sub HabillerLaGrilleDuJournal()

        If dgvJournal.Columns.Contains("Cle") Then dgvJournal.Columns("Cle").Visible = False
        If dgvJournal.Columns.Contains("AncienneValeur") Then dgvJournal.Columns("AncienneValeur").Visible = False

        Renommer("DateModification", "Date", 130)
        Renommer("Intitule", "Ce qui a changé", 220)
        Renommer("Avant", "Avant", 0)
        Renommer("NouvelleValeur", "Après", 0)
        Renommer("ModifiePar", "Par", 110)
        Renommer("Poste", "Poste", 120)

        If dgvJournal.Columns.Contains("DateModification") Then
            dgvJournal.Columns("DateModification").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm"
        End If
    End Sub

    ''' <summary>Renomme une colonne et lui donne sa largeur. Largeur 0 : la colonne s'étire.</summary>
    Private Sub Renommer(nom As String, entete As String, largeur As Integer)

        If Not dgvJournal.Columns.Contains(nom) Then Return

        dgvJournal.Columns(nom).HeaderText = entete

        If largeur = 0 Then
            dgvJournal.Columns(nom).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Return
        End If

        dgvJournal.Columns(nom).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgvJournal.Columns(nom).Width = largeur
    End Sub

#End Region

#Region "Les codes par défaut"

    Private Sub btnDefaut_Click(sender As Object, e As EventArgs) Handles btnDefaut.Click

        txtDetbsjrnl.Text = ConstantesWU.CB_DETBSJRNL
        txtBrn.Text = ConstantesWU.CB_BRN
        txtSrccode.Text = ConstantesWU.CB_SRCCODE
        txtTxncdDebit.Text = ConstantesWU.CB_TXNCD_DEBIT
        txtTxncdCredit.Text = ConstantesWU.CB_TXNCD_CREDIT
        txtAgenceSiege.Text = ConstantesWU.CB_AGENCE_SIEGE

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = "Codes par défaut rétablis à l'écran — pas encore enregistrés. " &
                         "Les comptes du siège, eux, ne sont pas touchés."
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        ' LA CELLULE EN COURS DE SAISIE EST VALIDÉE D'ABORD. Sans cela, cliquer Enregistrer
        ' sans avoir quitté la case lirait l'ANCIENNE valeur : le compte que la banque vient
        ' de taper serait perdu, et l'écran annoncerait un enregistrement réussi.
        dgvComptes.EndEdit()

        Dim messageErreur As String = String.Empty

        For Each zone As TextBox In Zones()
            If Not FormatCoreBankingWU.ControlerLeCode(zone.Text, IntituleDe(zone), messageErreur) Then
                Refuser("Le format n'est pas enregistrable.", messageErreur)
                zone.Focus()
                Return
            End If
        Next

        Dim comptes As List(Of String) = ComptesSaisis()

        If Not FormatCoreBankingWU.ControlerLesComptes(comptes, messageErreur) Then
            Refuser("Les comptes du siège ne sont pas enregistrables.", messageErreur)
            ongletsFormat.SelectedTab = ongComptes
            Return
        End If

        If Not Confirmer(comptes) Then Return

        Cursor = Cursors.WaitCursor
        Dim ecrit As Boolean
        Try
            ecrit = EcrireLeFormat(comptes, messageErreur)
        Finally
            Cursor = Cursors.Default
        End Try

        If Not ecrit Then
            ' LA BASE PEUT ÊTRE À MOITIÉ ÉCRITE, et c'est le vrai danger de cet écran : les
            ' sept valeurs partent l'une après l'autre, et l'échec peut survenir à la
            ' quatrième. L'écran est donc relu en base AVANT de rendre la main, pour montrer
            ' les codes réellement enregistrés et non ceux qui viennent d'être tapés. Sans
            ' cela, l'administrateur repartirait convaincu d'avoir un format qu'il n'a pas.
            Charger()
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Format enregistré INCOMPLET. Les codes affichés sont ceux que " &
                             "porte la base : corrigez, puis réenregistrez."
            Return
        End If

        UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                          "Format du fichier core banking modifié")

        ' Tout est relu plutôt que supposé : le journal vient de gagner des lignes.
        Charger()

        lblStatut.ForeColor = Drawing.Color.DarkGreen
        lblStatut.Text = "Format enregistré. Il vaut pour les fichiers À VENIR ; ceux déjà " &
                         "chargés portent les codes en vigueur ce jour-là."
    End Sub

    ''' <summary>
    ''' Écrit les sept valeurs du format, et s'arrête à la première qui résiste.
    '''
    ''' CHAQUE ÉCRITURE EST ATOMIQUE AVEC SON JOURNAL — OptionsWU.Enregistrer s'en charge —
    ''' MAIS LES SEPT NE LE SONT PAS ENTRE ELLES. C'est assumé : une transaction unique
    ''' tiendrait sept valeurs de paramétrage pour un écran que seul l'administrateur ouvre,
    ''' là où relire la base après l'échec dit déjà la vérité à celui qui saisit.
    ''' </summary>
    Private Function EcrireLeFormat(comptes As List(Of String),
                                    ByRef messageErreur As String) As Boolean

        If Not Ecrire(OptionsWU.CLE_CB_DETBSJRNL, txtDetbsjrnl.Text,
                      OptionsWU.LIBELLE_CB_DETBSJRNL, messageErreur) Then Return False
        If Not Ecrire(OptionsWU.CLE_CB_BRN, txtBrn.Text,
                      OptionsWU.LIBELLE_CB_BRN, messageErreur) Then Return False
        If Not Ecrire(OptionsWU.CLE_CB_SRCCODE, txtSrccode.Text,
                      OptionsWU.LIBELLE_CB_SRCCODE, messageErreur) Then Return False
        If Not Ecrire(OptionsWU.CLE_CB_TXNCD_DEBIT, txtTxncdDebit.Text,
                      OptionsWU.LIBELLE_CB_TXNCD_DEBIT, messageErreur) Then Return False
        If Not Ecrire(OptionsWU.CLE_CB_TXNCD_CREDIT, txtTxncdCredit.Text,
                      OptionsWU.LIBELLE_CB_TXNCD_CREDIT, messageErreur) Then Return False
        If Not Ecrire(OptionsWU.CLE_CB_AGENCE_SIEGE, txtAgenceSiege.Text,
                      OptionsWU.LIBELLE_CB_AGENCE_SIEGE, messageErreur) Then Return False

        ' LA LISTE PART EN UNE SEULE VALEUR, et le journal en garde une ligne lisible :
        ' « avant : 379100319,379200585 ». C'est ce qui a fait préférer une option à une
        ' table pour deux numéros de compte sans attribut.
        Return Ecrire(OptionsWU.CLE_CB_COMPTES_SIEGE,
                      FormatCoreBankingWU.TexteDepuisComptes(comptes),
                      OptionsWU.LIBELLE_CB_COMPTES_SIEGE, messageErreur)
    End Function

    ''' <summary>
    ''' Écrit une option, et dit ce qui a échoué.
    '''
    ''' LE DROIT EST VÉRIFIÉ PAR OptionsWU.Enregistrer, pas ici : l'administrateur seul peut
    ''' écrire, et c'est contrôlé à l'écriture plutôt qu'à l'affichage.
    ''' </summary>
    Private Function Ecrire(cle As String, valeur As String, libelle As String,
                            ByRef messageErreur As String) As Boolean

        If OptionsWU.Enregistrer(cle, If(valeur, String.Empty).Trim(), libelle, messageErreur) Then
            Return True
        End If

        Refuser($"Le code {cle} n'a pas été enregistré.", messageErreur)
        Return False
    End Function

    ''' <summary>
    ''' Fait relire les codes avant de les enregistrer.
    '''
    ''' CE N'EST PAS UNE PRÉCAUTION DE FORME. Ces codes partiront sur TOUTES les lignes de TOUS
    ''' les lots à venir, et un code faux ne se voit pas : le fichier se produit, se charge, et
    ''' c'est la comptabilité qui découvre des écritures du mauvais type.
    ''' </summary>
    Private Function Confirmer(comptes As List(Of String)) As Boolean

        Dim detail As String =
            $"DETBSJRNL   {txtDetbsjrnl.Text.Trim()}" & Environment.NewLine &
            $"BRN         {txtBrn.Text.Trim()}" & Environment.NewLine &
            $"SRCCODE     {txtSrccode.Text.Trim()}" & Environment.NewLine &
            $"TXNCD D     {txtTxncdDebit.Text.Trim()}" & Environment.NewLine &
            $"TXNCD C     {txtTxncdCredit.Text.Trim()}" & Environment.NewLine &
            $"ACBRN siège {txtAgenceSiege.Text.Trim()}"

        Dim listeComptes As String =
            If(comptes.Count = 0, "aucun",
               String.Join(", ", comptes) & $"  ({comptes.Count})")

        Return MessageBox.Show(Me,
            "Enregistrer ce format de fichier core banking ?" & Environment.NewLine & Environment.NewLine &
            detail & Environment.NewLine & Environment.NewLine &
            "Comptes rattachés au siège : " & listeComptes & Environment.NewLine & Environment.NewLine &
            "Une ligne sortira ainsi :" & Environment.NewLine & lblApercu.Text & Environment.NewLine &
            Environment.NewLine &
            "Ces codes partent sur toutes les lignes de tous les lots à venir. Les fichiers " &
            "déjà chargés portent ceux qui étaient en vigueur ce jour-là, et ne sont pas " &
            "réécrits.",
            "Confirmer le format", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    ''' <summary>
    ''' Dit en une ligne que rien n'a été enregistré, et ouvre le détail.
    '''
    ''' Le paramètre ne s'appelle PAS « resume » : Visual Basic réserve ce mot à l'instruction
    ''' Resume Next, et une variable qui le porte est refusée (BC30183).
    ''' </summary>
    Private Sub Refuser(constat As String, detail As String)

        lblStatut.ForeColor = Drawing.Color.Firebrick
        lblStatut.Text = constat

        MessageBox.Show(Me, detail, constat, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
