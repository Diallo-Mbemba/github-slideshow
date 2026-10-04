Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' Produit la PIÈCE COMPTABLE des écarts de change à partir d'un rapport de règlement, et la
''' conserve.
'''
''' CE QUI DISTINGUE CET ÉCRAN DE FrmControleChange. Celui-ci ÉCRIT en base ; l'autre ne fait
''' que vérifier. Le calcul est le même — c'est le même ChangeService — mais ici la parité
''' vient du PARAMÉTRAGE et non d'une saisie, les comptes viennent de SystemeWU, et la pièce
''' produite est destinée à être passée. Ce sont deux usages d'un seul calcul, et c'est
''' délibéré : si l'écran de contrôle avait son propre calcul, il ne contrôlerait rien.
'''
''' L'ENCHAÎNEMENT EST IMPOSÉ, ET CHAQUE ÉTAPE DÉVERROUILLE LA SUIVANTE.
'''
'''     Charger  ->  Calculer  ->  Produire la pièce  ->  Conserver
'''
''' On ne peut pas conserver ce qu'on n'a pas vu : le bouton reste gris jusqu'à ce que la
''' pièce existe et soit équilibrée. Une comptabilisation ne doit jamais tenir à un seul clic.
'''
''' LA PARITÉ VIENT DE LA BASE, ET SON ORIGINE EST AFFICHÉE. Le paramétrage l'historise avec
''' sa date d'effet ; l'écran lit celle qui s'appliquait au jour du règlement, et dit laquelle
''' il a employée. Un rapport qui chevaucherait un changement de parité est REFUSÉ plutôt que
''' calculé avec l'une des deux au hasard — le cas ne s'est jamais présenté, la parité du
''' franc CFA n'ayant pas bougé depuis 1999, mais c'est le genre de cas qu'on ne découvre
''' jamais à temps.
''' </summary>
Public Class FrmEcartsChange

    ''' <summary>Noms des colonnes de la grille du détail.</summary>
    Private Const COL_MTCN As String = "MTCN"
    Private Const COL_DATE As String = "Règlement"
    Private Const COL_SENS As String = "Sens"
    Private Const COL_PRODUIT As String = "Produit"
    Private Const COL_LOCAL As String = "Montant local"
    Private Const COL_DEVISE As String = "Montant devise"
    Private Const COL_CONTREVALEUR As String = "Contre-valeur"
    Private Const COL_ECART As String = "Écart"

    Private _rapport As DataTable
    Private _nomDuRapport As String = String.Empty
    Private _parite As Decimal = ConstantesWU.TAUX_CONVERSION
    Private _origineDeLaParite As String = String.Empty
    Private _resultat As ResultatChangeWU
    Private _pieces As List(Of PieceChangeWU)

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Chargement de l'écran"

    Private Sub FrmEcartsChange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RemplirLesDecoupages()
        MettreAJourEtatBoutons()

        ' LES COMPTES SONT RELUS À L'OUVERTURE, et ce n'est pas une précaution de confort.
        ' ComptesSystemeWU.Actuels n'est renseigné depuis la base que par l'écran de traitement
        ' de la compense : un agent qui ouvrirait directement celui-ci travaillerait sur les
        ' valeurs PAR DÉFAUT du code — où les deux comptes de change sont vides — et l'écran
        ' lui annoncerait un paramétrage absent alors qu'il est en place.
        Dim motif As String = String.Empty
        Dim comptes As ComptesSystemeWU = WURepository.ChargerComptesSysteme(motif)
        If motif.Length = 0 Then ComptesSystemeWU.Actuels = comptes

        ' Le paramétrage des comptes est annoncé AVANT tout travail. Découvrir qu'un compte
        ' manque après avoir chargé un rapport et attendu un calcul sur 2 400 lignes est une
        ' perte de temps évitable.
        AnnoncerLeParametrage()
    End Sub

    ''' <summary>
    ''' Les trois découpages, et le nombre de pièces que chacun produit. Le libellé le dit :
    ''' c'est la seule information qui permette de choisir.
    ''' </summary>
    Private Sub RemplirLesDecoupages()

        cboDecoupage.Items.Clear()
        cboDecoupage.Items.Add("Une pièce par journée de règlement")
        cboDecoupage.Items.Add("Une pièce par journée et par sens")
        cboDecoupage.Items.Add("Une pièce par journée, sens et code produit")
        cboDecoupage.SelectedIndex = 0
    End Sub

    Private Function DecoupageChoisi() As DecoupageChangeWU

        Select Case cboDecoupage.SelectedIndex
            Case 1 : Return DecoupageChangeWU.ParJourneeEtSens
            Case 2 : Return DecoupageChangeWU.ParJourneeSensEtProduit
            Case Else : Return DecoupageChangeWU.ParJournee
        End Select
    End Function

    Private Sub AnnoncerLeParametrage()

        Dim manquants As List(Of String) = ComptesSystemeWU.Actuels.ComptesDeChangeManquants()

        If manquants.Count = 0 Then
            lblStatut.ForeColor = Drawing.SystemColors.ControlText
            lblStatut.Text = "Chargez un rapport de règlement, puis calculez." & Environment.NewLine &
                             $"Comptes de change : gain {ComptesSystemeWU.Actuels.CompteGainDeChange}, " &
                             $"perte {ComptesSystemeWU.Actuels.ComptePerteDeChange}, " &
                             $"liaison {ComptesSystemeWU.Actuels.CompteCourant}."
            Return
        End If

        lblStatut.ForeColor = Drawing.Color.FromArgb(183, 28, 28)
        lblStatut.Text = "Comptes de change non paramétrés : " & String.Join(", ", manquants) & "." &
                         Environment.NewLine &
                         "Le calcul reste consultable, mais aucune pièce ne pourra être produite. " &
                         "Voir Paramétrage > Comptes systèmes."
    End Sub

    Private Sub MettreAJourEtatBoutons()

        btnAfficher.Enabled = _rapport IsNot Nothing
        btnPiece.Enabled = _resultat IsNot Nothing AndAlso _resultat.Ecarts.Count > 0

        Dim piecesPretes As Boolean = _pieces IsNot Nothing AndAlso _pieces.Count > 0

        btnExporter.Enabled = piecesPretes
        cboPiece.Enabled = piecesPretes

        ' LE FICHIER CORE BANKING EXIGE L'ÉQUILIBRE DE TOUTES LES PIÈCES, et pas seulement de
        ' celle qu'on regarde : il couvre la journée entière. CoreBankingService le revérifie
        ' de son côté — ce fichier impacte des comptes réels, le contrôle est au dernier verrou.
        btnCoreBanking.Enabled = piecesPretes AndAlso
                                 _pieces.Where(Function(p) Not p.EstEquilibree).Count() = 0

        ' CONSERVER EXIGE TROIS CHOSES À LA FOIS : une pièce, le droit de traiter la compense,
        ' et des pièces toutes équilibrées. Les trois sont vérifiées ici, et de nouveau par le
        ' dépôt au moment d'écrire : un bouton grisé est un confort, pas une sécurité.
        btnEnregistrer.Enabled = _pieces IsNot Nothing AndAlso _pieces.Count > 0 AndAlso
                                 SessionWU.PeutTraiterLaCompense AndAlso
                                 _pieces.Where(Function(p) Not p.EstEquilibree).Count() = 0
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

        Dim absentes As List(Of String) = ChangeService.ColonnesManquantes(table)

        If absentes.Count > 0 Then
            MessageBox.Show(Me,
                            "Ce rapport ne porte pas les colonnes nécessaires au calcul des écarts " &
                            "de change :" & Environment.NewLine & Environment.NewLine &
                            "    " & String.Join(", ", absentes) & Environment.NewLine & Environment.NewLine &
                            "Il s'agit probablement d'un rapport d'ACTIVITÉ et non d'un rapport de " &
                            "RÈGLEMENT.",
                            "Colonnes manquantes", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        _rapport = table
        _nomDuRapport = IO.Path.GetFileName(ofdRapport.FileName)

        Vider()
        LireLaParite()

        lblFichier.Text = $"{_nomDuRapport} — {table.Rows.Count:N0} lignes."
        lblFichier.ForeColor = Drawing.SystemColors.ControlText

        lblStatut.ForeColor = Drawing.SystemColors.ControlText
        lblStatut.Text = $"Rapport chargé. {_origineDeLaParite}" & Environment.NewLine &
                         "Cliquez sur Calculer."

        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>
    ''' Lit la parité applicable au rapport, et retient d'où elle vient.
    '''
    ''' La journée de référence est la date de règlement du rapport, telle que
    ''' WUReportService la reconstitue depuis les colonnes SetDateLOC — la même date que celle
    ''' employée par le contrôle de cohérence de la compense, pour que les deux traitements ne
    ''' puissent pas se rattacher à deux jours différents.
    ''' </summary>
    Private Sub LireLaParite()

        Dim jour As Date? = WUReportService.ObtenirDateReglement(_rapport)
        Dim journee As Date = If(jour.HasValue, jour.Value, Date.Today)

        Dim motif As String = String.Empty
        _parite = ChangeRepository.PariteEnVigueur(journee, motif)

        If motif.Length = 0 Then
            _origineDeLaParite = $"Parité {Nombre(_parite)} en vigueur au {journee:dd/MM/yyyy}, " &
                                 "lue dans le paramétrage."
        Else
            _origineDeLaParite = $"Parité {Nombre(_parite)} : repli du code — {PremiereLigne(motif)}"
        End If

        lblParite.Text = $"Parité : {Nombre(_parite)}"
        lblParite.ForeColor = If(motif.Length = 0,
                                 Drawing.SystemColors.GrayText,
                                 Drawing.Color.FromArgb(230, 81, 0))
    End Sub

    ''' <summary>Efface tout résultat précédent : un rapport neuf ne garde rien de l'ancien.</summary>
    Private Sub Vider()

        _resultat = Nothing
        _pieces = Nothing

        dgvEcarts.DataSource = Nothing
        dgvPiece.DataSource = Nothing
        cboPiece.Items.Clear()
        txtSynthese.Text = String.Empty
        lblTotaux.Text = "Aucune pièce produite."
    End Sub

#End Region

#Region "Calcul"

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click

        If _rapport Is Nothing Then Return

        Dim options As New OptionsChangeWU()
        options.Parite = _parite
        options.InclureEnvoisEnAttente = OptionsWU.ChangeInclureEnvoisEnAttente

        Cursor = Cursors.WaitCursor
        Try
            _resultat = ChangeService.Calculer(_rapport, options, _nomDuRapport)
        Finally
            Cursor = Cursors.Default
        End Try

        _pieces = Nothing
        dgvPiece.DataSource = Nothing
        cboPiece.Items.Clear()
        lblTotaux.Text = "Calcul effectué. Cliquez sur « Produire la pièce »."

        RemplirLeDetail()
        txtSynthese.Text = Normaliser(_resultat.Synthese())
        txtSynthese.Select(0, 0)

        AfficherLeStatutDuCalcul()
        MettreAJourEtatBoutons()
    End Sub

    Private Sub AfficherLeStatutDuCalcul()

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")

        Dim controle As String = If(_resultat.EstCoherent,
                                    "contrôle arithmétique OK",
                                    "CONTRÔLE EN ÉCHEC — ne comptabilisez rien")

        lblStatut.Text =
            $"{_resultat.Ecarts.Count:N0} transactions retenues, {_resultat.Exclusions.Count:N0} écartées — " &
            $"gains {_resultat.Gains.ToString("N0", fr)}, pertes {_resultat.Pertes.ToString("N0", fr)}, " &
            $"net {_resultat.Net.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA} ({controle})." &
            Environment.NewLine & _origineDeLaParite

        lblStatut.ForeColor = If(_resultat.EstCoherent,
                                 Drawing.SystemColors.ControlText,
                                 Drawing.Color.FromArgb(183, 28, 28))
    End Sub

#End Region

#Region "Pièce comptable"

    Private Sub btnPiece_Click(sender As Object, e As EventArgs) Handles btnPiece.Click

        If _resultat Is Nothing Then Return
        If Not VerifierLaPariteDeLaPeriode() Then Return

        Dim messageErreur As String = String.Empty

        Dim pieces As List(Of PieceChangeWU) =
            PieceChangeService.Generer(_resultat, ComptesSystemeWU.Actuels, DecoupageChoisi(), messageErreur)

        If pieces Is Nothing Then
            MessageBox.Show(Me, messageErreur, "Pièce de change impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Un message rendu AVEC des pièces est un avertissement, et non un refus : c'est le
        ' cas des transactions sans date de règlement, qui n'entrent dans aucune pièce.
        If messageErreur.Length > 0 Then
            MessageBox.Show(Me, messageErreur, "Avertissement",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        _pieces = pieces

        RemplirLaPiece()

        txtSynthese.Text = Normaliser(PieceChangeService.Synthese(_pieces, DecoupageChoisi()) &
                                      Environment.NewLine & Environment.NewLine &
                                      _resultat.Synthese())
        txtSynthese.Select(0, 0)

        onglets.SelectedTab = pagePiece
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>
    ''' Refuse de produire une pièce si le rapport chevauche un changement de parité.
    '''
    ''' DEUX PARITÉS SUR UNE MÊME PIÈCE N'AURAIENT AUCUN SENS : les contre-valeurs auraient été
    ''' calculées avec l'une, les écritures portées sous l'autre, et rien dans la pièce ne le
    ''' dirait. Le cas est théorique — la parité du franc CFA n'a pas bougé depuis 1999 — mais
    ''' s'il se présentait, il faudrait traiter le rapport journée par journée, et c'est ce que
    ''' le message demande.
    ''' </summary>
    Private Function VerifierLaPariteDeLaPeriode() As Boolean

        If Not _resultat.PremiereDate.HasValue OrElse Not _resultat.DerniereDate.HasValue Then Return True

        Dim motif As String = String.Empty
        Dim auDebut As Decimal = ChangeRepository.PariteEnVigueur(_resultat.PremiereDate.Value, motif)
        Dim aLaFin As Decimal = ChangeRepository.PariteEnVigueur(_resultat.DerniereDate.Value, motif)

        If auDebut = aLaFin Then Return True

        MessageBox.Show(Me,
                        "Ce rapport couvre une période sur laquelle la parité a changé :" &
                        Environment.NewLine & Environment.NewLine &
                        $"    au {_resultat.PremiereDate.Value:dd/MM/yyyy} : {Nombre(auDebut)}" & Environment.NewLine &
                        $"    au {_resultat.DerniereDate.Value:dd/MM/yyyy} : {Nombre(aLaFin)}" &
                        Environment.NewLine & Environment.NewLine &
                        "Une pièce ne peut pas porter deux parités. Traitez ce rapport journée par " &
                        "journée, ou corrigez la date d'effet dans le paramétrage de la parité.",
                        "Changement de parité sur la période", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Return False
    End Function

    ''' <summary>
    ''' Remplit la liste des pièces produites, et affiche la première.
    '''
    ''' UNE PIÈCE À LA FOIS, ET C'EST CE QUE LA BANQUE A DEMANDÉ. Afficher toutes les pièces
    ''' bout à bout dans une seule grille obligeait à ajouter des colonnes pour dire de
    ''' quelle pièce chaque ligne venait — journée, sens, produit — c'est-à-dire exactement
    ''' les colonnes dont une pièce comptable ne doit pas s'encombrer. Une pièce se choisit
    ''' donc, et ne montre que ses quatre colonnes.
    ''' </summary>
    Private Sub RemplirLaPiece()

        cboPiece.Items.Clear()

        For Each piece As PieceChangeWU In _pieces
            cboPiece.Items.Add(LibelleDuChoix(piece))
        Next

        ' L'affectation déclenche SelectedIndexChanged, donc l'affichage de la pièce.
        If cboPiece.Items.Count > 0 Then cboPiece.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' Ce que la liste des pièces affiche : de quoi reconnaître une pièce sans l'ouvrir.
    ''' </summary>
    Private Shared Function LibelleDuChoix(piece As PieceChangeWU) As String

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")

        Dim qualificatifs As New List(Of String)
        If piece.Sens.Length > 0 Then qualificatifs.Add(PieceChangeService.LibelleDuSens(piece.Sens))
        If piece.CodeProduit.Length > 0 Then qualificatifs.Add(piece.CodeProduit)

        Dim precision As String = If(qualificatifs.Count = 0,
                                     String.Empty,
                                     " " & String.Join(" ", qualificatifs))

        Return $"{piece.DateReglement:dd/MM/yyyy}{precision} — {piece.NombreTransactions} trx — " &
               $"net {piece.Net.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}"
    End Function

    ''' <summary>La pièce choisie dans la liste, ou Nothing.</summary>
    Private Function PieceChoisie() As PieceChangeWU

        If _pieces Is Nothing Then Return Nothing
        If cboPiece.SelectedIndex < 0 OrElse cboPiece.SelectedIndex >= _pieces.Count Then Return Nothing

        Return _pieces(cboPiece.SelectedIndex)
    End Function

    Private Sub cboPiece_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles cboPiece.SelectedIndexChanged

        AfficherLaPieceChoisie()
    End Sub

    ''' <summary>
    ''' Affiche la pièce choisie : QUATRE COLONNES, et rien d'autre.
    '''
    ''' CodeAgence reste dans les données et disparaît de l'écran. Elle n'est pas décorative —
    ''' c'est elle qui aiguille chaque écriture vers son agence dans le fichier core banking —
    ''' mais elle n'a rien à faire sur une pièce que lit un comptable. La pièce principale la
    ''' masque de la même façon, au même endroit de son code.
    ''' </summary>
    Private Sub AfficherLaPieceChoisie()

        Dim piece As PieceChangeWU = PieceChoisie()

        If piece Is Nothing OrElse piece.Lignes Is Nothing Then
            dgvPiece.DataSource = Nothing
            lblTotaux.Text = "Aucune pièce produite."
            Return
        End If

        dgvPiece.DataSource = piece.Lignes

        If dgvPiece.Columns.Count > 0 Then

            dgvPiece.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            If dgvPiece.Columns.Contains("CodeAgence") Then
                dgvPiece.Columns("CodeAgence").Visible = False
            End If

            AlignerUnMontant(dgvPiece, "Debit", "N0")
            AlignerUnMontant(dgvPiece, "Credit", "N0")

            GrilleWU.LargeurFixe(dgvPiece, "Compte", 140)
        End If

        AfficherLesTotaux(piece)
    End Sub

    ''' <summary>
    ''' Les totaux de la pièce affichée, puis ceux de l'ensemble.
    '''
    ''' LES DEUX SONT NÉCESSAIRES, et pour deux lectures différentes : la pièce qu'on passe
    ''' au journal, et le rapport qu'on rapproche du calcul. N'afficher que la première
    ''' laisserait croire, sur un rapport de trois journées, que le gain de change du jour est
    ''' le gain de change du rapport.
    ''' </summary>
    Private Sub AfficherLesTotaux(piece As PieceChangeWU)

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")

        Dim debits As Long = _pieces.Sum(Function(p) p.TotalDebit)
        Dim credits As Long = _pieces.Sum(Function(p) p.TotalCredit)
        Dim gains As Decimal = _pieces.Sum(Function(p) p.Gains)
        Dim pertes As Decimal = _pieces.Sum(Function(p) p.Pertes)
        Dim desequilibrees As Integer = _pieces.Where(Function(p) Not p.EstEquilibree).Count()

        Dim verdict As String = If(desequilibrees = 0,
                                   "équilibre OK",
                                   $"{desequilibrees} PIÈCE(S) DÉSÉQUILIBRÉE(S)")

        lblTotaux.Text =
            $"Cette pièce : {piece.Lignes.Rows.Count} écritures, " &
            $"débit {piece.TotalDebit.ToString("N0", fr)} = crédit {piece.TotalCredit.ToString("N0", fr)} " &
            $"{ConstantesWU.DEVISE_FCFA}." & Environment.NewLine &
            $"Les {_pieces.Count} pièce(s) : gains {gains.ToString("N0", fr)}, " &
            $"pertes {pertes.ToString("N0", fr)}, débits {debits.ToString("N0", fr)}, " &
            $"crédits {credits.ToString("N0", fr)} ({verdict})."

        lblTotaux.ForeColor = If(desequilibrees = 0,
                                 Drawing.SystemColors.ControlText,
                                 Drawing.Color.FromArgb(183, 28, 28))
    End Sub

#End Region

#Region "Détail"

    Private Sub RemplirLeDetail()

        Dim table As New DataTable("Detail")

        table.Columns.Add(COL_MTCN, GetType(String))
        table.Columns.Add(COL_DATE, GetType(Date))
        table.Columns.Add(COL_SENS, GetType(String))
        table.Columns.Add(COL_PRODUIT, GetType(String))
        table.Columns.Add(COL_LOCAL, GetType(Decimal))
        table.Columns.Add(COL_DEVISE, GetType(Decimal))
        table.Columns.Add(COL_CONTREVALEUR, GetType(Decimal))
        table.Columns.Add(COL_ECART, GetType(Decimal))

        For Each ligne As EcartChangeWU In _resultat.Ecarts

            Dim row As DataRow = table.NewRow()

            row(COL_MTCN) = ligne.Mtcn
            row(COL_DATE) = If(ligne.DateReglement.HasValue,
                               CType(ligne.DateReglement.Value, Object), DBNull.Value)
            row(COL_SENS) = ligne.SensLisible
            row(COL_PRODUIT) = ligne.CodeProduit
            row(COL_LOCAL) = ligne.MontantLocal
            row(COL_DEVISE) = ligne.MontantEnDevise
            row(COL_CONTREVALEUR) = ligne.ContreValeur
            row(COL_ECART) = ligne.Ecart

            table.Rows.Add(row)
        Next

        dgvEcarts.DataSource = table

        If dgvEcarts.Columns.Count = 0 Then Return

        dgvEcarts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        AlignerUnMontant(dgvEcarts, COL_LOCAL, "N0")
        AlignerUnMontant(dgvEcarts, COL_DEVISE, "N2")
        AlignerUnMontant(dgvEcarts, COL_CONTREVALEUR, "N0")
        AlignerUnMontant(dgvEcarts, COL_ECART, "N0")

        dgvEcarts.Columns(COL_DATE).DefaultCellStyle.Format = "dd/MM/yyyy"

        GrilleWU.LargeurFixe(dgvEcarts, COL_DATE, 90)
        GrilleWU.LargeurFixe(dgvEcarts, COL_SENS, 80)
        GrilleWU.LargeurFixe(dgvEcarts, COL_PRODUIT, 70)
    End Sub

    Private Shared Sub AlignerUnMontant(grille As DataGridView, nomColonne As String, format As String)

        If Not grille.Columns.Contains(nomColonne) Then Return

        Dim colonne As DataGridViewColumn = grille.Columns(nomColonne)
        colonne.DefaultCellStyle.Format = format
        colonne.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    ''' <summary>Colore l'écart : vert le gain, rouge la perte. Voir FrmControleChange, même règle.</summary>
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

#End Region

#Region "Export de la pièce"

    ''' <summary>
    ''' Exporte LA PIÈCE AFFICHÉE, et ses quatre colonnes.
    '''
    ''' QUATRE COLONNES, PAS DIX. La première version du classeur portait aussi la clé de la
    ''' pièce, la journée, le sens, le code produit, le nombre de transactions et le code
    ''' agence : nécessaire pour distinguer les pièces d'un classeur qui les contenait toutes,
    ''' inutile dès lors qu'on en exporte UNE. Une pièce comptable a quatre colonnes — compte,
    ''' libellé, débit, crédit — et tout le reste est du contexte, qui appartient au nom du
    ''' fichier et à l'en-tête, pas aux écritures.
    ''' </summary>
    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        Dim piece As PieceChangeWU = PieceChoisie()
        If piece Is Nothing Then Return

        sfdExport.FileName = NomDuClasseur(piece)

        If sfdExport.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim table As DataTable = TableDeLaPiece(piece)

        Cursor = Cursors.WaitCursor
        Try
            ' Debit et Credit sont déclarés NUMÉRIQUES : sans cela Excel les écrirait en texte
            ' et la somme de contrôle du comptable rendrait zéro. Compte reste du TEXTE, faute
            ' de quoi un numéro de compte perdrait ses zéros de tête.
            ExcelExportService.ExporterTableBrute(table,
                                                  New String() {"Debit", "Credit"},
                                                  "Pièce de change",
                                                  sfdExport.FileName,
                                                  True)

            lblStatut.ForeColor = Drawing.SystemColors.ControlText
            lblStatut.Text = $"Pièce exportée : {sfdExport.FileName}"

        Catch ex As InvalidOperationException
            MessageBox.Show(Me, ex.Message, "Export impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As System.Runtime.InteropServices.COMException
            AvertirDExcel(ex)

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Les quatre colonnes de la pièce : compte, libellé, débit, crédit.
    '''
    ''' CodeAgence n'y est pas. Elle reste dans les données de la pièce, où le fichier core
    ''' banking la lit pour aiguiller chaque écriture, mais elle n'a rien à faire sur le
    ''' document que lit un comptable.
    ''' </summary>
    Private Shared Function TableDeLaPiece(piece As PieceChangeWU) As DataTable

        Dim table As New DataTable("PieceChange")

        table.Columns.Add("Compte", GetType(String))
        table.Columns.Add("Libelle", GetType(String))
        table.Columns.Add("Debit", GetType(Long))
        table.Columns.Add("Credit", GetType(Long))

        If piece.Lignes Is Nothing Then Return table

        For Each ecriture As DataRow In piece.Lignes.Rows

            table.Rows.Add(Convert.ToString(ecriture("Compte")),
                           Convert.ToString(ecriture("Libelle")),
                           Convert.ToInt64(ecriture("Debit")),
                           Convert.ToInt64(ecriture("Credit")))
        Next

        Return table
    End Function

    ''' <summary>
    ''' Nom proposé pour le classeur. Il porte ce que les colonnes ne portent plus : la
    ''' journée, et le sens et le produit quand le découpage les distingue.
    ''' </summary>
    Private Shared Function NomDuClasseur(piece As PieceChangeWU) As String

        Dim morceaux As New List(Of String)
        morceaux.Add("Piece_change")
        morceaux.Add(piece.DateReglement.ToString("yyyyMMdd"))

        If piece.Sens.Length > 0 Then morceaux.Add(piece.Sens)
        If piece.CodeProduit.Length > 0 Then morceaux.Add(piece.CodeProduit)

        Return String.Join("_", morceaux) & ".xlsx"
    End Function

#End Region

#Region "Fichier core banking des écarts de change"

    ''' <summary>
    ''' Produit le fichier d'interface core banking des écarts de change.
    '''
    ''' IL EST SÉPARÉ DE CELUI DE LA COMPENSATION, ET SOUS SON PROPRE NUMÉRO DE LOT. Le
    ''' numéro de lot est dérivé de la journée : les deux fichiers d'un même jour seraient
    ''' arrivés sous le même numéro, et le core banking n'a que ce numéro pour reconnaître un
    ''' lot déjà chargé — il aurait rejeté le second comme doublon du premier. Le lot de change
    ''' porte donc la lettre « c » : « c7ob » là où la compensation du 27/03/2026 donne
    ''' « 07ob ». Voir CoreBankingService.NumeroDeLotDeChange.
    '''
    ''' LE FICHIER COUVRE UNE JOURNÉE, ET NON UNE PIÈCE. Un lot du core banking est une
    ''' journée ; si le découpage produit plusieurs pièces pour le 27/03 — par sens, par
    ''' produit — elles appartiennent toutes au même lot et au même fichier. Le fichier
    ''' reprend donc TOUTES les pièces de change de la journée de la pièce affichée.
    '''
    ''' LA DATE DE VALEUR EST LA JOURNÉE DE RÈGLEMENT, et non le jour de la production. Un
    ''' même fichier réexporté la semaine suivante doit être identique au premier : avec la
    ''' date du jour, il aurait porté le même numéro de lot et un contenu différent, ce qui
    ''' est exactement ce qu'un contrôle de doublon ne sait pas démêler.
    ''' </summary>
    Private Sub btnCoreBanking_Click(sender As Object, e As EventArgs) Handles btnCoreBanking.Click

        Dim piece As PieceChangeWU = PieceChoisie()
        If piece Is Nothing Then Return

        Dim journee As Date = piece.DateReglement.Date

        Dim duJour As List(Of PieceChangeWU) =
            _pieces.Where(Function(p) p.DateReglement.Date = journee).ToList()

        Dim lot As String = CoreBankingService.NumeroDeLotDeChange(journee)

        Dim messageErreur As String = String.Empty

        Dim fichier As DataTable = CoreBankingService.ConstruireSousLot(
            PieceChangeService.ConsoliderLesLignes(duJour), lot, journee, messageErreur)

        If fichier Is Nothing Then
            MessageBox.Show(Me, messageErreur, "Fichier core banking impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim reponse As DialogResult = MessageBox.Show(Me,
            $"Produire le fichier core banking des écarts de change du {journee:dd/MM/yyyy} ?" &
            Environment.NewLine & Environment.NewLine &
            $"    pièces de la journée .... {duJour.Count}" & Environment.NewLine &
            $"    lignes du fichier ....... {fichier.Rows.Count}" & Environment.NewLine &
            $"    numéro de lot ........... {lot}" & Environment.NewLine &
            $"    date de valeur .......... {journee:dd/MM/yyyy}" & Environment.NewLine & Environment.NewLine &
            "Ce lot est distinct de celui de la compensation de la même journée : les deux " &
            "fichiers se chargent l'un après l'autre sans se confondre.",
            "Fichier core banking des écarts de change", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then Return

        sfdExport.FileName = CoreBankingService.NomDeFichierDeChange(journee)
        If sfdExport.ShowDialog(Me) <> DialogResult.OK Then Return

        Cursor = Cursors.WaitCursor
        Try
            ' Seul AMOUNT est écrit en nombre : tout le reste est du texte, sans quoi Excel
            ' réinterpréterait les numéros de compte et le numéro de lot — « c7ob » resterait
            ' du texte quand un lot tout en chiffres deviendrait un nombre.
            ExcelExportService.ExporterTableBrute(fichier, New String() {"AMOUNT"},
                                                  "CoreBanking", sfdExport.FileName, True)

            lblStatut.ForeColor = Drawing.SystemColors.ControlText
            lblStatut.Text = $"Fichier core banking produit : {IO.Path.GetFileName(sfdExport.FileName)} " &
                             $"— lot {lot}, {fichier.Rows.Count} lignes."

        Catch ex As InvalidOperationException
            MessageBox.Show(Me, ex.Message, "Production impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As System.Runtime.InteropServices.COMException
            AvertirDExcel(ex)

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Excel absent du poste, ou refusant de démarrer. Le message d'Interop est illisible
    ''' pour un agent : on le remplace par ce qu'il doit faire.
    ''' </summary>
    Private Sub AvertirDExcel(ex As System.Runtime.InteropServices.COMException)

        MessageBox.Show(Me,
                        "Microsoft Excel n'a pas pu être démarré sur ce poste." & Environment.NewLine &
                        Environment.NewLine &
                        "La pièce reste consultable à l'écran, et la synthèse se copie depuis " &
                        "l'onglet « Synthèse »." & Environment.NewLine & Environment.NewLine &
                        $"Détail technique : {ex.Message}",
                        "Export impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

#End Region

#Region "Conservation"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        If _pieces Is Nothing OrElse _pieces.Count = 0 Then Return

        If Not SessionWU.PeutTraiterLaCompense Then
            MessageBox.Show(Me, "La conservation d'une pièce est réservée aux utilisateurs " &
                                "habilités à traiter la compense.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' CE QUI EST DÉJÀ EN BASE EST DEMANDÉ AVANT D'ÉCRIRE. Le dépôt refuserait de toute
        ' façon — un index unique y veille — mais son refus arriverait sur un MTCN quelconque
        ' et ne dirait pas combien de transactions sont concernées. Demander d'abord permet de
        ' distinguer « ce rapport est déjà enregistré en entier » de « trois de ses
        ' transactions le sont », qui n'appellent pas la même réponse.
        Dim motif As String = String.Empty
        Dim deja As Integer = ChangeRepository.TransactionsDejaConservees(_resultat, motif)

        If deja < 0 Then
            MessageBox.Show(Me, motif, "Vérification impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If deja > 0 Then
            MessageBox.Show(Me,
                            $"{deja:N0} des {_resultat.Ecarts.Count:N0} transactions de ce rapport sont " &
                            "DÉJÀ conservées en base." & Environment.NewLine & Environment.NewLine &
                            If(deja = _resultat.Ecarts.Count,
                               "Ce rapport a donc déjà été enregistré en entier : il n'y a rien à faire.",
                               "Le rapport chevauche un enregistrement précédent. Enregistrer de nouveau " &
                               "doublerait l'écart de change de ces transactions.") &
                            Environment.NewLine & Environment.NewLine &
                            "Rien n'a été enregistré.",
                            "Déjà conservé", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")

        Dim gains As String = _pieces.Sum(Function(p) p.Gains).ToString("N0", fr)
        Dim pertes As String = _pieces.Sum(Function(p) p.Pertes).ToString("N0", fr)

        Dim reponse As DialogResult = MessageBox.Show(Me,
            $"Conserver {_pieces.Count} pièce(s) de change ?" & Environment.NewLine & Environment.NewLine &
            $"    transactions ....... {_resultat.Ecarts.Count:N0}" & Environment.NewLine &
            $"    gains .............. {gains} {ConstantesWU.DEVISE_FCFA}" & Environment.NewLine &
            $"    pertes ............. {pertes} {ConstantesWU.DEVISE_FCFA}" & Environment.NewLine &
            $"    parité employée .... {Nombre(_parite)}" & Environment.NewLine & Environment.NewLine &
            "L'enregistrement est définitif : la base refusera ensuite le même rapport, pour que " &
            "l'écart de change ne puisse pas être compté deux fois.",
            "Conserver la pièce de change", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty

            If Not ChangeRepository.Conserver(_resultat, _pieces, messageErreur) Then
                MessageBox.Show(Me, messageErreur, "Enregistrement impossible",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.ForeColor = Drawing.SystemColors.ControlText
            Dim net As String = _pieces.Sum(Function(p) p.Net).ToString("N0", fr)

            lblStatut.Text = $"{_pieces.Count} pièce(s) de change conservée(s) — " &
                             $"net {net} {ConstantesWU.DEVISE_FCFA}."

            MessageBox.Show(Me,
                            "La pièce de change et son détail sont conservés." & Environment.NewLine &
                            Environment.NewLine &
                            "Le détail transaction par transaction reste consultable : c'est lui qui " &
                            "permettra de répondre si un montant est contesté.",
                            "Enregistrement effectué", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Après enregistrement, la conservation n'a plus d'objet : la base refuserait, et
            ' laisser le bouton actif inviterait à un second clic inutile.
            btnEnregistrer.Enabled = False

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

#End Region

#Region "Utilitaires"

    ''' <summary>La parité, écrite avec ses décimales significatives et sans zéro inutile.</summary>
    Private Shared Function Nombre(valeur As Decimal) As String
        Return valeur.ToString("0.######", Globalization.CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' La première ligne d'un message multiligne. Les messages des dépôts expliquent
    ''' longuement quoi faire ; un libellé d'écran n'a la place que de la première phrase, et
    ''' le reste se lit dans la fenêtre qui s'ouvre au besoin.
    ''' </summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim normalise As String = Normaliser(If(texte, String.Empty))
        Dim fin As Integer = normalise.IndexOf(Environment.NewLine, StringComparison.Ordinal)

        Return If(fin < 0, normalise, normalise.Substring(0, fin))
    End Function

    ''' <summary>Normalise les sauts de ligne : une zone de texte Windows attend vbCrLf.</summary>
    Private Shared Function Normaliser(texte As String) As String
        Return texte.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
    End Function

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
