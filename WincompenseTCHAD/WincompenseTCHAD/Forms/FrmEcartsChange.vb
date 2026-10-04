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
        btnExporter.Enabled = _pieces IsNot Nothing AndAlso _pieces.Count > 0

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

    Private Sub RemplirLaPiece()

        dgvPiece.DataSource = PieceChangeService.ConsoliderLesLignes(_pieces)

        If dgvPiece.Columns.Count > 0 Then

            dgvPiece.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            AlignerUnMontant(dgvPiece, "Debit", "N0")
            AlignerUnMontant(dgvPiece, "Credit", "N0")

            GrilleWU.LargeurFixe(dgvPiece, "Compte", 120)
            GrilleWU.LargeurFixe(dgvPiece, "CodeAgence", 90)
        End If

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
            $"{_pieces.Count} pièce(s), {dgvPiece.Rows.Count} écritures — " &
            $"gains {gains.ToString("N0", fr)}, pertes {pertes.ToString("N0", fr)} — " &
            $"débits {debits.ToString("N0", fr)}, crédits {credits.ToString("N0", fr)} ({verdict})."

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

#Region "Export"

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        If _pieces Is Nothing OrElse _pieces.Count = 0 Then Return

        sfdExport.FileName = $"Piece_change_{_pieces(0).DateReglement:yyyyMMdd}.xlsx"

        If sfdExport.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim table As DataTable = TableExportable()

        Cursor = Cursors.WaitCursor
        Try
            ' Debit et Credit sont déclarés NUMÉRIQUES : sans cela Excel les écrirait en texte
            ' et la somme de contrôle du comptable rendrait zéro. Compte et CodeAgence restent
            ' du TEXTE, faute de quoi un compte perdrait ses zéros de tête.
            ExcelExportService.ExporterTableBrute(table,
                                                  New String() {"Debit", "Credit", "NombreTransactions"},
                                                  "Pièce de change",
                                                  sfdExport.FileName,
                                                  True)

            lblStatut.ForeColor = Drawing.SystemColors.ControlText
            lblStatut.Text = $"Pièce exportée : {sfdExport.FileName}"

        Catch ex As InvalidOperationException
            MessageBox.Show(Me, ex.Message, "Export impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As System.Runtime.InteropServices.COMException
            ' Excel absent du poste, ou refusant de démarrer. Le message d'Interop est
            ' illisible pour un agent : on le remplace par ce qu'il doit faire.
            MessageBox.Show(Me,
                            "Microsoft Excel n'a pas pu être démarré sur ce poste." & Environment.NewLine &
                            Environment.NewLine &
                            "La pièce reste consultable à l'écran, et la synthèse se copie depuis " &
                            "l'onglet « Synthèse »." & Environment.NewLine & Environment.NewLine &
                            $"Détail technique : {ex.Message}",
                            "Export impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' La table exportée : les écritures, précédées des colonnes qui disent de quelle pièce
    ''' chacune vient. La grille les masque — elle n'affiche qu'une pièce à la fois dans
    ''' l'esprit du lecteur — mais un classeur qui les perdrait serait inexploitable dès qu'il
    ''' porte plus d'une pièce.
    ''' </summary>
    Private Function TableExportable() As DataTable

        Dim table As New DataTable("PieceChange")

        table.Columns.Add("Piece", GetType(String))
        table.Columns.Add("DateReglement", GetType(String))
        table.Columns.Add("Sens", GetType(String))
        table.Columns.Add("CodeProduit", GetType(String))
        table.Columns.Add("NombreTransactions", GetType(Integer))
        table.Columns.Add("Compte", GetType(String))
        table.Columns.Add("Libelle", GetType(String))
        table.Columns.Add("Debit", GetType(Long))
        table.Columns.Add("Credit", GetType(Long))
        table.Columns.Add("CodeAgence", GetType(String))

        For Each piece As PieceChangeWU In _pieces

            If piece.Lignes Is Nothing Then Continue For

            For Each ecriture As DataRow In piece.Lignes.Rows

                Dim row As DataRow = table.NewRow()

                row("Piece") = piece.CleGroupe
                row("DateReglement") = piece.DateReglement.ToString("dd/MM/yyyy")
                row("Sens") = piece.Sens
                row("CodeProduit") = piece.CodeProduit
                row("NombreTransactions") = piece.NombreTransactions
                row("Compte") = Convert.ToString(ecriture("Compte"))
                row("Libelle") = Convert.ToString(ecriture("Libelle"))
                row("Debit") = Convert.ToInt64(ecriture("Debit"))
                row("Credit") = Convert.ToInt64(ecriture("Credit"))
                row("CodeAgence") = Convert.ToString(ecriture("CodeAgence"))

                table.Rows.Add(row)
            Next
        Next

        Return table
    End Function

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
