Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Linq
Imports System.Windows.Forms

''' <summary>
''' Consultation des pièces comptables déjà produites.
'''
''' CE QUE CET ÉCRAN N'EST PAS : un recalcul. Il relit ce qui a été écrit le jour de la
''' comptabilisation, ligne à ligne, sans toucher au paramétrage courant. Une pièce ancienne
''' recalculée avec les taux d'aujourd'hui ne serait pas celle qui a été visée et signée.
'''
''' POURQUOI UNE JOURNÉE À LA FOIS, ET NON UNE PLAGE
''' Les rapports d'activité se consultent sur une période parce qu'ils cumulent. Une pièce, non :
''' c'est un justificatif daté, rattaché aux deux rapports Western Union d'UNE journée. Les
''' additionner sur une plage produirait un document qui ne correspond à aucun téléchargement
''' de la plateforme — et donc à rien de vérifiable.
'''
''' L'écran ne fait qu'aiguiller : la pièce s'ouvre dans FrmPieceComptable et le fichier dans
''' FrmFichierCoreBanking, les mêmes écrans que le jour de la compense. L'agent y retrouve ses
''' repères, et il n'y a qu'une façon d'afficher une pièce dans toute l'application.
''' </summary>
Public Class FrmPiecesArchivees

    Private _journees As List(Of PieceRepository.JourneeConservee)
    Private _piece As DataTable
    ''' <summary>
    ''' Ce qui est actuellement affiché : la date NE SUFFIT PAS, puisqu'une même journée
    ''' peut figurer une fois en vigueur et plusieurs fois en archive.
    ''' </summary>
    Private _cleAffichee As String = String.Empty

    ''' <summary>
    ''' Police barrée des journées annulées. Conservée d'une actualisation à l'autre : en
    ''' recréer une à chaque fois laisserait autant d'objets graphiques derrière soi.
    ''' </summary>
    Private _policeBarree As Drawing.Font

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
        _journees = New List(Of PieceRepository.JourneeConservee)()
        _piece = New DataTable()
    End Sub

#Region "Ouverture et chargement"

    Private Sub FrmPiecesArchivees_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Charger()
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        Charger()
    End Sub

    ''' <summary>Relit la liste des journées conservées, et présente la plus récente.</summary>
    Private Sub Charger()

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        _journees = PieceRepository.ListerJournees(messageErreur)
        Cursor = Cursors.Default

        dgvJournees.DataSource = TableDesJournees()
        FormaterLesJournees()

        If messageErreur.Length > 0 Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = PremiereLigne(messageErreur)
            FrmDiagnostic.Afficher(Me, "Pièces conservées", messageErreur)
            AfficherLaPiece(Nothing)
            Return
        End If

        lblStatut.ForeColor = Drawing.SystemColors.GrayText

        If _journees.Count = 0 Then
            lblStatut.Text = "Aucune pièce conservée pour l'instant."
            lblSousTitre.Text = "Les journées comptabilisées AVANT la mise en service de cette " &
                                "conservation n'ont pas de pièce : elles ne peuvent pas être retrouvées."
            AfficherLaPiece(Nothing)
            Return
        End If

        lblStatut.Text = $"{_journees.Count} journée(s) conservée(s)."

        ' La plus récente est en tête : c'est celle qu'on vient de comptabiliser, donc celle
        ' qu'on vient consulter neuf fois sur dix.
        If dgvJournees.Rows.Count > 0 Then dgvJournees.Rows(0).Selected = True
        PresenterLaSelection()
    End Sub

    ''' <summary>La liste des journées, sous la forme qu'attend une grille.</summary>
    Private Function TableDesJournees() As DataTable

        Dim table As New DataTable("Journees")
        table.Columns.Add("Journee", GetType(String))
        table.Columns.Add("Ecritures", GetType(Integer))
        table.Columns.Add("Debit", GetType(Long))
        table.Columns.Add("Credit", GetType(Long))
        table.Columns.Add("Equilibre", GetType(String))
        table.Columns.Add("Conservee", GetType(String))
        table.Columns.Add("Par", GetType(String))
        table.Columns.Add("Etat", GetType(String))

        For Each journee As PieceRepository.JourneeConservee In _journees
            table.Rows.Add(journee.DateActivite.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
                           journee.NombreEcritures, journee.TotalDebit, journee.TotalCredit,
                           If(journee.Equilibree, "oui", "NON"),
                           journee.DateEnregistrement.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture),
                           journee.EnregistrePar,
                           journee.LibelleEtat)
        Next

        Return table
    End Function

    Private Sub FormaterLesJournees()

        dgvJournees.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Entete(dgvJournees, "Journee", "Journée", 80)
        Entete(dgvJournees, "Ecritures", "Écr.", 45)
        Entete(dgvJournees, "Debit", "Débit", 95)
        Entete(dgvJournees, "Credit", "Crédit", 95)
        Entete(dgvJournees, "Equilibre", "Équil.", 45)
        Entete(dgvJournees, "Conservee", "Conservée le", 115)
        Entete(dgvJournees, "Par", "Par", 90)
        Entete(dgvJournees, "Etat", "État", 230)

        For Each nom As String In New String() {"Debit", "Credit", "Ecritures"}
            If Not dgvJournees.Columns.Contains(nom) Then Continue For
            dgvJournees.Columns(nom).DefaultCellStyle.Format = "N0"
            dgvJournees.Columns(nom).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next

        MarquerLesJourneesAnnulees()
    End Sub

    ''' <summary>
    ''' Une journée annulée ne se lit pas comme une journée en vigueur : elle est grise et
    ''' barrée. La couleur seule ne suffirait pas — la colonne État écrit ANNULÉE et le
    ''' motif, et c'est elle qui fait foi.
    ''' </summary>
    Private Sub MarquerLesJourneesAnnulees()

        If _policeBarree Is Nothing Then
            _policeBarree = New Drawing.Font(dgvJournees.Font, Drawing.FontStyle.Strikeout)
        End If

        For rang As Integer = 0 To dgvJournees.Rows.Count - 1

            If rang >= _journees.Count Then Exit For
            If Not _journees(rang).Annulee Then Continue For

            dgvJournees.Rows(rang).DefaultCellStyle.ForeColor = Drawing.SystemColors.GrayText
            dgvJournees.Rows(rang).DefaultCellStyle.Font = _policeBarree
        Next
    End Sub

    Private Shared Sub Entete(grille As DataGridView, colonne As String, libelle As String, largeur As Integer)
        If Not grille.Columns.Contains(colonne) Then Return
        grille.Columns(colonne).HeaderText = libelle
        GrilleWU.LargeurFixe(grille, colonne, largeur)
    End Sub

#End Region

#Region "Présentation d'une journée"

    Private Sub dgvJournees_SelectionChanged(sender As Object, e As EventArgs) Handles dgvJournees.SelectionChanged
        PresenterLaSelection()
    End Sub

    ''' <summary>La journée sélectionnée, ou Nothing.</summary>
    Private Function JourneeSelectionnee() As PieceRepository.JourneeConservee

        If dgvJournees.CurrentRow Is Nothing Then Return Nothing

        Dim rang As Integer = dgvJournees.CurrentRow.Index
        If rang < 0 OrElse rang >= _journees.Count Then Return Nothing

        Return _journees(rang)
    End Function

    Private Sub PresenterLaSelection()

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()

        If journee Is Nothing Then
            AfficherLaPiece(Nothing)
            Return
        End If

        ' Inutile de relire la base à chaque déplacement du curseur dans la même ligne :
        ' SelectionChanged se déclenche plus souvent que la sélection ne change vraiment.
        Dim cle As String = CleDe(journee)
        If String.Equals(_cleAffichee, cle, StringComparison.Ordinal) Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor

        ' Une journée annulée se charge par son identifiant d'annulation : la date ne dirait
        ' pas laquelle, puisqu'une même journée peut avoir été annulée plusieurs fois.
        Dim piece As DataTable = If(journee.Annulee,
                                    PieceRepository.ChargerAnnulee(journee.IdAnnulation, messageErreur),
                                    PieceRepository.Charger(journee.DateActivite, messageErreur))
        Cursor = Cursors.Default

        If messageErreur.Length > 0 Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = PremiereLigne(messageErreur)
            AfficherLaPiece(Nothing)
            Return
        End If

        _cleAffichee = cle
        AfficherLaPiece(piece)
        AnnoncerLaJournee(journee)
    End Sub

    ''' <summary>Ce qui distingue une ligne de la liste : sa date ET son annulation.</summary>
    Private Shared Function CleDe(journee As PieceRepository.JourneeConservee) As String
        Return journee.DateActivite.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) &
               "|" & journee.IdAnnulation.ToString(CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' Titre de la pièce affichée. Une journée annulée le dit ici aussi, et pas seulement
    ''' dans la liste : on consulte une pièce en la regardant, pas en regardant à côté.
    ''' </summary>
    Private Sub AnnoncerLaJournee(journee As PieceRepository.JourneeConservee)

        If Not journee.Annulee Then
            lblPiece.ForeColor = Drawing.SystemColors.ControlText
            lblPiece.Text = $"Écritures de la journée du {journee.DateActivite:dd/MM/yyyy}"
            Return
        End If

        Dim quand As String = If(journee.DateAnnulation.HasValue,
                                 journee.DateAnnulation.Value.ToString("dd/MM/yyyy",
                                                                       CultureInfo.InvariantCulture),
                                 "?")

        lblPiece.ForeColor = Drawing.Color.Firebrick
        lblPiece.Text = $"Journée du {journee.DateActivite:dd/MM/yyyy} — ANNULÉE le {quand}" &
                        If(journee.AnnuleePar.Length > 0, $" par {journee.AnnuleePar}", String.Empty) &
                        $"     {AnnulationWU.LibelleDepuisCode(journee.MotifAnnulation)}"
    End Sub

    Private Sub AfficherLaPiece(piece As DataTable)

        _piece = If(piece, New DataTable())

        If piece Is Nothing Then
            _cleAffichee = String.Empty
            lblPiece.ForeColor = Drawing.SystemColors.ControlText
            lblPiece.Text = "Écritures de la journée"
        End If

        dgvPiece.DataSource = _piece
        FormaterLaPiece()
        AfficherLesTotaux()

        Dim disponible As Boolean = _piece.Rows.Count > 0
        btnPiece.Enabled = disponible
        btnCoreBanking.Enabled = disponible

        ReglerLeBoutonAnnuler()
    End Sub

    ''' <summary>
    ''' Le bouton d'annulation ne s'offre que là où il a un sens : sur une journée en
    ''' vigueur, et pour un utilisateur qui a la fonction d'inputer.
    '''
    ''' Une journée DÉJÀ annulée ne se réannule pas : il n'y a plus rien dans les tables
    ''' vivantes, et l'autorisation refuserait de créer une archive vide. Mieux vaut un
    ''' bouton éteint qu'un refus après coup.
    ''' </summary>
    Private Sub ReglerLeBoutonAnnuler()

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()

        Dim annulee As Boolean = journee IsNot Nothing AndAlso journee.Annulee

        btnAnnuler.Enabled = journee IsNot Nothing AndAlso
                             Not annulee AndAlso
                             SessionWU.PeutSaisirLesPointsDeVente

        ' On ne reconstruit pas le fichier core banking d'une journée annulée : ce fichier
        ' passe des écritures, et repasser celles d'une journée retirée serait exactement
        ' l'inverse de ce qu'on a voulu. La pièce, elle, reste consultable et exportable :
        ' un justificatif se relit, il n'engage rien.
        If annulee Then btnCoreBanking.Enabled = False

        ' Le bordereau reste accessible sur une journée annulée : comprendre pourquoi elle a
        ' été retirée suppose de pouvoir relire comment elle avait été traitée.
        btnBordereau.Enabled = journee IsNot Nothing
    End Sub

    Private Sub FormaterLaPiece()

        For Each nomColonne As String In New String() {"Debit", "Credit"}
            If Not dgvPiece.Columns.Contains(nomColonne) Then Continue For
            dgvPiece.Columns(nomColonne).DefaultCellStyle.Format = "#,##0;-#,##0;"
            dgvPiece.Columns(nomColonne).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GrilleWU.LargeurFixe(dgvPiece, nomColonne, 110)
        Next

        GrilleWU.LargeurFixe(dgvPiece, "Compte", 120)
        If dgvPiece.Columns.Contains("Libelle") Then
            dgvPiece.Columns("Libelle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvPiece.Columns("Libelle").HeaderText = "Libellé"
        End If

        ' Le code agence n'appartient pas à la pièce : il n'alimente que la colonne ACBRN du
        ' fichier core banking. L'afficher ferait croire à une colonne comptable de plus.
        If dgvPiece.Columns.Contains("CodeAgence") Then dgvPiece.Columns("CodeAgence").Visible = False
    End Sub

    Private Sub AfficherLesTotaux()

        If _piece.Rows.Count = 0 Then
            lblTotaux.Text = String.Empty
            Return
        End If

        Dim totalDebit As Long = _piece.AsEnumerable().Sum(Function(r) Convert.ToInt64(r("Debit")))
        Dim totalCredit As Long = _piece.AsEnumerable().Sum(Function(r) Convert.ToInt64(r("Credit")))

        lblTotaux.Text = $"{_piece.Rows.Count} écritures     TOTAL DÉBIT : {totalDebit:N0} FCFA     " &
                         $"TOTAL CRÉDIT : {totalCredit:N0} FCFA"

        ' Une pièce conservée s'équilibre toujours — le contrôle précède l'historisation. Le
        ' constater vaut mieux que le croire : c'est l'objet même d'une consultation d'archive.
        lblTotaux.ForeColor = If(totalDebit = totalCredit, Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
    End Sub

#End Region

#Region "Ouverture de la pièce et du fichier core banking"

    ''' <summary>
    ''' Présente la pièce dans l'écran habituel, d'où elle s'exporte au formulaire de la banque.
    '''
    ''' La liste des points de vente n'est pas transmise : elle n'est pas conservée. Le classeur
    ''' exporté porte donc la pièce globale, sans les onglets individuels. Mieux vaut une pièce
    ''' fidèle sans ses détails qu'un détail reconstitué au paramétrage d'aujourd'hui.
    ''' </summary>
    Private Sub btnPiece_Click(sender As Object, e As EventArgs) Handles btnPiece.Click

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()
        If journee Is Nothing OrElse _piece.Rows.Count = 0 Then Return

        Dim titre As String = If(journee.Annulee,
                                 $"Pièce ANNULÉE — journée du {journee.DateActivite:dd/MM/yyyy}",
                                 $"Pièce comptable conservée — journée du {journee.DateActivite:dd/MM/yyyy}")

        Dim sousTitre As String =
            $"Conservée le {journee.DateEnregistrement:dd/MM/yyyy à HH:mm}" &
            If(journee.EnregistrePar.Length > 0, $" par {journee.EnregistrePar}", String.Empty) &
            "     Telle qu'elle a été produite : aucun recalcul."

        If journee.Annulee Then
            sousTitre &= Environment.NewLine &
                         $"ANNULÉE — {AnnulationWU.LibelleDepuisCode(journee.MotifAnnulation)}" &
                         If(journee.AnnuleePar.Length > 0, $" — par {journee.AnnuleePar}", String.Empty) &
                         "     Cette journée ne figure plus dans les rapports d'activité."
        End If

        Using apercu As New FrmPieceComptable(_piece, titre, sousTitre)

            apercu.DateActivite = journee.DateActivite
            apercu.NomFichierPropose = PieceComptableService.NomDeFichier(journee.DateActivite)

            apercu.ShowDialog(Me)

            If apercu.PieceExportee Then
                lblStatut.ForeColor = Drawing.SystemColors.GrayText
                lblStatut.Text = "Pièce exportée : " & IO.Path.GetFileName(apercu.CheminExporte)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Reconstruit le fichier destiné au core banking à partir de la pièce conservée.
    '''
    ''' Le fichier n'est pas stocké, et n'a pas à l'être : il DÉRIVE entièrement de la pièce,
    ''' par les mêmes règles. Le conserver serait garder deux fois la même chose, avec le
    ''' risque que les deux copies divergent.
    '''
    ''' La date de valeur, elle, se redemande : c'est le jour où les écritures sont réellement
    ''' passées, donc aujourd'hui — et non la date qu'on avait retenue à l'époque. Un fichier
    ''' rejoué aujourd'hui porte la date d'aujourd'hui, comme le montrera le relevé de compte.
    ''' </summary>
    Private Sub btnCoreBanking_Click(sender As Object, e As EventArgs) Handles btnCoreBanking.Click

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()
        If journee Is Nothing OrElse _piece.Rows.Count = 0 Then Return

        ' Même règle que sur l'écran de la compense, et écrite au même endroit : un fichier
        ' rejoué depuis l'archive engage les comptes autant que le premier.
        If Not VisaWU.AutoriserLeCoreBanking(Me, journee.DateActivite) Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Fichier core banking abandonné : journée non visée."
            Return
        End If

        Dim dateValeur As Date = Date.Today
        If Not ConfirmerLaDateDeValeur(dateValeur) Then
            lblStatut.Text = "Fichier core banking abandonné : date de valeur à vérifier."
            Return
        End If

        Dim messageErreur As String = String.Empty
        Dim fichier As DataTable = CoreBankingService.Construire(_piece, journee.DateActivite,
                                                                 dateValeur, messageErreur)

        If fichier Is Nothing Then
            MessageBox.Show(Me, messageErreur, "Fichier non produit",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Fichier core banking non produit."
            Return
        End If

        Dim numeroLot As String = CoreBankingService.NumeroDeLot(journee.DateActivite)

        Using apercu As New FrmFichierCoreBanking(fichier, journee.DateActivite, dateValeur, numeroLot)

            apercu.ShowDialog(Me)

            If apercu.FichierProduit Then
                lblStatut.ForeColor = Drawing.SystemColors.GrayText
                lblStatut.Text = "Fichier produit : " & IO.Path.GetFileName(apercu.CheminProduit)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Fait confirmer une date de valeur chômée ou incertaine. Même contrôle que l'écran de
    ''' compense : un fichier rejoué depuis l'archive engage les comptes autant que le premier.
    ''' </summary>
    Private Function ConfirmerLaDateDeValeur(dateValeur As Date) As Boolean

        Dim alerte As String

        If Not CalendrierWU.EstJourOuvre(dateValeur) Then
            alerte = $"Le {dateValeur:dd/MM/yyyy} est un jour chômé pour la banque." & Environment.NewLine &
                     "Le core banking peut rejeter le fichier, ou reporter d'office les écritures."
        Else
            alerte = CalendrierWU.Avertissement(dateValeur.Year)
            If alerte.Length = 0 Then Return True
        End If

        Return MessageBox.Show(Me,
            alerte & Environment.NewLine & Environment.NewLine &
            $"Date de valeur retenue : {dateValeur:dd/MM/yyyy}." & Environment.NewLine &
            "Produire le fichier tout de même ?",
            "Date de valeur à vérifier", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

#Region "Annulation d'une comptabilisation"

    ''' <summary>
    ''' Dépose une demande d'annulation sur la journée sélectionnée.
    '''
    ''' CE BOUTON N'ANNULE RIEN. Il ouvre l'écran de demande, qui dépose dans la file du
    ''' double regard. La journée reste comptabilisée et continue de figurer dans les
    ''' rapports tant qu'un authorizer — qui ne peut pas être le demandeur — ne l'a pas
    ''' autorisée. C'est pourquoi la liste n'est pas rechargée avec un air de réussite :
    ''' rien n'a changé dans la base, et le faire croire serait le pire des services.
    ''' </summary>
    Private Sub btnAnnuler_Click(sender As Object, e As EventArgs) Handles btnAnnuler.Click

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()
        If journee Is Nothing OrElse journee.Annulee Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim contenu As AnnulationRepository.ContenuJournee =
            AnnulationRepository.Decrire(journee.DateActivite, messageErreur)
        Cursor = Cursors.Default

        If messageErreur.Length > 0 Then
            FrmDiagnostic.Afficher(Me, "Annulation impossible", messageErreur)
            Return
        End If

        ' Une journée sans rien à retirer ne se demande pas en annulation : l'autorisation
        ' la refuserait, et l'agent aurait attendu pour rien.
        If contenu.EstVide Then
            MessageBox.Show(Me,
                $"La journée du {journee.DateActivite:dd/MM/yyyy} n'a plus rien de comptabilisé." &
                Environment.NewLine & Environment.NewLine &
                "Il n'y a donc rien à retirer. Actualisez la liste : elle a peut-être été " &
                "annulée entre-temps par quelqu'un d'autre.",
                "Rien à annuler", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using demande As New FrmAnnulerComptabilisation(journee.DateActivite, contenu)

            demande.ShowDialog(Me)

            If Not demande.DemandeDeposee Then Return

            lblStatut.ForeColor = Drawing.SystemColors.GrayText
            lblStatut.Text = $"Annulation demandée pour le {journee.DateActivite:dd/MM/yyyy} — " &
                             "en attente d'autorisation."
        End Using
    End Sub

#End Region

#Region "Fermeture"

    ''' <summary>
    ''' Le bordereau de la journée sélectionnée : ce qui se signe, là où la pièce est ce qui
    ''' se comptabilise.
    '''
    ''' Il s'ouvre aussi de lui-même en fin de comptabilisation. Ce bouton existe pour les
    ''' deux autres usages : le réimprimer, et le VISER — c'est ici que le chef de service
    ''' vient relire une journée.
    '''
    ''' Disponible même sur une journée annulée : comprendre POURQUOI une journée a été
    ''' retirée suppose de pouvoir relire comment elle avait été traitée.
    ''' </summary>
    Private Sub btnBordereau_Click(sender As Object, e As EventArgs) Handles btnBordereau.Click

        Dim journee As PieceRepository.JourneeConservee = JourneeSelectionnee()
        If journee Is Nothing Then Return

        Using bordereau As New FrmBordereauJournee(journee.DateActivite)
            bordereau.ShowDialog(Me)
        End Using

        ' Le visa a pu changer pendant ce temps : la liste se relit.
        Charger()
    End Sub

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
