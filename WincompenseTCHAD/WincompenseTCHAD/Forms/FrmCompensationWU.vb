Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO
Imports System.Linq

''' <summary>
''' Formulaire principal de compensation Western Union J+1 (Tchad).
''' Ce formulaire ne contient QUE de l'orchestration d'événements : toute la logique
''' métier (lecture, agrégation, accès SQL, calculs, pièce comptable) est déléguée
''' aux services WUReportService / WURepository / WUCalculationService / PieceComptableService.
''' </summary>
Public Class FrmCompensationWU

    ''' <summary>Constructeur requis par le Concepteur Windows Forms : initialise tous les contrôles.</summary>
    Public Sub New()
        InitializeComponent()
    End Sub

#Region "État interne du formulaire"

    Private _cheminActivite As String = String.Empty
    Private _cheminReglement As String = String.Empty

    ''' <summary>
    ''' Ce que le NOM de chaque fichier chargé a permis d'établir (type de rapport, période).
    ''' Conservé pour contrôler la concordance des deux rapports dès la seconde sélection,
    ''' sans avoir à relire les fichiers. Nothing tant qu'aucun fichier n'est sélectionné.
    ''' </summary>
    Private _infosActivite As InfosFichierRapport
    Private _infosReglement As InfosFichierRapport

    ''' <summary>
    ''' Date d'activité de la journée calculée, lue dans le rapport lui-même. Mémorisée pour
    ''' l'historisation : c'est elle qui classe la journée dans le temps, et non la date du jour.
    ''' </summary>
    Private _dateActivite As Date?

    ''' <summary>
    ''' Détail des transactions de la journée, identifiées par leur MTCN. Extrait au moment du
    ''' calcul — le rapport n'est plus en mémoire au moment de générer la pièce — et historisé
    ''' avec l'agrégat pour permettre de retrouver une opération précise.
    ''' </summary>
    Private _transactions As New List(Of TransactionWU)
    Private _listeCalculs As List(Of CalculWU)
    Private _dtPieceGeneree As DataTable

    ''' <summary>
    ''' Message technique du dernier échec de connexion à SQL Server, vide si la connexion a
    ''' abouti. Mémorisé pour être présenté à l'utilisateur : sans ce détail, un « base SQL
    ''' inaccessible » ne permet pas de distinguer un service arrêté d'une base absente.
    ''' </summary>
    Private _messageErreurConnexionSql As String = String.Empty

#End Region

#Region "Démarrage"

    ''' <summary>
    ''' Au démarrage, les comptes comptables de la pièce sont lus dans la table SystemeWU.
    ''' Si la base est inaccessible, les comptes par défaut de l'application prennent le relais
    ''' et l'utilisateur en est averti dans la barre d'état : la comptabilisation reste possible,
    ''' mais avec un paramétrage qui n'est peut-être plus celui de la Direction Comptable.
    ''' </summary>
    Private Sub FrmCompensationWU_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutTraiterLaCompense Then
            MessageBox.Show("Le traitement de la compense est réservé aux agents de la compense et aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If


        Dim messageErreur As String = String.Empty
        ComptesSystemeWU.Actuels = WURepository.ChargerComptesSysteme(messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then
            tsslStatut.Text = "Comptes comptables PAR DÉFAUT (table SystemeWU non lue)."
        End If
    End Sub

    ''' <summary>
    ''' Enregistre la journée dans l'historique (tables T_HistoriqueWU et T_HistoriqueMTCN),
    ''' source des rapports d'activité sur une période.
    '''
    ''' L'échec de l'historisation n'annule JAMAIS la pièce comptable : celle-ci est déjà
    ''' générée et équilibrée, elle reste la priorité. L'utilisateur est simplement averti que
    ''' la journée ne figurera pas dans les rapports tant qu'elle n'aura pas été regénérée.
    ''' </summary>
    Private Sub HistoriserLaJournee()

        If Not _dateActivite.HasValue Then
            ' Sans date d'activité exploitable, la journée ne peut pas être classée dans le temps.
            tsslStatut.Text &= "  |  Journée non historisée : date d'activité indéterminée."
            Return
        End If

        Dim nombreEnregistrees As Integer = 0
        Dim nombreTransactions As Integer = 0
        Dim messageErreur As String = String.Empty

        If HistoriqueRepository.EnregistrerJournee(_dateActivite.Value, _listeCalculs, _transactions,
                                                   nombreEnregistrees, nombreTransactions, messageErreur) Then
            tsslStatut.Text &= $"  |  Journée du {_dateActivite.Value:dd/MM/yyyy} historisée " &
                               $"({nombreEnregistrees} point(s) de vente, {nombreTransactions} transaction(s))."
            Return
        End If

        MessageBox.Show(
            messageErreur & Environment.NewLine & Environment.NewLine &
            "La pièce comptable, elle, est bien générée et reste utilisable." & Environment.NewLine &
            "Seul le rapport d'activité ignorera cette journée tant qu'elle n'aura pas été " &
            "regénérée une fois le problème corrigé.",
            "Journée non historisée", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        tsslStatut.Text &= "  |  Journée NON historisée."
    End Sub

#End Region

#Region "Chargement des fichiers"

    ''' <summary>
    ''' Sélection du rapport d'activité. Le fichier retenu peut être l'archive ZIP livrée par
    ''' Western Union (qui porte le nom du rapport qu'elle contient) comme le fichier texte déjà
    ''' décompressé : la décompression éventuelle est faite à la lecture, par WUReportService.
    ''' </summary>
    Private Sub btnActivite_Click(sender As Object, e As EventArgs) Handles btnActivite.Click

        If ofdActivite.ShowDialog() <> DialogResult.OK Then Return

        Dim infos As InfosFichierRapport = WUFichierService.AnalyserFichier(ofdActivite.FileName)

        ' Contrôles de sécurité sur le nom : type attendu, puis concordance avec l'autre rapport.
        If Not ValiderSelection(infos, TypeRapportWU.Activite, _infosReglement) Then Return

        _infosActivite = infos
        _cheminActivite = infos.Chemin
        lblActivite.Text = infos.NomFichier
        lblActivite.Tag = _cheminActivite

        ReinitialiserResultats()
        MettreAJourEtatBoutons()
        AfficherPeriodeChargee()
    End Sub

    ''' <summary>
    ''' Sélection du rapport de règlement. Comme pour l'activité, l'archive ZIP est acceptée
    ''' directement et décompressée en mémoire au moment de la lecture.
    ''' </summary>
    Private Sub btnReglement_Click(sender As Object, e As EventArgs) Handles btnReglement.Click

        If ofdReglement.ShowDialog() <> DialogResult.OK Then Return

        Dim infos As InfosFichierRapport = WUFichierService.AnalyserFichier(ofdReglement.FileName)

        If Not ValiderSelection(infos, TypeRapportWU.Reglement, _infosActivite) Then Return

        _infosReglement = infos
        _cheminReglement = infos.Chemin
        lblReglement.Text = infos.NomFichier
        lblReglement.Tag = _cheminReglement

        ReinitialiserResultats()
        MettreAJourEtatBoutons()
        AfficherPeriodeChargee()
    End Sub

    ''' <summary>
    ''' Contrôles de sécurité appliqués dès la sélection d'un fichier, sur son seul nom :
    '''
    '''   1. le fichier n'est pas le rapport de l'autre type (activité chargée à la place du
    '''      règlement, ou l'inverse) ;
    '''   2. sa période est celle de l'autre rapport déjà chargé.
    '''
    ''' Un rapport du mauvais TYPE est purement et simplement refusé. Une PÉRIODE différente
    ''' de celle de l'autre rapport déjà chargé ouvre en revanche un choix explicite : soit
    ''' abandonner le fichier, soit changer de journée de traitement, l'autre rapport étant
    ''' alors retiré. Dans les deux cas, deux rapports de périodes différentes ne peuvent
    ''' jamais être chargés ensemble.
    '''
    ''' Un nom de fichier non standard, dont on ne peut rien déduire, ne déclenche jamais de
    ''' refus ici : les contrôles sur le CONTENU des rapports, eux systématiques, prendront
    ''' le relais au moment du calcul.
    ''' </summary>
    ''' <param name="infos">Fichier que l'utilisateur vient de choisir.</param>
    ''' <param name="typeAttendu">Type de rapport attendu par le bouton utilisé.</param>
    ''' <param name="infosAutreRapport">Autre rapport déjà chargé, ou Nothing.</param>
    ''' <returns>True si le fichier peut être retenu.</returns>
    Private Function ValiderSelection(infos As InfosFichierRapport,
                                      typeAttendu As TypeRapportWU,
                                      infosAutreRapport As InfosFichierRapport) As Boolean

        ' --- 1. Type de rapport ---
        Dim messageType As String = String.Empty
        If Not WUFichierService.VerifierTypeAttendu(infos, typeAttendu, messageType) Then
            MessageBox.Show(messageType, "Rapport incorrect", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tsslStatut.Text = "Fichier refusé : ce n'est pas le rapport " &
                              WUFichierService.LibelleType(typeAttendu) & "."
            Return False
        End If

        ' --- 2. Concordance des périodes ---
        Dim infosActivite As InfosFichierRapport = If(typeAttendu = TypeRapportWU.Activite, infos, infosAutreRapport)
        Dim infosReglement As InfosFichierRapport = If(typeAttendu = TypeRapportWU.Reglement, infos, infosAutreRapport)

        Dim messagePeriode As String = String.Empty
        If Not WUFichierService.VerifierMemePeriode(infosActivite, infosReglement, messagePeriode) Then

            ' Refuser sèchement le fichier enfermerait l'utilisateur : une fois les deux rapports
            ' d'une journée chargés, aucun fichier d'une autre journée ne pourrait plus jamais être
            ' sélectionné, ni d'un côté ni de l'autre. On lui laisse donc le choix explicite de
            ' changer de journée — auquel cas l'autre rapport, devenu hors période, est retiré.
            ' Dans les deux cas, deux rapports de périodes différentes ne peuvent jamais coexister.
            Dim reponse As DialogResult = MessageBox.Show(
                messagePeriode & Environment.NewLine & Environment.NewLine &
                "Souhaitez-vous changer de journée de traitement ?" & Environment.NewLine & Environment.NewLine &
                "   Oui  : ce fichier est conservé, l'autre rapport est retiré (à recharger)." & Environment.NewLine &
                "   Non  : ce fichier est abandonné, la sélection précédente est conservée.",
                "Périodes différentes", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If reponse <> DialogResult.Yes Then
                tsslStatut.Text = "Fichier abandonné : période différente de celle de l'autre rapport."
                Return False
            End If

            RetirerAutreRapport(typeAttendu)
        End If

        Return True
    End Function

    ''' <summary>
    ''' Retire de la sélection le rapport qui n'est PAS du type indiqué, lorsque l'utilisateur
    ''' change de journée de traitement. Garantit qu'à aucun moment deux rapports de périodes
    ''' différentes ne sont chargés ensemble.
    ''' </summary>
    ''' <param name="typeConserve">Type du rapport que l'utilisateur vient de sélectionner.</param>
    Private Sub RetirerAutreRapport(typeConserve As TypeRapportWU)

        If typeConserve = TypeRapportWU.Activite Then
            _infosReglement = Nothing
            _cheminReglement = String.Empty
            lblReglement.Text = "(aucun fichier sélectionné)"
            lblReglement.Tag = Nothing
        Else
            _infosActivite = Nothing
            _cheminActivite = String.Empty
            lblActivite.Text = "(aucun fichier sélectionné)"
            lblActivite.Tag = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Affiche dans la barre d'état la période des rapports chargés, ou la raison pour laquelle
    ''' elle n'a pas pu être déterminée à partir des noms de fichiers.
    ''' </summary>
    Private Sub AfficherPeriodeChargee()

        Dim message As String = String.Empty
        WUFichierService.VerifierMemePeriode(_infosActivite, _infosReglement, message)

        If Not String.IsNullOrEmpty(message) Then
            tsslStatut.Text = message
        ElseIf _infosActivite IsNot Nothing AndAlso _infosActivite.Periode IsNot Nothing Then
            tsslStatut.Text = $"Rapport d'activité — période : {_infosActivite.Periode}."
        ElseIf _infosReglement IsNot Nothing AndAlso _infosReglement.Periode IsNot Nothing Then
            tsslStatut.Text = $"Rapport de règlement — période : {_infosReglement.Periode}."
        End If
    End Sub

    ''' <summary>Active btnAfficher uniquement lorsque les deux rapports sont sélectionnés.</summary>
    Private Sub MettreAJourEtatBoutons()
        btnAfficher.Enabled = Not String.IsNullOrWhiteSpace(_cheminActivite) AndAlso
                               Not String.IsNullOrWhiteSpace(_cheminReglement)
        btnGenererPiece.Enabled = False
        btnPieceAccount.Enabled = False
    End Sub

    ''' <summary>Réinitialise les résultats de calcul lorsqu'un nouveau fichier est sélectionné.</summary>
    Private Sub ReinitialiserResultats()
        _listeCalculs = Nothing
        _dtPieceGeneree = Nothing
        _dateActivite = Nothing
        _transactions = New List(Of TransactionWU)
        btnPieceAccount.Enabled = False
        dgvControle.DataSource = Nothing
        progressBarTraitement.Value = 0
        tsslLignesActivite.Text = "Lignes activité : 0"
        tsslLignesReglement.Text = "Lignes règlement : 0"
        tsslNombreAccounts.Text = "Accounts : 0"
        tsslStatut.Text = "Prêt."
    End Sub

#End Region

#Region "Affichage / Calcul (btnAfficher)"

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click

        Cursor = Cursors.WaitCursor
        btnAfficher.Enabled = False
        btnGenererPiece.Enabled = False
        progressBarTraitement.Value = 0

        Try
            tsslStatut.Text = "Lecture des rapports en cours..."
            Application.DoEvents()

            Dim dtActivite As DataTable = WUReportService.LireRapportWU(_cheminActivite)
            Dim dtReglement As DataTable = WUReportService.LireRapportWU(_cheminReglement)
            progressBarTraitement.Value = 15

            ' Contrôle de sécurité : chaque fichier est-il bien le rapport qu'il prétend être ?
            ' Fondé sur les colonnes réellement présentes, ce contrôle démasque un fichier
            ' renommé, que le contrôle sur le nom fait au chargement ne peut pas détecter.
            WUFichierService.VerifierTypeRapport(dtActivite, TypeRapportWU.Activite, Path.GetFileName(_cheminActivite))
            WUFichierService.VerifierTypeRapport(dtReglement, TypeRapportWU.Reglement, Path.GetFileName(_cheminReglement))

            ' Vérification structurelle avant tout calcul (section 15).
            WUReportService.VerifierColonnesRapport(dtActivite, ConstantesWU.ColonnesRapportActivite, "activité")
            WUReportService.VerifierColonnesRapport(dtReglement, ConstantesWU.ColonnesRapportReglement, "règlement")
            progressBarTraitement.Value = 30

            ' Date de la journée traitée, retenue pour l'historisation.
            _dateActivite = WUReportService.ObtenirDateActivite(dtActivite)

            ' Détail des transactions, extrait tant que le rapport est en mémoire.
            _transactions = If(_dateActivite.HasValue,
                               WUReportService.ExtraireTransactions(dtActivite, _dateActivite.Value),
                               New List(Of TransactionWU))

            ' Validation de la cohérence des dates entre les deux rapports (section 16).
            Dim messageDate As String = String.Empty
            If Not WUReportService.ValiderCoherenceDates(dtActivite, dtReglement, messageDate) Then
                MessageBox.Show(messageDate, "Incohérence de date", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                tsslStatut.Text = "Traitement bloqué : incohérence de date entre les deux rapports."
                Return
            End If
            progressBarTraitement.Value = 45

            ' Agrégation des deux rapports par Account.
            Dim aggActivite As Dictionary(Of String, ActiviteAgregat) = WUReportService.CalculerActivite(dtActivite)
            Dim aggReglement As Dictionary(Of String, ReglementAgregat) = WUReportService.CalculerReglement(dtReglement)
            progressBarTraitement.Value = 60

            tsslStatut.Text = "Récupération des paramètres SQL Server..."
            Application.DoEvents()

            ' Récupération des paramètres SQL Server + application des formules métier.
            _listeCalculs = ConstruireCalculsParAccount(aggActivite, aggReglement)
            progressBarTraitement.Value = 85

            ' Construction et affichage de la grille de contrôle.
            Dim dtControle As DataTable = PieceComptableService.CreerTableControle(_listeCalculs)
            dgvControle.DataSource = dtControle
            FormaterColonnesNumeriques()
            MasquerColonnesTechniques()
            MettreEnEvidenceAnomalies()

            tsslLignesActivite.Text = $"Lignes activité : {dtActivite.Rows.Count}"
            tsslLignesReglement.Text = $"Lignes règlement : {dtReglement.Rows.Count}"
            tsslNombreAccounts.Text = $"Accounts : {_listeCalculs.Count}"
            tsslStatut.Text = $"Calcul terminé. {messageDate}"

            btnGenererPiece.Enabled = _listeCalculs.Count > 0
            btnPieceAccount.Enabled = _listeCalculs.Count > 0
            progressBarTraitement.Value = 100

            ' La base est facultative pour calculer, mais indispensable pour identifier les points
            ' de vente : sans elle, tous les Accounts restent INCONNU et la pièce comptable ne peut
            ' pas utiliser leurs comptes de compensation et de commission. L'utilisateur doit donc
            ' en être averti explicitement, avec le détail technique permettant de diagnostiquer.
            If Not String.IsNullOrEmpty(_messageErreurConnexionSql) Then
                tsslStatut.Text = "Calcul effectué SANS les paramètres SQL Server (base inaccessible)."
                MessageBox.Show(
                    _messageErreurConnexionSql & Environment.NewLine & Environment.NewLine &
                    "Le calcul a tout de même été effectué, mais aucun Account n'a pu être identifié :" &
                    Environment.NewLine &
                    "ils apparaissent tous en INCONNU et la pièce comptable utilisera le compte courant WU" & Environment.NewLine &
                    "au lieu des comptes de compensation et de commission des sous-agents." & Environment.NewLine & Environment.NewLine &
                    "À vérifier :" & Environment.NewLine &
                    "  1. le service SQL Server (SQLEXPRESS) est démarré ;" & Environment.NewLine &
                    "  2. le nom de l'instance est correct dans le fichier App.config ;" & Environment.NewLine &
                    "  3. la base GWC_WINCOMPENSE_ETD existe (scripts du dossier Scripts\ exécutés) ;" & Environment.NewLine &
                    "  4. votre compte Windows a accès à cette base.",
                    "Base SQL Server inaccessible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As RapportInvalideException
            MessageBox.Show(ex.Message, "Anomalie de rapport", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tsslStatut.Text = "Erreur : " & ex.Message

        Catch ex As Exception
            MessageBox.Show("Erreur inattendue lors du calcul : " & ex.Message, "Erreur",
                             MessageBoxButtons.OK, MessageBoxIcon.Error)
            tsslStatut.Text = "Erreur inattendue lors du calcul."

        Finally
            Cursor = Cursors.Default
            btnAfficher.Enabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Construit la liste des CalculWU pour tous les Accounts rencontrés dans l'un ou l'autre
    ''' rapport, récupère leurs paramètres SQL Server, applique les formules et répartit les
    ''' commissions. Si la connexion SQL Server est totalement indisponible, le traitement
    ''' continue sans plantage : chaque Account est marqué en anomalie (ErreurSQL = True).
    ''' </summary>
    Private Function ConstruireCalculsParAccount(comptesActivite As Dictionary(Of String, ActiviteAgregat),
                                                  comptesReglement As Dictionary(Of String, ReglementAgregat)) As List(Of CalculWU)

        Dim tousLesAccounts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each cle As String In comptesActivite.Keys
            tousLesAccounts.Add(cle)
        Next
        For Each cle As String In comptesReglement.Keys
            tousLesAccounts.Add(cle)
        Next

        Dim liste As New List(Of CalculWU)

        Dim connexion As SqlConnection = Nothing
        Dim connexionOuverte As Boolean = False
        Dim messageErreurConnexion As String = String.Empty
        _messageErreurConnexionSql = String.Empty

        Try
            connexion = WURepository.CreerConnexion()
            connexion.Open()
            connexionOuverte = True
        Catch ex As Exception
            ' Base SQL inaccessible (section 15) : on ne bloque pas le traitement, on continue en mode dégradé.
            messageErreurConnexion = $"Connexion à SQL Server ({WURepository.ObtenirChaineConnexion()}) impossible : {ex.Message}"
            _messageErreurConnexionSql = messageErreurConnexion
        End Try

        Try
            For Each account As String In tousLesAccounts

                Dim calc As CalculWU

                If connexionOuverte Then
                    calc = WURepository.ChargerParametresAccount(account, connexion)
                Else
                    calc = New CalculWU() With {
                        .Account = account,
                        .TypePdv = "INCONNU",
                        .Designation = "BASE SQL INACCESSIBLE",
                        .ErreurSQL = True,
                        .MessageErreurSQL = messageErreurConnexion
                    }
                End If

                Dim agregatActivite As ActiviteAgregat = Nothing
                If comptesActivite.TryGetValue(account, agregatActivite) Then
                    calc.PrincipalEnvoi = agregatActivite.PrincipalEnvoi
                    calc.PrincipalPaye = agregatActivite.PrincipalPaye
                    calc.ChargeEnvoi = agregatActivite.ChargeEnvoi
                    calc.Taxes = agregatActivite.Taxes
                    calc.NombreEnvois = agregatActivite.NombreEnvois
                    calc.NombrePaiements = agregatActivite.NombrePaiements
                    calc.NombreAnnulations = agregatActivite.NombreAnnulations
                Else
                    calc.DonneesManquantes = True ' Account présent en règlement mais absent de l'activité.
                End If

                Dim agregatReglement As ReglementAgregat = Nothing
                If comptesReglement.TryGetValue(account, agregatReglement) Then
                    calc.CommissionPaiement = agregatReglement.CommissionPaiement
                End If

                WUCalculationService.AppliquerFormules(calc)
                WUCalculationService.RepartirCommissions(calc)
                WUCalculationService.CalculerTotaux(calc)

                liste.Add(calc)
            Next

        Finally
            If connexionOuverte Then
                connexion.Close()
                connexion.Dispose()
            End If
        End Try

        Return liste.OrderBy(Function(c) c.Account, StringComparer.OrdinalIgnoreCase).ToList()
    End Function

#End Region

#Region "Mise en forme de la grille de contrôle"

    Private Sub FormaterColonnesNumeriques()
        Dim colonnesMontant As String() = {
            "PrincipalEnvoi", "PrincipalPaye", "ChargeEnvoi", "Taxes",
            "CommissionTransfert", "CommissionPaiement", "CommissionEnvoi",
            "TVA", "TTAEnvoi", "TTAReception", "TaxeEnvoi",
            "CommissionTransfertBanque", "CommissionPaiementBanque", "CommissionEnvoiBanque",
            "CommissionTransfertSA", "CommissionPaiementSA", "CommissionEnvoiSA",
            "TotalDebit", "TotalCredit", "Solde"
        }

        For Each nomColonne As String In colonnesMontant
            If dgvControle.Columns.Contains(nomColonne) Then
                dgvControle.Columns(nomColonne).DefaultCellStyle.Format = "N2"
                dgvControle.Columns(nomColonne).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
        Next

        If dgvControle.Columns.Contains("TauxSA") Then
            dgvControle.Columns("TauxSA").DefaultCellStyle.Format = "P1"
        End If

        ' L'écart d'arrondi est un montant entier en FCFA (0 ou ±1 en fonctionnement normal).
        If dgvControle.Columns.Contains("EcartArrondi") Then
            dgvControle.Columns("EcartArrondi").DefaultCellStyle.Format = "N0"
            dgvControle.Columns("EcartArrondi").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    ''' <summary>Masque les colonnes techniques utilisées uniquement pour la mise en évidence des anomalies.</summary>
    Private Sub MasquerColonnesTechniques()
        For Each nomColonne As String In {"ErreurSQL", "DonneesManquantes"}
            If dgvControle.Columns.Contains(nomColonne) Then
                dgvControle.Columns(nomColonne).Visible = False
            End If
        Next
    End Sub

    ''' <summary>
    ''' Met en évidence visuellement (section 12) : Type INCONNU, Solde ≠ 0, erreur SQL, données manquantes.
    ''' L'ordre ci-dessous reflète la priorité de gravité (erreur SQL prioritaire sur le reste).
    ''' </summary>
    Private Sub MettreEnEvidenceAnomalies()
        For Each ligne As DataGridViewRow In dgvControle.Rows

            Dim erreurSql As Boolean = Convert.ToBoolean(ligne.Cells("ErreurSQL").Value)
            Dim donneesManquantes As Boolean = Convert.ToBoolean(ligne.Cells("DonneesManquantes").Value)
            Dim type As String = Convert.ToString(ligne.Cells("Type").Value)
            Dim solde As Decimal = Convert.ToDecimal(ligne.Cells("Solde").Value)
            Dim ecartArrondi As Long = Convert.ToInt64(ligne.Cells("EcartArrondi").Value)

            If erreurSql Then
                ligne.DefaultCellStyle.BackColor = Color.MistyRose
                ligne.DefaultCellStyle.ForeColor = Color.DarkRed
            ElseIf String.Equals(type, "INCONNU", StringComparison.OrdinalIgnoreCase) Then
                ligne.DefaultCellStyle.BackColor = Color.LightYellow
            ElseIf donneesManquantes Then
                ligne.DefaultCellStyle.BackColor = Color.Gainsboro
            ElseIf Math.Abs(ecartArrondi) > ConstantesWU.SEUIL_ECART_LIGNE_ANORMAL Then
                ' Écart trop important pour un simple arrondi : paramétrage probablement incomplet.
                ligne.DefaultCellStyle.BackColor = Color.LightSalmon
            ElseIf solde <> 0D Then
                ligne.DefaultCellStyle.BackColor = Color.Khaki
            Else
                ligne.DefaultCellStyle.BackColor = dgvControle.DefaultCellStyle.BackColor
            End If
        Next
    End Sub

#End Region

#Region "Génération de la pièce comptable (btnGenererPiece)"


    ''' <summary>
    ''' Produit le fichier à treize colonnes chargé dans le core banking.
    '''
    ''' Il dérive de la pièce déjà générée, et non d'un nouveau calcul : ce qui est chargé doit
    ''' être exactement ce que le comptable a vu et validé à l'écran.
    ''' </summary>
    Private Sub btnCoreBanking_Click(sender As Object, e As EventArgs) Handles btnCoreBanking.Click

        If _dtPieceGeneree Is Nothing OrElse _dtPieceGeneree.Rows.Count = 0 Then
            MessageBox.Show("Générez d'abord la pièce comptable : le fichier en dérive." & Environment.NewLine &
                            "Ce qui est chargé dans le core banking doit être exactement ce que vous avez validé.",
                            "Action impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not _dateActivite.HasValue Then
            MessageBox.Show("La journée traitée n'a pas pu être déterminée : impossible de dater les écritures.",
                            "Action impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Compensation J+1 : l'activité du jour J est portée en valeur au lendemain.
        Dim dateCompensation As Date = _dateActivite.Value.Date.AddDays(1)

        Dim messageErreur As String = String.Empty
        Dim fichier As DataTable = CoreBankingService.Construire(_dtPieceGeneree, dateCompensation, messageErreur)

        If fichier Is Nothing Then
            MessageBox.Show(messageErreur, "Fichier non produit", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tsslStatut.Text = "Fichier core banking non produit : " & messageErreur
            Return
        End If

        Dim chemin As String = DemanderLeChemin(CoreBankingService.NomDeFichier(dateCompensation))
        If chemin.Length = 0 Then Return

        Try
            Cursor = Cursors.WaitCursor

            ' Seul AMOUNT est écrit en nombre : tout le reste est du texte, sans quoi Excel
            ' réinterpréterait les numéros de compte et les numéros de lot.
            ExcelExportService.ExporterTableBrute(fichier, New String() {"AMOUNT"}, "CoreBanking", chemin)

            Dim numeroLot As String = CoreBankingService.NumeroDeLot(dateCompensation)

            MessageBox.Show(
                $"Fichier produit : {IO.Path.GetFileName(chemin)}" & Environment.NewLine & Environment.NewLine &
                $"    lignes         : {fichier.Rows.Count}" & Environment.NewLine &
                $"    date de valeur : {dateCompensation:dd/MM/yyyy}" & Environment.NewLine &
                $"    numéro de lot  : {numeroLot}" & Environment.NewLine & Environment.NewLine &
                "Le numéro de lot est dérivé de la date : réexporter cette journée redonnera le " &
                "même numéro, ce qui permet au core banking de reconnaître un double chargement.",
                "Fichier core banking", MessageBoxButtons.OK, MessageBoxIcon.Information)

            tsslStatut.Text = $"Fichier core banking produit : {fichier.Rows.Count} ligne(s), lot {numeroLot}."

        Catch ex As Exception
            MessageBox.Show("Production du fichier impossible : " & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>Demande où enregistrer le fichier. Chaîne vide si l'utilisateur renonce.</summary>
    Private Function DemanderLeChemin(nomPropose As String) As String

        Using dialogue As New SaveFileDialog()

            dialogue.Title = "Enregistrer le fichier destiné au core banking"
            dialogue.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
            dialogue.FileName = nomPropose
            dialogue.OverwritePrompt = True

            If dialogue.ShowDialog(Me) <> DialogResult.OK Then Return String.Empty
            Return dialogue.FileName
        End Using
    End Function

    Private Sub btnGenererPiece_Click(sender As Object, e As EventArgs) Handles btnGenererPiece.Click

        If _listeCalculs Is Nothing OrElse _listeCalculs.Count = 0 Then
            MessageBox.Show("Veuillez d'abord charger les rapports et lancer le calcul (Afficher / Calculer).",
                             "Action impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Cursor = Cursors.WaitCursor

            Dim dtPiece As DataTable = PieceComptableService.GenererPieceComptable(_listeCalculs)
            Dim messageControle As String = String.Empty
            Dim pieceUtilisable As Boolean = PieceComptableService.VerifierEquilibrePiece(dtPiece, messageControle)

            If Not pieceUtilisable Then
                MessageBox.Show(messageControle, "Anomalie d'équilibrage", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tsslStatut.Text = "Génération bloquée : " & messageControle
                Return
            End If

            _dtPieceGeneree = dtPiece
            tsslStatut.Text = messageControle

            ' La journée est comptabilisée : c'est le moment de l'historiser, et pas avant.
            ' Un simple affichage ne doit rien laisser dans l'historique.
            HistoriserLaJournee()

            OuvrirPieceDansExcel()

        Catch ex As Exception
            MessageBox.Show("Erreur lors de la génération de la pièce comptable : " & ex.Message,
                             "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Affiche la pièce comptable du seul point de vente sélectionné dans la grille de contrôle.
    ''' La pièce est construite par le même service que la pièce globale, en ne lui transmettant
    ''' que l'Account concerné : les écritures sont donc rigoureusement identiques à celles que
    ''' produira la pièce globale pour ce point de vente.
    '''
    ''' VerifierEquilibrePiece n'est volontairement PAS appelée ici : le compte d'attente ne doit
    ''' s'appliquer qu'à la pièce globale, jamais point de vente par point de vente (section 14).
    ''' L'écart d'arrondi propre à cet Account est simplement affiché, sans être corrigé.
    ''' </summary>
    Private Sub btnPieceAccount_Click(sender As Object, e As EventArgs) Handles btnPieceAccount.Click
        AfficherPieceDuPdvSelectionne()
    End Sub

    ''' <summary>Double-cliquer une ligne de la grille ouvre la pièce de ce point de vente.</summary>
    Private Sub dgvControle_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvControle.CellDoubleClick
        If e.RowIndex >= 0 Then
            AfficherPieceDuPdvSelectionne()
        End If
    End Sub

    Private Sub AfficherPieceDuPdvSelectionne()

        If _listeCalculs Is Nothing OrElse _listeCalculs.Count = 0 Then
            MessageBox.Show("Veuillez d'abord charger les rapports et lancer le calcul (Afficher / Calculer).",
                             "Action impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If dgvControle.CurrentRow Is Nothing Then
            MessageBox.Show("Sélectionnez d'abord une ligne (un Account) dans la grille de contrôle.",
                             "Aucune sélection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim account As String = Convert.ToString(dgvControle.CurrentRow.Cells("Account").Value)
        Dim calc As CalculWU = _listeCalculs.FirstOrDefault(
            Function(c) String.Equals(c.Account, account, StringComparison.OrdinalIgnoreCase))

        If calc Is Nothing Then
            MessageBox.Show($"Aucun calcul trouvé pour l'Account {account}.",
                             "Account introuvable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Cursor = Cursors.WaitCursor

            Dim dtPieceAccount As DataTable = PieceComptableService.GenererPieceComptable(New CalculWU() {calc})

            If dtPieceAccount.Rows.Count = 0 Then
                MessageBox.Show($"L'Account {calc.Account} ne génère aucune écriture (tous ses montants sont nuls).",
                                 "Pièce vide", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim titre As String = $"Pièce comptable — {calc.Account} — {calc.Designation}"
            Dim sousTitre As String = $"Type : {calc.TypePdv}" &
                                      If(String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase),
                                         $"     Taux sous-agent : {calc.TauxSA:P2}", String.Empty) &
                                      $"     Compte de compensation : {If(String.IsNullOrWhiteSpace(calc.CompteCompense), "(non paramétré)", calc.CompteCompense)}" &
                                      $"     Compte de commission : {If(String.IsNullOrWhiteSpace(calc.CompteCommission), "(non paramétré)", calc.CompteCommission)}"

            Using formulaire As New FrmPieceComptable(dtPieceAccount, titre, sousTitre)
                formulaire.ShowDialog(Me)
            End Using

            tsslStatut.Text = $"Pièce comptable affichée pour l'Account {calc.Account}."

        Catch ex As Exception
            MessageBox.Show("Erreur lors de la génération de la pièce du point de vente : " & ex.Message,
                             "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Ouvre directement la pièce comptable générée dans Microsoft Excel (fenêtre visible),
    ''' sans boîte de dialogue d'enregistrement — le classeur est sauvegardé dans un fichier
    ''' temporaire et laissé ouvert pour consultation/impression/enregistrement immédiat.
    ''' Un échec (Excel absent du poste, etc.) n'invalide jamais la pièce comptable déjà
    ''' générée en mémoire : un message d'avertissement est affiché à la place.
    ''' </summary>
    Private Sub OuvrirPieceDansExcel()
        If _dtPieceGeneree Is Nothing OrElse _dtPieceGeneree.Rows.Count = 0 Then Return

        Try
            Dim cheminTemp As String = PieceComptableService.OuvrirPieceComptableExcel(_dtPieceGeneree)
            tsslStatut.Text = $"Pièce comptable ouverte dans Excel ({_dtPieceGeneree.Rows.Count} lignes) : {cheminTemp}"
        Catch ex As Exception
            MessageBox.Show(
                "Impossible d'ouvrir la pièce comptable dans Excel : " & ex.Message & Environment.NewLine &
                "La pièce comptable reste disponible (DataTable dtPiece en mémoire).",
                "Ouverture Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

#End Region

End Class
