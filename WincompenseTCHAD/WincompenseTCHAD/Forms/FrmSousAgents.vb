Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Gestion des sous-agents Western Union (table T_Pdv_SA) : consultation, création,
''' modification et suppression.
'''
''' Le formulaire n'exécute aucune requête SQL lui-même : tout passe par PdvRepository.
''' Il travaille en deux modes, matérialisés par _enCreation :
'''   - CONSULTATION/MODIFICATION : une fiche existante est sélectionnée dans la liste ;
'''     l'Account est affiché mais verrouillé — c'est la clé sous laquelle les rapports
'''     Western Union désignent le point de vente, la renommer romprait le lien avec
'''     l'historique. Pour changer un Account, il faut supprimer puis recréer la fiche ;
'''   - CRÉATION : les champs sont vides et l'Account est saisissable.
''' </summary>
Public Class FrmSousAgents

    ''' <summary>Liste actuellement affichée, telle que renvoyée par la base.</summary>
    Private _liste As List(Of PointDeVenteSA)

    ''' <summary>Vrai lorsque l'utilisateur saisit une NOUVELLE fiche (non encore enregistrée).</summary>
    Private _enCreation As Boolean = False

    ''' <summary>Empêche le remplissage des champs pendant un rechargement de la liste.</summary>
    Private _chargementEnCours As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

#Region "Chargement et affichage de la liste"

    Private Sub FrmSousAgents_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Habillage et adaptation à l'écran du poste, avant tout autre traitement : la fenêtre
        ' prend sa taille définitive une fois pour toutes, et les contrôles ancrés suivent.
        ThemeWU.Appliquer(Me, "Sous-agents", btnEnregistrer)
        DimensionsWU.Adapter(Me, True)

        If Not SessionWU.PeutGererLesPointsDeVente Then
            MessageBox.Show("Le paramétrage des sous-agents est réservé aux commerciaux et aux administrateurs.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        ChargerGroupes()
        ChargerListe()
    End Sub

    ''' <summary>Recharge la liste depuis la base en appliquant le filtre de recherche courant.</summary>
    Private Sub ChargerListe(Optional codeASelectionner As String = Nothing)

        Cursor = Cursors.WaitCursor
        _chargementEnCours = True

        Try
            Dim messageErreur As String = String.Empty
            _liste = PdvRepository.ListerSousAgents(txtRecherche.Text, messageErreur)

            dgvListe.DataSource = Nothing
            dgvListe.DataSource = _liste
            RenommerColonnes()

            If Not String.IsNullOrEmpty(messageErreur) Then
                lblStatut.Text = messageErreur
                lblNombre.Text = String.Empty
                MessageBox.Show(messageErreur, "Base SQL Server", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                lblNombre.Text = $"{_liste.Count} sous-agent(s)"
                lblStatut.Text = If(_liste.Count = 0,
                                    "Aucun sous-agent ne correspond à la recherche.",
                                    "Sélectionnez une fiche pour la modifier, ou cliquez sur « Nouveau ».")
            End If

        Finally
            _chargementEnCours = False
            Cursor = Cursors.Default
        End Try

        ' Après un enregistrement, on revient sur la fiche concernée plutôt que sur la première :
        ' l'utilisateur garde ainsi sous les yeux ce qu'il vient de saisir.
        If Not String.IsNullOrEmpty(codeASelectionner) Then
            SelectionnerCode(codeASelectionner)
        End If

        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>Donne aux colonnes de la grille des intitulés lisibles plutôt que les noms de propriétés.</summary>
    Private Sub RenommerColonnes()

        DefinirEntete("CodePdv", "Account")
        DefinirEntete("Designation", "Désignation")
        DefinirEntete("GroupeStatistique", "Groupe statistique")
        DefinirEntete("Taux", "Taux")
        DefinirEntete("CompteCompense", "Compte compensation")
        DefinirEntete("CompteCommission", "Compte commission")
        DefinirEntete("CodeAgence", "Code agence")

        If dgvListe.Columns.Contains("Taux") Then
            dgvListe.Columns("Taux").DefaultCellStyle.Format = "P0"
            dgvListe.Columns("Taux").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If

        dgvListe.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub DefinirEntete(nomColonne As String, intitule As String)
        If dgvListe.Columns.Contains(nomColonne) Then
            dgvListe.Columns(nomColonne).HeaderText = intitule
        End If
    End Sub

    ''' <summary>Sélectionne dans la grille la ligne portant l'Account indiqué.</summary>
    Private Sub SelectionnerCode(code As String)

        For Each ligne As DataGridViewRow In dgvListe.Rows
            Dim fiche As PointDeVenteSA = TryCast(ligne.DataBoundItem, PointDeVenteSA)
            If fiche IsNot Nothing AndAlso String.Equals(fiche.CodePdv, code, StringComparison.OrdinalIgnoreCase) Then
                ligne.Selected = True
                dgvListe.CurrentCell = ligne.Cells(0)
                Return
            End If
        Next
    End Sub

    ''' <summary>Fiche actuellement sélectionnée dans la grille, ou Nothing.</summary>
    Private Function FicheSelectionnee() As PointDeVenteSA

        If dgvListe.CurrentRow Is Nothing Then Return Nothing
        Return TryCast(dgvListe.CurrentRow.DataBoundItem, PointDeVenteSA)
    End Function

    Private Sub dgvListe_SelectionChanged(sender As Object, e As EventArgs) Handles dgvListe.SelectionChanged
        If _chargementEnCours Then Return
        _enCreation = False
        AfficherFicheSelectionnee()
        MettreAJourEtatBoutons()
    End Sub

    ''' <summary>
    ''' Recopie la fiche sélectionnée dans les champs de saisie.
    '''
    ''' Les valeurs affichées sont celles de la FICHE, jamais celles de son groupe : si elles
    ''' divergent — ce que la règle d'héritage exclut, mais que des données antérieures peuvent
    ''' présenter — la divergence doit rester visible, pas être masquée par une recopie
    ''' silencieuse qui finirait par l'enregistrer.
    ''' </summary>
    Private Sub AfficherFicheSelectionnee()

        If _enCreation Then Return

        Dim fiche As PointDeVenteSA = FicheSelectionnee()

        If fiche Is Nothing Then
            ViderChamps()
            Return
        End If

        _affichageFiche = True
        Try
            txtCodePdv.Text = fiche.CodePdv
            txtDesignation.Text = fiche.Designation
            cboGroupeStatistique.Text = fiche.GroupeStatistique
            txtTaux.Text = fiche.Taux.ToString("0.00", Globalization.CultureInfo.CurrentCulture)
            txtCompteCompense.Text = fiche.CompteCompense
            txtCompteCommission.Text = fiche.CompteCommission
            txtCodeAgence.Text = fiche.CodeAgence
        Finally
            _affichageFiche = False
        End Try

        VerrouillerChampsHerites()
        SignalerDivergenceAvecLeGroupe(fiche)
    End Sub

    ''' <summary>
    ''' Signale, dans la barre d'état, qu'une fiche ne porte pas les valeurs de son groupe.
    ''' Simple information : rien n'est corrigé tant que l'utilisateur n'enregistre pas.
    ''' </summary>
    Private Sub SignalerDivergenceAvecLeGroupe(fiche As PointDeVenteSA)

        Dim groupe As GroupeStatistiqueWU = TrouverGroupe(fiche.GroupeStatistique)
        If groupe Is Nothing Then Return

        If groupe.CorrespondA(fiche) Then
            lblStatut.Text = $"Groupe « {groupe.Nom} » : {groupe.NombreSousAgents} sous-agent(s), " &
                             $"taux {groupe.Taux:0.00}, activité {groupe.CompteActivite}, " &
                             $"commission {groupe.CompteCommission}."
            Return
        End If

        lblStatut.Text = $"ATTENTION — cette fiche ne porte pas les valeurs de son groupe " &
                         $"« {groupe.Nom} » (taux {groupe.Taux:0.00}, activité {groupe.CompteActivite}, " &
                         $"commission {groupe.CompteCommission})."
    End Sub

    Private Sub ViderChamps()
        txtCodePdv.Text = String.Empty
        txtDesignation.Text = String.Empty
        cboGroupeStatistique.Text = String.Empty
        txtTaux.Text = String.Empty
        txtCompteCompense.Text = String.Empty
        txtCompteCommission.Text = String.Empty
        txtCodeAgence.Text = String.Empty
    End Sub

    ''' <summary>
    ''' L'Account n'est saisissable qu'en création ; la suppression n'a de sens que sur une
    ''' fiche existante.
    ''' </summary>
    Private Sub MettreAJourEtatBoutons()
        txtCodePdv.ReadOnly = Not _enCreation
        btnSupprimer.Enabled = Not _enCreation AndAlso FicheSelectionnee() IsNot Nothing
        grpDetail.Text = If(_enCreation, "Nouveau sous-agent", "Fiche du sous-agent")
        VerrouillerChampsHerites()
    End Sub

#End Region

#Region "Groupes statistiques et héritage"

    ''' <summary>
    ''' Groupes statistiques lus dans T_GroupeStatistique au dernier rafraîchissement. Ils
    ''' alimentent la liste déroulante et portent les trois valeurs dont le sous-agent hérite.
    ''' </summary>
    Private _groupes As New List(Of GroupeStatistiqueWU)

    ''' <summary>
    ''' Empêche l'héritage de se déclencher pendant l'affichage d'une fiche existante : une
    ''' fiche doit s'afficher avec SES valeurs, même si elles ont dérivé de celles de son
    ''' groupe. Écraser silencieusement une dérive reviendrait à la masquer, puis à l'enregistrer.
    ''' </summary>
    Private _affichageFiche As Boolean = False

    ''' <summary>Recharge les groupes et la liste déroulante depuis T_GroupeStatistique.</summary>
    Private Sub ChargerGroupes()

        Dim messageErreur As String = String.Empty
        Dim groupes As List(Of GroupeStatistiqueWU) = PdvRepository.ListerGroupes(String.Empty, messageErreur)

        If Not String.IsNullOrEmpty(messageErreur) Then
            ' Le plus souvent : la table n'existe pas encore (migration non jouée). Le dire une
            ' fois, sans empêcher de consulter les sous-agents.
            _groupes = New List(Of GroupeStatistiqueWU)
            cboGroupeStatistique.Items.Clear()
            lblStatut.Text = "Groupes statistiques indisponibles."
            MessageBox.Show(messageErreur, "Groupes statistiques", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        _groupes = groupes

        Dim saisieEnCours As String = cboGroupeStatistique.Text

        cboGroupeStatistique.BeginUpdate()
        Try
            cboGroupeStatistique.Items.Clear()
            For Each groupe As GroupeStatistiqueWU In _groupes
                cboGroupeStatistique.Items.Add(groupe.Nom)
            Next
        Finally
            cboGroupeStatistique.EndUpdate()
        End Try

        cboGroupeStatistique.Text = saisieEnCours
    End Sub

    ''' <summary>
    ''' Recherche un groupe par son libellé, sans tenir compte de la casse ni des espaces de
    ''' bordure : « reseau » saisi désigne bien le groupe « RESEAU ».
    ''' </summary>
    Private Function TrouverGroupe(saisie As String) As GroupeStatistiqueWU

        If String.IsNullOrWhiteSpace(saisie) Then Return Nothing

        Dim recherche As String = saisie.Trim()

        For Each groupe As GroupeStatistiqueWU In _groupes
            If String.Equals(groupe.Nom, recherche, StringComparison.OrdinalIgnoreCase) Then
                Return groupe
            End If
        Next

        Return Nothing
    End Function

    Private Sub cboGroupeStatistique_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGroupeStatistique.SelectedIndexChanged
        AppliquerHeritageGroupe()
    End Sub

    ''' <summary>Un libellé peut aussi être frappé au clavier : l'héritage s'applique à la sortie du champ.</summary>
    Private Sub cboGroupeStatistique_Leave(sender As Object, e As EventArgs) Handles cboGroupeStatistique.Leave
        AppliquerHeritageGroupe()
    End Sub

    ''' <summary>
    ''' Recopie dans les zones de saisie le compte d'activité, le compte de commission et le
    ''' taux du groupe sélectionné. Ces trois champs restent TOUJOURS en lecture seule : leurs
    ''' valeurs appartiennent au groupe, et se modifient dans l'écran des groupes statistiques,
    ''' où le changement s'applique d'un coup à tous les sous-agents concernés.
    ''' </summary>
    Private Sub AppliquerHeritageGroupe()

        ' Pendant l'affichage d'une fiche, ses propres valeurs priment (voir _affichageFiche).
        If _affichageFiche OrElse _chargementEnCours Then Return

        Dim groupe As GroupeStatistiqueWU = TrouverGroupe(cboGroupeStatistique.Text)

        If groupe Is Nothing Then
            If cboGroupeStatistique.Text.Trim().Length > 0 Then
                lblStatut.Text = $"Groupe « {cboGroupeStatistique.Text.Trim()} » inconnu : " &
                                 "cliquez sur « ... » pour le créer."
            End If
            Return
        End If

        txtCompteCompense.Text = groupe.CompteActivite
        txtCompteCommission.Text = groupe.CompteCommission
        txtTaux.Text = groupe.Taux.ToString("0.00", Globalization.CultureInfo.CurrentCulture)

        lblStatut.Text = $"Groupe « {groupe.Nom} » : {groupe.NombreSousAgents} sous-agent(s), " &
                         $"taux {groupe.Taux:0.00}, activité {groupe.CompteActivite}, " &
                         $"commission {groupe.CompteCommission}." &
                         If(groupe.EstEnregistre, String.Empty,
                            "  [groupe hérité, pas encore enregistré dans la table des groupes]")
    End Sub

    ''' <summary>
    ''' Les trois champs hérités sont en lecture seule, sans exception : ils appartiennent au
    ''' groupe. Le fond grisé le rend visible d'un coup d'oeil.
    ''' </summary>
    Private Sub VerrouillerChampsHerites()

        txtCompteCompense.ReadOnly = True
        txtCompteCommission.ReadOnly = True
        txtTaux.ReadOnly = True

        txtCompteCompense.BackColor = Drawing.SystemColors.Control
        txtCompteCommission.BackColor = Drawing.SystemColors.Control
        txtTaux.BackColor = Drawing.SystemColors.Control
    End Sub

    ''' <summary>
    ''' Ouvre l'écran des groupes statistiques, en création si le libellé saisi est inconnu.
    ''' Au retour, la liste est rechargée et le groupe nouvellement créé aussitôt appliqué.
    ''' </summary>
    Private Sub btnNouveauGroupe_Click(sender As Object, e As EventArgs) Handles btnNouveauGroupe.Click

        Dim saisie As String = cboGroupeStatistique.Text.Trim()
        Dim aCreer As String = If(TrouverGroupe(saisie) Is Nothing, saisie, String.Empty)

        Using formulaire As New FrmGroupesStatistiques(aCreer)
            formulaire.ShowDialog(Me)
        End Using

        ChargerGroupes()

        ' Le groupe vient peut-être d'être créé : on l'applique sans faire ressaisir son nom.
        If saisie.Length > 0 Then
            cboGroupeStatistique.Text = saisie
            AppliquerHeritageGroupe()
        End If
    End Sub

    ''' <summary>
    ''' Valide le groupe statistique saisi et retourne le libellé à enregistrer.
    '''
    ''' Le groupe est OBLIGATOIRE : c'est de lui que le sous-agent tient son compte d'activité,
    ''' son compte de commission et son taux.
    '''
    ''' Trois cas :
    '''   - groupe enregistré dans T_GroupeStatistique : retenu tel quel ;
    '''   - groupe HÉRITÉ, encore porté par les seuls sous-agents de T_Pdv_SA : accepté, avec
    '''     proposition de l'enregistrer au passage. Le refus n'est jamais bloquant — sans quoi
    '''     une migration non aboutie empêcherait tout rattachement à un groupe existant ;
    '''   - libellé totalement inconnu : l'écran des groupes s'ouvre, pré-rempli, car un groupe
    '''     se définit par trois valeurs et non par un simple nom.
    '''
    ''' Retourne Nothing si le groupe n'est finalement pas déterminé : l'enregistrement doit
    ''' alors être abandonné.
    ''' </summary>
    Private Function ValiderGroupeStatistique() As String

        Dim saisie As String = cboGroupeStatistique.Text.Trim()

        If saisie.Length = 0 Then
            MessageBox.Show(
                "Le groupe statistique est obligatoire." & Environment.NewLine & Environment.NewLine &
                "C'est de lui que le sous-agent hérite son compte d'activité, son compte de " &
                "commission et son taux." & Environment.NewLine & Environment.NewLine &
                "Choisissez un groupe dans la liste, ou cliquez sur « ... » pour en créer un.",
                "Groupe statistique obligatoire", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            cboGroupeStatistique.Focus()
            Return Nothing
        End If

        Dim existant As GroupeStatistiqueWU = TrouverGroupe(saisie)

        If existant IsNot Nothing Then

            ' Groupe hérité de l'ancien paramétrage : on propose de le régulariser, sans l'exiger.
            If Not existant.EstEnregistre Then
                ProposerEnregistrementGroupeHerite(existant)
                existant = If(TrouverGroupe(saisie), existant)
            End If

            ' Aligne la casse sur celle enregistrée : « reseau » saisi rejoint « RESEAU ».
            If Not String.Equals(existant.Nom, saisie, StringComparison.Ordinal) Then
                cboGroupeStatistique.Text = existant.Nom
            End If

            Return existant.Nom
        End If

        Dim reponse As DialogResult = MessageBox.Show(
            $"Le groupe statistique « {saisie} » n'existe pas." & Environment.NewLine & Environment.NewLine &
            "Un groupe se définit par un compte d'activité, un compte de commission et un taux : " &
            "il ne peut donc pas être créé d'un simple nom." & Environment.NewLine & Environment.NewLine &
            "Ouvrir l'écran des groupes statistiques pour le créer maintenant ?",
            "Groupe inconnu", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then
            cboGroupeStatistique.DroppedDown = True
            cboGroupeStatistique.Focus()
            Return Nothing
        End If

        Using formulaire As New FrmGroupesStatistiques(saisie)
            formulaire.ShowDialog(Me)
        End Using

        ChargerGroupes()

        existant = TrouverGroupe(saisie)
        If existant Is Nothing Then
            ' Le groupe n'a finalement pas été créé : rien à enregistrer.
            cboGroupeStatistique.Focus()
            Return Nothing
        End If

        cboGroupeStatistique.Text = existant.Nom
        AppliquerHeritageGroupe()
        Return existant.Nom
    End Function

    ''' <summary>
    ''' Propose d'inscrire dans T_GroupeStatistique un groupe hérité de l'ancien paramétrage,
    ''' avec les valeurs que portent déjà ses sous-agents. Un refus n'empêche rien : le
    ''' sous-agent est rattaché au groupe dans tous les cas.
    ''' </summary>
    Private Sub ProposerEnregistrementGroupeHerite(groupe As GroupeStatistiqueWU)

        Dim reponse As DialogResult = MessageBox.Show(
            $"Le groupe « {groupe.Nom} » provient de l'ancien paramétrage : il est porté par " &
            $"{groupe.NombreSousAgents} sous-agent(s) mais ne figure pas encore dans la table des groupes." &
            Environment.NewLine & Environment.NewLine &
            "L'y enregistrer maintenant, avec ses valeurs actuelles ?" & Environment.NewLine & Environment.NewLine &
            $"    compte d'activité    : {groupe.CompteActivite}" & Environment.NewLine &
            $"    compte de commission : {groupe.CompteCommission}" & Environment.NewLine &
            $"    taux                 : {groupe.Taux:0.00}" & Environment.NewLine & Environment.NewLine &
            "Répondre « Non » n'empêche pas d'y rattacher ce sous-agent.",
            "Groupe hérité non enregistré", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reponse <> DialogResult.Yes Then
            lblStatut.Text = $"Groupe « {groupe.Nom} » utilisé sans être enregistré dans la table des groupes."
            Return
        End If

        Dim messageErreur As String = String.Empty

        If PdvRepository.AjouterGroupe(groupe, messageErreur) Then
            lblStatut.Text = $"Groupe « {groupe.Nom} » enregistré dans la table des groupes."
            ChargerGroupes()
            Return
        End If

        ' Échec le plus fréquent : un de ses comptes appartient déjà à un autre groupe. Le
        ' rattachement du sous-agent reste possible, le paramétrage est simplement à arbitrer.
        MessageBox.Show(
            messageErreur & Environment.NewLine & Environment.NewLine &
            $"Le sous-agent peut tout de même être rattaché au groupe « {groupe.Nom} ».",
            "Groupe non enregistré", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

#End Region

#Region "Recherche"

    ''' <summary>
    ''' La recherche se déclenche sur Entrée (ou sur « Actualiser »), jamais à chaque caractère
    ''' frappé : inutile d'interroger la base à chaque touche.
    ''' </summary>
    Private Sub txtRecherche_KeyDown(sender As Object, e As KeyEventArgs) Handles txtRecherche.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            _enCreation = False
            ChargerListe()
        End If
    End Sub

    Private Sub btnActualiser_Click(sender As Object, e As EventArgs) Handles btnActualiser.Click
        _enCreation = False
        ChargerGroupes()
        ChargerListe()
    End Sub

#End Region

#Region "Création, modification, suppression"

    Private Sub btnNouveau_Click(sender As Object, e As EventArgs) Handles btnNouveau.Click

        _enCreation = True
        dgvListe.ClearSelection()
        ViderChamps()
        txtTaux.Text = "0,00"
        MettreAJourEtatBoutons()

        lblStatut.Text = "Nouveau sous-agent : choisissez son groupe statistique, dont il héritera " &
                         "le compte d'activité, le compte de commission et le taux."
        txtCodePdv.Focus()
    End Sub

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim fiche As PointDeVenteSA = LireChamps()
        If fiche Is Nothing Then Return ' Taux illisible : message déjà affiché.

        fiche.Normaliser()

        Dim anomalies As List(Of String) = fiche.Anomalies()
        If anomalies.Count > 0 Then
            MessageBox.Show(
                "Saisie incomplète ou incorrecte :" & Environment.NewLine & Environment.NewLine &
                "    " & String.Join(Environment.NewLine & "    ", anomalies),
                "Fiche non enregistrable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Le groupe statistique est obligatoire et détermine à lui seul le compte d'activité, le
        ' compte de commission et le taux. Ce contrôle vient APRÈS les autres : inutile d'ouvrir
        ' l'écran des groupes si la fiche est de toute façon incomplète par ailleurs.
        Dim nomGroupe As String = ValiderGroupeStatistique()
        If nomGroupe Is Nothing Then Return ' Groupe non déterminé : message déjà affiché.

        Dim groupe As GroupeStatistiqueWU = TrouverGroupe(nomGroupe)
        If groupe Is Nothing Then Return ' Ne devrait pas se produire : ValiderGroupeStatistique l'a vérifié.

        ' Les valeurs viennent du groupe, jamais de la fiche. Si la fiche en portait d'autres —
        ' une dérive héritée du paramétrage antérieur — elle est réalignée, mais jamais sans le dire.
        If Not groupe.CorrespondA(fiche) Then
            Dim retour As DialogResult = MessageBox.Show(
                $"Cette fiche ne porte pas les valeurs de son groupe « {groupe.Nom} » :" &
                Environment.NewLine & Environment.NewLine &
                $"    compte d'activité    : {fiche.CompteCompense}  ->  {groupe.CompteActivite}" & Environment.NewLine &
                $"    compte de commission : {fiche.CompteCommission}  ->  {groupe.CompteCommission}" & Environment.NewLine &
                $"    taux                 : {fiche.Taux:0.00}  ->  {groupe.Taux:0.00}" & Environment.NewLine &
                Environment.NewLine &
                "Un Account hérite des valeurs de son groupe : l'enregistrement les adoptera." &
                Environment.NewLine & Environment.NewLine &
                "Pour modifier les valeurs elles-mêmes, passez par l'écran des groupes " &
                "(bouton « ... ») : le changement s'appliquera alors à tout le groupe." &
                Environment.NewLine & Environment.NewLine & "Continuer ?",
                "Alignement sur le groupe", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If retour <> DialogResult.Yes Then Return
        End If

        groupe.AppliquerA(fiche)


        ' Un même Account dans les deux tables serait silencieusement traité comme un sous-agent :
        ' la recherche interroge T_Pdv_SA en premier. L'utilisateur doit le savoir avant de valider.
        If _enCreation AndAlso PdvRepository.ExisteDansAutreTable(fiche.CodePdv, estSousAgent:=True) Then
            Dim reponse As DialogResult = MessageBox.Show(
                $"L'Account « {fiche.CodePdv} » est déjà enregistré comme AGENCE PROPRE (T_Pdv_EC)." &
                Environment.NewLine & Environment.NewLine &
                "S'il est également créé comme sous-agent, c'est la fiche sous-agent qui sera retenue " &
                "lors de la comptabilisation : la fiche agence deviendrait sans effet." & Environment.NewLine &
                Environment.NewLine & "Créer tout de même ce sous-agent ?",
                "Account déjà utilisé comme agence", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If reponse <> DialogResult.Yes Then Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            Dim reussi As Boolean = If(_enCreation,
                                       PdvRepository.AjouterSousAgent(fiche, messageErreur),
                                       PdvRepository.ModifierSousAgent(fiche, messageErreur))

            If Not reussi Then
                MessageBox.Show(messageErreur, "Enregistrement impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = If(_enCreation,
                                $"Sous-agent « {fiche.CodePdv} » créé.",
                                $"Sous-agent « {fiche.CodePdv} » modifié.")
            _enCreation = False


            ' Un groupe qui vient d'être créé n'existait dans aucune fiche : il devient
            ' sélectionnable dès maintenant pour les saisies suivantes.
            ChargerGroupes()

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe(fiche.CodePdv)
    End Sub

    Private Sub btnSupprimer_Click(sender As Object, e As EventArgs) Handles btnSupprimer.Click

        Dim fiche As PointDeVenteSA = FicheSelectionnee()
        If fiche Is Nothing Then Return

        Dim reponse As DialogResult = MessageBox.Show(
            $"Supprimer définitivement le sous-agent « {fiche.CodePdv} » ({fiche.Designation}) ?" &
            Environment.NewLine & Environment.NewLine &
            "Les rapports Western Union portant cet Account ne seront plus rattachés à aucun point " &
            "de vente : ils apparaîtront en INCONNU et leur pièce comptable utilisera le compte " &
            "courant WU au lieu du compte de compensation." & Environment.NewLine & Environment.NewLine &
            "Cette suppression est irréversible.",
            "Confirmer la suppression", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If reponse <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        Try
            Dim messageErreur As String = String.Empty
            If Not PdvRepository.SupprimerSousAgent(fiche.CodePdv, messageErreur) Then
                MessageBox.Show(messageErreur, "Suppression impossible", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            lblStatut.Text = $"Sous-agent « {fiche.CodePdv} » supprimé."

            ' Supprimer le dernier sous-agent d'un groupe fait disparaître ce groupe : la liste
            ' déroulante doit cesser de le proposer.
            ChargerGroupes()

        Finally
            Cursor = Cursors.Default
        End Try

        ChargerListe()
    End Sub

    ''' <summary>
    ''' Construit une fiche à partir des champs saisis. Retourne Nothing si le taux n'est pas
    ''' un nombre exploitable — seule vérification faite ici, les autres relevant de
    ''' PointDeVenteSA.Anomalies.
    ''' </summary>
    Private Function LireChamps() As PointDeVenteSA

        Dim taux As Decimal = 0D

        If Not String.IsNullOrWhiteSpace(txtTaux.Text) AndAlso
           Not PointDeVenteSA.EssayerAnalyserTaux(txtTaux.Text, taux) Then

            MessageBox.Show(
                $"Le taux saisi (« {txtTaux.Text} ») n'est pas un nombre." & Environment.NewLine & Environment.NewLine &
                "Saisissez une fraction : 0,70 pour 70 %.",
                "Taux incorrect", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTaux.Focus()
            Return Nothing
        End If

        Return New PointDeVenteSA() With {
            .CodePdv = txtCodePdv.Text,
            .Designation = txtDesignation.Text,
            .GroupeStatistique = cboGroupeStatistique.Text,
            .Taux = taux,
            .CompteCompense = txtCompteCompense.Text,
            .CompteCommission = txtCompteCommission.Text,
            .CodeAgence = txtCodeAgence.Text
        }
    End Function

    ''' <summary>
    ''' Ouvre l'état de restitution « sous-agents par groupe statistique » : récapitulatif des
    ''' groupes et détail de leurs sous-agents. Écran de consultation seule, ouvert en modeless
    ''' afin de pouvoir être laissé à l'écran pendant la saisie d'une fiche.
    ''' </summary>
    Private Sub btnParGroupe_Click(sender As Object, e As EventArgs) Handles btnParGroupe.Click

        ' Dans le conteneur MDI, l'état s'ouvre comme fenêtre fille : il se range avec les
        ' autres et se retrouve dans le menu Fenêtres. Hors MDI — si l'écran a été ouvert en
        ' boîte de dialogue — il reste une fenêtre flottante non modale, l'utilisateur devant
        ' pouvoir consulter la liste tout en modifiant une fiche.
        Dim principal As FrmPrincipal = TryCast(Me.MdiParent, FrmPrincipal)

        If principal IsNot Nothing Then
            principal.AfficherEnfant(Of FrmSousAgentsParGroupe)()
            Return
        End If

        ' Pas de Using : la durée de vie du formulaire dépasse celle de cette méthode.
        Dim etat As New FrmSousAgentsParGroupe()
        etat.Show(Me)
    End Sub

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
