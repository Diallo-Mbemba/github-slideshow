Option Strict On
Option Explicit On

Imports System.Data
Imports System.Windows.Forms

''' <summary>
''' LA NARRATIVE COMPTABLE : l'écran où la banque écrit elle-même les libellés que porteront
''' ses écritures sur la pièce comptable, et le texte que son core banking recevra dans la
''' colonne ADDLTEXT.
'''
''' POURQUOI CET ÉCRAN EXISTE
'''
''' Cette phrase a changé quatre fois en quatre livraisons. Chaque mot a coûté une
''' modification du code, une compilation, un commit, un pull et un redéploiement sur les
''' postes de la banque — pour un texte qui n'entre dans aucun calcul. Ce qui se lit dans le
''' grand livre appartient à la Direction Comptable ; ce qui s'y calcule reste au code.
'''
''' DEUX TEXTES, ET ILS NE VONT PAS AU MÊME ENDROIT
'''
'''   — LE MODÈLE DU POINT DE VENTE, premier onglet : « LD WU ACTIVITE <point de vente>
'''     <période> ». C'est TOUJOURS lui que le core banking reçoit dans ADDLTEXT, pour les
'''     douze lignes du point de vente.
'''   — LES LIBELLÉS PAR NATURE, deuxième onglet : « LD COMPTE COURANT WESTERN UNION ETD »,
'''     « LD TVA SUR COMMISSION WU »… C'est ce que porte LA PIÈCE COMPTABLE, ligne à ligne,
'''     et c'est le mode par défaut — celui de la pièce manuelle de la banque, rétabli par
'''     son rectificatif du 06/10/2026.
'''
''' LA BASCULE NE CONCERNE QUE LA PIÈCE. « Un seul libellé » étend le modèle du point de vente
''' aux douze lignes de la pièce, où ce que chaque ligne EST se lit alors dans son seul numéro
''' de compte. Le fichier core banking, lui, ne change pas : il porte ce modèle dans les deux
''' cas.
'''
''' RIEN À SAISIR POUR RETROUVER LA PIÈCE HABITUELLE. Les treize natures s'affichent avec le
''' modèle qu'elles appliquent réellement — leur libellé historique tant qu'elles n'ont pas
''' été touchées — et seules celles que la banque modifie sont conservées. Une bascule qui
''' réécrirait treize libellés d'un coup serait un piège : on ne découvre pas au grand livre
''' ce qu'un bouton radio a décidé.
'''
''' L'ÉCRAN NE SE CONTENTE PAS DE RECUEILLIR UNE SAISIE
'''
''' Il la MESURE. Le core banking refuse un narratif de plus de cent cinquante caractères, et
''' l'application refuse alors de produire le fichier — un refus qui tomberait le matin de la
''' compense. L'aperçu est donc calculé sur le PIRE CAS RÉEL : le point de vente dont le nom
''' est le plus long dans le référentiel, l'Account le plus long, le code agence le plus long,
''' et la période la plus longue que l'application sache écrire. Ce qui est validé ici passera
''' en production, et chaque nature est mesurée séparément.
'''
''' ET IL GARDE TRACE. Le troisième onglet montre le journal : qui a changé quoi, quand, depuis
''' quel poste, et ce qu'il y avait avant.
''' </summary>
Public Class FrmNarrative

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Le pire cas, lu dans le référentiel"

    ''' <summary>Désignation la plus longue du référentiel, agences et sous-agents confondus.</summary>
    Private _pireDesignation As String = String.Empty

    ''' <summary>Account le plus long. Ce n'est pas forcément celui du point de vente ci-dessus.</summary>
    Private _pireAccount As String = String.Empty

    ''' <summary>Code agence le plus long.</summary>
    Private _pireCodeAgence As String = String.Empty

    ''' <summary>
    ''' La période la plus longue que SuffixeDePeriode sache produire : celle qui est à cheval
    ''' sur deux mois ET deux années, donc écrite « DU 28/12/2026 AU 04/01/2027 ».
    ''' </summary>
    Private _pirePeriode As String = String.Empty

    ''' <summary>
    ''' Faux si le référentiel n'a pas pu être lu. L'aperçu le dit alors, au lieu de faire
    ''' croire qu'un modèle a été mesuré sur des données réelles.
    ''' </summary>
    Private _referentielLu As Boolean = False

    ''' <summary>
    ''' Cherche les valeurs les plus longues du référentiel. C'est ce qui rend le contrôle des
    ''' cent cinquante caractères honnête : mesurer le modèle sur un nom moyen laisserait
    ''' passer celui qui ne tient pas pour l'agence au nom le plus long.
    '''
    ''' LES TROIS VALEURS SONT CHERCHÉES SÉPARÉMENT, et non prises sur un même point de vente.
    ''' Le pire cas d'un modèle qui emploie les trois repères est bien la somme des trois
    ''' maximums, même si aucun point de vente ne les réunit : sinon un modèle validé ici
    ''' dépasserait en production.
    ''' </summary>
    Private Sub ChargerLePireCas()

        _pirePeriode = PieceComptableService.SuffixeDePeriode(New Date(2026, 12, 28),
                                                              New Date(2027, 1, 4))

        Dim messageAgences As String = String.Empty
        Dim messageSousAgents As String = String.Empty

        Dim agences As List(Of PointDeVenteEC) = PdvRepository.ListerAgences(String.Empty, messageAgences)
        Dim sousAgents As List(Of PointDeVenteSA) = PdvRepository.ListerSousAgents(String.Empty, messageSousAgents)

        _referentielLu = messageAgences.Length = 0 AndAlso messageSousAgents.Length = 0

        If agences IsNot Nothing Then
            For Each agence As PointDeVenteEC In agences
                Retenir(agence.Designation, agence.CodeSite, agence.CodeAgenceVoyager)
            Next
        End If

        If sousAgents IsNot Nothing Then
            For Each sousAgent As PointDeVenteSA In sousAgents
                Retenir(sousAgent.Designation, sousAgent.CodePdv, sousAgent.CodeAgence)
            Next
        End If
    End Sub

    ''' <summary>Garde la plus longue des valeurs rencontrées, repère par repère.</summary>
    Private Sub Retenir(designation As String, account As String, codeAgence As String)

        Dim nom As String = If(designation, String.Empty).Trim()
        Dim compte As String = If(account, String.Empty).Trim()
        Dim agence As String = If(codeAgence, String.Empty).Trim()

        If nom.Length > _pireDesignation.Length Then _pireDesignation = nom
        If compte.Length > _pireAccount.Length Then _pireAccount = compte
        If agence.Length > _pireCodeAgence.Length Then _pireCodeAgence = agence
    End Sub

    ''' <summary>Le modèle rendu sur le pire cas du référentiel, tel qu'il partira au core banking.</summary>
    Private Function RenduPireCas(modele As String) As String

        Return ModeleNarrativeWU.Appliquer(modele, _pireDesignation, _pireAccount,
                                           _pireCodeAgence, _pirePeriode)
    End Function

#End Region

#Region "État de l'écran"

    ''' <summary>
    ''' Faux si T_NarrativeNatureWU est absente — script 23 non exécuté. Le mode « par nature »
    ''' est alors proposé en lecture seule avec le script à lancer, plutôt que offert puis
    ''' incapable de rien retenir.
    ''' </summary>
    Private _naturesDisponibles As Boolean = False

    ''' <summary>Vrai si le journal est tenu : son absence n'empêche rien, elle se dit.</summary>
    Private _journalDisponible As Boolean = False

    ''' <summary>
    ''' Les modèles effectifs au chargement, nature par nature. Ils servent à n'écrire que ce
    ''' qui a CHANGÉ : réenregistrer treize natures identiques ferait treize transactions pour
    ''' rien, et l'écran est ouvert pour changer une ligne, pas treize.
    ''' </summary>
    Private ReadOnly _naturesChargees As New Dictionary(Of NatureMouvementWU, String)

    ''' <summary>
    ''' Le modèle global tel qu'il était au chargement.
    '''
    ''' IL EST RETENU À PART, et non déduit de la nature « Mouvement » : celle-là peut être
    ''' personnalisée, et servir de référence au modèle global ferait suivre les douze autres
    ''' à une valeur qui n'est pas le global.
    ''' </summary>
    Private _modeleGlobalCharge As String = String.Empty

    ''' <summary>
    ''' Vrai pendant le chargement de la grille et des boutons radio, pour que leurs événements
    ''' ne déclenchent pas d'aperçu sur un écran encore à moitié rempli.
    ''' </summary>
    Private _enChargement As Boolean = False

#End Region

#Region "Ouverture"

    Private Sub FrmNarrative_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesComptesSystemes Then
            MessageBox.Show("La narrative comptable est réservée à l'administrateur." &
                            Environment.NewLine & Environment.NewLine &
                            "Ces textes partent dans le grand livre de la banque, sur toutes les " &
                            "écritures de toutes les journées : ils ne se changent pas depuis le " &
                            "guichet.",
                            "Accès refusé", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' La fermeture est différée : un formulaire ne peut pas se fermer pendant son
            ' propre chargement, la fenêtre resterait affichée et vide.
            BeginInvoke(New Action(AddressOf Close))
            Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            Charger()
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Charger()

        _enChargement = True
        Try
            txtModele.MaxLength = ConstantesWU.NARRATIVE_MODELE_LONGUEUR_MAX
            AfficherLaLegende()
            ChargerLePireCas()

            ' Les valeurs sont relues en base, et non prises dans le cache : on vient
            ' précisément les modifier, et quelqu'un d'autre a pu le faire avant.
            OptionsWU.Oublier()
            NarrativeRepository.Oublier()
            JournalParametreWU.Oublier()

            _naturesDisponibles = NarrativeRepository.Disponible()
            _journalDisponible = JournalParametreWU.Disponible()

            _modeleGlobalCharge = OptionsWU.NarrativeModele
            txtModele.Text = _modeleGlobalCharge

            AfficherLeMode(OptionsWU.NarrativeMode)
            ChargerLesNatures()
            ChargerLeJournal()

            lblDerniereModification.Text = Historique()
            lblStatut.ForeColor = Drawing.SystemColors.GrayText
            lblStatut.Text = String.Empty

        Finally
            _enChargement = False
        End Try

        AfficherLApercu()
        AfficherLApercuDeLaNature()
    End Sub

    ''' <summary>
    ''' Coche le mode en vigueur, et dit ce que l'absence de la table des natures empêche.
    '''
    ''' LES DEUX MODES RESTENT OFFERTS, MÊME SANS LA TABLE. Le mode « par nature » ne dépend
    ''' plus d'elle : une nature absente retombe sur son libellé historique, que le code porte
    ''' en constante. La table ne sert qu'à PERSONNALISER un libellé — c'est donc la grille
    ''' qui se ferme, pas le mode.
    '''
    ''' ET L'ÉCRAN DIT POURQUOI. Une grille inerte sans explication se lit comme un défaut de
    ''' l'application, et l'informatique de la banque n'a aucun moyen de deviner qu'il lui
    ''' manque un script.
    ''' </summary>
    Private Sub AfficherLeMode(mode As ModeNarrativeWU)

        Dim parNature As Boolean = (mode = ModeNarrativeWU.ParNature)

        rdoModeParNature.Checked = parNature
        rdoModeUnique.Checked = Not parNature

        If _naturesDisponibles Then
            lblModeIndisponible.Visible = False
            Return
        End If

        lblModeIndisponible.Visible = True
        lblModeIndisponible.Text =
            "Les libellés par nature s'appliquent, mais ne sont pas modifiables : la table " &
            "T_NarrativeNatureWU est absente de la base. Faites exécuter " &
            "Scripts\00_InstallationComplete.sql par l'informatique — il la crée, et crée aussi " &
            "le journal des modifications."
    End Sub

    ''' <summary>
    ''' Dit si la banque a enregistré un modèle, ou si celui du code s'applique encore. La
    ''' distinction compte : un modèle jamais enregistré se remet d'aplomb en fermant l'écran,
    ''' un modèle enregistré par quelqu'un d'autre se discute avec lui.
    ''' </summary>
    Private Function Historique() As String

        Dim trace As String = OptionsWU.DerniereModification(OptionsWU.CLE_NARRATIVE_MODELE)

        If trace.Length = 0 Then
            Return "Aucun modèle n'est enregistré dans la base : celui qui s'affiche est le " &
                   "modèle par défaut de l'application."
        End If

        Return trace
    End Function

    ''' <summary>
    ''' La légende des repères, lue dans ModeleNarrativeWU. L'écran n'écrit pas la sienne :
    ''' elle divergerait au premier repère ajouté, et la banque saisirait un repère que la
    ''' légende annonce mais que la substitution ignore.
    ''' </summary>
    Private Sub AfficherLaLegende()

        Dim lignes As New List(Of String)

        For Each jeton As ModeleNarrativeWU.Jeton In ModeleNarrativeWU.Jetons()
            lignes.Add($"{jeton.Nom}   {jeton.Description}")
        Next

        lblJetons.Text = String.Join(Environment.NewLine, lignes)
    End Sub

#End Region

#Region "Le modèle global"

    Private Sub txtModele_TextChanged(sender As Object, e As EventArgs) Handles txtModele.TextChanged

        If _enChargement Then Return

        AfficherLApercu()

        ' Les natures qui reprennent le modèle global suivent la saisie en direct : sans cela
        ' la grille afficherait l'ancien modèle, et la banque croirait devoir les corriger une
        ' à une.
        RafraichirLesNaturesSuivantLeGlobal()
    End Sub

    ''' <summary>
    ''' Montre la phrase telle qu'elle sortira, et la mesure.
    '''
    ''' L'APERÇU EST CALCULÉ PAR LE MÊME CODE QUE LA PIÈCE — ModeleNarrativeWU.Appliquer, et
    ''' non une imitation locale. Un aperçu qui ressemble au résultat sans en être le produit
    ''' est un aperçu qui mentira un jour, et c'est précisément le jour où la banque s'y sera
    ''' fiée.
    ''' </summary>
    Private Sub AfficherLApercu()

        Dim rendu As String = RenduPireCas(txtModele.Text)

        lblApercu.Text = rendu
        Mesurer(lblLongueur, rendu)

        lblPireCas.Text = DescriptionDuPireCas()

        ' UN MODÈLE SANS AUCUN REPÈRE N'EST PAS REFUSÉ, mais il est dit : toutes les écritures
        ' de tous les points de vente de toutes les journées porteraient le même texte, et le
        ' grand livre ne dirait plus ni qui ni quand.
        If Not ModeleNarrativeWU.EmploieUnRepere(txtModele.Text) Then
            lblPireCas.ForeColor = Drawing.Color.Firebrick
            lblPireCas.Text =
                "Ce modèle n'emploie aucun repère : toutes les écritures de tous les points " &
                "de vente et de toutes les journées porteront ce même texte. Le grand livre " &
                "ne dira plus ni quel point de vente ni quelle période."
            Return
        End If

        lblPireCas.ForeColor = Drawing.SystemColors.GrayText
    End Sub

    ''' <summary>Écrit la longueur rendue, et la passe au rouge si le core banking la refuserait.</summary>
    Private Sub Mesurer(etiquette As Label, rendu As String)

        Dim tient As Boolean = rendu.Length <= ConstantesWU.CB_NARRATIF_LONGUEUR_MAX

        etiquette.ForeColor = If(tient, Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
        etiquette.Text =
            $"{rendu.Length} caractères sur les {ConstantesWU.CB_NARRATIF_LONGUEUR_MAX} " &
            "que le core banking accepte." &
            If(tient, String.Empty, " Ce modèle ne peut pas être enregistré.")
    End Sub

    ''' <summary>
    ''' Dit sur quoi l'aperçu a été calculé. Sans cette phrase, l'agent verrait un nom
    ''' d'agence qu'il n'a pas demandé et croirait à une erreur.
    ''' </summary>
    Private Function DescriptionDuPireCas() As String

        If Not _referentielLu Then
            Return "Le référentiel des points de vente n'a pas pu être lu : l'aperçu ne porte " &
                   "donc sur aucun nom réel, et la longueur affichée est incomplète. Le " &
                   "contrôle sera refait au moment de produire le fichier core banking."
        End If

        If _pireDesignation.Length = 0 Then
            Return "Le référentiel ne contient encore aucun point de vente : l'aperçu ne porte " &
                   "sur aucun nom réel."
        End If

        Return $"Calculé sur le cas le plus long du référentiel — désignation « {_pireDesignation} », " &
               $"Account « {_pireAccount} », code agence « {_pireCodeAgence} » et une période " &
               $"à cheval sur deux années (« {_pirePeriode} »)."
    End Function

    Private Sub btnDefaut_Click(sender As Object, e As EventArgs) Handles btnDefaut.Click

        txtModele.Text = ModeleNarrativeWU.ModeleParDefaut
        txtModele.Focus()

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = "Modèle par défaut rétabli à l'écran — pas encore enregistré."
    End Sub

#End Region

#Region "Le mode"

    Private Sub rdoModeUnique_CheckedChanged(sender As Object, e As EventArgs) Handles rdoModeUnique.CheckedChanged
        AppliquerLeMode()
    End Sub

    Private Sub rdoModeParNature_CheckedChanged(sender As Object, e As EventArgs) Handles rdoModeParNature.CheckedChanged
        AppliquerLeMode()
    End Sub

    ''' <summary>
    ''' Ouvre ou ferme la grille des natures selon le mode choisi, et redit ce que le mode fait.
    '''
    ''' LA GRILLE RESTE VISIBLE EN MODE UNIQUE, mais en lecture seule : la banque doit pouvoir
    ''' REGARDER ce qu'elle abandonne avant de basculer. Un onglet qui disparaît ne se consulte
    ''' pas.
    '''
    ''' ELLE EST AUSSI EN LECTURE SEULE SANS LA TABLE DES NATURES : les libellés s'appliquent,
    ''' mais rien ne pourrait retenir une modification. Une case qu'on peut remplir et qui
    ''' perd la saisie à la fermeture est pire qu'une case grisée.
    ''' </summary>
    Private Sub AppliquerLeMode()

        Dim parNature As Boolean = rdoModeParNature.Checked
        Dim modifiable As Boolean = parNature AndAlso _naturesDisponibles

        dgvNatures.ReadOnly = Not modifiable
        dgvNatures.DefaultCellStyle.BackColor = If(modifiable, Drawing.SystemColors.Window,
                                                   Drawing.SystemColors.Control)

        lblAideNatures.Text =
            If(parNature,
               "Chaque ligne de la pièce portera le libellé de sa nature. Une case laissée " &
               "telle quelle garde son libellé historique : seules les natures que vous " &
               "modifiez sont conservées.",
               "Mode « un seul libellé » : ces libellés ne sont PAS appliqués à la pièce, qui " &
               "portera le modèle du premier onglet sur ses douze lignes. Ils sont conservés " &
               "et redeviennent actifs si vous revenez au mode par nature. L'écart d'arrondi " &
               "fait exception — il garde toujours son libellé propre, dans les deux modes.")

        If _enChargement Then Return
        AfficherLApercuDeLaNature()
    End Sub

#End Region

#Region "Les libellés par nature"

    Private Const COL_CODE As String = "Code"
    Private Const COL_NATURE As String = "Nature"
    Private Const COL_MODELE As String = "Modele"

    ''' <summary>
    ''' Remplit la grille avec le modèle EFFECTIF de chaque nature — celui qui serait posé si
    ''' le mode « par nature » était actif.
    '''
    ''' POURQUOI L'EFFECTIF, ET NON LA SEULE SAISIE. Une case vide pour la TVA signifierait
    ''' « suit son libellé historique », ce qui est vrai mais illisible : la banque ne verrait
    ''' pas ce que la ligne porterait, ni qu'elle peut le reprendre à son compte. Montrer
    ''' l'effectif dit la vérité, et l'enregistrement ne garde que ce qui s'en écarte.
    ''' </summary>
    Private Sub ChargerLesNatures()

        Dim narratives As NarrativesWU = NarrativeRepository.EnVigueur()

        Dim table As New DataTable("natures")
        table.Columns.Add(COL_CODE, GetType(String))
        table.Columns.Add(COL_NATURE, GetType(String))
        table.Columns.Add(COL_MODELE, GetType(String))

        _naturesChargees.Clear()

        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()

            Dim modele As String = narratives.Modele(nature)
            _naturesChargees(nature) = modele

            table.Rows.Add(NaturesMouvementWU.Code(nature),
                           NaturesMouvementWU.Intitule(nature),
                           modele)
        Next

        dgvNatures.DataSource = table
        HabillerLaGrilleDesNatures()
    End Sub

    Private Sub HabillerLaGrilleDesNatures()

        If Not dgvNatures.Columns.Contains(COL_CODE) Then Return

        dgvNatures.Columns(COL_CODE).Visible = False

        dgvNatures.Columns(COL_NATURE).HeaderText = "Nature du mouvement"
        dgvNatures.Columns(COL_NATURE).ReadOnly = True
        dgvNatures.Columns(COL_NATURE).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgvNatures.Columns(COL_NATURE).Width = 260

        dgvNatures.Columns(COL_MODELE).HeaderText = "Libellé de la ligne"
        dgvNatures.Columns(COL_MODELE).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        AppliquerLeMode()
    End Sub

    Private Sub dgvNatures_SelectionChanged(sender As Object, e As EventArgs) Handles dgvNatures.SelectionChanged
        AfficherLApercuDeLaNature()
    End Sub

    Private Sub dgvNatures_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvNatures.CellEndEdit
        AfficherLApercuDeLaNature()
    End Sub

    ''' <summary>
    ''' Mesure la nature sélectionnée, séparément des autres.
    '''
    ''' CHAQUE NATURE A SON PROPRE PIRE CAS. Mesurer le seul modèle global ne dirait rien d'un
    ''' libellé de TVA que la banque aurait rallongé : c'est lui qui ferait échouer la
    ''' production du fichier, et sur sa seule ligne.
    ''' </summary>
    Private Sub AfficherLApercuDeLaNature()

        Dim nature As NatureMouvementWU? = NatureSelectionnee()

        If Not nature.HasValue Then
            lblApercuNature.Text = String.Empty
            lblLongueurNature.Text = String.Empty
            Return
        End If

        Dim modele As String = ModeleSaisi(nature.Value)
        Dim rendu As String = RenduPireCas(modele)

        lblApercuNature.Text = $"{NaturesMouvementWU.Intitule(nature.Value)} : {rendu}"
        Mesurer(lblLongueurNature, rendu)
    End Sub

    ''' <summary>La nature de la ligne sélectionnée, ou Nothing si la grille est vide.</summary>
    Private Function NatureSelectionnee() As NatureMouvementWU?

        If dgvNatures.CurrentRow Is Nothing Then Return Nothing

        Dim cellule As DataGridViewCell = dgvNatures.CurrentRow.Cells(COL_CODE)
        If cellule Is Nothing OrElse cellule.Value Is Nothing Then Return Nothing

        Return NaturesMouvementWU.DepuisCode(Convert.ToString(cellule.Value))
    End Function

    ''' <summary>Le modèle actuellement saisi dans la grille pour cette nature.</summary>
    Private Function ModeleSaisi(nature As NatureMouvementWU) As String

        Dim recherche As String = NaturesMouvementWU.Code(nature)

        For Each ligne As DataGridViewRow In dgvNatures.Rows

            If ligne.IsNewRow Then Continue For

            Dim code As String = Convert.ToString(ligne.Cells(COL_CODE).Value)
            If Not String.Equals(code, recherche, StringComparison.OrdinalIgnoreCase) Then Continue For

            Return Convert.ToString(ligne.Cells(COL_MODELE).Value)
        Next

        Return String.Empty
    End Function

    ''' <summary>
    ''' Fait suivre la saisie du modèle global aux natures qui le reprennent, c'est-à-dire
    ''' celles dont la case montre encore le modèle chargé.
    '''
    ''' EN MODE « PAR NATURE », IL N'Y A RIEN À FAIRE SUIVRE : une nature non personnalisée y
    ''' montre son libellé historique, et non le modèle global. La frappe du premier onglet ne
    ''' concerne alors que le fichier core banking, et la grille n'a pas à bouger.
    '''
    ''' L'ÉCART D'ARRONDI N'EST JAMAIS TOUCHÉ : il ne suit pas le modèle global, et le faire
    ''' suivre ici le remplacerait par « LD WU ACTIVITE » dès la première frappe.
    ''' </summary>
    Private Sub RafraichirLesNaturesSuivantLeGlobal()

        If dgvNatures.DataSource Is Nothing Then Return
        If rdoModeParNature.Checked Then
            _modeleGlobalCharge = txtModele.Text
            Return
        End If

        Dim ancienGlobal As String = _modeleGlobalCharge
        Dim nouveauGlobal As String = txtModele.Text

        For Each ligne As DataGridViewRow In dgvNatures.Rows

            If ligne.IsNewRow Then Continue For

            Dim nature As NatureMouvementWU? =
                NaturesMouvementWU.DepuisCode(Convert.ToString(ligne.Cells(COL_CODE).Value))

            If Not nature.HasValue OrElse nature.Value = NatureMouvementWU.EcartArrondi Then Continue For

            ' Seules les cases restées sur l'ancien modèle global suivent : une nature que la
            ' banque a personnalisée n'est pas écrasée par une frappe dans l'onglet voisin.
            If Not String.Equals(Convert.ToString(ligne.Cells(COL_MODELE).Value),
                                 ancienGlobal, StringComparison.Ordinal) Then Continue For

            ligne.Cells(COL_MODELE).Value = nouveauGlobal
        Next

        ' L'ancien global devient le nouveau : sans cela, une deuxième frappe ne trouverait
        ' plus aucune case à faire suivre.
        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()
            If nature = NatureMouvementWU.EcartArrondi Then Continue For
            If Not _naturesChargees.ContainsKey(nature) Then Continue For
            If Not String.Equals(_naturesChargees(nature), ancienGlobal, StringComparison.Ordinal) Then Continue For
            _naturesChargees(nature) = nouveauGlobal
        Next

        _modeleGlobalCharge = nouveauGlobal

        AfficherLApercuDeLaNature()
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
            $"(les {LIMITE_JOURNAL} dernières). Ce journal ne peut être ni modifié ni effacé " &
            "depuis l'application, et aucun droit d'écriture n'est accordé dessus en base."
    End Sub

    ''' <summary>
    ''' Combien de modifications la grille charge. Un journal se consulte et ne se télécharge
    ''' pas : au bout de deux ans, tout charger mettrait l'écran plusieurs secondes à s'ouvrir
    ''' pour montrer trente lignes qu'on ne fera jamais défiler.
    ''' </summary>
    Private Const LIMITE_JOURNAL As Integer = 200

    Private Sub HabillerLaGrilleDuJournal()

        ' La grille se lie à une liste d'objets : ses colonnes portent les noms des propriétés,
        ' et celles qui n'intéressent pas le lecteur sont retirées plutôt que renommées.
        If dgvJournal.Columns.Contains("Cle") Then dgvJournal.Columns("Cle").Visible = False
        If dgvJournal.Columns.Contains("AncienneValeur") Then dgvJournal.Columns("AncienneValeur").Visible = False

        Renommer("DateModification", "Date", 130)
        Renommer("Intitule", "Ce qui a changé", 200)
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

        If largeur <= 0 Then
            dgvJournal.Columns(nom).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Return
        End If

        dgvJournal.Columns(nom).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgvJournal.Columns(nom).Width = largeur
    End Sub

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        ' LA CELLULE EN COURS DE SAISIE EST VALIDÉE D'ABORD. Sans cela, cliquer Enregistrer
        ' sans avoir quitté la case lirait l'ANCIENNE valeur : le libellé que la banque vient
        ' de taper serait perdu, et l'écran annoncerait un enregistrement réussi.
        dgvNatures.EndEdit()

        Dim modele As String = txtModele.Text.Trim()
        Dim messageErreur As String = String.Empty

        If Not ModeleNarrativeWU.Controler(modele, _pireDesignation, _pireAccount,
                                            _pireCodeAgence, _pirePeriode, messageErreur) Then
            Refuser("Le modèle global n'est pas enregistrable.", messageErreur)
            Return
        End If

        Dim parNature As Boolean = rdoModeParNature.Checked

        ' CHAQUE NATURE EST CONTRÔLÉE SÉPARÉMENT, et seulement si elle va servir. Contrôler des
        ' libellés que le mode unique n'appliquera pas empêcherait d'enregistrer un modèle
        ' global parfaitement valable à cause d'une case qui ne sort nulle part.
        If parNature AndAlso Not ControlerLesNatures(messageErreur) Then
            Refuser("Un libellé par nature n'est pas enregistrable.", messageErreur)
            Return
        End If

        If Not Confirmer(modele, parNature) Then Return

        Cursor = Cursors.WaitCursor
        Try
            If Not OptionsWU.Enregistrer(OptionsWU.CLE_NARRATIVE_MODELE, modele,
                                         OptionsWU.LIBELLE_NARRATIVE_MODELE, messageErreur) Then
                Refuser("Modèle global non enregistré.", messageErreur)
                Return
            End If

            Dim mode As ModeNarrativeWU =
                If(parNature, ModeNarrativeWU.ParNature, ModeNarrativeWU.ModeleUnique)

            If Not OptionsWU.Enregistrer(OptionsWU.CLE_NARRATIVE_MODE, NarrativesWU.CodeDeMode(mode),
                                         OptionsWU.LIBELLE_NARRATIVE_MODE, messageErreur) Then
                Refuser("Mode non enregistré.", messageErreur)
                Return
            End If

            If _naturesDisponibles AndAlso Not EnregistrerLesNatures(messageErreur) Then
                Refuser("Libellés par nature partiellement enregistrés.", messageErreur)
                Return
            End If

        Finally
            Cursor = Cursors.Default
        End Try

        UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                          "Narrative comptable : " & NarrativesWU.IntituleDeMode(
                                              If(parNature, ModeNarrativeWU.ParNature,
                                                 ModeNarrativeWU.ModeleUnique)))

        ' Tout est relu plutôt que supposé : le journal vient de gagner des lignes, et la
        ' grille des natures a pu perdre celles qui ont rejoint leur libellé historique.
        Charger()

        lblStatut.ForeColor = Drawing.Color.DarkGreen
        lblStatut.Text = "Paramétrage enregistré."
    End Sub

    ''' <summary>
    ''' Dit en une ligne que rien n'a été enregistré, et ouvre le détail.
    '''
    ''' Le paramètre ne s'appelle PAS « resume » : Visual Basic réserve ce mot à l'instruction
    ''' Resume Next, et une variable qui le porte est refusée (BC30183).
    ''' </summary>
    Private Sub Refuser(constat As String, detail As String)

        lblStatut.ForeColor = Drawing.Color.Firebrick
        lblStatut.Text = constat
        FrmDiagnostic.Afficher(Me, "Narrative comptable", detail)
    End Sub

    ''' <summary>
    ''' Contrôle les treize libellés, et nomme celui qui ne passe pas.
    '''
    ''' « Un modèle dépasse » ne serait pas une information : il y en a treize, et l'agent
    ''' chercherait lequel case par case.
    ''' </summary>
    Private Function ControlerLesNatures(ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()

            Dim modele As String = ModeleSaisi(nature).Trim()
            Dim detail As String = String.Empty

            If ModeleNarrativeWU.Controler(modele, _pireDesignation, _pireAccount,
                                           _pireCodeAgence, _pirePeriode, detail) Then Continue For

            messageErreur = $"Nature « {NaturesMouvementWU.Intitule(nature)} » :" &
                            Environment.NewLine & Environment.NewLine & detail
            Return False
        Next

        Return True
    End Function

    ''' <summary>
    ''' Enregistre les natures qui ont CHANGÉ, et seulement elles.
    '''
    ''' UNE NATURE REVENUE À SON REPLI EST EFFACÉE, pas enregistrée à l'identique : c'est ce
    ''' qui garde vivant ce repli, et ce qui fait qu'un futur changement le lui fera suivre au
    ''' lieu de la laisser sur l'ancienne phrase.
    '''
    ''' LE REPLI D'UNE NATURE DÉPEND DU MODE, exactement comme dans NarrativesWU.Modele : son
    ''' LIBELLÉ HISTORIQUE en mode « par nature », le MODÈLE GLOBAL en mode « un seul
    ''' libellé ». L'écart d'arrondi fait exception dans les deux modes — il ne suit jamais le
    ''' modèle global, qui le réduirait à « LD WU ACTIVITE ».
    '''
    ''' LES DEUX DOIVENT S'ACCORDER, sans quoi l'écran effacerait une ligne que la pièce
    ''' rendrait ensuite autrement : la banque verrait son libellé disparaître à la
    ''' réouverture de l'écran, ou pire, le verrait rester et la pièce en porter un autre.
    ''' </summary>
    Private Function EnregistrerLesNatures(ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Dim modeleGlobal As String = txtModele.Text.Trim()
        Dim parNature As Boolean = rdoModeParNature.Checked

        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()

            Dim saisi As String = ModeleSaisi(nature).Trim()

            Dim repli As String
            If nature = NatureMouvementWU.EcartArrondi Then
                repli = ConstantesWU.NARRATIVE_ECART_DEFAUT
            ElseIf parNature Then
                repli = NaturesMouvementWU.ModeleParDefaut(nature)
            Else
                repli = modeleGlobal
            End If

            Dim aEcrire As String = If(String.Equals(saisi, repli, StringComparison.Ordinal),
                                       String.Empty, saisi)

            Dim chargee As String = String.Empty
            If _naturesChargees.ContainsKey(nature) Then chargee = _naturesChargees(nature)

            ' Inchangée : ni écriture, ni transaction, ni ligne de journal.
            If String.Equals(saisi, chargee.Trim(), StringComparison.Ordinal) Then Continue For

            If Not NarrativeRepository.Enregistrer(nature, aEcrire, messageErreur) Then Return False
        Next

        Return True
    End Function

    ''' <summary>
    ''' Fait relire la phrase avant de l'enregistrer.
    '''
    ''' CE N'EST PAS UNE PRÉCAUTION DE FORME. Ce texte partira sur TOUTES les écritures de
    ''' TOUTES les journées à venir, dans le grand livre de la banque, et les pièces déjà
    ''' produites ne changeront pas — elles garderont l'ancien. Relire la phrase rendue, et
    ''' non le gabarit, est le seul moment où l'on voit ce qu'on décide.
    ''' </summary>
    Private Function Confirmer(modele As String, parNature As Boolean) As Boolean

        Dim mode As ModeNarrativeWU =
            If(parNature, ModeNarrativeWU.ParNature, ModeNarrativeWU.ModeleUnique)

        Dim detail As String

        If parNature Then
            detail = "Chaque ligne de la pièce portera le libellé de sa nature. Le fichier " &
                     "core banking, lui, portera celui-ci :" & Environment.NewLine &
                     Environment.NewLine & RenduPireCas(modele)
        Else
            detail = "Les lignes de la pièce et le fichier core banking porteront tous ce " &
                     "texte :" & Environment.NewLine & Environment.NewLine & RenduPireCas(modele)
        End If

        Return MessageBox.Show(Me,
            "Enregistrer ce paramétrage de narrative ?" & Environment.NewLine & Environment.NewLine &
            "Mode : " & NarrativesWU.IntituleDeMode(mode) & "." & Environment.NewLine & Environment.NewLine &
            detail & Environment.NewLine & Environment.NewLine &
            "Ces textes partent sur toutes les pièces à venir, et dans la colonne ADDLTEXT du " &
            "fichier core banking. Les pièces déjà produites gardent le leur : elles ne sont " &
            "pas réécrites." &
            If(ModeleNarrativeWU.EmploieUnRepere(modele), String.Empty,
               Environment.NewLine & Environment.NewLine &
               "ATTENTION : le modèle du point de vente n'emploie aucun repère. Le fichier " &
               "core banking portera le même texte sur toutes ses lignes, sans nommer le " &
               "point de vente ni la période."),
            "Confirmer la narrative", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
