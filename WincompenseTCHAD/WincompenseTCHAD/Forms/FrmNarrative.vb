Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' LA NARRATIVE COMPTABLE : l'écran où la banque écrit elle-même la phrase que portera chaque
''' ligne de ses pièces, et que son core banking recevra dans la colonne ADDLTEXT.
'''
''' POURQUOI CET ÉCRAN EXISTE
'''
''' Cette phrase a changé quatre fois en quatre livraisons. Chaque mot a coûté une
''' modification du code, une compilation, un commit, un pull et un redéploiement sur les
''' postes de la banque — pour un texte qui n'entre dans aucun calcul. Ce qui se lit dans le
''' grand livre appartient à la Direction Comptable ; ce qui s'y calcule reste au code.
'''
''' L'ÉCRAN NE SE CONTENTE PAS DE RECUEILLIR UNE SAISIE
'''
''' Il la MESURE. Le core banking refuse un narratif de plus de cent cinquante caractères, et
''' l'application refuse alors de produire le fichier — un refus qui tomberait le matin de la
''' compense, à l'heure où personne n'a le temps. L'aperçu est donc calculé sur le PIRE CAS
''' RÉEL : le point de vente dont le nom est le plus long dans le référentiel, l'Account le
''' plus long, le code agence le plus long, et la période la plus longue que l'application
''' sache écrire. Ce qui est validé ici passera en production.
'''
''' CE QU'IL NE PERMET PAS DE FAIRE
'''
''' Saisir un repère que l'application ne connaît pas : il partirait au grand livre avec ses
''' accolades. Saisir un modèle vide : douze écritures sans libellé partiraient sans que rien
''' n'avertisse. Et il n'ouvre pas la saisie des majuscules — elles sont posées par le code,
''' parce qu'une journée saisie en minuscules ne ressemblerait pas aux autres.
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

#End Region

#Region "Ouverture"

    Private Sub FrmNarrative_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not SessionWU.PeutGererLesComptesSystemes Then
            MessageBox.Show("La narrative comptable est réservée à l'administrateur." &
                            Environment.NewLine & Environment.NewLine &
                            "Ce texte part dans le grand livre de la banque, sur toutes les " &
                            "écritures de toutes les journées : il ne se change pas depuis le " &
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

        txtModele.MaxLength = ConstantesWU.NARRATIVE_MODELE_LONGUEUR_MAX

        AfficherLaLegende()
        ChargerLePireCas()

        ' La valeur est relue en base, et non prise dans le cache : on vient précisément la
        ' modifier, et quelqu'un d'autre a pu le faire avant.
        OptionsWU.Oublier()
        txtModele.Text = OptionsWU.NarrativeModele

        lblDerniereModification.Text = Historique()

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = String.Empty

        AfficherLApercu()
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

#Region "Aperçu"

    Private Sub txtModele_TextChanged(sender As Object, e As EventArgs) Handles txtModele.TextChanged
        AfficherLApercu()
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

        Dim rendu As String = ModeleNarrativeWU.Appliquer(txtModele.Text, _pireDesignation,
                                                           _pireAccount, _pireCodeAgence,
                                                           _pirePeriode)

        lblApercu.Text = rendu

        Dim tient As Boolean = rendu.Length <= ConstantesWU.CB_NARRATIF_LONGUEUR_MAX

        lblLongueur.ForeColor = If(tient, Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
        lblLongueur.Text =
            $"{rendu.Length} caractères sur les {ConstantesWU.CB_NARRATIF_LONGUEUR_MAX} " &
            "que le core banking accepte." &
            If(tient, String.Empty, " Ce modèle ne peut pas être enregistré.")

        lblPireCas.Text = DescriptionDuPireCas()

        ' UN MODÈLE SANS AUCUN REPÈRE N'EST PAS REFUSÉ, mais il est dit : toutes les
        ' écritures de tous les points de vente de toutes les journées porteraient le même
        ' texte, et le grand livre ne dirait plus ni qui ni quand.
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

#End Region

#Region "Enregistrement"

    Private Sub btnDefaut_Click(sender As Object, e As EventArgs) Handles btnDefaut.Click

        txtModele.Text = ModeleNarrativeWU.ModeleParDefaut
        txtModele.Focus()

        lblStatut.ForeColor = Drawing.SystemColors.GrayText
        lblStatut.Text = "Modèle par défaut rétabli à l'écran — pas encore enregistré."
    End Sub

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        Dim modele As String = txtModele.Text.Trim()
        Dim messageErreur As String = String.Empty

        If Not ModeleNarrativeWU.Controler(modele, _pireDesignation, _pireAccount,
                                            _pireCodeAgence, _pirePeriode, messageErreur) Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Modèle non enregistré."
            FrmDiagnostic.Afficher(Me, "Narrative comptable", messageErreur)
            Return
        End If

        If Not Confirmer(modele) Then Return

        Cursor = Cursors.WaitCursor
        Dim enregistre As Boolean = OptionsWU.Enregistrer(
            OptionsWU.CLE_NARRATIVE_MODELE, modele,
            OptionsWU.LIBELLE_NARRATIVE_MODELE, messageErreur)
        Cursor = Cursors.Default

        If Not enregistre Then
            lblStatut.ForeColor = Drawing.Color.Firebrick
            lblStatut.Text = "Modèle non enregistré."
            FrmDiagnostic.Afficher(Me, "Narrative comptable", messageErreur)
            Return
        End If

        UtilisateurRepository.Journaliser(SessionWU.Identifiant, True,
                                          "Narrative comptable : " & modele)

        txtModele.Text = modele
        lblDerniereModification.Text = Historique()

        lblStatut.ForeColor = Drawing.Color.DarkGreen
        lblStatut.Text = "Modèle enregistré."
    End Sub

    ''' <summary>
    ''' Fait relire la phrase avant de l'enregistrer.
    '''
    ''' CE N'EST PAS UNE PRÉCAUTION DE FORME. Ce texte partira sur TOUTES les écritures de
    ''' TOUTES les journées à venir, dans le grand livre de la banque, et les pièces déjà
    ''' produites ne changeront pas — elles garderont l'ancien. Relire la phrase rendue, et
    ''' non le gabarit, est le seul moment où l'on voit ce qu'on décide.
    ''' </summary>
    Private Function Confirmer(modele As String) As Boolean

        Dim exemple As String = ModeleNarrativeWU.Appliquer(modele, _pireDesignation,
                                                             _pireAccount, _pireCodeAgence,
                                                             _pirePeriode)

        Return MessageBox.Show(Me,
            "Enregistrer ce modèle de narrative ?" & Environment.NewLine & Environment.NewLine &
            modele & Environment.NewLine & Environment.NewLine &
            "Les écritures sortiront désormais ainsi :" & Environment.NewLine & Environment.NewLine &
            exemple & Environment.NewLine & Environment.NewLine &
            "Ce texte part sur toutes les lignes de toutes les pièces à venir, et dans la " &
            "colonne ADDLTEXT du fichier core banking. Les pièces déjà produites gardent le " &
            "leur : elles ne sont pas réécrites." &
            If(ModeleNarrativeWU.EmploieUnRepere(modele), String.Empty,
               Environment.NewLine & Environment.NewLine &
               "ATTENTION : ce modèle n'emploie aucun repère. Toutes les écritures porteront " &
               "le même texte, sans nommer le point de vente ni la période."),
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
