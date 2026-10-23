Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' Liste du réseau Western Union, destinée à l'impression.
'''
''' ÉCRAN DE RESTITUTION : il ne fait que lire. Aucun bouton n'y crée, ne modifie ni ne
''' supprime quoi que ce soit — c'est l'état qu'on sort pour une revue, un audit ou un
''' classement, et rien d'autre.
'''
''' CE QUI S'IMPRIME EST CE QUI S'AFFICHE. Les sections sont construites une seule fois par
''' ListeReseauService ; la grille les met à plat, le PDF et le classeur en font des blocs.
''' Les deux rendus ne peuvent donc pas diverger : ils partent du même objet.
'''
''' LA DATE DE CRÉATION N'EST PAS DANS LE RÉFÉRENTIEL. Ni T_Pdv_SA ni T_Pdv_EC ne la portent.
''' Elle est reconstituée depuis les demandes autorisées, et manque pour toute fiche créée
''' hors de ce circuit : l'état écrit alors « inconnue », et une période demandée écarte ces
''' fiches plutôt que de les inclure sans preuve.
''' </summary>
Public Class FrmListeReseau

    ''' <summary>Le réseau lu en base, conservé pour éviter une relecture par changement de filtre.</summary>
    Private _reseau As New List(Of LigneReseauWU)

    ''' <summary>Les groupes statistiques, pour les comptes et le taux affichés en titre de section.</summary>
    Private _groupes As New List(Of GroupeStatistiqueWU)

    ''' <summary>Les sections telles qu'affichées : c'est exactement ce qui part à l'impression.</summary>
    Private _sections As New List(Of SectionReseauWU)

    ''' <summary>Faux quand T_DemandeWU est absente : aucune date ne peut alors être retrouvée.</summary>
    Private _historiqueDisponible As Boolean = True

    Public Sub New()
        InitializeComponent()
        IconesWU.Habiller(Me)
    End Sub

#Region "Ouverture"

    Private Sub FrmListeReseau_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Du 1er janvier de l'année en cours à aujourd'hui : la période la plus souvent
        ' demandée est l'exercice courant. Les deux champs restent grisés tant que la case
        ' n'est pas cochée — proposer des dates n'est pas les appliquer.
        dtpDebut.Value = New Date(Date.Today.Year, 1, 1)
        dtpFin.Value = Date.Today

        AjusterLesFiltres()
        Charger()
    End Sub

    ''' <summary>
    ''' Lit le réseau et les groupes, puis affiche.
    '''
    ''' LES DEUX LECTURES SONT FAITES ICI ET PAS AILLEURS : changer de portée ou de période ne
    ''' relit rien. Un référentiel de points de vente tient en quelques dizaines de lignes, et
    ''' le filtrer en mémoire évite qu'un même état, affiché puis imprimé, repose sur deux
    ''' lectures séparées par une modification faite entre-temps.
    ''' </summary>
    Private Sub Charger()

        Cursor = Cursors.WaitCursor
        Try
            _historiqueDisponible = PdvRepository.HistoriqueDesCreationsDisponible()

            Dim messageErreur As String = String.Empty
            _reseau = PdvRepository.ListerLeReseau(messageErreur)

            If messageErreur.Length > 0 Then
                Annoncer(messageErreur, True)
                Return
            End If

            _groupes = PdvRepository.ListerGroupes(String.Empty, messageErreur)

            ' Les groupes ne servent qu'à enrichir les titres de section : s'ils manquent,
            ' l'état reste exact, les titres sont seulement plus courts.
            ChargerLaListeDesGroupes()
            Afficher()

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ChargerLaListeDesGroupes()

        Dim noms As New List(Of String)

        For Each groupe As GroupeStatistiqueWU In _groupes
            Dim nom As String = groupe.Nom.Trim()
            If nom.Length > 0 AndAlso Not noms.Contains(nom) Then noms.Add(nom)
        Next

        noms.Sort(Function(gauche, droite) String.Compare(gauche, droite,
                                                          StringComparison.OrdinalIgnoreCase))

        cboGroupe.DataSource = noms
        If noms.Count > 0 Then cboGroupe.SelectedIndex = 0

        ' Sans aucun groupe enregistré, demander « un groupe statistique » n'a pas de sens.
        rdoGroupe.Enabled = noms.Count > 0
    End Sub

#End Region

#Region "Filtres"

    Private Sub rdoTout_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTout.CheckedChanged
        AjusterLesFiltres()
    End Sub

    Private Sub chkPeriode_CheckedChanged(sender As Object, e As EventArgs) Handles chkPeriode.CheckedChanged
        AjusterLesFiltres()
    End Sub

    ''' <summary>Grise ce qui ne s'applique pas : un champ actif qu'on ignore est un piège.</summary>
    Private Sub AjusterLesFiltres()

        cboGroupe.Enabled = rdoGroupe.Checked
        dtpDebut.Enabled = chkPeriode.Checked
        dtpFin.Enabled = chkPeriode.Checked
        lblEt.Enabled = chkPeriode.Checked
    End Sub

    ''' <summary>Le groupe demandé, ou chaîne vide pour tout le réseau.</summary>
    Private Function GroupeChoisi() As String

        If Not rdoGroupe.Checked Then Return String.Empty
        Return Convert.ToString(cboGroupe.SelectedItem, Globalization.CultureInfo.InvariantCulture).Trim()
    End Function

    ''' <summary>Le début de période, ou Nothing si la case n'est pas cochée.</summary>
    Private Function DebutChoisi() As Date?
        If Not chkPeriode.Checked Then Return Nothing
        Return dtpDebut.Value.Date
    End Function

    Private Function FinChoisie() As Date?
        If Not chkPeriode.Checked Then Return Nothing
        Return dtpFin.Value.Date
    End Function

#End Region

#Region "Affichage"

    Private Sub btnAfficher_Click(sender As Object, e As EventArgs) Handles btnAfficher.Click
        Afficher()
    End Sub

    Private Sub Afficher()

        ' Une période à l'envers ne ramènerait rien, et l'écran annoncerait « aucun point de
        ' vente » là où c'est la saisie qui est fautive.
        If chkPeriode.Checked AndAlso dtpFin.Value.Date < dtpDebut.Value.Date Then
            Annoncer("La date de fin est antérieure à la date de début.", True)
            dgvReseau.DataSource = Nothing
            _sections = New List(Of SectionReseauWU)
            Return
        End If

        _sections = ListeReseauService.Construire(_reseau, _groupes, GroupeChoisi(),
                                                  DebutChoisi(), FinChoisie())

        dgvReseau.DataSource = ListeReseauService.TableAPlat(_sections)
        HabillerLaGrille()

        ' L'absence d'historique est dite ICI AUSSI, et pas seulement dans le document : celui
        ' qui lit « inconnue » sur toutes les lignes doit savoir pourquoi avant d'imprimer.
        Dim total As String = ListeReseauService.Total(_sections)

        If _historiqueDisponible Then
            Annoncer(total, False)
        Else
            Annoncer(total & " — les dates de création ne sont pas disponibles sur cette base : " &
                     "la table des demandes est absente.", True)
        End If
    End Sub

    ''' <summary>
    ''' Donne à chaque colonne la place qui lui revient.
    '''
    ''' La section prend le plus de largeur : c'est elle qui porte le groupe, l'effectif, les
    ''' comptes et le taux. Les codes, eux, tiennent en quelques caractères.
    ''' </summary>
    Private Sub HabillerLaGrille()

        dgvReseau.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Poids(ListeReseauService.COL_SECTION, 230)
        Poids(ListeReseauService.COL_ACCOUNT, 90)
        Poids(ListeReseauService.COL_DESIGNATION, 150)
        Poids(ListeReseauService.COL_CODE_AGENCE, 80)
        Poids(ListeReseauService.COL_CREE_LE, 80)
        Poids(ListeReseauService.COL_CREE_PAR, 110)
    End Sub

    Private Sub Poids(nomColonne As String, valeur As Integer)
        If Not dgvReseau.Columns.Contains(nomColonne) Then Return
        dgvReseau.Columns(nomColonne).FillWeight = valeur
    End Sub

    Private Sub Annoncer(texte As String, enAlerte As Boolean)

        lblTotal.ForeColor = If(enAlerte, Drawing.Color.Firebrick, Drawing.SystemColors.ControlText)
        lblTotal.Text = texte
    End Sub

#End Region

#Region "Impression et export"

    Private Sub btnImprimer_Click(sender As Object, e As EventArgs) Handles btnImprimer.Click

        If RienAImprimer() Then Return

        sfdEtat.Filter = "Document PDF (*.pdf)|*.pdf"
        sfdEtat.FileName = $"Liste-reseau-{Date.Today:yyyyMMdd}.pdf"
        sfdEtat.OverwritePrompt = True

        If sfdEtat.ShowDialog(Me) <> DialogResult.OK Then Return

        Produire(Sub(progression As ProgressionWU)
                     ExcelExportService.ExporterEnPdf(
                         ListeReseauService.TITRE_DOCUMENT, SousTitres(),
                         ListeReseauService.Blocs(_sections), ListeReseauService.NOM_FEUILLE,
                         sfdEtat.FileName, progression:=progression)
                 End Sub,
                 "Impression de la liste du réseau",
                 "PDF produit : " & IO.Path.GetFileName(sfdEtat.FileName) & " — il s'ouvre aussitôt.")
    End Sub

    Private Sub btnExporter_Click(sender As Object, e As EventArgs) Handles btnExporter.Click

        If RienAImprimer() Then Return

        sfdEtat.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
        sfdEtat.FileName = $"Liste-reseau-{Date.Today:yyyyMMdd}.xlsx"
        sfdEtat.OverwritePrompt = True

        If sfdEtat.ShowDialog(Me) <> DialogResult.OK Then Return

        Produire(Sub(progression As ProgressionWU)
                     ExcelExportService.ExporterEtOuvrir(
                         ListeReseauService.TITRE_DOCUMENT, SousTitres(),
                         ListeReseauService.Blocs(_sections), ListeReseauService.NOM_FEUILLE,
                         sfdEtat.FileName, progression)
                 End Sub,
                 "Export Excel de la liste du réseau",
                 "Classeur produit : " & IO.Path.GetFileName(sfdEtat.FileName))
    End Sub

    ''' <summary>
    ''' Exécute une production de document : fenêtre d'avancement, curseur, et récupération
    ''' d'une erreur sans faire tomber l'écran.
    '''
    ''' LES DEUX BOUTONS PARTAGENT CE CORPS parce que seul l'appel au service les distingue.
    ''' Dupliquer l'enveloppe aurait garanti qu'un correctif n'en touche qu'une moitié.
    ''' </summary>
    Private Sub Produire(travail As Action(Of ProgressionWU), titreAvancement As String,
                         messageReussite As String)

        Dim avancement As FrmProgression = Nothing

        Try
            Cursor = Cursors.WaitCursor
            avancement = FrmProgression.Ouvrir(Me, titreAvancement)

            travail(avancement.Progression)

            avancement.Fermer()
            Annoncer(messageReussite, False)

        Catch ex As InvalidOperationException
            If avancement IsNot Nothing Then avancement.Fermer()
            Echouer(ex.Message)

        Catch ex As Runtime.InteropServices.COMException
            If avancement IsNot Nothing Then avancement.Fermer()
            Echouer(ex.Message)

        Catch ex As IO.IOException
            If avancement IsNot Nothing Then avancement.Fermer()
            Echouer(ex.Message)

        Finally
            If avancement IsNot Nothing Then avancement.Fermer()
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Echouer(detail As String)

        Annoncer("Document non produit.", True)
        MessageBox.Show(Me, "Le document n'a pas pu être produit : " & detail,
                        "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>Vrai, et le dit, quand il n'y a rien à mettre sur le papier.</summary>
    Private Function RienAImprimer() As Boolean

        If _sections.Count > 0 Then Return False

        MessageBox.Show(Me, "Aucun point de vente ne correspond à ce qui est demandé : " &
                            "il n'y a rien à imprimer.",
                        "Liste vide", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Return True
    End Function

    Private Function SousTitres() As List(Of SousTitreExcel)
        Return ListeReseauService.SousTitres(GroupeChoisi(), DebutChoisi(), FinChoisie(),
                                             _sections, _historiqueDisponible)
    End Function

#End Region

#Region "Fermeture"

    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Close()
    End Sub

#End Region

End Class
