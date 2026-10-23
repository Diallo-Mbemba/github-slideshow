Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>
''' Compose l'état « Liste du réseau » : les sections, leur contenu, leur ordre et leurs
''' intitulés.
'''
''' POURQUOI UN SERVICE ET NON LE FORMULAIRE. Ce qui s'affiche et ce qui s'imprime doivent
''' être la même chose. Un découpage fait dans l'écran, puis refait pour l'export, finirait
''' par diverger au premier correctif porté d'un seul côté : l'écran dirait vingt-quatre
''' points de vente, le PDF vingt-trois, et personne ne saurait lequel croire. Ici, les
''' sections sont construites UNE fois ; l'écran les met à plat, l'export en fait des blocs.
'''
''' CE SERVICE NE LIT RIEN ET N'ÉCRIT RIEN. Il reçoit le réseau déjà lu et se contente de
''' l'ordonner : c'est ce qui permet de le faire tourner sans base.
''' </summary>
Public NotInheritable Class ListeReseauService

    Private Sub New()
    End Sub

    ''' <summary>Titre du document, en tête de l'état imprimé comme du classeur.</summary>
    Public Const TITRE_DOCUMENT As String = "LISTE DU RÉSEAU WESTERN UNION"

    ''' <summary>Nom de l'onglet du classeur Excel.</summary>
    Public Const NOM_FEUILLE As String = "Liste du réseau"

    ''' <summary>Intitulé de la section qui recueille les sous-agents sans groupe statistique.</summary>
    Public Const SANS_GROUPE As String = "SOUS-AGENTS SANS GROUPE STATISTIQUE"

    ''' <summary>Intitulé de la section des agences propres.</summary>
    Public Const AGENCES_PROPRES As String = "AGENCES PROPRES"

    ' Les en-têtes servent de noms de colonnes : la grille et le classeur les affichent tels
    ' quels, et aucune table de correspondance ne peut se désaccorder avec eux.
    Public Const COL_SECTION As String = "Section"
    Public Const COL_ACCOUNT As String = "Account"
    Public Const COL_DESIGNATION As String = "Désignation"
    Public Const COL_CODE_AGENCE As String = "Code agence"
    Public Const COL_CREE_LE As String = "Créé le"
    Public Const COL_CREE_PAR As String = "Créé par"

#Region "Construction des sections"

    ''' <summary>
    ''' Range le réseau en sections, dans l'ordre où l'état les présente.
    '''
    ''' L'ORDRE EST CELUI DE LA LECTURE, et il est voulu : les groupes d'abord, par ordre
    ''' alphabétique, puis les sous-agents que personne n'a rattachés — ce sont ceux qui
    ''' n'héritent d'aucun compte ni d'aucun taux, et qu'il faut voir — puis les agences
    ''' propres, qui ne relèvent d'aucun groupe.
    '''
    ''' UNE SECTION VIDE N'EST PAS RENDUE. Un groupe dont tous les sous-agents sont écartés
    ''' par la période n'a rien à dire ; son titre seul, suivi d'un tableau vide, ferait
    ''' croire à une perte de données.
    ''' </summary>
    ''' <param name="reseau">Le réseau lu en base, sous-agents et agences mêlés.</param>
    ''' <param name="groupes">Les groupes statistiques, pour leurs comptes et leur taux.</param>
    ''' <param name="groupeChoisi">Un groupe précis, ou chaîne vide pour tout le réseau.</param>
    ''' <param name="debut">Début de la période de création, ou Nothing si elle n'est pas posée.</param>
    ''' <param name="fin">Fin de la période, bornes comprises.</param>
    Public Shared Function Construire(reseau As List(Of LigneReseauWU),
                                      groupes As List(Of GroupeStatistiqueWU),
                                      groupeChoisi As String,
                                      debut As Date?, fin As Date?) As List(Of SectionReseauWU)

        Dim sections As New List(Of SectionReseauWU)
        If reseau Is Nothing Then Return sections

        Dim retenues As List(Of LigneReseauWU) = DansLaPeriode(reseau, debut, fin)
        Dim unSeulGroupe As String = If(groupeChoisi, String.Empty).Trim()
        Dim toutLeReseau As Boolean = unSeulGroupe.Length = 0

        For Each nom As String In NomsDesGroupes(retenues, groupes)

            If Not toutLeReseau AndAlso
               Not String.Equals(nom, unSeulGroupe, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim lignes As List(Of LigneReseauWU) =
                Trier(retenues.FindAll(Function(ligne) ligne.EstSousAgent AndAlso
                                                       MemeGroupe(ligne.Groupe, nom)))

            If lignes.Count = 0 Then Continue For

            sections.Add(New SectionReseauWU() With {
                .Titre = TitreDeGroupe(nom, lignes.Count, groupes),
                .Lignes = lignes
            })
        Next

        ' Les deux dernières sections n'existent que pour l'état complet : demander un groupe
        ' précis, c'est demander ce groupe et rien d'autre.
        If Not toutLeReseau Then Return sections

        Dim sansGroupe As List(Of LigneReseauWU) =
            Trier(retenues.FindAll(Function(ligne) ligne.EstSousAgent AndAlso
                                                   ligne.Groupe.Trim().Length = 0))

        If sansGroupe.Count > 0 Then
            sections.Add(New SectionReseauWU() With {
                .Titre = $"{SANS_GROUPE} – {sansGroupe.Count} sous-agent(s)",
                .Lignes = sansGroupe
            })
        End If

        Dim agences As List(Of LigneReseauWU) =
            Trier(retenues.FindAll(Function(ligne) Not ligne.EstSousAgent))

        If agences.Count > 0 Then
            sections.Add(New SectionReseauWU() With {
                .Titre = $"{AGENCES_PROPRES} – {agences.Count} agence(s)",
                .Lignes = agences,
                .EstSectionDAgences = True
            })
        End If

        Return sections
    End Function

    ''' <summary>
    ''' Applique la période de création.
    '''
    ''' UNE FICHE SANS DATE EST ÉCARTÉE dès qu'une période est demandée, et ce n'est pas un
    ''' oubli : on ne peut pas affirmer qu'elle a été créée dans l'intervalle. L'écarter est
    ''' le seul choix honnête — la garder reviendrait à l'inclure sans preuve, et c'est
    ''' justement ce que la période sert à trancher.
    ''' </summary>
    Private Shared Function DansLaPeriode(reseau As List(Of LigneReseauWU),
                                          debut As Date?, fin As Date?) As List(Of LigneReseauWU)

        If Not debut.HasValue OrElse Not fin.HasValue Then Return New List(Of LigneReseauWU)(reseau)

        Dim premier As Date = debut.Value.Date
        Dim dernier As Date = fin.Value.Date

        Return reseau.FindAll(Function(ligne) ligne.DateConnue AndAlso
                                              ligne.DateCreation.Value.Date >= premier AndAlso
                                              ligne.DateCreation.Value.Date <= dernier)
    End Function

    ''' <summary>
    ''' Trie une section : par date de création, puis par Account.
    '''
    ''' LES DATES INCONNUES PASSENT EN TÊTE. Ce sont les fiches les plus anciennes — celles
    ''' d'avant le circuit des demandes — et surtout celles dont l'état ne peut rien dire :
    ''' les reléguer en fin de tableau reviendrait à les cacher.
    ''' </summary>
    Private Shared Function Trier(lignes As List(Of LigneReseauWU)) As List(Of LigneReseauWU)

        lignes.Sort(Function(gauche, droite)

                        If gauche.DateConnue <> droite.DateConnue Then
                            Return If(gauche.DateConnue, 1, -1)
                        End If

                        If gauche.DateConnue Then
                            Dim ecart As Integer = Date.Compare(gauche.DateCreation.Value,
                                                                droite.DateCreation.Value)
                            If ecart <> 0 Then Return ecart
                        End If

                        Return String.Compare(gauche.Account, droite.Account,
                                              StringComparison.OrdinalIgnoreCase)
                    End Function)

        Return lignes
    End Function

    ''' <summary>
    ''' Les noms de groupes à parcourir, par ordre alphabétique.
    '''
    ''' Ils viennent des DEUX sources : la table des groupes, et les groupes portés par les
    ''' sous-agents eux-mêmes. Un groupe que la migration n'a pas repris n'existe que dans
    ''' T_Pdv_SA, et ses sous-agents doivent tout de même s'imprimer.
    ''' </summary>
    Private Shared Function NomsDesGroupes(reseau As List(Of LigneReseauWU),
                                           groupes As List(Of GroupeStatistiqueWU)) As List(Of String)

        Dim noms As New List(Of String)

        If groupes IsNot Nothing Then
            For Each groupe As GroupeStatistiqueWU In groupes
                Ajouter(noms, groupe.Nom)
            Next
        End If

        For Each ligne As LigneReseauWU In reseau
            If ligne.EstSousAgent Then Ajouter(noms, ligne.Groupe)
        Next

        noms.Sort(Function(gauche, droite) String.Compare(gauche, droite,
                                                          StringComparison.OrdinalIgnoreCase))
        Return noms
    End Function

    Private Shared Sub Ajouter(noms As List(Of String), nom As String)

        Dim net As String = If(nom, String.Empty).Trim()
        If net.Length = 0 Then Return
        If noms.Exists(Function(connu) MemeGroupe(connu, net)) Then Return

        noms.Add(net)
    End Sub

    Private Shared Function MemeGroupe(gauche As String, droite As String) As Boolean
        Return String.Equals(If(gauche, String.Empty).Trim(), If(droite, String.Empty).Trim(),
                             StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Compose le titre d'une section de groupe : son nom, son effectif, ses comptes, son taux.
    '''
    ''' CE QUI MANQUE N'EST PAS ÉCRIT. Un groupe hérité, que la migration n'a pas repris, n'a
    ''' ni compte ni taux enregistrés : son titre s'arrête à l'effectif plutôt que d'annoncer
    ''' « activité  » suivi d'un blanc.
    '''
    ''' LE COMPTE DE RÈGLEMENT N'Y FIGURE PAS : T_GroupeStatistique n'en porte pas. Un groupe
    ''' a un compte d'activité, un compte de commission et un taux, et c'est tout.
    ''' </summary>
    Private Shared Function TitreDeGroupe(nom As String, effectif As Integer,
                                          groupes As List(Of GroupeStatistiqueWU)) As String

        Dim parties As New List(Of String) From {
            $"GROUPE {nom}",
            $"{effectif} sous-agent(s)"
        }

        Dim groupe As GroupeStatistiqueWU = Nothing
        If groupes IsNot Nothing Then
            groupe = groupes.Find(Function(candidat) MemeGroupe(candidat.Nom, nom))
        End If

        If groupe IsNot Nothing Then

            If groupe.CompteActivite.Trim().Length > 0 Then
                parties.Add($"activité {groupe.CompteActivite.Trim()}")
            End If

            If groupe.CompteCommission.Trim().Length > 0 Then
                parties.Add($"commission {groupe.CompteCommission.Trim()}")
            End If

            If groupe.Taux > 0D Then
                parties.Add("taux " & groupe.Taux.ToString("0 %", CultureInfo.InvariantCulture))
            End If
        End If

        Return String.Join(" – ", parties)
    End Function

#End Region

#Region "Mise en tableau"

    ''' <summary>
    ''' Les sections mises à plat pour la grille : une colonne Section rappelle, sur chaque
    ''' ligne, de quelle section elle relève.
    '''
    ''' À L'ÉCRAN SEULEMENT. Une grille ne sait pas porter de titres intercalaires ; à
    ''' l'impression, la section redevient un titre au-dessus de son tableau, et la colonne
    ''' disparaît — répéter soixante caractères sur chaque ligne d'un document imprimé
    ''' gâcherait une colonne entière pour une information déjà écrite en tête.
    ''' </summary>
    Public Shared Function TableAPlat(sections As List(Of SectionReseauWU)) As DataTable

        Dim table As New DataTable("reseau")
        table.Columns.Add(COL_SECTION, GetType(String))
        AjouterLesColonnes(table)

        If sections Is Nothing Then Return table

        For Each section As SectionReseauWU In sections
            For Each ligne As LigneReseauWU In section.Lignes
                table.Rows.Add(section.Titre, ligne.Account, ligne.Designation,
                               ligne.CodeAgence, ligne.DateLisible, ligne.AuteurLisible)
            Next
        Next

        Return table
    End Function

    ''' <summary>Le tableau d'une seule section, tel qu'il s'imprime sous son titre.</summary>
    Public Shared Function TableDeSection(section As SectionReseauWU) As DataTable

        Dim table As New DataTable("section")
        AjouterLesColonnes(table)

        If section Is Nothing Then Return table

        For Each ligne As LigneReseauWU In section.Lignes
            table.Rows.Add(ligne.Account, ligne.Designation, ligne.CodeAgence,
                           ligne.DateLisible, ligne.AuteurLisible)
        Next

        Return table
    End Function

    Private Shared Sub AjouterLesColonnes(table As DataTable)

        table.Columns.Add(COL_ACCOUNT, GetType(String))
        table.Columns.Add(COL_DESIGNATION, GetType(String))
        table.Columns.Add(COL_CODE_AGENCE, GetType(String))

        ' La date part en TEXTE, déjà mise en forme : c'est le seul moyen d'écrire « inconnue »
        ' dans la même colonne qu'une date, et l'ordre du tableau est déjà fixé par le tri.
        table.Columns.Add(COL_CREE_LE, GetType(String))
        table.Columns.Add(COL_CREE_PAR, GetType(String))
    End Sub

    ''' <summary>
    ''' Les blocs du document : une section, un bloc, son titre au-dessus de son tableau.
    ''' </summary>
    Public Shared Function Blocs(sections As List(Of SectionReseauWU)) As List(Of BlocExcel)

        Dim liste As New List(Of BlocExcel)
        If sections Is Nothing Then Return liste

        For Each section As SectionReseauWU In sections
            liste.Add(New BlocExcel(section.Titre, TableDeSection(section)) With {
                .TitreEnExergue = True
            })
        Next

        Return liste
    End Function

#End Region

#Region "En-tête du document"

    ''' <summary>
    ''' Les lignes d'information sous le titre : ce qui a été demandé, et par qui.
    '''
    ''' L'ABSENCE D'HISTORIQUE EST DITE, et en exergue. Sur une base antérieure au circuit des
    ''' demandes, toutes les dates sont inconnues : sans cette ligne, l'état laisserait croire
    ''' à un réseau dont personne n'a jamais rien créé.
    ''' </summary>
    Public Shared Function SousTitres(groupeChoisi As String, debut As Date?, fin As Date?,
                                      sections As List(Of SectionReseauWU),
                                      historiqueDisponible As Boolean) As List(Of SousTitreExcel)

        Dim unSeulGroupe As String = If(groupeChoisi, String.Empty).Trim()
        Dim toutLeReseau As Boolean = unSeulGroupe.Length = 0

        Dim entete As New List(Of SousTitreExcel) From {
            New SousTitreExcel(If(toutLeReseau,
                                  "Tout le réseau : groupes statistiques, sous-agents et agences propres",
                                  $"Groupe statistique : {unSeulGroupe}"), Not toutLeReseau),
            New SousTitreExcel(LibellePeriode(debut, fin)),
            New SousTitreExcel(Total(sections))
        }

        If Not historiqueDisponible Then
            entete.Add(New SousTitreExcel(
                "Les dates de création ne sont pas disponibles sur cette base : la table des " &
                "demandes est absente.", True))
        End If

        entete.Add(New SousTitreExcel(
            "Édité le " &
            Date.Now.ToString("dd/MM/yyyy 'à' HH:mm", CultureInfo.InvariantCulture) &
            $" par {SessionWU.Auteur}"))

        Return entete
    End Function

    ''' <summary>La période telle qu'elle s'écrit en tête de l'état.</summary>
    Public Shared Function LibellePeriode(debut As Date?, fin As Date?) As String

        If Not debut.HasValue OrElse Not fin.HasValue Then
            Return "Toutes les dates de création, connues ou non"
        End If

        Return "Points de vente créés entre le " &
               debut.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) & " et le " &
               fin.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' Le dénombrement, sous la grille comme en tête du document.
    '''
    ''' Ce qui n'existe pas n'est pas compté à zéro : l'état d'un seul groupe ne contient
    ''' aucune agence propre, et annoncer « 0 agence(s) propre(s) » ferait chercher ce qui
    ''' n'a jamais été demandé.
    ''' </summary>
    Public Shared Function Total(sections As List(Of SectionReseauWU)) As String

        Dim sousAgents As Integer = 0
        Dim agences As Integer = 0

        If sections IsNot Nothing Then
            For Each section As SectionReseauWU In sections
                If section.EstSectionDAgences Then
                    agences += section.Nombre
                Else
                    sousAgents += section.Nombre
                End If
            Next
        End If

        Dim nombre As Integer = sousAgents + agences
        If nombre = 0 Then Return "Aucun point de vente à imprimer."

        Dim details As New List(Of String)
        If sousAgents > 0 Then details.Add($"{sousAgents} sous-agent(s)")
        If agences > 0 Then details.Add($"{agences} agence(s) propre(s)")

        Return $"{nombre} point(s) de vente : {String.Join(", ", details)}."
    End Function

#End Region

End Class
