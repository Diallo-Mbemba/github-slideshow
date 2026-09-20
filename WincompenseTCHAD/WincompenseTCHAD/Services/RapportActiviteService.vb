Option Strict On
Option Explicit On

Imports System.Data

''' <summary>
''' Construction des quatre états du rapport d'activité sur une période, à partir des lignes
''' d'historique (une par journée et par point de vente).
'''
''' Les quatre états sont bâtis sur LA MÊME liste de lignes, agrégée différemment : leurs
''' totaux sont donc nécessairement identiques, ce qui fait du rapprochement entre pages un
''' contrôle de cohérence et non une coïncidence.
'''
''' Aucun accès à la base ici, aucune mise en forme : uniquement des DataTable, que le
''' formulaire affiche et que ExcelExportService exporte.
''' </summary>
Public NotInheritable Class RapportActiviteService

    Private Sub New()
    End Sub

    ''' <summary>Libellé des lignes de total, reconnaissable pour la mise en évidence.</summary>
    Public Const LIBELLE_TOTAL As String = "TOTAL"

    ''' <summary>Libellé des sous-totaux par nature de point de vente.</summary>
    Public Const LIBELLE_SOUS_TOTAL As String = "Sous-total"

    ''' <summary>
    ''' Niveau de chaque ligne de l'état par point de vente. Il pilote le pliage à l'écran et
    ''' la mise en forme, et permet d'exclure le détail d'un export sans reconstruire l'état.
    ''' </summary>
    Public Const NIVEAU_CATEGORIE As Integer = 0
    Public Const NIVEAU_POINT_DE_VENTE As Integer = 1
    Public Const NIVEAU_TRANSACTION As Integer = 2
    Public Const NIVEAU_SOUS_TOTAL As Integer = 3
    Public Const NIVEAU_TOTAL As Integer = 4

    ''' <summary>Signes de la colonne de pliage.</summary>
    Public Const SIGNE_PLIE As String = "+"
    Public Const SIGNE_DEPLIE As String = "-"

    ''' <summary>Libellé des points de vente sans groupe statistique.</summary>
    Public Const SANS_GROUPE As String = "(sans groupe statistique)"

    ''' <summary>Agence de rattachement inconnue : l'Account n'est pas au référentiel des agences.</summary>
    Public Const SANS_AGENCE As String = "(agence non rattachée)"

#Region "Page 1 — Synthèse"

    ''' <summary>
    ''' Synthèse de la période : volumes, montants d'envoi et de paiement, commissions et taxes.
    ''' Présentée en rubriques plutôt qu'en colonnes : c'est une page de lecture, destinée à
    ''' être posée sur un bureau.
    ''' </summary>
    Public Shared Function ConstruireSynthese(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("Synthese")
        table.Columns.Add("Rubrique", GetType(String))
        table.Columns.Add("Valeur", GetType(Decimal))

        Dim cumul As LigneHistoriqueWU = Cumuler(lignes)
        Dim jours As Integer = CompterJours(lignes)
        Dim pointsDeVente As Integer = CompterPointsDeVente(lignes)

        table.Rows.Add("VOLUMES", DBNull.Value)
        table.Rows.Add("    Journées comptabilisées", CDec(jours))
        table.Rows.Add("    Points de vente actifs", CDec(pointsDeVente))
        table.Rows.Add("    Transactions d'envoi", CDec(cumul.NombreEnvois))
        table.Rows.Add("    Transactions de paiement", CDec(cumul.NombrePaiements))
        table.Rows.Add("    Transactions annulées", CDec(cumul.NombreAnnulations))

        table.Rows.Add("ENVOIS (FCFA)", DBNull.Value)
        table.Rows.Add("    Principal envoyé", cumul.PrincipalEnvoi)
        table.Rows.Add("    Frais d'envoi", cumul.ChargeEnvoi)
        table.Rows.Add("    Taxes perçues sur envoi", cumul.Taxes)

        table.Rows.Add("PAIEMENTS (FCFA)", DBNull.Value)
        table.Rows.Add("    Principal payé", cumul.PrincipalPaye)

        table.Rows.Add("COMMISSIONS (FCFA)", DBNull.Value)
        table.Rows.Add("    Commission Envoi", cumul.CommissionEnvoi)
        table.Rows.Add("    Commission Paiement", cumul.CommissionPaiement)
        table.Rows.Add("    Commission Transfert", cumul.CommissionTransfert)
        table.Rows.Add("    TOTAL COMMISSIONS", cumul.TotalCommissions)

        ' Répartition sous-agents / agences propres : ce que la banque réalise par son propre
        ' réseau ne se confond pas avec ce qu'elle réalise par ses sous-agents.
        table.Rows.Add("RÉPARTITION PAR NATURE DE POINT DE VENTE", DBNull.Value)

        For Each typePdv As String In New String() {TYPE_SOUS_AGENT, TYPE_AGENCE, TYPE_NON_PARAMETRE}

            Dim duType As New List(Of LigneHistoriqueWU)
            Dim comptes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
                If String.Equals(LibelleType(ligne.TypePdv), typePdv, StringComparison.Ordinal) Then
                    duType.Add(ligne)
                    comptes.Add(ligne.Account)
                End If
            Next

            ' Une nature absente de la période n'a pas à encombrer l'état d'une ligne à zéro.
            If duType.Count = 0 Then Continue For

            Dim cumulType As LigneHistoriqueWU = CumulerLignes(duType)
            table.Rows.Add($"    {typePdv} — points de vente", CDec(comptes.Count))
            table.Rows.Add($"    {typePdv} — principal envoyé", cumulType.PrincipalEnvoi)
            table.Rows.Add($"    {typePdv} — principal payé", cumulType.PrincipalPaye)
            table.Rows.Add($"    {typePdv} — commissions", cumulType.TotalCommissions)
        Next

        table.Rows.Add("TAXES (FCFA)", DBNull.Value)
        table.Rows.Add("    TVA collectée", cumul.TVA)
        table.Rows.Add("    TTA sur envoi", cumul.TTAEnvoi)
        table.Rows.Add("    TTA sur réception", cumul.TTAReception)
        table.Rows.Add("    Taxe sur envoi", cumul.TaxeEnvoi)
        table.Rows.Add("    TOTAL TAXES", cumul.TotalTaxes)

        Return table
    End Function

#End Region

#Region "Page 2 — Jour par jour"

    ''' <summary>
    ''' Une ligne par journée comptabilisée, cumul en bas. C'est la page qui fait ressortir
    ''' une journée anormale.
    ''' </summary>
    Public Shared Function ConstruireParJour(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As DataTable = CreerTableDetail("ParJour", "Date", GetType(String))

        Dim cumuls As New SortedDictionary(Of Date, LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If Not cumuls.ContainsKey(ligne.DateActivite) Then
                cumuls(ligne.DateActivite) = New LigneHistoriqueWU()
            End If
            cumuls(ligne.DateActivite).Cumuler(ligne)
        Next

        For Each jour As Date In cumuls.Keys
            AjouterLigneDetail(table, jour.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture), cumuls(jour))
        Next

        AjouterLigneTotal(table, lignes)
        Return table
    End Function

#End Region

#Region "Page 3 — Par point de vente"

    ''' <summary>
    ''' Une ligne par point de vente, cumulée sur toute la période, classée par principal
    ''' envoyé décroissant : les points de vente qui pèsent apparaissent en tête.
    ''' </summary>
    ''' <param name="agencesConnues">
    ''' Agences propres du paramétrage (T_Pdv_EC). Celles qui n'ont eu AUCUNE activité sur la
    ''' période y figurent tout de même, à zéro : une agence qui n'a rien fait est une
    ''' information de gestion, et l'omettre reviendrait à la rendre invisible au moment même
    ''' où elle mérite d'être regardée. Facultatif.
    ''' </param>
    ''' <param name="transactions">
    ''' Détail des transactions, rattaché à chaque point de vente : dérouler un Account fait
    ''' apparaître ses envois et ses paiements, chacun avec son MTCN. Facultatif.
    ''' </param>
    Public Shared Function ConstruireParPointDeVente(lignes As List(Of LigneHistoriqueWU),
                                                     Optional agencesConnues As List(Of PointDeVenteEC) = Nothing,
                                                     Optional transactions As List(Of TransactionWU) = Nothing) As DataTable

        Dim table As New DataTable("ParPointDeVente")

        ' Colonne de pliage : "+" quand le point de vente a du détail, vide sinon.
        table.Columns.Add("Deroule", GetType(String))
        table.Columns.Add("Account", GetType(String))
        table.Columns.Add("Designation", GetType(String))
        table.Columns.Add("Groupe", GetType(String))
        table.Columns.Add("MTCN", GetType(String))
        table.Columns.Add("Statut", GetType(String))
        AjouterColonnesChiffrees(table)

        ' Niveau de la ligne, masqué à l'affichage : il pilote le pliage et l'export.
        table.Columns.Add("Niveau", GetType(Integer))

        ' Transactions regroupées par Account, triées par date puis par MTCN.
        Dim detailParAccount As Dictionary(Of String, List(Of TransactionWU)) = RegrouperTransactions(transactions)

        Dim cumuls As New Dictionary(Of String, LigneHistoriqueWU)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)

            If Not cumuls.ContainsKey(ligne.Account) Then
                cumuls(ligne.Account) = New LigneHistoriqueWU() With {.Account = ligne.Account}
            End If

            Dim cumul As LigneHistoriqueWU = cumuls(ligne.Account)

            ' Les lignes arrivant triées par date, la dernière valeur rencontrée est la plus
            ' récente : c'est l'identification à jour du point de vente qui est retenue.
            cumul.Designation = ligne.Designation
            cumul.GroupeStatistique = ligne.GroupeStatistique
            cumul.TypePdv = ligne.TypePdv
            cumul.Cumuler(ligne)
        Next

        ' Complément : les agences propres restées sans activité sur la période.
        If agencesConnues IsNot Nothing Then
            For Each agence As PointDeVenteEC In agencesConnues

                If agence Is Nothing OrElse String.IsNullOrWhiteSpace(agence.CodeSite) Then Continue For
                If cumuls.ContainsKey(agence.CodeSite) Then Continue For

                cumuls(agence.CodeSite) = New LigneHistoriqueWU() With {
                    .Account = agence.CodeSite,
                    .Designation = agence.Designation,
                    .TypePdv = "EC"
                }
            Next
        End If

        ' Sous-agents et agences propres sont restitués SÉPARÉMENT, chacun avec son sous-total :
        ' ce que la banque réalise par son propre réseau ne se confond pas avec ce qu'elle
        ' réalise par ses sous-agents.
        For Each typePdv As String In New String() {TYPE_SOUS_AGENT, TYPE_AGENCE, TYPE_NON_PARAMETRE}

            Dim duType As New List(Of LigneHistoriqueWU)

            For Each cumul As LigneHistoriqueWU In cumuls.Values
                If String.Equals(LibelleType(cumul.TypePdv), typePdv, StringComparison.Ordinal) Then
                    duType.Add(cumul)
                End If
            Next

            If duType.Count = 0 Then Continue For

            ' Les points de vente qui pèsent apparaissent en tête de leur catégorie ; ceux
            ' restés sans activité, tous à zéro, se rangent alors par Account plutôt que dans
            ' un ordre arbitraire.
            duType.Sort(Function(x, y)
                            Dim parPoids As Integer = y.PrincipalEnvoi.CompareTo(x.PrincipalEnvoi)
                            If parPoids <> 0 Then Return parPoids
                            Return String.Compare(x.Account, y.Account, StringComparison.OrdinalIgnoreCase)
                        End Function)

            Dim entete As DataRow = table.NewRow()
            entete("Deroule") = String.Empty
            entete("Account") = typePdv
            entete("Niveau") = NIVEAU_CATEGORIE
            table.Rows.Add(entete)

            For Each cumul As LigneHistoriqueWU In duType

                Dim operations As List(Of TransactionWU) = Nothing
                detailParAccount.TryGetValue(cumul.Account, operations)
                Dim aDuDetail As Boolean = operations IsNot Nothing AndAlso operations.Count > 0

                Dim enregistrement As DataRow = table.NewRow()
                enregistrement("Deroule") = If(aDuDetail, SIGNE_PLIE, String.Empty)
                enregistrement("Account") = cumul.Account
                enregistrement("Designation") = cumul.Designation
                enregistrement("Groupe") = LibelleGroupe(cumul.GroupeStatistique)
                enregistrement("MTCN") = String.Empty
                enregistrement("Statut") = String.Empty
                RemplirColonnesChiffrees(enregistrement, cumul)
                enregistrement("Niveau") = NIVEAU_POINT_DE_VENTE
                table.Rows.Add(enregistrement)

                If aDuDetail Then
                    AjouterLignesTransactions(table, operations)
                End If
            Next

            Dim sansActivite As Integer = 0
            For Each cumul As LigneHistoriqueWU In duType
                If cumul.NombreEnvois = 0 AndAlso cumul.NombrePaiements = 0 AndAlso
                   cumul.NombreAnnulations = 0 Then sansActivite += 1
            Next

            Dim sousTotal As DataRow = table.NewRow()
            sousTotal("Deroule") = String.Empty
            sousTotal("Account") = LIBELLE_SOUS_TOTAL
            sousTotal("Designation") = $"{typePdv} — {duType.Count} point(s) de vente" &
                                       If(sansActivite > 0, $", dont {sansActivite} sans activité", String.Empty)
            RemplirColonnesChiffrees(sousTotal, CumulerLignes(duType))
            sousTotal("Niveau") = NIVEAU_SOUS_TOTAL
            table.Rows.Add(sousTotal)
        Next

        Dim total As DataRow = table.NewRow()
        total("Deroule") = String.Empty
        total("Account") = LIBELLE_TOTAL
        total("Designation") = $"{cumuls.Count} point(s) de vente"
        RemplirColonnesChiffrees(total, Cumuler(lignes))
        total("Niveau") = NIVEAU_TOTAL
        table.Rows.Add(total)

        Return table
    End Function

    ''' <summary>
    ''' Écrit le détail des transactions d'un point de vente, une ligne chacune.
    '''
    ''' Une transaction est UNE unité de volume : son montant alimente la colonne d'envoi ou de
    ''' paiement selon son sens, et son compteur la colonne correspondante. Les cumuls du point
    ''' de vente au-dessus sont donc exactement la somme des lignes en dessous — le déroulé
    ''' justifie le total, il ne se contente pas de l'accompagner.
    ''' </summary>
    Private Shared Sub AjouterLignesTransactions(table As DataTable, operations As List(Of TransactionWU))

        For Each operation As TransactionWU In operations

            Dim enregistrement As DataRow = table.NewRow()
            enregistrement("Deroule") = String.Empty

            ' L'indentation place visuellement la transaction sous son point de vente.
            enregistrement("Account") = "    " & operation.DateActivite.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture)
            enregistrement("Designation") = operation.Sens
            enregistrement("Groupe") = String.Empty
            enregistrement("MTCN") = operation.MTCN
            enregistrement("Statut") = LibelleStatut(operation.Statut)

            If operation.EstAnnulee Then
                ' Une transaction annulée ne pèse dans aucun montant : elle ne compte que comme
                ' annulation, exactement comme dans l'agrégat de la journée.
                enregistrement("NbAnnulations") = 1
            ElseIf String.Equals(operation.Sens, TransactionWU.SENS_ENVOI, StringComparison.Ordinal) Then
                enregistrement("NbEnvois") = 1
                enregistrement("PrincipalEnvoi") = operation.Montant
            Else
                enregistrement("NbPaiements") = 1
                enregistrement("PrincipalPaye") = operation.Montant
            End If

            enregistrement("Niveau") = NIVEAU_TRANSACTION
            table.Rows.Add(enregistrement)
        Next
    End Sub

    ''' <summary>Regroupe les transactions par Account, triées par date puis par MTCN.</summary>
    Private Shared Function RegrouperTransactions(transactions As List(Of TransactionWU)) As Dictionary(Of String, List(Of TransactionWU))

        Dim parAccount As New Dictionary(Of String, List(Of TransactionWU))(StringComparer.OrdinalIgnoreCase)

        If transactions Is Nothing Then Return parAccount

        For Each operation As TransactionWU In transactions

            If operation Is Nothing OrElse String.IsNullOrWhiteSpace(operation.Account) Then Continue For

            If Not parAccount.ContainsKey(operation.Account) Then
                parAccount(operation.Account) = New List(Of TransactionWU)
            End If

            parAccount(operation.Account).Add(operation)
        Next

        For Each operations As List(Of TransactionWU) In parAccount.Values
            operations.Sort(Function(x, y)
                                Dim parDate As Integer = x.DateActivite.CompareTo(y.DateActivite)
                                If parDate <> 0 Then Return parDate
                                Return String.Compare(x.MTCN, y.MTCN, StringComparison.OrdinalIgnoreCase)
                            End Function)
        Next

        Return parAccount
    End Function

#End Region

#Region "Page 4 — Par groupe statistique"

    ''' <summary>
    ''' Une ligne par groupe statistique, cumulée sur la période. Les points de vente sans
    ''' groupe forment une ligne distincte plutôt que d'être omis.
    ''' </summary>
    Public Shared Function ConstruireParGroupe(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("ParGroupe")
        table.Columns.Add("Groupe", GetType(String))
        table.Columns.Add("PointsDeVente", GetType(Integer))
        AjouterColonnesChiffrees(table)

        Dim cumuls As New SortedDictionary(Of String, LigneHistoriqueWU)(StringComparer.OrdinalIgnoreCase)
        Dim comptesParGroupe As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)

            Dim groupe As String = LibelleGroupe(ligne.GroupeStatistique)

            If Not cumuls.ContainsKey(groupe) Then
                cumuls(groupe) = New LigneHistoriqueWU()
                comptesParGroupe(groupe) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            End If

            cumuls(groupe).Cumuler(ligne)
            comptesParGroupe(groupe).Add(ligne.Account)
        Next

        For Each groupe As String In cumuls.Keys
            Dim enregistrement As DataRow = table.NewRow()
            enregistrement("Groupe") = groupe
            enregistrement("PointsDeVente") = comptesParGroupe(groupe).Count
            RemplirColonnesChiffrees(enregistrement, cumuls(groupe))
            table.Rows.Add(enregistrement)
        Next

        Dim total As DataRow = table.NewRow()
        total("Groupe") = LIBELLE_TOTAL
        total("PointsDeVente") = CompterPointsDeVente(lignes)
        RemplirColonnesChiffrees(total, Cumuler(lignes))
        table.Rows.Add(total)

        Return table
    End Function

#End Region

#Region "Types de point de vente et filtrage"

    ''' <summary>Libellés des trois natures de point de vente, dans l'ordre où elles sont restituées.</summary>
    Public Const TYPE_SOUS_AGENT As String = "SOUS-AGENTS"
    Public Const TYPE_AGENCE As String = "AGENCES PROPRES"
    Public Const TYPE_NON_PARAMETRE As String = "NON PARAMÉTRÉS"

    ''' <summary>
    ''' Nature du point de vente, telle qu'enregistrée le jour de la comptabilisation.
    ''' Les Accounts non paramétrés forment une catégorie à part : les ranger avec les agences
    ''' propres reviendrait à affirmer qu'ils en sont, ce que précisément on ignore.
    ''' </summary>
    Public Shared Function LibelleType(typePdv As String) As String

        Select Case If(typePdv, String.Empty).Trim().ToUpperInvariant()
            Case "SA" : Return TYPE_SOUS_AGENT
            Case "EC" : Return TYPE_AGENCE
            Case Else : Return TYPE_NON_PARAMETRE
        End Select
    End Function

    ''' <summary>
    ''' Restreint les lignes à un groupe statistique. Un libellé vide retourne la liste
    ''' entière : c'est le cas « tous les groupes ».
    ''' </summary>
    Public Shared Function Filtrer(lignes As List(Of LigneHistoriqueWU), groupe As String) As List(Of LigneHistoriqueWU)

        Dim toutes As List(Of LigneHistoriqueWU) = SansNothing(lignes)

        If String.IsNullOrWhiteSpace(groupe) Then Return toutes

        Dim retenues As New List(Of LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In toutes
            If String.Equals(LibelleGroupe(ligne.GroupeStatistique), groupe, StringComparison.OrdinalIgnoreCase) Then
                retenues.Add(ligne)
            End If
        Next

        Return retenues
    End Function

    ''' <summary>Groupes statistiques présents dans la période, triés. Lus dans l'historique lui-même.</summary>
    ''' <remarks>
    ''' Et non dans T_GroupeStatistique : un groupe supprimé depuis reste présent dans l'historique,
    ''' et doit continuer d'être consultable.
    ''' </remarks>
    Public Shared Function ListerGroupesPresents(lignes As List(Of LigneHistoriqueWU)) As List(Of String)

        Dim groupes As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            groupes.Add(LibelleGroupe(ligne.GroupeStatistique))
        Next

        Return New List(Of String)(groupes)
    End Function

#End Region

#Region "Page — Évolution des commissions"

    ''' <summary>
    ''' Évolution des commissions jour après jour : le détail des trois commissions, leur total,
    ''' la variation par rapport à la journée précédente et le cumul depuis le début de la période.
    '''
    ''' La variation est exprimée en pourcentage du jour précédent. Elle est laissée VIDE pour la
    ''' première journée et lorsque la veille est à zéro : une variation depuis zéro n'a pas de
    ''' sens, et afficher 100 % ou l'infini induirait en erreur.
    ''' </summary>
    Public Shared Function ConstruireEvolutionCommissions(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("EvolutionCommissions")
        table.Columns.Add("Date", GetType(String))
        table.Columns.Add("CommissionEnvoi", GetType(Decimal))
        table.Columns.Add("CommissionPaiement", GetType(Decimal))
        table.Columns.Add("CommissionTransfert", GetType(Decimal))
        table.Columns.Add("TotalCommissions", GetType(Decimal))
        table.Columns.Add("Variation", GetType(Decimal))
        table.Columns.Add("Cumul", GetType(Decimal))

        Dim cumuls As New SortedDictionary(Of Date, LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If Not cumuls.ContainsKey(ligne.DateActivite) Then
                cumuls(ligne.DateActivite) = New LigneHistoriqueWU()
            End If
            cumuls(ligne.DateActivite).Cumuler(ligne)
        Next

        Dim cumulPeriode As Decimal = 0D
        Dim veille As Decimal = -1D ' -1 : aucune journée précédente

        For Each jour As Date In cumuls.Keys

            Dim cumulJour As LigneHistoriqueWU = cumuls(jour)
            Dim totalJour As Decimal = cumulJour.TotalCommissions
            cumulPeriode += totalJour

            Dim enregistrement As DataRow = table.NewRow()
            enregistrement("Date") = jour.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture)
            enregistrement("CommissionEnvoi") = cumulJour.CommissionEnvoi
            enregistrement("CommissionPaiement") = cumulJour.CommissionPaiement
            enregistrement("CommissionTransfert") = cumulJour.CommissionTransfert
            enregistrement("TotalCommissions") = totalJour
            enregistrement("Cumul") = cumulPeriode

            If veille > 0D Then
                enregistrement("Variation") = (totalJour - veille) / veille
            Else
                enregistrement("Variation") = DBNull.Value
            End If

            table.Rows.Add(enregistrement)
            veille = totalJour
        Next

        ' Ligne de total : la variation et le cumul n'y ont pas de sens, laissés vides.
        If cumuls.Count > 0 Then
            Dim total As LigneHistoriqueWU = Cumuler(lignes)
            Dim ligneTotal As DataRow = table.NewRow()
            ligneTotal("Date") = LIBELLE_TOTAL
            ligneTotal("CommissionEnvoi") = total.CommissionEnvoi
            ligneTotal("CommissionPaiement") = total.CommissionPaiement
            ligneTotal("CommissionTransfert") = total.CommissionTransfert
            ligneTotal("TotalCommissions") = total.TotalCommissions
            ligneTotal("Variation") = DBNull.Value
            ligneTotal("Cumul") = DBNull.Value
            table.Rows.Add(ligneTotal)
        End If

        Return table
    End Function

    ''' <summary>
    ''' Évolution des taxes jour après jour, détaillée par taxe.
    ''' 
    ''' Le jumeau de l'évolution des commissions, et pour la même raison : ce que la banque
    ''' REVERSE à l'État se suit comme ce qu'elle GAGNE. Une déclaration fiscale se prépare sur
    ''' des totaux par nature de taxe, pas sur une colonne « total taxes » dont on ne sait plus
    ''' ce qu'elle recouvre.
    ''' 
    ''' Les quatre taxes n'ont ni la même assiette ni le même destinataire :
    ''' 
    '''   TVA collectée        19,25 % des frais d'envoi
    '''   TTA sur envoi         0,2 % du principal envoyé
    '''   TTA sur paiement      0,2 % du principal payé — SOUS-AGENTS SEULEMENT, une agence
    '''                         propre n'en doit pas ; la colonne est donc structurellement plus
    '''                         basse que celle des envois, et ce n'est pas une anomalie
    '''   Impôts et taxe envoi  25 % du solde de taxes
    ''' 
    ''' La variation suit la même règle que pour les commissions : laissée vide pour la première
    ''' journée et lorsque la veille est à zéro — une variation depuis zéro n'a pas de sens.
    ''' </summary>
#Region "Ce que la banque a gardé"

    ''' <summary>Libellé de la population, tel qu'il apparaît en colonne.</summary>
    Public Const POPULATION_SOUS_AGENTS As String = "Sous-agents"
    Public Const POPULATION_AGENCES As String = "Agences propres"

    ''' <summary>
    ''' Répartition de la commission encaissée par la banque, par nature et par population.
    '''
    ''' La banque tire sa commission de deux sources qui n'obéissent pas à la même règle :
    ''' sur un sous-agent elle garde ce qui n'est pas rétrocédé, sur une agence propre elle
    ''' garde tout. Le total des deux est ce qu'elle a réellement encaissé — le chiffre que
    ''' l'onglet « Évolution des commissions » ne donne PAS, puisqu'il montre la commission
    ''' générée par l'activité, rétrocession comprise.
    '''
    ''' La dernière colonne rappelle ce qui est parti aux sous-agents : ce n'est pas un
    ''' produit de la banque, mais l'afficher permet de vérifier d'un coup d'œil que les
    ''' deux parts font bien la commission totale.
    ''' </summary>
    Public Shared Function ConstruireCommissionsBanque(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("CommissionsBanque")
        table.Columns.Add("Nature", GetType(String))
        table.Columns.Add("SousAgents", GetType(Decimal))
        table.Columns.Add("AgencesPropres", GetType(Decimal))
        table.Columns.Add("TotalBanque", GetType(Decimal))
        table.Columns.Add("Retrocede", GetType(Decimal))
        table.Columns.Add("CommissionTotale", GetType(Decimal))

        Dim sousAgents As LigneHistoriqueWU = CumulerPopulation(lignes, True)
        Dim agences As LigneHistoriqueWU = CumulerPopulation(lignes, False)

        AjouterNature(table, "Commission sur envoi",
                      sousAgents.CommissionEnvoiBanque, agences.CommissionEnvoiBanque,
                      sousAgents.CommissionEnvoi + agences.CommissionEnvoi)

        AjouterNature(table, "Commission sur paiement",
                      sousAgents.CommissionPaiementBanque, agences.CommissionPaiementBanque,
                      sousAgents.CommissionPaiement + agences.CommissionPaiement)

        AjouterNature(table, "Commission sur transfert",
                      sousAgents.CommissionTransfertBanque, agences.CommissionTransfertBanque,
                      sousAgents.CommissionTransfert + agences.CommissionTransfert)

        AjouterNature(table, "TOTAL",
                      sousAgents.TotalPartBanque, agences.TotalPartBanque,
                      sousAgents.TotalCommissions + agences.TotalCommissions)

        Return table
    End Function

    Private Shared Sub AjouterNature(table As DataTable, nature As String,
                                     partSousAgents As Decimal, partAgences As Decimal,
                                     commissionTotale As Decimal)

        Dim totalBanque As Decimal = partSousAgents + partAgences

        table.Rows.Add(nature, partSousAgents, partAgences, totalBanque,
                       commissionTotale - totalBanque, commissionTotale)
    End Sub

    ''' <summary>
    ''' Cumule les lignes d'une population. Un sous-agent est reconnu à son type, conservé
    ''' tel qu'il était le jour de la comptabilisation : un point de vente qui aurait changé
    ''' de nature depuis ne réécrit pas le passé.
    ''' </summary>
    Private Shared Function CumulerPopulation(lignes As List(Of LigneHistoriqueWU),
                                              sousAgents As Boolean) As LigneHistoriqueWU

        Dim cumul As New LigneHistoriqueWU()

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If ligne.EstSousAgent <> sousAgents Then Continue For
            cumul.Cumuler(ligne)
        Next

        Return cumul
    End Function

    ''' <summary>
    ''' Ce que la banque a gardé, jour par jour.
    '''
    ''' La colonne « Documentée » ne décore pas : une journée comptabilisée avant que la
    ''' répartition ne soit conservée porte un zéro côté sous-agents qui n'est pas un
    ''' chiffre mais une absence. L'écran doit pouvoir le dire.
    ''' </summary>
    Public Shared Function ConstruireCommissionsBanqueParJour(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("CommissionsBanqueParJour")
        table.Columns.Add("Date", GetType(String))
        table.Columns.Add("SousAgents", GetType(Decimal))
        table.Columns.Add("AgencesPropres", GetType(Decimal))
        table.Columns.Add("TotalBanque", GetType(Decimal))
        table.Columns.Add("Retrocede", GetType(Decimal))
        table.Columns.Add("CommissionTotale", GetType(Decimal))
        table.Columns.Add("Documentee", GetType(String))

        Dim parJour As New SortedDictionary(Of Date, LigneHistoriqueWU)()
        Dim parJourSA As New SortedDictionary(Of Date, LigneHistoriqueWU)()

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)

            If Not parJour.ContainsKey(ligne.DateActivite) Then
                parJour(ligne.DateActivite) = New LigneHistoriqueWU()
                parJourSA(ligne.DateActivite) = New LigneHistoriqueWU()
            End If

            parJour(ligne.DateActivite).Cumuler(ligne)
            If ligne.EstSousAgent Then parJourSA(ligne.DateActivite).Cumuler(ligne)
        Next

        For Each jour As Date In parJour.Keys

            Dim total As LigneHistoriqueWU = parJour(jour)
            Dim sousAgents As LigneHistoriqueWU = parJourSA(jour)

            table.Rows.Add(jour.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture),
                           sousAgents.TotalPartBanque,
                           total.TotalPartBanque - sousAgents.TotalPartBanque,
                           total.TotalPartBanque,
                           total.TotalPartSousAgent,
                           total.TotalCommissions,
                           If(total.RepartitionConnue, "oui", "NON"))
        Next

        Return table
    End Function

    ''' <summary>
    ''' La première journée dont la répartition est connue, ou Nothing s'il n'y en a aucune.
    ''' Sert à écrire « la répartition commence au … » plutôt qu'à laisser croire à un zéro.
    ''' </summary>
    Public Shared Function PremierJourDocumente(lignes As List(Of LigneHistoriqueWU)) As Date?

        Dim premier As Date? = Nothing

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If Not ligne.RepartitionConnue Then Continue For
            If Not premier.HasValue OrElse ligne.DateActivite < premier.Value Then
                premier = ligne.DateActivite
            End If
        Next

        Return premier
    End Function

    ''' <summary>Les journées dont la répartition n'est pas connue, les plus anciennes d'abord.</summary>
    Public Shared Function JourneesNonDocumentees(lignes As List(Of LigneHistoriqueWU)) As List(Of Date)

        Dim journees As New List(Of Date)()

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If ligne.RepartitionConnue Then Continue For
            If Not journees.Contains(ligne.DateActivite) Then journees.Add(ligne.DateActivite)
        Next

        journees.Sort()
        Return journees
    End Function

    ''' <summary>Ce que la banque a gardé sur toute la période.</summary>
    Public Shared Function TotalPartBanque(lignes As List(Of LigneHistoriqueWU)) As Decimal

        Dim total As Decimal = 0D

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            total += ligne.TotalPartBanque
        Next

        Return total
    End Function

#End Region

    Public Shared Function ConstruireEvolutionTaxes(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("EvolutionTaxes")
        table.Columns.Add("Date", GetType(String))
        table.Columns.Add("TVA", GetType(Decimal))
        table.Columns.Add("TTAEnvoi", GetType(Decimal))
        table.Columns.Add("TTAReception", GetType(Decimal))
        table.Columns.Add("TaxeEnvoi", GetType(Decimal))
        table.Columns.Add("TotalTaxes", GetType(Decimal))
        table.Columns.Add("Variation", GetType(Decimal))
        table.Columns.Add("Cumul", GetType(Decimal))

        Dim cumuls As New SortedDictionary(Of Date, LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            If Not cumuls.ContainsKey(ligne.DateActivite) Then
                cumuls(ligne.DateActivite) = New LigneHistoriqueWU()
            End If
            cumuls(ligne.DateActivite).Cumuler(ligne)
        Next

        Dim cumulPeriode As Decimal = 0D
        Dim veille As Decimal = -1D ' -1 : aucune journée précédente

        For Each jour As Date In cumuls.Keys

            Dim cumulJour As LigneHistoriqueWU = cumuls(jour)
            Dim totalJour As Decimal = cumulJour.TotalTaxes
            cumulPeriode += totalJour

            Dim enregistrement As DataRow = table.NewRow()
            enregistrement("Date") = jour.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture)
            enregistrement("TVA") = cumulJour.TVA
            enregistrement("TTAEnvoi") = cumulJour.TTAEnvoi
            enregistrement("TTAReception") = cumulJour.TTAReception
            enregistrement("TaxeEnvoi") = cumulJour.TaxeEnvoi
            enregistrement("TotalTaxes") = totalJour
            enregistrement("Cumul") = cumulPeriode

            If veille > 0D Then
                enregistrement("Variation") = (totalJour - veille) / veille
            Else
                enregistrement("Variation") = DBNull.Value
            End If

            table.Rows.Add(enregistrement)
            veille = totalJour
        Next

        ' Ligne de total : la variation et le cumul n'y ont pas de sens, laissés vides.
        If cumuls.Count > 0 Then
            Dim total As LigneHistoriqueWU = Cumuler(lignes)
            Dim ligneTotal As DataRow = table.NewRow()
            ligneTotal("Date") = LIBELLE_TOTAL
            ligneTotal("TVA") = total.TVA
            ligneTotal("TTAEnvoi") = total.TTAEnvoi
            ligneTotal("TTAReception") = total.TTAReception
            ligneTotal("TaxeEnvoi") = total.TaxeEnvoi
            ligneTotal("TotalTaxes") = total.TotalTaxes
            ligneTotal("Variation") = DBNull.Value
            ligneTotal("Cumul") = DBNull.Value
            table.Rows.Add(ligneTotal)
        End If

        Return table
    End Function

    ''' <summary>
    ''' Restreint les lignes à la population visée.
    ''' 
    ''' Les Accounts NON PARAMÉTRÉS suivent les sous-agents, et non les agences : un Account
    ''' inconnu n'est pas une agence propre, il n'est rien encore. C'est dans l'écran des
    ''' sous-agents qu'on ira le créer, autant qu'il y soit visible.
    ''' </summary>
    Public Shared Function FiltrerParPortee(lignes As List(Of LigneHistoriqueWU),
                                            portee As PorteeRapportWU) As List(Of LigneHistoriqueWU)

        Dim retenues As New List(Of LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)

            Dim estAgence As Boolean = String.Equals(ligne.TypePdv, "EC", StringComparison.OrdinalIgnoreCase)

            If (portee = PorteeRapportWU.AgencesPropres) = estAgence Then retenues.Add(ligne)
        Next

        Return retenues
    End Function

    ''' <summary>Libellé d'une agence dans les listes et les filtres : « 001 — AGENCE SIEGE ».</summary>
    Public Shared Function LibelleAgence(code As String, designation As String) As String

        Dim codeNet As String = If(code, String.Empty).Trim()
        Dim nomNet As String = If(designation, String.Empty).Trim()

        If codeNet.Length = 0 AndAlso nomNet.Length = 0 Then Return SANS_AGENCE
        If codeNet.Length = 0 Then Return nomNet
        If nomNet.Length = 0 Then Return codeNet

        Return codeNet & " — " & nomNet
    End Function

    ''' <summary>
    ''' Rattache chaque Account à son agence, d'après le référentiel COURANT.
    ''' 
    ''' L'historique ne porte pas le code agence : il garde l'Account, sa désignation, son groupe
    ''' et son type. Le rattachement se fait donc au paramétrage d'aujourd'hui — un Account qui
    ''' changerait d'agence emporterait tout son passé avec lui.
    ''' 
    ''' C'est acceptable ICI et ne l'était pas pour la pièce comptable : un rapport de gestion
    ''' n'est pas un justificatif, et le rattachement d'un Account à son agence ne bouge
    ''' quasiment jamais. Le jour où cela deviendrait gênant, une colonne CodeAgence dans
    ''' T_HistoriqueWU figerait le passé.
    ''' </summary>
    Private Shared Function AgenceParAccount(agencesConnues As List(Of PointDeVenteEC)) As Dictionary(Of String, String)

        Dim rattachement As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        If agencesConnues Is Nothing Then Return rattachement

        For Each agence As PointDeVenteEC In agencesConnues

            If agence Is Nothing OrElse String.IsNullOrWhiteSpace(agence.CodeSite) Then Continue For

            rattachement(agence.CodeSite.Trim()) =
                LibelleAgence(agence.CodeAgenceVoyager, agence.Designation)
        Next

        Return rattachement
    End Function

    ''' <summary>Les agences présentes sur la période, pour le filtre de l'écran.</summary>
    Public Shared Function ListerAgencesPresentes(lignes As List(Of LigneHistoriqueWU),
                                                  agencesConnues As List(Of PointDeVenteEC)) As List(Of String)

        Dim rattachement As Dictionary(Of String, String) = AgenceParAccount(agencesConnues)
        Dim agences As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            agences.Add(AgenceDe(ligne.Account, rattachement))
        Next

        ' Les agences sans la moindre activité doivent tout de même pouvoir être choisies :
        ' constater qu'une agence n'a rien fait est un résultat, pas une absence de résultat.
        If agencesConnues IsNot Nothing Then
            For Each agence As PointDeVenteEC In agencesConnues
                If agence Is Nothing Then Continue For
                agences.Add(LibelleAgence(agence.CodeAgenceVoyager, agence.Designation))
            Next
        End If

        Return New List(Of String)(agences)
    End Function

    ''' <summary>
    ''' Performance du réseau propre : les agences classées, et sous chacune ses Accounts.
    ''' 
    ''' DEUX NIVEAUX, PARCE QU'UNE AGENCE A PLUSIEURS ACCOUNTS. Dans T_Pdv_EC, l'Account est la
    ''' clé primaire mais le code agence ne porte AUCUNE contrainte d'unicité : un même guichet
    ''' peut tenir plusieurs points Western Union. Un classement à un seul niveau mélangerait
    ''' donc des guichets et des comptes, et la banque comparerait des agences à des fractions
    ''' d'agences.
    ''' 
    ''' LE CLASSEMENT SE FAIT SUR LES COMMISSIONS, décroissantes : c'est ce que la banque gagne,
    ''' et c'est la seule colonne qui réponde à « quelle agence rapporte le plus ». Les volumes
    ''' et les montants restent affichés à côté — une agence peut faire du volume sans marge, et
    ''' c'est précisément ce qu'un classement doit laisser voir.
    ''' 
    ''' LA PART est celle de l'agence dans les commissions de TOUT le réseau propre. Sans elle,
    ''' une liste de performance n'est qu'une liste : on voit qui est en tête, pas de combien.
    ''' 
    ''' LES AGENCES SANS ACTIVITÉ FIGURENT, à zéro. Constater qu'une agence n'a rien fait de la
    ''' période est un résultat, pas une absence de résultat — et c'est même celui qu'on cherche.
    ''' </summary>
    Public Shared Function ConstruirePerformanceAgences(lignes As List(Of LigneHistoriqueWU),
                                                        agencesConnues As List(Of PointDeVenteEC)) As DataTable

        Dim table As New DataTable("PerformanceAgences")
        table.Columns.Add("Rang", GetType(String))
        table.Columns.Add("Agence", GetType(String))
        table.Columns.Add("Account", GetType(String))
        table.Columns.Add("Designation", GetType(String))
        AjouterColonnesChiffrees(table)
        table.Columns.Add("Part", GetType(Decimal))
        table.Columns.Add("Niveau", GetType(Integer))

        Dim rattachement As Dictionary(Of String, String) = AgenceParAccount(agencesConnues)

        ' Cumul par Account, puis regroupement par agence : deux passes, parce que le
        ' rattachement d'un Account à son agence vient du référentiel et non de l'historique.
        Dim parAccount As New Dictionary(Of String, LigneHistoriqueWU)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)

            If Not parAccount.ContainsKey(ligne.Account) Then
                parAccount(ligne.Account) = New LigneHistoriqueWU() With {.Account = ligne.Account}
            End If

            Dim cumul As LigneHistoriqueWU = parAccount(ligne.Account)
            cumul.Designation = ligne.Designation
            cumul.TypePdv = ligne.TypePdv
            cumul.Cumuler(ligne)
        Next

        ' Les Accounts du référentiel restés sans activité : ils doivent figurer à zéro.
        If agencesConnues IsNot Nothing Then
            For Each agence As PointDeVenteEC In agencesConnues

                If agence Is Nothing OrElse String.IsNullOrWhiteSpace(agence.CodeSite) Then Continue For
                If parAccount.ContainsKey(agence.CodeSite) Then Continue For

                parAccount(agence.CodeSite) = New LigneHistoriqueWU() With {
                    .Account = agence.CodeSite,
                    .Designation = agence.Designation,
                    .TypePdv = "EC"
                }
            Next
        End If

        Dim parAgence As New Dictionary(Of String, List(Of LigneHistoriqueWU))(StringComparer.OrdinalIgnoreCase)

        For Each cumul As LigneHistoriqueWU In parAccount.Values

            Dim agence As String = AgenceDe(cumul.Account, rattachement)

            If Not parAgence.ContainsKey(agence) Then parAgence(agence) = New List(Of LigneHistoriqueWU)()
            parAgence(agence).Add(cumul)
        Next

        Dim totalReseau As Decimal = 0D
        For Each cumul As LigneHistoriqueWU In parAccount.Values
            totalReseau += cumul.TotalCommissions
        Next

        ' Le cumul de chaque agence est calculé UNE FOIS, avant le tri : le recalculer dans le
        ' comparateur le referait autant de fois que la comparaison est appelée.
        Dim cumulParAgence As New Dictionary(Of String, LigneHistoriqueWU)(StringComparer.OrdinalIgnoreCase)
        For Each paire As KeyValuePair(Of String, List(Of LigneHistoriqueWU)) In parAgence
            cumulParAgence(paire.Key) = CumulerLignes(paire.Value)
        Next

        ' Les agences qui pèsent en tête. À commissions égales — deux agences à zéro, par
        ' exemple — le nom départage, pour que deux affichages successifs donnent le même ordre.
        Dim classement As List(Of String) = New List(Of String)(parAgence.Keys)
        classement.Sort(Function(gauche, droite)
                            Dim ecart As Integer = cumulParAgence(droite).TotalCommissions.CompareTo(
                                                   cumulParAgence(gauche).TotalCommissions)
                            If ecart <> 0 Then Return ecart
                            Return String.Compare(gauche, droite, StringComparison.OrdinalIgnoreCase)
                        End Function)

        Dim rang As Integer = 0

        For Each agence As String In classement

            rang += 1

            Dim comptes As List(Of LigneHistoriqueWU) = parAgence(agence)
            Dim cumulAgence As LigneHistoriqueWU = cumulParAgence(agence)

            Dim enteteAgence As DataRow = table.NewRow()
            enteteAgence("Rang") = rang.ToString(Globalization.CultureInfo.InvariantCulture)
            enteteAgence("Agence") = agence
            enteteAgence("Account") = String.Empty
            enteteAgence("Designation") = $"{comptes.Count} Account(s)"
            RemplirColonnesChiffrees(enteteAgence, cumulAgence)
            enteteAgence("Part") = Part(cumulAgence.TotalCommissions, totalReseau)
            enteteAgence("Niveau") = NIVEAU_CATEGORIE
            table.Rows.Add(enteteAgence)

            comptes.Sort(Function(gauche, droite)
                             Dim ecart As Integer = droite.TotalCommissions.CompareTo(gauche.TotalCommissions)
                             If ecart <> 0 Then Return ecart
                             Return String.Compare(gauche.Account, droite.Account, StringComparison.OrdinalIgnoreCase)
                         End Function)

            For Each compte As LigneHistoriqueWU In comptes

                Dim enregistrement As DataRow = table.NewRow()
                enregistrement("Rang") = String.Empty
                enregistrement("Agence") = String.Empty
                enregistrement("Account") = compte.Account
                enregistrement("Designation") = compte.Designation
                RemplirColonnesChiffrees(enregistrement, compte)
                enregistrement("Part") = Part(compte.TotalCommissions, totalReseau)
                enregistrement("Niveau") = NIVEAU_POINT_DE_VENTE
                table.Rows.Add(enregistrement)
            Next
        Next

        If table.Rows.Count = 0 Then Return table

        Dim total As DataRow = table.NewRow()
        total("Rang") = String.Empty
        total("Agence") = LIBELLE_TOTAL
        total("Account") = String.Empty
        total("Designation") = $"{classement.Count} agence(s), {parAccount.Count} Account(s)"
        RemplirColonnesChiffrees(total, CumulerLignes(New List(Of LigneHistoriqueWU)(parAccount.Values)))
        total("Part") = If(totalReseau = 0D, 0D, 1D)
        total("Niveau") = NIVEAU_TOTAL
        table.Rows.Add(total)

        Return table
    End Function

    ''' <summary>Part d'un montant dans un total. Zéro si le total l'est : rien ne se compare à rien.</summary>
    Private Shared Function Part(montant As Decimal, total As Decimal) As Decimal
        If total = 0D Then Return 0D
        Return montant / total
    End Function

    ''' <summary>Restreint les lignes à une agence. Chaîne vide : toutes.</summary>
    Public Shared Function FiltrerParAgence(lignes As List(Of LigneHistoriqueWU), agence As String,
                                            agencesConnues As List(Of PointDeVenteEC)) As List(Of LigneHistoriqueWU)

        Dim toutes As List(Of LigneHistoriqueWU) = SansNothing(lignes)
        If String.IsNullOrWhiteSpace(agence) Then Return toutes

        Dim rattachement As Dictionary(Of String, String) = AgenceParAccount(agencesConnues)
        Dim retenues As New List(Of LigneHistoriqueWU)

        For Each ligne As LigneHistoriqueWU In toutes
            If String.Equals(AgenceDe(ligne.Account, rattachement), agence, StringComparison.OrdinalIgnoreCase) Then
                retenues.Add(ligne)
            End If
        Next

        Return retenues
    End Function

    Private Shared Function AgenceDe(account As String, rattachement As Dictionary(Of String, String)) As String

        Dim nom As String = Nothing
        If rattachement.TryGetValue(If(account, String.Empty).Trim(), nom) Then Return nom

        Return SANS_AGENCE
    End Function

#End Region

#Region "Statut des transactions"

    ''' <summary>
    ''' Traduit le statut Western Union en libellé lisible. Un statut inconnu est restitué tel
    ''' quel plutôt que masqué : mieux vaut un code brut qu'une information perdue.
    ''' </summary>
    Public Shared Function LibelleStatut(statut As String) As String

        Select Case If(statut, String.Empty).Trim().ToUpperInvariant()
            Case "S" : Return "Réglée"
            Case "W" : Return "En attente"
            Case "C" : Return "ANNULÉE"
            Case "" : Return String.Empty
            Case Else : Return statut.Trim()
        End Select
    End Function

#End Region

#Region "Utilitaires de construction"


    ''' <summary>Table de détail à colonne d'en-tête libre (date, par exemple) puis colonnes chiffrées.</summary>
    Private Shared Function CreerTableDetail(nom As String, premiereColonne As String, type As Type) As DataTable

        Dim table As New DataTable(nom)
        table.Columns.Add(premiereColonne, type)
        AjouterColonnesChiffrees(table)
        Return table
    End Function

    ''' <summary>Colonnes chiffrées communes aux pages 2, 3 et 4 : mêmes intitulés, mêmes totaux.</summary>
    Private Shared Sub AjouterColonnesChiffrees(table As DataTable)

        table.Columns.Add("NbEnvois", GetType(Integer))
        table.Columns.Add("PrincipalEnvoi", GetType(Decimal))
        table.Columns.Add("NbPaiements", GetType(Integer))
        table.Columns.Add("PrincipalPaye", GetType(Decimal))
        ' Les annulations figuraient en synthèse mais manquaient au détail : sans elles, il était
        ' impossible de savoir QUI annule, alors que c'est précisément ce qu'on veut suivre.
        table.Columns.Add("NbAnnulations", GetType(Integer))
        table.Columns.Add("Commissions", GetType(Decimal))
        table.Columns.Add("TVA", GetType(Decimal))
        table.Columns.Add("TTA", GetType(Decimal))
        table.Columns.Add("TotalTaxes", GetType(Decimal))
    End Sub

    Private Shared Sub RemplirColonnesChiffrees(enregistrement As DataRow, cumul As LigneHistoriqueWU)

        enregistrement("NbEnvois") = cumul.NombreEnvois
        enregistrement("PrincipalEnvoi") = cumul.PrincipalEnvoi
        enregistrement("NbPaiements") = cumul.NombrePaiements
        enregistrement("PrincipalPaye") = cumul.PrincipalPaye
        enregistrement("NbAnnulations") = cumul.NombreAnnulations
        enregistrement("Commissions") = cumul.TotalCommissions
        enregistrement("TVA") = cumul.TVA
        ' Les deux TTA sont regroupées : envoi et réception relèvent de la même taxe, seule
        ' l'assiette diffère. Le détail reste disponible en page de synthèse.
        enregistrement("TTA") = cumul.TTAEnvoi + cumul.TTAReception
        enregistrement("TotalTaxes") = cumul.TotalTaxes
    End Sub

    Private Shared Sub AjouterLigneDetail(table As DataTable, libelle As String, cumul As LigneHistoriqueWU)

        Dim enregistrement As DataRow = table.NewRow()
        enregistrement(0) = libelle
        RemplirColonnesChiffrees(enregistrement, cumul)
        table.Rows.Add(enregistrement)
    End Sub

    Private Shared Sub AjouterLigneTotal(table As DataTable, lignes As List(Of LigneHistoriqueWU))
        AjouterLigneDetail(table, LIBELLE_TOTAL, Cumuler(lignes))
    End Sub

    ''' <summary>Cumul de toutes les lignes de la période.</summary>
    Public Shared Function Cumuler(lignes As List(Of LigneHistoriqueWU)) As LigneHistoriqueWU
        Return CumulerLignes(SansNothing(lignes))
    End Function

    ''' <summary>Cumul d'un ensemble quelconque de lignes déjà constitué.</summary>
    Private Shared Function CumulerLignes(lignes As List(Of LigneHistoriqueWU)) As LigneHistoriqueWU

        Dim total As New LigneHistoriqueWU()

        If lignes Is Nothing Then Return total

        For Each ligne As LigneHistoriqueWU In lignes
            If ligne IsNot Nothing Then total.Cumuler(ligne)
        Next

        Return total
    End Function

    ''' <summary>Nombre de journées distinctes présentes dans l'historique de la période.</summary>
    Public Shared Function CompterJours(lignes As List(Of LigneHistoriqueWU)) As Integer

        Dim jours As New HashSet(Of Date)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            jours.Add(ligne.DateActivite)
        Next

        Return jours.Count
    End Function

    ''' <summary>Nombre de points de vente distincts ayant eu une activité sur la période.</summary>
    Public Shared Function CompterPointsDeVente(lignes As List(Of LigneHistoriqueWU)) As Integer

        Dim comptes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each ligne As LigneHistoriqueWU In SansNothing(lignes)
            comptes.Add(ligne.Account)
        Next

        Return comptes.Count
    End Function

    Private Shared Function LibelleGroupe(groupe As String) As String
        Dim valeur As String = If(groupe, String.Empty).Trim()
        Return If(valeur.Length = 0, SANS_GROUPE, valeur)
    End Function

    ''' <summary>Énumère la liste en ignorant les éléments absents, pour ne pas semer de tests partout.</summary>
    Private Shared Function SansNothing(lignes As List(Of LigneHistoriqueWU)) As List(Of LigneHistoriqueWU)

        Dim retenues As New List(Of LigneHistoriqueWU)

        If lignes Is Nothing Then Return retenues

        For Each ligne As LigneHistoriqueWU In lignes
            If ligne IsNot Nothing Then retenues.Add(ligne)
        Next

        Return retenues
    End Function

#End Region

End Class
