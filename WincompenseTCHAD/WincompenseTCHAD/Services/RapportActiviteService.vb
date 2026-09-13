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

    ''' <summary>Libellé des points de vente sans groupe statistique.</summary>
    Public Const SANS_GROUPE As String = "(sans groupe statistique)"

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
    Public Shared Function ConstruireParPointDeVente(lignes As List(Of LigneHistoriqueWU)) As DataTable

        Dim table As New DataTable("ParPointDeVente")
        table.Columns.Add("Account", GetType(String))
        table.Columns.Add("Designation", GetType(String))
        table.Columns.Add("Groupe", GetType(String))
        AjouterColonnesChiffrees(table)

        ' Cumul par Account, en conservant sa nature : sous-agent, agence propre ou non paramétré.
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

            ' Les points de vente qui pèsent apparaissent en tête de leur catégorie.
            duType.Sort(Function(x, y) y.PrincipalEnvoi.CompareTo(x.PrincipalEnvoi))

            Dim entete As DataRow = table.NewRow()
            entete("Account") = typePdv
            entete("Designation") = String.Empty
            entete("Groupe") = String.Empty
            table.Rows.Add(entete)

            For Each cumul As LigneHistoriqueWU In duType
                Dim enregistrement As DataRow = table.NewRow()
                enregistrement("Account") = cumul.Account
                enregistrement("Designation") = cumul.Designation
                enregistrement("Groupe") = LibelleGroupe(cumul.GroupeStatistique)
                RemplirColonnesChiffrees(enregistrement, cumul)
                table.Rows.Add(enregistrement)
            Next

            Dim sousTotal As DataRow = table.NewRow()
            sousTotal("Account") = LIBELLE_SOUS_TOTAL
            sousTotal("Designation") = $"{typePdv} — {duType.Count} point(s) de vente"
            sousTotal("Groupe") = String.Empty
            RemplirColonnesChiffrees(sousTotal, CumulerLignes(duType))
            table.Rows.Add(sousTotal)
        Next

        Dim total As DataRow = table.NewRow()
        total("Account") = LIBELLE_TOTAL
        total("Designation") = $"{cumuls.Count} point(s) de vente"
        total("Groupe") = String.Empty
        RemplirColonnesChiffrees(total, Cumuler(lignes))
        table.Rows.Add(total)

        Return table
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
