Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>
''' Sous-agent Western Union : une ligne de la table T_Pdv_SA.
''' L'Account (Code_Pdv) est l'identifiant métier unique et la clé primaire de la table.
''' </summary>
Public Class PointDeVenteSA

    ''' <summary>Account du sous-agent (clé primaire). Jamais codeagence.</summary>
    Public Property CodePdv As String = String.Empty

    ''' <summary>Désignation de l'agence, reprise dans le libellé des écritures comptables.</summary>
    Public Property Designation As String = String.Empty

    ''' <summary>Groupe statistique. Non exploité par les calculs, mais présent dans la table.</summary>
    Public Property GroupeStatistique As String = String.Empty

    ''' <summary>
    ''' Quote-part du sous-agent dans les commissions, exprimée en fraction : 0,70 = 70 %.
    ''' La colonne SQL étant de type DECIMAL(4,2), elle ne retient que DEUX décimales.
    ''' </summary>
    Public Property Taux As Decimal = 0D

    ''' <summary>Compte de compensation du sous-agent (ligne de mouvement de la pièce comptable).</summary>
    Public Property CompteCompense As String = String.Empty

    ''' <summary>Compte de commission du sous-agent (rétrocession des commissions).</summary>
    Public Property CompteCommission As String = String.Empty

    ''' <summary>Code agence. Informatif uniquement : jamais utilisé comme clé de recherche.</summary>
    Public Property CodeAgence As String = String.Empty

    ''' <summary>Retire les espaces parasites de tous les champs texte.</summary>
    Public Sub Normaliser()
        CodePdv = If(CodePdv, String.Empty).Trim()
        Designation = If(Designation, String.Empty).Trim()
        GroupeStatistique = If(GroupeStatistique, String.Empty).Trim()
        CompteCompense = If(CompteCompense, String.Empty).Trim()
        CompteCommission = If(CompteCommission, String.Empty).Trim()
        CodeAgence = If(CodeAgence, String.Empty).Trim()
    End Sub

    ''' <summary>
    ''' Contrôle de saisie. Retourne la liste des anomalies BLOQUANTES — vide si la fiche est
    ''' enregistrable. Les champs contrôlés sont ceux déclarés NOT NULL dans T_Pdv_SA, plus
    ''' le taux dont une valeur aberrante fausserait toutes les répartitions de commissions.
    ''' </summary>
    Public Function Anomalies() As List(Of String)

        Dim messages As New List(Of String)

        If String.IsNullOrWhiteSpace(CodePdv) Then messages.Add("L'Account (Code_Pdv) est obligatoire.")
        If String.IsNullOrWhiteSpace(Designation) Then messages.Add("La désignation de l'agence est obligatoire.")
        If String.IsNullOrWhiteSpace(CompteCompense) Then messages.Add("Le compte de compensation est obligatoire.")
        If String.IsNullOrWhiteSpace(CompteCommission) Then messages.Add("Le compte de commission est obligatoire.")

        ' Un taux hors [0 ; 1] n'a aucun sens métier — la part du sous-agent ne peut pas être
        ' négative ni dépasser 100 % — et la confusion « 70 » pour « 0,70 » multiplierait par
        ' cent toutes les commissions rétrocédées. Le blocage est donc délibéré.
        If Taux < 0D OrElse Taux > 1D Then
            messages.Add($"Le taux doit être compris entre 0 et 1 (saisir 0,70 pour 70 %). Valeur saisie : {Taux}.")
        End If

        ' DECIMAL(4,2) : au-delà de deux décimales, SQL Server arrondirait en silence.
        If Decimal.Round(Taux, 2) <> Taux Then
            messages.Add($"Le taux ne peut comporter que deux décimales (la base arrondirait {Taux} à {Decimal.Round(Taux, 2)}).")
        End If

        Return messages
    End Function

    ''' <summary>
    ''' Analyse un taux saisi au clavier, en acceptant indifféremment la virgule et le point
    ''' comme séparateur décimal : selon le poste, la culture Windows n'est pas la même et
    ''' l'utilisateur ne doit pas avoir à s'en préoccuper.
    ''' </summary>
    ''' <returns>False si le texte n'est pas un nombre exploitable.</returns>
    Public Shared Function EssayerAnalyserTaux(texte As String, ByRef taux As Decimal) As Boolean

        taux = 0D

        If String.IsNullOrWhiteSpace(texte) Then Return False

        Dim valeur As String = texte.Trim().Replace(" ", "")

        If Decimal.TryParse(valeur, NumberStyles.Number, CultureInfo.CurrentCulture, taux) Then Return True
        If Decimal.TryParse(valeur, NumberStyles.Number, CultureInfo.InvariantCulture, taux) Then Return True

        ' Dernier essai : le séparateur saisi est celui de l'autre convention.
        Return Decimal.TryParse(valeur.Replace(",", "."), NumberStyles.Number, CultureInfo.InvariantCulture, taux)
    End Function

End Class

''' <summary>
''' Groupe statistique de sous-agents.
'''
''' Règle métier : un groupe porte UN compte d'activité, UN compte de commission et UN taux.
''' Tout Account appartient à un seul groupe et en hérite automatiquement ces trois valeurs ;
''' un même couple de comptes ne peut appartenir qu'à un seul groupe.
'''
''' Il n'existe pas de table de groupes : un groupe n'a d'existence que par les sous-agents
''' qui le portent, et ses valeurs sont donc celles de ses membres. Rien n'empêche alors, au
''' niveau de la base, que deux membres d'un même groupe portent des valeurs différentes :
''' c'est pourquoi cette classe transporte aussi le constat d'une éventuelle INCOHÉRENCE,
''' que l'application signale sans jamais la corriger d'autorité.
''' </summary>
Public Class GroupeStatistiqueWU

    ''' <summary>Libellé du groupe, tel qu'enregistré.</summary>
    Public Property Nom As String = String.Empty

    ''' <summary>
    ''' Compte d'activité du groupe (colonne CompteCompense). C'est lui qui porte la ligne de
    ''' mouvement « CCS_... ACTIVITE WU » de la pièce comptable.
    ''' </summary>
    Public Property CompteActivite As String = String.Empty

    ''' <summary>Compte de commission du groupe (colonne CompteCommission).</summary>
    Public Property CompteCommission As String = String.Empty

    ''' <summary>Quote-part du groupe dans les commissions (0,70 = 70 %).</summary>
    Public Property Taux As Decimal = 0D

    ''' <summary>Nombre de sous-agents rattachés au groupe.</summary>
    Public Property NombreSousAgents As Integer = 0

    ''' <summary>
    ''' Vrai si les sous-agents du groupe ne portent pas tous les mêmes valeurs — situation
    ''' que la règle métier exclut, mais que des données antérieures peuvent présenter.
    ''' </summary>
    Public Property EstIncoherent As Boolean = False

    ''' <summary>Description des valeurs divergentes, vide si le groupe est cohérent.</summary>
    Public Property DetailIncoherence As String = String.Empty

    ''' <summary>Vrai si les trois valeurs héritées de ce groupe sont celles de la fiche indiquée.</summary>
    Public Function CorrespondA(pdv As PointDeVenteSA) As Boolean

        If pdv Is Nothing Then Return False

        Return String.Equals(CompteActivite, If(pdv.CompteCompense, String.Empty).Trim(), StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(CompteCommission, If(pdv.CompteCommission, String.Empty).Trim(), StringComparison.OrdinalIgnoreCase) AndAlso
               Decimal.Round(Taux, 2) = Decimal.Round(pdv.Taux, 2)
    End Function

    Public Overrides Function ToString() As String
        Return Nom
    End Function

End Class

''' <summary>
''' Agence propre Ecobank : une ligne de la table T_Pdv_EC.
''' L'Account (Codesite) est l'identifiant métier unique et la clé primaire de la table.
''' Une agence propre ne rétrocède aucune commission : la banque en conserve 100 %.
''' </summary>
Public Class PointDeVenteEC

    ''' <summary>Account de l'agence (clé primaire).</summary>
    Public Property CodeSite As String = String.Empty

    ''' <summary>Désignation de l'agence, reprise dans le libellé des écritures comptables.</summary>
    Public Property Designation As String = String.Empty

    ''' <summary>Code agence Voyager (colonne [CodeAgenc-Voyager]).</summary>
    Public Property CodeAgenceVoyager As String = String.Empty

    Public Sub Normaliser()
        CodeSite = If(CodeSite, String.Empty).Trim()
        Designation = If(Designation, String.Empty).Trim()
        CodeAgenceVoyager = If(CodeAgenceVoyager, String.Empty).Trim()
    End Sub

    ''' <summary>Anomalies bloquantes : seul l'Account est indispensable pour identifier l'agence.</summary>
    Public Function Anomalies() As List(Of String)

        Dim messages As New List(Of String)

        If String.IsNullOrWhiteSpace(CodeSite) Then messages.Add("L'Account (Codesite) est obligatoire.")

        Return messages
    End Function

    ''' <summary>
    ''' Avertissements non bloquants. La désignation est facultative dans la table, mais elle
    ''' est reprise telle quelle dans le libellé des écritures (« CCS_... ACTIVITE WU ») :
    ''' la laisser vide produit une pièce comptable difficilement lisible.
    ''' </summary>
    Public Function Avertissements() As List(Of String)

        Dim messages As New List(Of String)

        If String.IsNullOrWhiteSpace(Designation) Then
            messages.Add("La désignation est vide : le libellé des écritures comptables de cette agence sera incomplet.")
        End If

        Return messages
    End Function

End Class
