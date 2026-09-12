Option Strict On
Option Explicit On

''' <summary>
''' Représente l'ensemble des données agrégées et calculées pour un Account
''' (sous-agent, agence propre ou compte non paramétré) dans le cadre de la
''' compensation Western Union J+1.
''' Une instance de cette classe correspond à une ligne de la grille de contrôle
''' et alimente la génération de la pièce comptable.
''' </summary>
Public Class CalculWU

#Region "Identification"

    ''' <summary>Identifiant Account (clé métier unique, jamais codeagence).</summary>
    Public Property Account As String = String.Empty

    ''' <summary>Désignation de l'agence / du sous-agent.</summary>
    Public Property Designation As String = String.Empty

    ''' <summary>Type de point de vente : "SA" (sous-agent), "EC" (agence propre) ou "INCONNU".</summary>
    Public Property TypePdv As String = "INCONNU"

    ''' <summary>Code agence associé (codeagence), à titre informatif uniquement.</summary>
    Public Property CodeAgence As String = String.Empty

    ''' <summary>Compte de compensation du sous-agent (T_Pdv_SA.CompteCompense).</summary>
    Public Property CompteCompense As String = String.Empty

    ''' <summary>Compte de commission du sous-agent (T_Pdv_SA.CompteCommission).</summary>
    Public Property CompteCommission As String = String.Empty

    ''' <summary>Taux de rétrocession du sous-agent (ex : 0.70). Non applicable pour les agences propres.</summary>
    Public Property TauxSA As Decimal = 0D

#End Region

#Region "Agrégats bruts (activité + règlement)"

    Public Property PrincipalEnvoi As Decimal = 0D
    Public Property PrincipalPaye As Decimal = 0D
    Public Property ChargeEnvoi As Decimal = 0D
    Public Property Taxes As Decimal = 0D

#End Region

#Region "Commissions et taxes calculées"

    Public Property CommissionEnvoi As Decimal = 0D
    Public Property CommissionPaiement As Decimal = 0D
    Public Property TVA As Decimal = 0D
    Public Property TTAEnvoi As Decimal = 0D
    Public Property TTAReception As Decimal = 0D
    Public Property SoldeTaxes As Decimal = 0D
    Public Property CommissionTransfert As Decimal = 0D
    Public Property TaxeEnvoi As Decimal = 0D

    ''' <summary>
    ''' Valeur de contrôle intermédiaire "CommissionTransfertBase" = Round(SoldeTaxes), telle
    ''' que produite par le classeur bancaire d'origine (section 8). Non utilisée dans les
    ''' calculs qui suivent (voir WUCalculationService.AppliquerFormules) ; conservée uniquement
    ''' à titre de traçabilité / rapprochement avec le classeur Excel de référence.
    ''' </summary>
    Public Property CommissionTransfertBaseControle As Decimal = 0D

#End Region

#Region "Répartition Banque / Sous-agent"

    Public Property CommissionTransfertBanque As Decimal = 0D
    Public Property CommissionPaiementBanque As Decimal = 0D
    Public Property CommissionEnvoiBanque As Decimal = 0D

    Public Property CommissionTransfertSA As Decimal = 0D
    Public Property CommissionPaiementSA As Decimal = 0D
    Public Property CommissionEnvoiSA As Decimal = 0D

#End Region

#Region "Totaux de contrôle"

    Public Property TotalDebit As Decimal = 0D
    Public Property TotalCredit As Decimal = 0D
    Public Property Solde As Decimal = 0D

    ''' <summary>
    ''' Écart d'arrondi généré par cet Account dans la pièce comptable : différence entre la
    ''' ligne de mouvement arrondie et la somme de ses contreparties arrondies (compte courant
    ''' WU, commissions, taxes). En valeurs exactes cet écart est nul par construction ; il ne
    ''' provient que des arrondis FCFA effectués ligne par ligne (ArrondiFCFA).
    ''' La SOMME de ces écarts sur tous les Accounts est égale à l'écart global de la pièce,
    ''' celui absorbé par le compte d'attente (section 14).
    ''' Renseigné par PieceComptableService (CreerTableControle et GenererPieceComptable).
    ''' </summary>
    Public Property EcartArrondi As Long = 0L

#End Region

#Region "Indicateurs d'anomalie (mise en évidence dans la grille)"

    ''' <summary>Vrai si une erreur SQL est survenue lors de la recherche des paramètres de cet Account.</summary>
    Public Property ErreurSQL As Boolean = False

    ''' <summary>Message d'erreur SQL éventuel, à titre de diagnostic.</summary>
    Public Property MessageErreurSQL As String = String.Empty

    ''' <summary>Vrai si des données essentielles sont manquantes ou incohérentes pour cet Account.</summary>
    Public Property DonneesManquantes As Boolean = False

#End Region

    ''' <summary>Indique si l'Account n'est rattaché à aucun paramétrage connu.</summary>
    Public ReadOnly Property EstInconnu As Boolean
        Get
            Return String.Equals(TypePdv, "INCONNU", StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

End Class
