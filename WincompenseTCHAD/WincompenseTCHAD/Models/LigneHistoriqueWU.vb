Option Strict On
Option Explicit On

''' <summary>
''' Une ligne de l'historique : le résultat d'UNE journée comptabilisée pour UN point de vente
''' (table T_HistoriqueWU).
'''
''' L'identification du point de vente — désignation, groupe, type — est recopiée telle qu'elle
''' était le jour de la comptabilisation, et non rattachée par clé au paramétrage courant :
''' un point de vente peut changer de groupe, ce qui ne doit pas réécrire le passé.
''' </summary>
Public Class LigneHistoriqueWU

#Region "Identification"

    Public Property DateActivite As Date
    Public Property Account As String = String.Empty
    Public Property Designation As String = String.Empty
    Public Property GroupeStatistique As String = String.Empty
    Public Property TypePdv As String = String.Empty

#End Region

#Region "Volumes"

    Public Property NombreEnvois As Integer = 0
    Public Property NombrePaiements As Integer = 0
    Public Property NombreAnnulations As Integer = 0

#End Region

#Region "Montants issus des rapports"

    Public Property PrincipalEnvoi As Decimal = 0D
    Public Property ChargeEnvoi As Decimal = 0D
    Public Property Taxes As Decimal = 0D
    Public Property PrincipalPaye As Decimal = 0D

#End Region

#Region "Commissions et taxes calculées"

    Public Property CommissionEnvoi As Decimal = 0D
    Public Property CommissionPaiement As Decimal = 0D
    Public Property CommissionTransfert As Decimal = 0D

    Public Property TVA As Decimal = 0D
    Public Property TTAEnvoi As Decimal = 0D
    Public Property TTAReception As Decimal = 0D
    Public Property TaxeEnvoi As Decimal = 0D

#End Region

#Region "Totaux"

    ''' <summary>Total des commissions : envoi, paiement et transfert.</summary>
    Public ReadOnly Property TotalCommissions As Decimal
        Get
            Return CommissionEnvoi + CommissionPaiement + CommissionTransfert
        End Get
    End Property

    ''' <summary>Total des taxes : TVA, TTA envoi, TTA réception et taxe sur envoi.</summary>
    Public ReadOnly Property TotalTaxes As Decimal
        Get
            Return TVA + TTAEnvoi + TTAReception + TaxeEnvoi
        End Get
    End Property

#End Region

    ''' <summary>Ajoute les valeurs d'une autre ligne : sert à cumuler par jour, par PDV ou par groupe.</summary>
    Public Sub Cumuler(autre As LigneHistoriqueWU)

        If autre Is Nothing Then Return

        NombreEnvois += autre.NombreEnvois
        NombrePaiements += autre.NombrePaiements
        NombreAnnulations += autre.NombreAnnulations

        PrincipalEnvoi += autre.PrincipalEnvoi
        ChargeEnvoi += autre.ChargeEnvoi
        Taxes += autre.Taxes
        PrincipalPaye += autre.PrincipalPaye

        CommissionEnvoi += autre.CommissionEnvoi
        CommissionPaiement += autre.CommissionPaiement
        CommissionTransfert += autre.CommissionTransfert

        TVA += autre.TVA
        TTAEnvoi += autre.TTAEnvoi
        TTAReception += autre.TTAReception
        TaxeEnvoi += autre.TaxeEnvoi
    End Sub

    ''' <summary>Construit une ligne d'historique à partir du calcul d'une journée.</summary>
    ''' <param name="calc">Résultat du calcul pour un Account.</param>
    ''' <param name="jour">Journée d'activité comptabilisée.</param>
    Public Shared Function DepuisCalcul(calc As CalculWU, jour As Date) As LigneHistoriqueWU

        Return New LigneHistoriqueWU() With {
            .DateActivite = jour.Date,
            .Account = calc.Account,
            .Designation = calc.Designation,
            .GroupeStatistique = calc.GroupeStatistique,
            .TypePdv = calc.TypePdv,
            .NombreEnvois = calc.NombreEnvois,
            .NombrePaiements = calc.NombrePaiements,
            .NombreAnnulations = calc.NombreAnnulations,
            .PrincipalEnvoi = calc.PrincipalEnvoi,
            .ChargeEnvoi = calc.ChargeEnvoi,
            .Taxes = calc.Taxes,
            .PrincipalPaye = calc.PrincipalPaye,
            .CommissionEnvoi = calc.CommissionEnvoi,
            .CommissionPaiement = calc.CommissionPaiement,
            .CommissionTransfert = calc.CommissionTransfert,
            .TVA = calc.TVA,
            .TTAEnvoi = calc.TTAEnvoi,
            .TTAReception = calc.TTAReception,
            .TaxeEnvoi = calc.TaxeEnvoi
        }
    End Function

End Class
