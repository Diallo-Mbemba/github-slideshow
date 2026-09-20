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

#Region "Ce que la banque a gardé"

    ''' <summary>
    ''' Part de la commission revenant à la banque, TELLE QU'ELLE A ÉTÉ CALCULÉE ce jour-là.
    '''
    ''' Elle n'est pas recalculée à la lecture : la part vaut « commission × (1 − taux) », et
    ''' le taux change. Un sous-agent passé de 70 % à 60 % ferait varier rétroactivement ce
    ''' que la banque a gagné le mois dernier — l'application refuse déjà cela pour les
    ''' pièces comptables, et pour la même raison.
    ''' </summary>
    Public Property CommissionEnvoiBanque As Decimal = 0D
    Public Property CommissionPaiementBanque As Decimal = 0D
    Public Property CommissionTransfertBanque As Decimal = 0D

    ''' <summary>Taux du sous-agent ce jour-là. Zéro pour une agence propre.</summary>
    Public Property TauxSA As Decimal = 0D

    ''' <summary>
    ''' Vrai si la répartition de cette ligne est connue.
    '''
    ''' Elle l'est dans deux cas : la part a été enregistrée en base, OU le point de vente
    ''' n'est pas un sous-agent — son taux valant zéro par construction, la banque a gardé
    ''' la totalité, et ce n'est pas une supposition mais la règle elle-même.
    '''
    ''' Elle vaut VRAI par défaut pour qu'une ligne de cumul, qui part de rien, ne se
    ''' déclare pas ignorante : c'est en lui ajoutant une ligne non documentée qu'elle le
    ''' devient.
    ''' </summary>
    Public Property RepartitionConnue As Boolean = True

    ''' <summary>Ce que la banque a gardé, toutes natures de commission confondues.</summary>
    Public ReadOnly Property TotalPartBanque As Decimal
        Get
            Return CommissionEnvoiBanque + CommissionPaiementBanque + CommissionTransfertBanque
        End Get
    End Property

    ''' <summary>
    ''' Ce qui a été rétrocédé au sous-agent : le reste. Il ne revient pas à la banque, mais
    ''' l'afficher à côté permet de vérifier d'un coup d'œil que les deux parts font bien la
    ''' commission totale.
    ''' </summary>
    Public ReadOnly Property TotalPartSousAgent As Decimal
        Get
            Return TotalCommissions - TotalPartBanque
        End Get
    End Property

    ''' <summary>Vrai si cette ligne décrit un sous-agent.</summary>
    Public ReadOnly Property EstSousAgent As Boolean
        Get
            Return String.Equals(TypePdv, "SA", StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    ''' <summary>
    ''' Pose la part de la banque à partir de ce que la base a rendu.
    '''
    ''' Une colonne absente ne devient pas zéro : sur une agence propre, la totalité revient
    ''' à la banque ; sur un sous-agent, la répartition est déclarée inconnue et l'état le
    ''' dira plutôt que d'afficher un zéro qui se lirait comme un chiffre.
    ''' </summary>
    Public Sub PoserLaPartBanque(envoi As Decimal?, paiement As Decimal?, transfert As Decimal?)

        If envoi.HasValue OrElse paiement.HasValue OrElse transfert.HasValue Then
            CommissionEnvoiBanque = envoi.GetValueOrDefault()
            CommissionPaiementBanque = paiement.GetValueOrDefault()
            CommissionTransfertBanque = transfert.GetValueOrDefault()
            RepartitionConnue = True
            Return
        End If

        If Not EstSousAgent Then
            CommissionEnvoiBanque = CommissionEnvoi
            CommissionPaiementBanque = CommissionPaiement
            CommissionTransfertBanque = CommissionTransfert
            RepartitionConnue = True
            Return
        End If

        CommissionEnvoiBanque = 0D
        CommissionPaiementBanque = 0D
        CommissionTransfertBanque = 0D
        RepartitionConnue = False
    End Sub

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

        CommissionEnvoiBanque += autre.CommissionEnvoiBanque
        CommissionPaiementBanque += autre.CommissionPaiementBanque
        CommissionTransfertBanque += autre.CommissionTransfertBanque

        ' Un cumul n'est documenté que si TOUT ce qu'il contient l'est. Une seule journée
        ' sans répartition suffit à rendre le total douteux, et il vaut mieux le dire.
        RepartitionConnue = RepartitionConnue AndAlso autre.RepartitionConnue

        ' Le taux ne se cumule pas : additionner deux taux n'aurait aucun sens. Il reste
        ' celui de la ligne d'origine, et vaut zéro sur un cumul de plusieurs points de vente.
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
            .TaxeEnvoi = calc.TaxeEnvoi,
            .CommissionEnvoiBanque = calc.CommissionEnvoiBanque,
            .CommissionPaiementBanque = calc.CommissionPaiementBanque,
            .CommissionTransfertBanque = calc.CommissionTransfertBanque,
            .TauxSA = calc.TauxSA,
            .RepartitionConnue = True
        }
    End Function

End Class
