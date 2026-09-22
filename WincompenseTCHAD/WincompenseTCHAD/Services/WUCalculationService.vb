Option Strict On
Option Explicit On

''' <summary>
''' Service de calcul métier : application des formules de commissions/taxes,
''' répartition Banque / Sous-agent et arrondi final en FCFA.
''' Aucune lecture de fichier ni accès SQL ici : uniquement des calculs sur CalculWU.
''' </summary>
Public NotInheritable Class WUCalculationService

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Applique les formules de commissions et taxes à partir des agrégats bruts déjà renseignés
    ''' sur l'objet (PrincipalEnvoi, PrincipalPaye, ChargeEnvoi, Taxes, CommissionPaiement).
    ''' Toutes les décimales sont conservées ; aucun arrondi intermédiaire, sauf lorsque la formule
    ''' du classeur bancaire l'exige explicitement (CommissionTransfertBase, cf. section 8).
    ''' </summary>
    Public Shared Sub AppliquerFormules(calc As CalculWU)

        If calc Is Nothing Then Return

        ' Commission Envoi = ChargeEnvoi * 20,5 %.
        calc.CommissionEnvoi = calc.ChargeEnvoi * ConstantesWU.TAUX_COMMISSION_ENVOI

        ' TVA PORTÉE AU COMPTE DE TVA = ChargeEnvoi * 18 %.
        '
        ' Le taux de TVA perçu reste 19,25 % (ConstantesWU.TAUX_TVA) : il n'est pas modifié.
        ' Ce qui est écrit ici, c'est la part de cette TVA qui va sur le COMPTE DE TVA, et
        ' elle ne vaut que 18 points. Le 1,25 point restant n'est pas de la TVA au sens
        ' comptable ; il n'est donc pas retranché du solde de taxes quelques lignes plus bas,
        ' il y demeure et s'y ventile 25 % / 75 % comme le reste du solde.
        '
        ' Ne pas remplacer TAUX_TVA_COMPTE_TVA par TAUX_TVA « pour revenir à 19,25 % » : le
        ' taux perçu EST déjà 19,25 %, la totalité du montant est bien dans la pièce, elle
        ' est simplement répartie sur trois comptes au lieu d'un seul. Écrire TAUX_TVA ici
        ' ne relèverait pas le taux, cela déplacerait 1,25 point des impôts et de la
        ' commission sur transfert vers le compte de TVA — c'est exactement l'écart que la
        ' banque nous a signalé sur AHB020013.
        calc.TVA = calc.ChargeEnvoi * ConstantesWU.TAUX_TVA_COMPTE_TVA

        ' TTA Envoi = PrincipalEnvoi * 0,2 %.
        calc.TTAEnvoi = calc.PrincipalEnvoi * ConstantesWU.TAUX_TTA

        ' TTA Réception = PrincipalPaye * 0,2 % — MAIS SEULEMENT POUR UN SOUS-AGENT.
        '
        ' RÈGLE DE LA BANQUE : une agence propre ne supporte pas de TTA sur paiement. La taxe
        ' n'est pas due, elle n'est donc pas calculée — et non pas calculée puis omise de la
        ' pièce. La différence n'est pas de forme : calculée, elle apparaîtrait dans la grille
        ' de contrôle, dans le rapport d'activité et dans l'historique comme une taxe que la
        ' banque devrait, alors qu'elle ne la doit pas.
        '
        ' C'EST LE SEUL ENDROIT À MODIFIER, et c'est voulu. Tout ce qui suit part de cette
        ' valeur : la ligne de la pièce, qu'AjouterLigneSiNonNul écarte d'elle-même à zéro ;
        ' la contrepartie sur le compte courant WU, qui augmente d'autant ; l'écart d'arrondi ;
        ' les totaux de contrôle ; l'historique ; le fichier destiné au core banking. Poser la
        ' règle plus bas obligerait à la répéter, et deux copies d'une règle finissent
        ' toujours par diverger.
        '
        ' INCONNU N'EST PAS TRAITÉ COMME EC ICI, contrairement à RepartirCommissions. Un
        ' Account non paramétré n'est pas une agence propre : il n'est rien encore. Il n'entre
        ' de toute façon pas dans la pièce ; la taxe reste donc calculée, et la grille de
        ' contrôle montre ce qu'il faudrait payer s'il s'avérait être un sous-agent. Mettre à
        ' zéro sur une supposition serait affirmer plus qu'on ne sait.
        If String.Equals(calc.TypePdv, "EC", StringComparison.OrdinalIgnoreCase) Then
            calc.TTAReception = 0D
        Else
            calc.TTAReception = calc.PrincipalPaye * ConstantesWU.TAUX_TTA
        End If

        ' Solde de taxes = Taxes - TVA portée au compte de TVA - TTAEnvoi.
        '
        ' La formule est inchangée, et c'est précisément ce qui fait qu'elle absorbe d'elle-même
        ' le 1,25 point non porté au compte de TVA : le solde étant un RÉSIDU, tout franc qui
        ' n'est pas retranché ici y reste et se retrouve ensuite réparti 25 % / 75 %.
        calc.SoldeTaxes = calc.Taxes - calc.TVA - calc.TTAEnvoi

        ' Reproduction de la logique du classeur bancaire (section 8) : une base arrondie
        ' "CommissionTransfertBase" est calculée explicitement (Round) à titre de valeur de
        ' référence/contrôle intermédiaire, telle que produite par le classeur d'origine.
        ' Elle n'est volontairement PAS réinjectée dans les calculs qui suivent : conformément
        ' à la règle générale « pas d'arrondi intermédiaire », la répartition Commission
        ' Transfert / Taxe Envoi s'effectue sur le SoldeTaxes exact (non arrondi).
        Dim commissionTransfertBase As Decimal = Math.Round(calc.SoldeTaxes, 0, MidpointRounding.AwayFromZero)
        calc.CommissionTransfertBaseControle = commissionTransfertBase

        calc.CommissionTransfert = calc.SoldeTaxes * ConstantesWU.TAUX_COM_TRANSFERT
        calc.TaxeEnvoi = calc.SoldeTaxes * ConstantesWU.TAUX_TAXE_ENVOI

        ' CommissionPaiement est déjà calculée en amont par WUReportService.CalculerReglement
        ' et affectée à calc.CommissionPaiement avant l'appel à cette méthode.
    End Sub

    ''' <summary>
    ''' Répartit les commissions entre la Banque et le Sous-agent selon le type de point de vente :
    ''' - "SA" : répartition selon TauxSA / TauxBanque = 1 - TauxSA.
    ''' - "EC" (agence propre) : la banque conserve 100 % des commissions.
    ''' - "INCONNU" : à défaut de paramétrage, traité comme une agence propre (100 % banque),
    '''   l'Account restant signalé comme anomalie dans la grille (TypePdv = INCONNU). À CONFIRMER.
    ''' </summary>
    Public Shared Sub RepartirCommissions(calc As CalculWU)

        If calc Is Nothing Then Return

        Select Case calc.TypePdv
            Case "SA"
                Dim tauxSA As Decimal = calc.TauxSA
                Dim tauxBanque As Decimal = 1D - tauxSA

                calc.CommissionTransfertBanque = calc.CommissionTransfert * tauxBanque
                calc.CommissionPaiementBanque = calc.CommissionPaiement * tauxBanque
                calc.CommissionEnvoiBanque = calc.CommissionEnvoi * tauxBanque

                calc.CommissionTransfertSA = calc.CommissionTransfert * tauxSA
                calc.CommissionPaiementSA = calc.CommissionPaiement * tauxSA
                calc.CommissionEnvoiSA = calc.CommissionEnvoi * tauxSA

            Case "EC", "INCONNU"
                ' Agence propre (ou Account non paramétré, par défaut) : la banque conserve 100 %.
                calc.CommissionTransfertBanque = calc.CommissionTransfert
                calc.CommissionPaiementBanque = calc.CommissionPaiement
                calc.CommissionEnvoiBanque = calc.CommissionEnvoi

                calc.CommissionTransfertSA = 0D
                calc.CommissionPaiementSA = 0D
                calc.CommissionEnvoiSA = 0D

            Case Else
                ' Cas théoriquement impossible (TypePdv toujours renseigné par WURepository).
                calc.CommissionTransfertBanque = calc.CommissionTransfert
                calc.CommissionPaiementBanque = calc.CommissionPaiement
                calc.CommissionEnvoiBanque = calc.CommissionEnvoi
                calc.DonneesManquantes = True
        End Select
    End Sub

    ''' <summary>
    ''' Calcule les totaux de contrôle par Account (grille de contrôle uniquement, hors pièce comptable).
    '''
    ''' Solde = TotalDebit − TotalCredit, et vaut toujours l'opposé de CalculWU.NetMouvement :
    ''' les commissions et taxes s'équilibrent par construction à ce niveau (partie double),
    ''' seul l'écart net entre ce que le point de vente a payé et ce qu'il a encaissé doit
    ''' apparaître.
    '''
    ''' LA TTA SUR RÉCEPTION EST DU CÔTÉ ENCAISSÉ, comme dans NetMouvement : la plateforme ne
    ''' la prélève pas, c'est le point de vente qui la reverse. L'oublier ici ferait afficher à
    ''' la grille de contrôle un solde qui ne serait plus celui de la pièce — deux chiffres pour
    ''' la même chose, et l'agent ne saurait plus lequel croire.
    '''
    ''' Conformément à la règle métier, aucun équilibrage n'est forcé Account par Account :
    ''' ce Solde est un indicateur de contrôle, pas une correction.
    ''' </summary>
    Public Shared Sub CalculerTotaux(calc As CalculWU)

        If calc Is Nothing Then Return

        Dim totalCommissionsEtTaxes As Decimal =
            calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
            calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
            calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

        calc.TotalDebit = calc.PrincipalPaye + totalCommissionsEtTaxes
        calc.TotalCredit = (calc.PrincipalEnvoi + calc.ChargeEnvoi + calc.Taxes + calc.TTAReception) + totalCommissionsEtTaxes
        calc.Solde = calc.TotalDebit - calc.TotalCredit
    End Sub

    ''' <summary>
    ''' Arrondit un montant en FCFA à l'unité la plus proche (arrondi commercial, AwayFromZero).
    ''' À utiliser UNIQUEMENT au moment de la génération de la pièce comptable définitive,
    ''' jamais lors des calculs intermédiaires.
    ''' </summary>
    Public Shared Function ArrondiFCFA(montant As Decimal) As Long
        Return CLng(Math.Round(montant, 0, MidpointRounding.AwayFromZero))
    End Function

End Class
