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
    ''' <summary>
    ''' Groupe statistique du point de vente, pour les sous-agents. Non utilisé par les calculs
    ''' comptables : repris dans l'historique, où il permet de ventiler l'activité par groupe.
    ''' </summary>
    Public Property GroupeStatistique As String = String.Empty

    Public Property CompteCompense As String = String.Empty

    ''' <summary>Compte de commission du sous-agent (T_Pdv_SA.CompteCommission).</summary>
    Public Property CompteCommission As String = String.Empty

    ''' <summary>Taux de rétrocession du sous-agent (ex : 0.70). Non applicable pour les agences propres.</summary>
    Public Property TauxSA As Decimal = 0D

#End Region

#Region "Agrégats bruts (activité + règlement)"

    ''' <summary>Nombre de transactions d'envoi de la journée pour cet Account.</summary>
    Public Property NombreEnvois As Integer = 0

    ''' <summary>Nombre de transactions de paiement de la journée pour cet Account.</summary>
    Public Property NombrePaiements As Integer = 0

    ''' <summary>Nombre de transactions annulées, exclues des montants mais comptées.</summary>
    Public Property NombreAnnulations As Integer = 0

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

    ''' <summary>
    ''' Indique si cet Account peut être comptabilisé, c'est-à-dire si la banque sait sur quels
    ''' comptes poser ses écritures.
    '''
    ''' Un Account non comptabilisable est écarté de la pièce comptable : il n'y apparaît pas
    ''' du tout. Le comptabiliser quand même reviendrait à poser son mouvement sur le compte
    ''' courant WU de la banque, c'est-à-dire sur un compte qui n'est pas le sien, et à laisser
    ''' la correction se faire à la main, écriture par écriture, une fois la pièce chargée.
    '''
    ''' Trois cas, et trois seulement :
    '''   - l'Account est absent de T_Pdv_SA comme de T_Pdv_EC : aucun compte n'est connu ;
    '''   - c'est un sous-agent sans compte de compensation : son mouvement n'a pas de compte ;
    '''   - c'est un sous-agent sans compte de commission : sa rétrocession n'a pas de compte,
    '''     et la pièce partirait en déséquilibre du montant de cette commission.
    '''
    ''' Une agence propre (EC) connue est toujours comptabilisable : ses écritures vont sur le
    ''' compte courant WU par règle métier, et non faute de mieux.
    ''' </summary>
    Public ReadOnly Property EstComptabilisable As Boolean
        Get
            Return MotifNonComptabilise.Length = 0
        End Get
    End Property

    ''' <summary>
    ''' Raison, en clair, pour laquelle l'Account ne peut pas être comptabilisé.
    ''' Chaîne vide s'il est comptabilisable. Destinée à être affichée telle quelle.
    ''' </summary>
    Public ReadOnly Property MotifNonComptabilise As String
        Get
            If ErreurSQL Then Return "paramétrage non lu (erreur SQL)"
            If EstInconnu Then Return "absent du paramétrage"

            If String.Equals(TypePdv, "SA", StringComparison.OrdinalIgnoreCase) Then
                If String.IsNullOrWhiteSpace(CompteCompense) Then Return "sous-agent sans compte de compensation"
                If String.IsNullOrWhiteSpace(CompteCommission) Then Return "sous-agent sans compte de commission"
            End If

            Return String.Empty
        End Get
    End Property

    ''' <summary>
    ''' Montant net porté sur le compte de compensation du point de vente : ce que ce point de
    ''' vente doit verser à la banque.
    '''
    '''     (principal envoyé + charges + taxes) − principal payé + TTA sur réception
    '''
    ''' POURQUOI LA TTA SUR RÉCEPTION S'AJOUTE, ALORS QU'ELLE NE FIGURE DANS AUCUN RAPPORT
    '''
    ''' La symétrie se lit dans les données de la plateforme. À l'ENVOI, la TTA est déjà
    ''' encaissée par Western Union : elle est l'une des trois composantes de Taxes (Tax3REC),
    ''' donc déjà dans la caisse du point de vente, et déjà comptée ici par le terme « taxes ».
    ''' À la RÉCEPTION, la plateforme ne prélève RIEN : la taxe est due au Trésor tchadien sur
    ''' une opération que Western Union ignore.
    '''
    ''' Sans ce terme, la pièce créditait le compte de TTA sur réception sans l'avoir encaissée
    ''' nulle part : la contrepartie tombait sur le compte courant Western Union, qui se
    ''' trouvait financer une taxe tchadienne qu'il ne doit pas. Ce n'était pas un choix, mais
    ''' un oubli — et il valait 19 860 F sur une seule semaine et un seul sous-agent.
    '''
    ''' Une agence propre ne retient pas de TTA sur paiement (règle de la banque) : sa
    ''' TTAReception vaut zéro, et ce terme est alors sans effet. La règle est posée une seule
    ''' fois, dans WUCalculationService.AppliquerFormules.
    '''
    ''' CETTE PROPRIÉTÉ EST LE SEUL ENDROIT OÙ LA FORMULE EST ÉCRITE. La pièce comptable et le
    ''' calcul de l'écart d'arrondi l'appellent au lieu de la recopier : deux copies d'une
    ''' formule finissent toujours par diverger, et celle-ci porte le montant que le point de
    ''' vente doit réellement verser.
    '''
    ''' Sert aussi à chiffrer ce qui n'est pas comptabilisé lorsque l'Account est écarté — un
    ''' nombre de lignes ne dit rien de l'enjeu.
    ''' </summary>
    Public ReadOnly Property NetMouvement As Decimal
        Get
            Return (PrincipalEnvoi + ChargeEnvoi + Taxes) - PrincipalPaye + TTAReception
        End Get
    End Property

End Class
