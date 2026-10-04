Option Strict On
Option Explicit On

''' <summary>
''' Le sens d'un écart de change : la banque gagne, la banque perd, ou rien.
'''
''' TROIS VALEURS, PAS DEUX. Un écart nul n'est pas un gain de zéro franc : c'est une
''' transaction dont la conversion tombe juste, et il y en a. Les compter à part évite de
''' lire « 1 200 gains » là où 400 lignes n'ont rien produit.
''' </summary>
Public Enum NatureEcartWU

    ''' <summary>La contre-valeur tombe exactement sur le montant local : rien à comptabiliser.</summary>
    Neutre = 0

    ''' <summary>Montant local supérieur à la contre-valeur : gain de change.</summary>
    Gain = 1

    ''' <summary>Montant local inférieur à la contre-valeur : perte de change.</summary>
    Perte = 2

End Enum

''' <summary>
''' L'écart de change d'UNE transaction : son montant local, son montant en devise, la parité
''' employée, et la différence entre les deux.
'''
''' POURQUOI L'ÉCART N'EST PAS UNE PROPRIÉTÉ QU'ON ÉCRIT. Trois grandeurs de cet objet se
''' déduisent des autres — la contre-valeur, l'écart, sa nature — et toutes trois sont en
''' lecture seule. Un service qui calculerait l'écart et l'affecterait pourrait écrire un
''' écart qui ne correspond pas aux montants de la même ligne : la pièce comptable serait
''' alors juste au total et fausse au détail, et c'est exactement le genre de faute que
''' personne ne retrouve. Ici l'incohérence est IMPOSSIBLE À ÉCRIRE : on pose le montant
''' local, le montant en devise et la parité ; le reste en découle.
'''
''' CE QUI EST CONSERVÉ POUR LA PREUVE, ET NON POUR LE CALCUL. ClearPrincipalLoc et
''' ClearFxLoc sont les deux colonnes du rapport dont la somme forme le montant en devise
''' d'un ENVOI ; pour un PAIEMENT, seule la première entre dans le calcul. L'écran de
''' contrôle les affiche toutes les deux : sans elles, un comptable qui conteste un écart
''' n'a aucun moyen de refaire le chemin depuis son rapport.
''' </summary>
Public Class EcartChangeWU

#Region "Identité de la transaction"

    ''' <summary>Numéro de contrôle du transfert (colonne MTCN) : la référence que la banque cite.</summary>
    Public Property Mtcn As String = String.Empty

    ''' <summary>
    ''' Date de règlement de la ligne, reconstituée depuis le triplet SetDateLOC. Nothing si
    ''' le rapport ne la porte pas : la date sert à présenter et à regrouper, jamais à calculer.
    ''' </summary>
    Public Property DateReglement As Date?

    ''' <summary>"S" pour un envoi, "P" pour un paiement (colonne SendPayIndicator).</summary>
    Public Property Sens As String = String.Empty

    ''' <summary>Code produit Western Union de la ligne (IMTR, FTSS, AVSS…). Informatif.</summary>
    Public Property CodeProduit As String = String.Empty

    ''' <summary>
    ''' Statut de la transaction dans le rapport de règlement (colonne TxnStatus) : "S" réglée,
    ''' "W" en attente. Conservé parce qu'une option décide du sort des envois en attente, et
    ''' que la décision doit rester lisible ligne par ligne.
    ''' </summary>
    Public Property Statut As String = String.Empty

    ''' <summary>Devise des montants LOC du rapport (colonne LOCCurrencyCode).</summary>
    Public Property DeviseLocale As String = String.Empty

    ''' <summary>Nom du rapport d'où la ligne vient, pour l'archivage de la pièce.</summary>
    Public Property FichierSource As String = String.Empty

#End Region

#Region "Les trois grandeurs posées"

    ''' <summary>
    ''' Montant en monnaie locale tel que le rapport l'exprime : RecPrincipalREC pour un envoi,
    ''' ClearPrincipalPAY pour un paiement. C'est le montant que la banque a réellement encaissé
    ''' ou décaissé en FCFA.
    ''' </summary>
    Public Property MontantLocal As Decimal

    ''' <summary>
    ''' Montant de la même transaction en devise de règlement (EUR) : ClearPrincipalLOC + ClearFXLOC
    ''' pour un envoi, ClearPrincipalLOC seul pour un paiement.
    ''' </summary>
    Public Property MontantEnDevise As Decimal

    ''' <summary>
    ''' Parité employée pour convertir le montant en devise. Elle est portée par la LIGNE et non
    ''' par le service : une pièce conservée doit pouvoir être relue dans dix ans avec la parité
    ''' de son jour, et non avec celle d'aujourd'hui.
    ''' </summary>
    Public Property Parite As Decimal

    ''' <summary>Les deux composantes du montant en devise, conservées pour la vérification.</summary>
    Public Property ClearPrincipalLoc As Decimal

    ''' <summary>Voir <see cref="ClearPrincipalLoc"/>.</summary>
    Public Property ClearFxLoc As Decimal

#End Region

#Region "Ce qui s'en déduit"

    ''' <summary>
    ''' Contre-valeur en monnaie locale du montant en devise, ARRONDIE AU FRANC.
    '''
    ''' L'arrondi est posé ici, et à la transaction. L'appliquer au total aurait donné un
    ''' résultat juste au centime près et faux au détail : la pièce comptable porte des francs
    ''' entiers, et la somme des écarts arrondis n'est pas l'arrondi de la somme des écarts.
    ''' AwayFromZero parce que le franc CFA ne connaît pas l'arrondi bancaire : 0,5 monte.
    ''' </summary>
    Public ReadOnly Property ContreValeur As Decimal
        Get
            Return Math.Round(MontantEnDevise * Parite, 0, MidpointRounding.AwayFromZero)
        End Get
    End Property

    ''' <summary>
    ''' L'écart : montant local moins contre-valeur. Positif, la banque gagne ; négatif, elle perd.
    ''' </summary>
    Public ReadOnly Property Ecart As Decimal
        Get
            Return MontantLocal - ContreValeur
        End Get
    End Property

    ''' <summary>Le sens de l'écart. Voir <see cref="NatureEcartWU"/>.</summary>
    Public ReadOnly Property Nature As NatureEcartWU
        Get
            If Ecart > 0D Then Return NatureEcartWU.Gain
            If Ecart < 0D Then Return NatureEcartWU.Perte
            Return NatureEcartWU.Neutre
        End Get
    End Property

    ''' <summary>La nature, écrite pour l'écran.</summary>
    Public ReadOnly Property NatureLisible As String
        Get
            Select Case Nature
                Case NatureEcartWU.Gain : Return "Gain"
                Case NatureEcartWU.Perte : Return "Perte"
                Case Else : Return "Neutre"
            End Select
        End Get
    End Property

    ''' <summary>Le sens, écrit pour l'écran.</summary>
    Public ReadOnly Property SensLisible As String
        Get
            Select Case If(Sens, String.Empty).Trim().ToUpperInvariant()
                Case ConstantesWU.SENS_ENVOI : Return "Envoi"
                Case ConstantesWU.SENS_PAIEMENT : Return "Paiement"
                Case Else : Return If(Sens, String.Empty)
            End Select
        End Get
    End Property

#End Region

End Class

''' <summary>
''' Une ligne du rapport ÉCARTÉE du calcul, et la raison de son exclusion.
'''
''' POURQUOI JOURNALISER UNE EXCLUSION. Vingt-huit lignes sur deux mille quatre cents ne
''' sont pas entrées dans le calcul du rapport étudié. Sans ce journal, la banque voit un
''' total et doit nous croire. Avec lui, elle lit « ajustement FOR FULL REFUND », reconnaît
''' ses propres remboursements, et vérifie. Une exclusion muette est une exclusion
''' indéfendable.
''' </summary>
Public Class ExclusionChangeWU

    ''' <summary>Numéro de la ligne dans le rapport, en comptant l'en-tête pour 1.</summary>
    Public Property NumeroDeLigne As Integer

    ''' <summary>MTCN de la ligne écartée, quand elle en porte un.</summary>
    Public Property Mtcn As String = String.Empty

    ''' <summary>La raison, en français, telle qu'elle s'affiche.</summary>
    Public Property Motif As String = String.Empty

    ''' <summary>Les valeurs lues qui ont motivé l'exclusion (type, statut, devise).</summary>
    Public Property Detail As String = String.Empty

End Class
