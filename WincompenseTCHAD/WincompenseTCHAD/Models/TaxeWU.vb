Option Strict On
Option Explicit On

''' <summary>
''' Le montant auquel le taux d'une taxe s'applique.
'''
''' IL Y EN A QUATRE, ET C'EST TOUT L'INTÉRÊT DE CETTE ÉNUMÉRATION. L'écran centrafricain
''' dont celui-ci s'inspire annonce « les taux s'appliquent à la CHARGE, jamais au
''' principal » : c'est vrai là-bas, c'est faux ici. Au Tchad, la TVA porte sur les charges
''' d'envoi, les deux TTA portent sur le PRINCIPAL, et les deux quotes-parts portent sur le
''' SOLDE DE TAXES. Écrire un barème sans dire l'assiette reviendrait à afficher cinq taux
''' que rien ne permettrait de vérifier.
''' </summary>
Public Enum AssietteWU

    ''' <summary>Charges d'envoi : les frais facturés au client. Assiette de la TVA.</summary>
    ChargeEnvoi = 1

    ''' <summary>Montant nominal envoyé. Assiette de la TTA sur envoi.</summary>
    PrincipalEnvoi = 2

    ''' <summary>Montant nominal payé. Assiette de la TTA sur réception.</summary>
    PrincipalPaye = 3

    ''' <summary>
    ''' Solde de taxes, c'est-à-dire le RÉSIDU : total des taxes lu dans le rapport, moins la
    ''' TVA, moins la TTA sur envoi. Assiette des deux quotes-parts, 25 % et 75 %.
    ''' </summary>
    SoldeTaxes = 4

End Enum

''' <summary>
''' Une ligne du barème : une taxe, son taux, son assiette et le compte où elle atterrit.
'''
''' CET OBJET DÉCRIT, IL NE CALCULE PAS LA PIÈCE. Le calcul de la compensation Western Union
''' reste entier dans WUCalculationService, à partir des constantes de ConstantesWU, et rien
''' ici ne l'alimente. C'est la règle posée par la banque : le barème Western Union est
''' réconcilié, il ne se modifie pas depuis un écran. Les produits ajoutés plus tard — Ria
''' d'abord — auront le leur, lu en base et modifiable ; ils emploieront alors la méthode
''' <see cref="Montant"/>, qui est ici la seule part calculatoire et qui n'attend qu'eux.
''' </summary>
Public Class TaxeWU

#Region "Identité"

    ''' <summary>
    ''' Code court, stable, non traduit. Il rattache la ligne à la formule qui l'emploie :
    ''' le libellé peut changer, le code non.
    ''' </summary>
    Public Property Code As String = String.Empty

    ''' <summary>Nom affiché. Celui-ci se change librement : il n'entre dans aucun calcul.</summary>
    Public Property Libelle As String = String.Empty

    ''' <summary>Rang d'affichage dans l'écran du barème.</summary>
    Public Property Ordre As Integer = 0

#End Region

#Region "Calcul"

    ''' <summary>Taux appliqué à l'assiette. 0,1925 pour 19,25 %.</summary>
    Public Property Taux As Decimal = 0D

    ''' <summary>Montant auquel le taux s'applique.</summary>
    Public Property Assiette As AssietteWU = AssietteWU.ChargeEnvoi

    ''' <summary>
    ''' Montant de la taxe pour une assiette donnée. Aucun arrondi : il n'intervient qu'à
    ''' l'écriture de la pièce, jamais dans un calcul intermédiaire.
    ''' </summary>
    ''' <param name="montantBase">Valeur de l'assiette, en FCFA.</param>
    Public Function Montant(montantBase As Decimal) As Decimal
        Return montantBase * Taux
    End Function

#End Region

#Region "Comptabilisation"

    ''' <summary>Compte comptable crédité, tel que le paramétrage en service le désigne.</summary>
    Public Property Compte As String = String.Empty

    ''' <summary>
    ''' La taxe pose-t-elle sa propre ligne sur la pièce ? Une taxe calculée mais sans ligne
    ''' se retrouverait par différence sur la contrepartie. Les cinq lignes du Tchad en
    ''' posent une ; la question est ouverte pour les produits à venir.
    ''' </summary>
    Public Property LignePosee As Boolean = True

    ''' <summary>
    ''' Vrai si la taxe n'est due que par un sous-agent.
    '''
    ''' UNE SEULE LIGNE LE PORTE : la TTA sur réception. Règle confirmée par la banque —
    ''' une agence propre ne retient pas de TTA sur paiement. Elle n'est pas calculée puis
    ''' omise, elle n'est pas due : voir WUCalculationService.AppliquerFormules.
    ''' </summary>
    Public Property SousAgentSeulement As Boolean = False

    ''' <summary>
    ''' Vrai si le compte crédité est un compte de PRODUIT et non de taxe.
    '''
    ''' La commission sur transfert — 25 % du solde — atterrit sur 728300148, un compte de
    ''' produit bancaire. Elle figure dans ce barème parce qu'elle se calcule sur le solde de
    ''' taxes comme sa jumelle à 75 %, et qu'on ne comprend ni l'une ni l'autre séparément.
    ''' Mais un écran nommé « Taxes » qui pilote une recette doit le dire, sans quoi un
    ''' comptable qui le relit se trompera sur la nature de la ligne.
    ''' </summary>
    Public Property EstUnProduitBancaire As Boolean = False

#End Region

#Region "Présentation"

    ''' <summary>L'assiette en clair, pour l'écran.</summary>
    Public ReadOnly Property LibelleDeLAssiette As String
        Get
            Select Case Assiette
                Case AssietteWU.ChargeEnvoi : Return "charges d'envoi"
                Case AssietteWU.PrincipalEnvoi : Return "principal envoyé"
                Case AssietteWU.PrincipalPaye : Return "principal payé"
                Case AssietteWU.SoldeTaxes : Return "solde de taxes"
                Case Else : Return String.Empty
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Le taux en pourcentage, sans zéro inutile : « 19,25 % », « 0,2 % », « 25 % ».
    ''' Trois décimales suffisent — le plus fin des taux du Tchad est le millième.
    ''' </summary>
    Public Function TauxEnTexte() As String

        Dim pourcentage As Decimal = Math.Round(Taux * 100D, 3)
        Return pourcentage.ToString("0.###", Globalization.CultureInfo.CurrentCulture) & " %"
    End Function

#End Region

#Region "Services"

    ''' <summary>Copie indépendante, pour éditer sans toucher au barème en service.</summary>
    Public Function Copier() As TaxeWU

        Return New TaxeWU() With {
            .Code = Code,
            .Libelle = Libelle,
            .Ordre = Ordre,
            .Taux = Taux,
            .Assiette = Assiette,
            .Compte = Compte,
            .LignePosee = LignePosee,
            .SousAgentSeulement = SousAgentSeulement,
            .EstUnProduitBancaire = EstUnProduitBancaire
        }
    End Function

#End Region

End Class
