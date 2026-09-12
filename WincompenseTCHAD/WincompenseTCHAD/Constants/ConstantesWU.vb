Option Strict On
Option Explicit On

''' <summary>
''' Constantes métier utilisées pour la compensation Western Union J+1 (Tchad).
''' Toutes les valeurs (taux, comptes comptables, libellés) sont centralisées ici
''' afin d'éviter les "valeurs magiques" dans le code métier et de faciliter
''' toute évolution future validée par la Direction Comptable.
''' </summary>
Public NotInheritable Class ConstantesWU

    Private Sub New()
        ' Classe statique : instanciation interdite.
    End Sub

#Region "Taux de conversion et de commissions"

    ''' <summary>Taux de conversion des montants en devise locale de règlement (LOC) vers le FCFA.</summary>
    Public Const TAUX_CONVERSION As Decimal = 655.957D

    ''' <summary>Taux de commission sur les envois (20,5 %).</summary>
    Public Const TAUX_COMMISSION_ENVOI As Decimal = 0.205D

    ''' <summary>Taux de TVA appliqué sur les frais d'envoi (19,25 %).</summary>
    Public Const TAUX_TVA As Decimal = 0.1925D

    ''' <summary>Taux de TTA (Taxe sur le Transfert d'Argent), appliqué à l'envoi comme à la réception (0,2 %).</summary>
    Public Const TAUX_TTA As Decimal = 0.002D

    ''' <summary>Quote-part de la commission sur transfert (25 % du solde de taxes).</summary>
    Public Const TAUX_COM_TRANSFERT As Decimal = 0.25D

    ''' <summary>Quote-part de la taxe sur envoi (75 % du solde de taxes).</summary>
    Public Const TAUX_TAXE_ENVOI As Decimal = 0.75D

#End Region

#Region "Pays de référence"

    ''' <summary>Pays de paiement retenu pour le calcul de la commission de paiement (règlement).</summary>
    Public Const PAYS_TCHAD As String = "TCHAD"

#End Region

#Region "Comptes comptables connus (banque)"

    ''' <summary>Compte courant / bilan de compensation Western Union.</summary>
    Public Const CPT_COMPTE_COURANT As String = "32100003292"

    ''' <summary>Commission sur Transfert / Envoi, part banque.</summary>
    Public Const CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE As String = "728300148"

    ''' <summary>Commission sur Paiement, part banque.</summary>
    Public Const CPT_COMMISSION_PAIEMENT_BANQUE As String = "728300149"

    ''' <summary>Impôts et taxe sur envoi.</summary>
    Public Const CPT_IMPOTS_TAXE_ENVOI As String = "434000147"

    ''' <summary>TVA collectée Western Union.</summary>
    Public Const CPT_TVA_COLLECTEE As String = "434000104"

    ''' <summary>TTA sur envoi de fonds.</summary>
    Public Const CPT_TTA_ENVOI As String = "434000145"

    ''' <summary>TTA sur réception de fonds.</summary>
    Public Const CPT_TTA_RECEPTION As String = "434000159"

    ''' <summary>
    ''' Compte d'attente provisoire utilisé pour absorber un écart résiduel d'arrondi
    ''' au niveau GLOBAL de la pièce comptable (jamais Account par Account).
    ''' À CONFIRMER / REMPLACER par le numéro de compte réel avant mise en production.
    ''' </summary>
    Public Const CPT_ATTENTE As String = "XXXXXXXXXX"

#End Region

#Region "Libellés comptables"

    Public Const LIB_COMPTE_COURANT As String = "COMPTE COURANT WESTERN UNION ETD"
    Public Const LIB_COMMISSION_TRANSFERT_BANQUE As String = "Commission sur Transfert_Ecobank"
    Public Const LIB_COMMISSION_PAIEMENT_BANQUE As String = "Commission sur Paiement_Ecobank"
    Public Const LIB_COMMISSION_ENVOI_BANQUE As String = "Commission sur Envoi_Ecobank"
    Public Const LIB_COMMISSION_TRANSFERT_SA As String = "Commission sur Transfert_Sous-agence"
    Public Const LIB_COMMISSION_PAIEMENT_SA As String = "Commission sur Paiement_Sous-agence"
    Public Const LIB_COMMISSION_ENVOI_SA As String = "Commission sur Envoi_Sous-agence"
    Public Const LIB_IMPOTS_TAXE_ENVOI As String = "IMPOTS ET TAXE SUR ENVOI"
    Public Const LIB_TVA As String = "TVA COLLECTEES WESTERN UNION"
    Public Const LIB_TTA_ENVOI As String = "TTA (TAXE SUR TRANSFER DE FONDS WU)"
    Public Const LIB_TTA_RECEPTION As String = "TTA (TAXE SUR RECEPTION DE FONDS WU)"
    Public Const LIB_ECART_ATTENTE As String = "ECART D'ARRONDI - COMPTE D'ATTENTE PROVISOIRE"

#End Region

#Region "Contrôle d'équilibrage"

    ''' <summary>Seuil de tolérance (en FCFA) au-delà duquel la génération de la pièce est bloquée.</summary>
    Public Const SEUIL_ECART_TOLERE As Decimal = 1000D

#End Region

#Region "Colonnes attendues dans les rapports"

    ''' <summary>Colonnes obligatoires du rapport d'activité Western Union.</summary>
    Public Shared ReadOnly ColonnesRapportActivite As String() = {
        "Account", "SendPayIndicator", "RecPrincipalREC", "TotalChargesREC",
        "TaxesREC", "TaxesPAY", "PayPrincipalPAY"
    }

    ''' <summary>Colonnes obligatoires du rapport de règlement Western Union.</summary>
    Public Shared ReadOnly ColonnesRapportReglement As String() = {
        "Account", "TransactionType", "PayCountry", "ClearChargesLOC",
        "ClearFXLOC", "SendPayIndicator"
    }

    ''' <summary>Nom de la colonne de date dans le rapport d'activité.</summary>
    Public Const COLONNE_DATE_ACTIVITE As String = "txnDateLOC"

    ''' <summary>
    ''' Nom présumé de la colonne de date dans le rapport de règlement.
    ''' À CONFIRMER : si cette colonne n'existe pas sous ce nom, la comparaison
    ''' de cohérence des dates est simplement ignorée (avec avertissement).
    ''' </summary>
    Public Const COLONNE_DATE_REGLEMENT As String = "txnDateLOC"

#End Region

End Class
