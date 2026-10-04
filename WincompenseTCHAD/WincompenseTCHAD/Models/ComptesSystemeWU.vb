Option Strict On
Option Explicit On

''' <summary>
''' Comptes comptables utilisés par la pièce comptable Western Union.
'''
''' Ces comptes ne sont plus figés dans le code : ils sont paramétrés dans la table SQL Server
''' SystemeWU et modifiables depuis le formulaire FrmComptesSysteme. Les constantes de
''' ConstantesWU en restent les valeurs PAR DÉFAUT, utilisées tant que la base n'a pas été lue
''' (ou si elle est inaccessible) : le comportement de l'application est ainsi strictement
''' identique à celui d'avant le paramétrage lorsque la table n'est pas exploitable.
'''
''' Correspondance avec les colonnes de la table SystemeWU (établie d'après le formulaire
''' « Comptes Systèmes WU » de la Direction Comptable et la ligne de paramétrage existante) :
'''
'''   Compte courant Western Union ETD ....... Cpte_PositionNette
'''   Compte inter bancaire .................. Cpte_attenteDEBIT / Cpte_attenteCREDIT
'''   Commission sur Transfert_Ecobank ....... Cpte_Produit
'''   Commission sur Envoi_Ecobank ........... Cpte_Produit_Envoi
'''   Commission sur Paiement_Ecobank ........ Cpte_Produit_Paiement
'''   Impôts et taxe sur envoi ............... Tthu
'''   TVA .................................... Tob
'''   TTA sur envoi WU ....................... Cpte_Envoi
'''   TTA sur paiement WU .................... Cpte_Paiement
'''   Gain de change ......................... Cpte_Gainde_Change
'''   Perte de change ........................ Cpte_Pertede_Change
'''
''' CPTE_GAINDE_CHANGE A CHANGÉ DE STATUT. Jusqu'à la pièce de change, elle figurait parmi
''' les colonnes que l'application ne lisait pas : la table la portait depuis l'origine, à
''' NULL, la banque ayant réservé la place. Elle est désormais lue et écrite, et
''' Cpte_Pertede_Change — ajoutée par le script 21 — l'accompagne.
'''
''' Les autres colonnes de SystemeWU (Passif, Actif, Cpte_Charge_Publicitaire,
''' Cpte_Envoi_agence, Cpte_Paiement_agence) ne sont toujours NI lues NI écrites : elles ne
''' concernent pas la pièce comptable Western Union et restent la propriété de ce qui les
''' utilise par ailleurs.
''' </summary>
Public Class ComptesSystemeWU

#Region "Comptes paramétrables (champs du formulaire)"

    ''' <summary>Compte courant / bilan de compensation Western Union (colonne Cpte_PositionNette).</summary>
    Public Property CompteCourant As String = ConstantesWU.CPT_COMPTE_COURANT

    ''' <summary>
    ''' Compte inter bancaire (colonnes Cpte_attenteDEBIT et Cpte_attenteCREDIT), soit
    ''' 381000101 « VIREMENTS INTERBANCAIRES ÉMISES » dans le paramétrage actuel.
    '''
    ''' Il sert à DEUX choses, et c'est bien le même compte dans les livres de la banque :
    '''   - il porte la ligne de mouvement des AGENCES PROPRES, qui n'ont pas de compte de
    '''     compensation à elles (voir PieceComptableService.GenererPieceComptable) ;
    '''   - il absorbe l'écart d'arrondi résiduel au niveau GLOBAL de la pièce comptable
    '''     (voir VerifierEquilibrePiece).
    '''
    ''' Le jour où la banque distinguerait ces deux usages, il suffirait d'ajouter ici une
    ''' seconde propriété : la pièce ne lit ce numéro qu'à ces deux endroits.
    ''' </summary>
    Public Property CompteInterBancaire As String = ConstantesWU.CPT_ATTENTE

    ''' <summary>Commission sur Transfert, part banque (colonne Cpte_Produit).</summary>
    Public Property CommissionTransfertBanque As String = ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE

    ''' <summary>Commission sur Envoi, part banque (colonne Cpte_Produit_Envoi).</summary>
    Public Property CommissionEnvoiBanque As String = ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE

    ''' <summary>Commission sur Paiement, part banque (colonne Cpte_Produit_Paiement).</summary>
    Public Property CommissionPaiementBanque As String = ConstantesWU.CPT_COMMISSION_PAIEMENT_BANQUE

    ''' <summary>Impôts et taxe sur envoi (colonne Tthu).</summary>
    Public Property ImpotsTaxeEnvoi As String = ConstantesWU.CPT_IMPOTS_TAXE_ENVOI

    ''' <summary>TVA collectée Western Union (colonne Tob).</summary>
    Public Property TVACollectee As String = ConstantesWU.CPT_TVA_COLLECTEE

    ''' <summary>TTA sur envoi de fonds (colonne Cpte_Envoi).</summary>
    Public Property TTAEnvoi As String = ConstantesWU.CPT_TTA_ENVOI

    ''' <summary>TTA sur réception / paiement de fonds (colonne Cpte_Paiement).</summary>
    Public Property TTAReception As String = ConstantesWU.CPT_TTA_RECEPTION

    ''' <summary>
    ''' Compte de GAIN de change (colonne Cpte_Gainde_Change).
    '''
    ''' CETTE COLONNE EXISTAIT DÉJÀ, ET N'ÉTAIT LUE PAR PERSONNE. La table SystemeWU la porte
    ''' depuis l'origine, à NULL : la banque avait réservé la place pour le jour où les écarts
    ''' de change seraient comptabilisés. Ce jour est venu, et le paramétrage n'a donc pas
    ''' changé de forme — il s'est rempli.
    '''
    ''' PAS DE VALEUR PAR DÉFAUT, contrairement à tous les comptes au-dessus. Les autres ont
    ''' une constante de repli parce que leur numéro est connu, réconcilié et en service. Ce
    ''' compte-ci, non : l'inventer reviendrait à comptabiliser un demi-million de francs sur
    ''' un numéro choisi par le développeur. Vide, il bloque la pièce de change et ne bloque
    ''' qu'elle.
    ''' </summary>
    Public Property CompteGainDeChange As String = String.Empty

    ''' <summary>
    ''' Compte de PERTE de change (colonne Cpte_Pertede_Change).
    '''
    ''' Celle-là n'existait pas : le script 21 l'ajoute. Même règle que le gain — aucune valeur
    ''' par défaut, et son absence ne bloque que la pièce de change.
    ''' </summary>
    Public Property ComptePerteDeChange As String = String.Empty

#End Region

#Region "Identification de la ligne de paramétrage"

    ''' <summary>
    ''' Valeur de la colonne « code » de la ligne lue dans SystemeWU. Sert à cibler exactement
    ''' cette ligne lors de l'enregistrement, la table pouvant en contenir plusieurs.
    ''' Vide si la ligne lue ne porte pas de code, ou si les valeurs par défaut sont utilisées.
    ''' </summary>
    Public Property CodeParametrage As String = String.Empty

    ''' <summary>
    ''' Vrai si ces comptes proviennent réellement de la table SystemeWU. Faux s'il s'agit des
    ''' valeurs par défaut du code (base inaccessible, table absente ou vide) : dans ce cas
    ''' l'enregistrement est impossible et l'utilisateur doit en être averti.
    ''' </summary>
    Public Property ChargeDepuisBase As Boolean = False

#End Region

#Region "Services"

    ''' <summary>
    ''' Comptes effectivement utilisés par l'application. Initialisés avec les valeurs par
    ''' défaut du code, puis remplacés au démarrage par ceux de la base (voir
    ''' WURepository.ChargerComptesSysteme). Point d'accès unique : la pièce comptable ne
    ''' référence plus jamais directement les constantes de comptes.
    ''' </summary>
    Public Shared Property Actuels As New ComptesSystemeWU()

    ''' <summary>Copie indépendante, pour éditer les comptes sans toucher à ceux en service tant que l'utilisateur n'a pas enregistré.</summary>
    Public Function Copier() As ComptesSystemeWU
        Return New ComptesSystemeWU() With {
            .CompteCourant = CompteCourant,
            .CompteInterBancaire = CompteInterBancaire,
            .CommissionTransfertBanque = CommissionTransfertBanque,
            .CommissionEnvoiBanque = CommissionEnvoiBanque,
            .CommissionPaiementBanque = CommissionPaiementBanque,
            .ImpotsTaxeEnvoi = ImpotsTaxeEnvoi,
            .TVACollectee = TVACollectee,
            .TTAEnvoi = TTAEnvoi,
            .TTAReception = TTAReception,
            .CompteGainDeChange = CompteGainDeChange,
            .ComptePerteDeChange = ComptePerteDeChange,
            .CodeParametrage = CodeParametrage,
            .ChargeDepuisBase = ChargeDepuisBase
        }
    End Function

    ''' <summary>
    ''' Vérifie que tous les comptes sont renseignés. Retourne la liste des libellés des champs
    ''' vides — vide si le paramétrage est complet. Un compte vide rendrait la pièce comptable
    ''' inexploitable : les écritures correspondantes partiraient sans numéro de compte.
    ''' </summary>
    Public Function ComptesManquants() As List(Of String)

        Dim manquants As New List(Of String)

        If String.IsNullOrWhiteSpace(CompteCourant) Then manquants.Add("Compte courant Western Union ETD")
        If String.IsNullOrWhiteSpace(CompteInterBancaire) Then manquants.Add("Compte inter bancaire")
        If String.IsNullOrWhiteSpace(CommissionTransfertBanque) Then manquants.Add("Commission sur Transfert_Ecobank")
        If String.IsNullOrWhiteSpace(CommissionEnvoiBanque) Then manquants.Add("Commission sur Envoi_Ecobank")
        If String.IsNullOrWhiteSpace(CommissionPaiementBanque) Then manquants.Add("Commission sur Paiement_Ecobank")
        If String.IsNullOrWhiteSpace(ImpotsTaxeEnvoi) Then manquants.Add("Impôts et taxe sur envoi")
        If String.IsNullOrWhiteSpace(TVACollectee) Then manquants.Add("TVA")
        If String.IsNullOrWhiteSpace(TTAEnvoi) Then manquants.Add("TTA sur envoi WU")
        If String.IsNullOrWhiteSpace(TTAReception) Then manquants.Add("TTA sur paiement WU")

        Return manquants
    End Function

    ''' <summary>
    ''' Les comptes de CHANGE manquants, séparément des autres.
    '''
    ''' POURQUOI UNE SECONDE FONCTION, ET NON DEUX LIGNES DANS ComptesManquants. Parce que
    ''' ComptesManquants BLOQUE LA PIÈCE PRINCIPALE, celle qui tourne en production depuis des
    ''' mois. Y ajouter les deux comptes de change aurait arrêté la compense du jour au
    ''' lendemain de la livraison, pour une pièce que la banque n'a pas encore paramétrée —
    ''' et pour une fonctionnalité dont elle n'a pas encore besoin ce matin-là.
    '''
    ''' Les deux paramétrages sont donc indépendants, et chacun ne bloque que sa propre pièce.
    ''' </summary>
    Public Function ComptesDeChangeManquants() As List(Of String)

        Dim manquants As New List(Of String)

        If String.IsNullOrWhiteSpace(CompteGainDeChange) Then manquants.Add("Gain de change")
        If String.IsNullOrWhiteSpace(ComptePerteDeChange) Then manquants.Add("Perte de change")

        ' Le compte de liaison n'est pas un compte propre au change : c'est le compte courant
        ' Western Union, déjà exigé par la pièce principale. Il est vérifié ici quand même,
        ' parce que la pièce de change s'en sert et qu'un diagnostic doit nommer TOUT ce qui
        ' lui manque, et non laisser découvrir le second défaut après avoir corrigé le premier.
        If String.IsNullOrWhiteSpace(CompteCourant) Then manquants.Add("Compte courant Western Union ETD (compte de liaison)")

        Return manquants
    End Function

    ''' <summary>Retire les espaces parasites de tous les comptes (saisie manuelle).</summary>
    Public Sub Normaliser()
        CompteCourant = If(CompteCourant, String.Empty).Trim()
        CompteInterBancaire = If(CompteInterBancaire, String.Empty).Trim()
        CommissionTransfertBanque = If(CommissionTransfertBanque, String.Empty).Trim()
        CommissionEnvoiBanque = If(CommissionEnvoiBanque, String.Empty).Trim()
        CommissionPaiementBanque = If(CommissionPaiementBanque, String.Empty).Trim()
        ImpotsTaxeEnvoi = If(ImpotsTaxeEnvoi, String.Empty).Trim()
        TVACollectee = If(TVACollectee, String.Empty).Trim()
        TTAEnvoi = If(TTAEnvoi, String.Empty).Trim()
        TTAReception = If(TTAReception, String.Empty).Trim()
        CompteGainDeChange = If(CompteGainDeChange, String.Empty).Trim()
        ComptePerteDeChange = If(ComptePerteDeChange, String.Empty).Trim()
    End Sub

    ''' <summary>
    ''' Liste des comptes qui ne sont pas exclusivement numériques. Tous les comptes du plan
    ''' comptable en service sont des suites de chiffres : un compte saisi autrement est
    ''' probablement une faute de frappe (espace, lettre), sans que ce soit une certitude —
    ''' l'utilisateur est donc averti, jamais bloqué.
    ''' </summary>
    Public Function ComptesNonNumeriques() As List(Of String)

        Dim suspects As New List(Of String)

        AjouterSiNonNumerique(suspects, "Compte courant Western Union ETD", CompteCourant)
        AjouterSiNonNumerique(suspects, "Compte inter bancaire", CompteInterBancaire)
        AjouterSiNonNumerique(suspects, "Commission sur Transfert_Ecobank", CommissionTransfertBanque)
        AjouterSiNonNumerique(suspects, "Commission sur Envoi_Ecobank", CommissionEnvoiBanque)
        AjouterSiNonNumerique(suspects, "Commission sur Paiement_Ecobank", CommissionPaiementBanque)
        AjouterSiNonNumerique(suspects, "Impôts et taxe sur envoi", ImpotsTaxeEnvoi)
        AjouterSiNonNumerique(suspects, "TVA", TVACollectee)
        AjouterSiNonNumerique(suspects, "TTA sur envoi WU", TTAEnvoi)
        AjouterSiNonNumerique(suspects, "TTA sur paiement WU", TTAReception)
        AjouterSiNonNumerique(suspects, "Gain de change", CompteGainDeChange)
        AjouterSiNonNumerique(suspects, "Perte de change", ComptePerteDeChange)

        Return suspects
    End Function

    Private Shared Sub AjouterSiNonNumerique(suspects As List(Of String), libelle As String, compte As String)

        If String.IsNullOrWhiteSpace(compte) Then Return ' Cas déjà traité par ComptesManquants.

        For Each caractere As Char In compte
            If Not Char.IsDigit(caractere) Then
                suspects.Add($"{libelle} : « {compte} »")
                Return
            End If
        Next
    End Sub

#End Region

End Class
