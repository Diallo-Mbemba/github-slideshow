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

    ''' <summary>
    ''' Taux de conversion EUR vers FCFA (parité fixe). Il ne s'applique QUE lorsque les montants
    ''' LOC du rapport de règlement sont libellés en EUR : c'était le cas des rapports Western
    ''' Union jusqu'à la refonte de leur format (LOCCurrencyCode = EUR). Depuis, les montants LOC
    ''' sont directement en XAF (LOCCurrencyCode = XAF) et aucune conversion ne doit être appliquée.
    ''' La devise réelle de chaque ligne pilote ce choix : voir WUReportService.ObtenirFacteurConversion.
    ''' </summary>
    Public Const TAUX_CONVERSION As Decimal = 655.957D

    ''' <summary>Taux de commission sur les envois (20,5 %).</summary>
    Public Const TAUX_COMMISSION_ENVOI As Decimal = 0.205D

    ''' <summary>
    ''' Taux de TVA perçu sur les frais d'envoi : 19,25 %. CE TAUX NE CHANGE PAS.
    ''' C'est lui, et lui seul, qui définit le montant de TVA perçu ; il reste la source
    ''' dont TAUX_TVA_COMPTE_TVA est dérivé. Le modifier ici suffit à tout recalculer.
    ''' </summary>
    Public Const TAUX_TVA As Decimal = 0.1925D

    ''' <summary>
    ''' Part du taux de TVA qui n'est PAS portée au compte de TVA : 1,25 point.
    '''
    ''' La TVA perçue (19,25 % des frais d'envoi) ne va pas en totalité sur le compte de
    ''' TVA. La banque n'y porte que 18 points ; le 1,25 point restant n'est pas de la TVA
    ''' au sens comptable, c'est une taxe additionnelle. Il n'est pas retranché du solde de
    ''' taxes, il y demeure — et s'y ventile donc 25 % / 75 % comme le reste du solde, sur
    ''' la commission sur transfert et sur les impôts et taxes sur envoi.
    '''
    ''' Rien n'est perdu ni ajouté : sur des frais d'envoi de 522 000 F, les 100 485 F de
    ''' TVA perçue à 19,25 % se retrouvent intégralement dans la pièce, répartis sur trois
    ''' comptes — 93 960 F au compte de TVA, 4 894 F aux impôts et taxes sur envoi,
    ''' 1 631 F à la commission sur transfert.
    ''' </summary>
    Public Const TAUX_TVA_HORS_COMPTE_TVA As Decimal = 0.0125D

    ''' <summary>
    ''' Part du taux de TVA effectivement portée au compte de TVA : 18 %.
    ''' Valeur DÉRIVÉE, jamais saisie en dur : 19,25 % - 1,25 point. Écrire 0,18 ici
    ''' romprait le lien avec TAUX_TVA et laisserait les deux valeurs diverger le jour
    ''' où le taux de TVA changerait.
    ''' </summary>
    Public Const TAUX_TVA_COMPTE_TVA As Decimal = TAUX_TVA - TAUX_TVA_HORS_COMPTE_TVA

    ''' <summary>Taux de TTA (Taxe sur le Transfert d'Argent), appliqué à l'envoi comme à la réception (0,2 %).</summary>
    Public Const TAUX_TTA As Decimal = 0.002D

    ''' <summary>Quote-part de la commission sur transfert (25 % du solde de taxes).</summary>
    Public Const TAUX_COM_TRANSFERT As Decimal = 0.25D

    ''' <summary>Quote-part de la taxe sur envoi (75 % du solde de taxes).</summary>
    Public Const TAUX_TAXE_ENVOI As Decimal = 0.75D

#End Region

#Region "Pays de référence"

    ''' <summary>
    ''' Libellés acceptés pour le pays de paiement local dans le rapport de règlement.
    ''' Western Union a renommé "TCHAD" en "CHAD" lors de la refonte du format des rapports :
    ''' les deux graphies sont acceptées afin de pouvoir traiter indifféremment un fichier
    ''' à l'ancien ou au nouveau format. Comparaison insensible à la casse.
    ''' </summary>
    Public Shared ReadOnly PaysPaiementLocal As String() = {"CHAD", "TCHAD"}

    ''' <summary>Nom de la colonne portant la devise des montants LOC du rapport de règlement.</summary>
    Public Const COLONNE_DEVISE_LOC As String = "LOCCurrencyCode"

    ''' <summary>Code ISO du franc CFA (BEAC) : montants LOC déjà exprimés en FCFA, aucune conversion.</summary>
    Public Const DEVISE_FCFA As String = "XAF"

    ''' <summary>Code ISO de l'euro : montants LOC à convertir en FCFA via TAUX_CONVERSION.</summary>
    Public Const DEVISE_EURO As String = "EUR"

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
    ''' Compte inter bancaire, utilisé pour absorber un écart résiduel d'arrondi au niveau
    ''' GLOBAL de la pièce comptable (jamais Account par Account). Il remplaçait jusqu'ici un
    ''' compte d'attente fictif ("XXXXXXXXXX") faute de numéro connu : la ligne de paramétrage
    ''' de la table SystemeWU (colonnes Cpte_attenteDEBIT / Cpte_attenteCREDIT) et le formulaire
    ''' « Comptes Systèmes WU » de la Direction Comptable le désignent tous deux comme étant
    ''' le compte inter bancaire 381000101.
    ''' Cette valeur n'est plus qu'un DÉFAUT : le compte réellement utilisé est celui lu dans
    ''' SystemeWU au démarrage (voir ComptesSystemeWU.CompteInterBancaire).
    ''' </summary>
    Public Const CPT_ATTENTE As String = "381000101"

#End Region

#Region "Libellés comptables"

    Public Const LIB_COMPTE_COURANT As String = "COMPTE COURANT WESTERN UNION ETD"

    ' ------------------------------------------------------------------------------------
    '  Libellés du formulaire de pièce comptable de la banque
    '
    '  Repris tels quels du modèle fourni (classe_bis.xlsx), majuscules et espacement
    '  compris : ce document est visé et signé à la main, et une pièce qui ne ressemble pas
    '  aux autres se fait renvoyer au guichet de la comptabilité.
    ' ------------------------------------------------------------------------------------
    Public Const PIECE_BANQUE As String = "ECOBANK TCHAD"
    Public Const PIECE_TITRE As String = "VERIFICATION   PIECE  COMPTABLE"
    Public Const PIECE_DATE As String = "DATE :"
    Public Const PIECE_DE As String = "DE :"
    Public Const PIECE_POUR As String = "POUR :"
    Public Const PIECE_AGENCE As String = "AGENCE:"
    Public Const PIECE_COMPTES As String = "N° DE COMPTES"
    Public Const PIECE_LIBELLES As String = "LIBELLES"
    Public Const PIECE_MONTANTS As String = "MONTANTS"
    Public Const PIECE_DEBIT As String = "DEBIT :"
    Public Const PIECE_CREDIT As String = "CREDIT :"
    Public Const PIECE_RAISON As String = "RAISON :"

    ''' <summary>
    ''' Libellé de la ligne de total, posée sous le dernier crédit.
    '''
    ''' La pièce manuelle de la banque porte ce total ; la nôtre ne le portait pas. C'est lui
    ''' qui permet au vérificateur de conclure d'un coup d'œil, sans additionner douze lignes.
    ''' </summary>
    Public Const PIECE_TOTAL_CREDITS As String = "TOTAL DES CRÉDITS"
    Public Const PIECE_SIGNATURES As String = "SIGNATURES REQUISES"
    Public Const PIECE_FCU As String = "FCU"
    Public Const PIECE_INITIE As String = "INITIE PAR :"
    Public Const PIECE_CONTROLE As String = "CONTRÔLE PAR"
    Public Const PIECE_APPROUVE As String = "APPROUVE PAR:"
    Public Const PIECE_ECRITURE As String = "ECRITURE"
    Public Const PIECE_OPS As String = "OPS"
    Public Const PIECE_PASSEE As String = "PASSEE PAR:"
    Public Const PIECE_AUTORISEE As String = "AUTORISEE PAR:"
    Public Const PIECE_DATE_ENREGISTREMENT As String = "DATE D'ENREGISTREMENT COMPTABLE"
    Public Const PIECE_NUMERO_SEQUENCE As String = "NUMERO DE SEQUENCE ENREGISTREE"

    ' « DE : » et « POUR : » disent DE LA PART DE QUI et POUR QUI la pièce est établie.
    '
    ' Le modèle les laisse vides — il est fait pour être rempli à la main. Une pièce produite
    ' par l'application, elle, sait d'où elle vient : le service émetteur est complété du nom
    ' de l'utilisateur connecté, celui-là même qui a lancé la compense.
    Public Const PIECE_SERVICE_EMETTEUR As String = "COMPENSATION WESTERN UNION"
    Public Const PIECE_SERVICE_DESTINATAIRE As String = "COMPTABILITE"

    ' Agence émettrice portée par l'en-tête, en face de « AGENCE: ».
    '
    ' Elle VARIE d'une pièce à l'autre : chaque pièce de point de vente porte l'agence de
    ' rattachement de ce point de vente. Cette constante n'est que le dernier recours — la
    ' pièce globale, qui n'appartient à aucune agence, et le point de vente dont l'agence
    ' n'est pas retrouvée dans le référentiel.
    Public Const PIECE_AGENCE_DEFAUT As String = "N'Djamena"

    ' Police du formulaire. Deux polices, comme le modèle : les intitulés en Arial Black,
    ' le contenu en Century Schoolbook.
    Public Const PIECE_POLICE_TITRE As String = "Arial Black"
    Public Const PIECE_POLICE_CORPS As String = "Century Schoolbook"
    Public Const LIB_COMMISSION_TRANSFERT_BANQUE As String = "Commission sur Transfert_Ecobank"
    Public Const LIB_COMMISSION_PAIEMENT_BANQUE As String = "Commission sur Paiement_Ecobank"
    Public Const LIB_COMMISSION_ENVOI_BANQUE As String = "Commission sur Envoi_Ecobank"
    Public Const LIB_COMMISSION_TRANSFERT_SA As String = "Commission sur Transfert_Sous-agence"
    Public Const LIB_COMMISSION_PAIEMENT_SA As String = "Commission sur Paiement_Sous-agence"
    Public Const LIB_COMMISSION_ENVOI_SA As String = "Commission sur Envoi_Sous-agence"
    Public Const LIB_IMPOTS_TAXE_ENVOI As String = "IMPOTS ET TAXE SUR ENVOI"
    Public Const LIB_TVA As String = "TVA COLLECTEES WESTERN UNION"
    Public Const LIB_TTA_ENVOI As String = "TTA (TAXE SUR TRANSFER DE FONDS WU)"
    ' Double espace avant "DE" : reproduit fidèlement le libellé du classeur de référence
    ' PieceComptabilsationTchad.xlsx (colonne LIBELLES, ligne TTA Réception).
    Public Const LIB_TTA_RECEPTION As String = "TTA (TAXE SUR RECEPTION  DE FONDS WU)"
    Public Const LIB_ECART_ATTENTE As String = "ECART D'ARRONDI - COMPTE INTER BANCAIRE"

    ''' <summary>
    ''' Gabarit du libellé de la ligne de mouvement (activité) du point de vente, tel que
    ''' constaté dans le classeur de référence : "CCS_BOLOLO ACTIVITE WU" pour l'agence BOLOLO.
    ''' {0} est remplacé par la Designation de l'Account.
    ''' </summary>
    Public Const LIB_MOUVEMENT_ACTIVITE_FORMAT As String = "CCS_{0} ACTIVITE WU"

    ''' <summary>
    ''' Préfixe du gabarit ci-dessus, isolé pour pouvoir constater qu'il est DÉJÀ présent.
    '''
    ''' La plupart des désignations commencent par « CCS », et le gabarit produisait alors
    ''' « CCS_CCS NGARTA RUE DE 40M ACTIVITE WU ». Voir PieceComptableService.LibelleDuMouvement.
    ''' </summary>
    Public Const PREFIXE_CCS As String = "CCS"

    ''' <summary>Fin du gabarit, employée quand le préfixe est déjà porté par la désignation.</summary>
    Public Const LIB_MOUVEMENT_ACTIVITE_SUFFIXE As String = "ACTIVITE WU"

#End Region

#Region "Contrôle d'équilibrage"

    ''' <summary>Seuil de tolérance (en FCFA) au-delà duquel la génération de la pièce est bloquée.</summary>
    Public Const SEUIL_ECART_TOLERE As Decimal = 1000D

    ''' <summary>
    ''' Seuil (en FCFA) au-delà duquel l'écart d'arrondi d'UNE ligne (un Account) est considéré
    ''' comme anormal et mis en évidence dans la grille de contrôle. Un Account génère au plus
    ''' une dizaine de lignes arrondies : son écart d'arrondi légitime reste de quelques FCFA.
    ''' Un écart nettement supérieur trahit un paramétrage incomplet (typiquement un sous-agent
    ''' sans CompteCommission, dont les commissions ne peuvent donc pas être passées).
    ''' </summary>
    Public Const SEUIL_ECART_LIGNE_ANORMAL As Long = 10L

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

    ''' <summary>
    ''' Nom de la colonne portant le statut de la transaction dans le rapport d'activité.
    ''' Absente des rapports à l'ancien format : son absence n'est donc jamais bloquante.
    ''' </summary>
    Public Const COLONNE_STATUS_ACTIVITE As String = "STATUS"

    ''' <summary>
    ''' Statuts dont les lignes sont EXCLUES de l'agrégation du rapport d'activité.
    ''' "C" = transaction annulée (CANCELLED) : règle métier validée avec la Direction Comptable.
    ''' Les autres statuts rencontrés ("S" = réglée, "W" = en attente) sont conservés.
    ''' Comparaison insensible à la casse. Voir WUReportService.InclureLigneActivite.
    ''' </summary>
    Public Shared ReadOnly StatutsActiviteExclus As String() = {"C"}

    ''' <summary>
    ''' Formats de date acceptés dans les rapports, essayés dans cet ordre avant de retomber
    ''' sur une analyse selon la culture. "yyyyMMdd" est le format effectivement produit par
    ''' Western Union (ex. 20260530), à l'ancien comme au nouveau format de rapport.
    ''' </summary>
    Public Shared ReadOnly FormatsDateRapport As String() = {"yyyyMMdd", "dd/MM/yyyy", "yyyy-MM-dd"}

    ''' <summary>Nom de la colonne de date dans le rapport d'activité.</summary>
    Public Const COLONNE_DATE_ACTIVITE As String = "txnDateLOC"

    ''' <summary>
    ''' Nom de la colonne de date dans le rapport de règlement. Cette colonne n'existe en réalité
    ''' dans aucun format de rapport de règlement connu : la date y est reconstituée à partir des
    ''' colonnes Année/Mois/Jour (voir PrefixesDateReglement). Le nom est conservé pour le cas où
    ''' Western Union ajouterait une colonne de date simple à un format ultérieur.
    ''' </summary>
    Public Const COLONNE_DATE_REGLEMENT As String = "txnDateLOC"

    ''' <summary>
    ''' Préfixes des triplets de colonnes Année/Mois/Jour permettant de reconstituer la date du
    ''' rapport d'ACTIVITÉ lorsque la colonne de date simple est absente ou illisible.
    ''' "TXNDATELOC" -> TXNDATELOCYEAR / TXNDATELOCMONTH / TXNDATELOCDAY (nouveau format) ;
    ''' "TxnDate"    -> TxnDateYear / TxnDateMonth / TxnDateDay (ancien format).
    ''' Essayés dans l'ordre, recherche de colonne insensible à la casse.
    ''' </summary>
    Public Shared ReadOnly PrefixesDateActivite As String() = {"TXNDATELOC", "TxnDate"}

    ''' <summary>
    ''' Préfixes des triplets de colonnes Année/Mois/Jour permettant de reconstituer la date du
    ''' rapport de RÈGLEMENT, qui ne comporte aucune colonne de date simple.
    ''' "SetDateLOC" (date de règlement locale) puis "RepDate" (date d'édition du rapport) :
    ''' ces deux dates portent une valeur unique correspondant à la journée traitée, dans
    ''' l'ancien comme dans le nouveau format, ce qui en fait les seules comparables à la date
    ''' d'activité. Les dates d'émission (RecDateLOC) et de paiement (PayDateLOC) sont au
    ''' contraire étalées sur plusieurs jours et ne conviennent donc pas à ce contrôle.
    ''' </summary>
    Public Shared ReadOnly PrefixesDateReglement As String() = {"SetDateLOC", "RepDate"}

    ''' <summary>Suffixes composant un triplet de colonnes de date (Année, Mois, Jour).</summary>
    Public Const SUFFIXE_DATE_ANNEE As String = "Year"
    Public Const SUFFIXE_DATE_MOIS As String = "Month"
    Public Const SUFFIXE_DATE_JOUR As String = "Day"

#End Region

#Region "Contrôles de sécurité sur les fichiers chargés"

    ''' <summary>
    ''' Fragments recherchés dans le NOM du fichier pour reconnaître un rapport d'activité.
    ''' Couvrent les deux nomenclatures rencontrées : celle de Western Union
    ''' ("RSP_TD383_ACTIVITY_REPORT_BY_ACCOUNT_...") et l'intitulé français des anciens
    ''' rapports ("Rapport d'activité par Site ... du 02 Jan 2021"). La recherche est faite
    ''' sur un nom normalisé (minuscules, accents retirés), la ponctuation pouvant varier.
    ''' </summary>
    Public Shared ReadOnly MarqueursNomActivite As String() = {"activity", "activite"}

    ''' <summary>
    ''' Fragments recherchés dans le NOM du fichier pour reconnaître un rapport de règlement :
    ''' "settlement" (nomenclature Western Union) et "reglement" (anciens intitulés français).
    ''' </summary>
    Public Shared ReadOnly MarqueursNomReglement As String() = {"settlement", "reglement"}

    ''' <summary>
    ''' Colonnes présentes dans TOUT rapport d'activité (ancien comme nouveau format) et dans
    ''' AUCUN rapport de règlement : elles constituent donc la signature permettant d'identifier
    ''' le type réel d'un fichier d'après son CONTENU, indépendamment de son nom — un fichier
    ''' pouvant toujours être renommé. Déterminées par comparaison des en-têtes des rapports
    ''' des 02/01/2021 et 30/05/2026. Comparaison insensible à la casse.
    ''' </summary>
    Public Shared ReadOnly SignatureRapportActivite As String() = {
        "TaxesREC", "TaxesPAY", "PayPrincipalPAY", "txnDateLOC"
    }

    ''' <summary>
    ''' Colonnes présentes dans TOUT rapport de règlement et dans AUCUN rapport d'activité.
    ''' Même méthode de détermination que SignatureRapportActivite.
    ''' </summary>
    Public Shared ReadOnly SignatureRapportReglement As String() = {
        "TransactionType", "PayCountry", "ClearChargesLOC", "ClearFXLOC", "SetDateLOCYear"
    }

    ''' <summary>
    ''' Nombre minimum de colonnes de signature devant être présentes pour conclure au type.
    ''' Fixé à 2 et non à la totalité : ainsi la détection résiste à la disparition d'une
    ''' colonne lors d'une future évolution du format, sans risque de confusion puisque aucune
    ''' de ces colonnes n'existe dans le rapport de l'autre type.
    ''' </summary>
    Public Const MIN_COLONNES_SIGNATURE As Integer = 2

    ''' <summary>
    ''' Mois acceptés dans le nom des anciens rapports ("... du 02 Jan 2021"), en français
    ''' comme en anglais, sous forme abrégée ou complète. Les formes les plus longues sont
    ''' placées en premier afin que "juillet" ne soit jamais confondu avec "juin". Le texte
    ''' comparé est préalablement normalisé (minuscules, accents retirés).
    ''' </summary>
    Public Shared ReadOnly MoisNommes As String() = {
        "janvier", "janv", "jan",
        "fevrier", "fevr", "fev", "feb",
        "mars", "mar",
        "avril", "avr", "apr",
        "mai", "may",
        "juillet", "juil", "jul",
        "juin", "jun",
        "aout", "aug", "aou",
        "septembre", "sept", "sep",
        "octobre", "oct",
        "novembre", "nov",
        "decembre", "dec"
    }

    ''' <summary>Numéro du mois associé à chaque entrée de MoisNommes, dans le même ordre.</summary>
    Public Shared ReadOnly NumerosMoisNommes As Integer() = {
        1, 1, 1,
        2, 2, 2, 2,
        3, 3,
        4, 4, 4,
        5, 5,
        7, 7, 7,
        6, 6,
        8, 8, 8,
        9, 9, 9,
        10, 10,
        11, 11,
        12, 12
    }

#End Region


#Region "Fichier d'interface vers le core banking"

    ' Le fichier chargé dans le core banking compte treize colonnes, dont neuf sont constantes
    ' ou déduites. Les valeurs ci-dessous sont celles de la banque : elles ne se devinent pas et
    ' ne doivent pas être modifiées sans son accord.

    Public Const CB_DETBSJRNL As String = "BBR"
    Public Const CB_BRN As String = "N01"
    Public Const CB_SRCCODE As String = "ECOSOURCE"
    Public Const CB_COST_CENTER As String = "10000"

    ''' <summary>Code transaction d'un débit.</summary>
    Public Const CB_TXNCD_DEBIT As String = "U24"

    ''' <summary>Code transaction d'un crédit.</summary>
    Public Const CB_TXNCD_CREDIT As String = "F15"

    ''' <summary>Sens débit, tel qu'attendu dans la colonne DRCR.</summary>
    Public Const CB_SENS_DEBIT As String = "D"

    ''' <summary>Sens crédit, tel qu'attendu dans la colonne DRCR.</summary>
    Public Const CB_SENS_CREDIT As String = "C"

    ''' <summary>
    ''' Agence du siège. Elle est portée par les deux comptes ci-dessous quel que soit le point
    ''' de vente qui les a mouvementés, et par les lignes qui ne se rattachent à aucun point de
    ''' vente — la ligne d'écart d'arrondi, notamment.
    ''' </summary>
    Public Const CB_AGENCE_SIEGE As String = "N01"

    ''' <summary>
    ''' Comptes rattachés d'office à l'agence du siège dans la colonne ACBRN.
    '''
    ''' Ils sont écrits ici et non dans SystemeWU : la banque les a donnés tels quels, et cette
    ''' liste ne décrit pas un paramétrage comptable mais une règle d'aiguillage propre au
    ''' format du core banking.
    ''' </summary>
    Public Shared ReadOnly CB_COMPTES_SIEGE As String() = New String() {"379100319", "379200585"}

    ''' <summary>
    ''' Origine du compteur de jours qui sert de numéro de lot. Choisie pour que les numéros
    ''' produits aujourd'hui aient la même forme que ceux de la banque : au 22 avril 2026, le
    ''' compteur vaut « 07p1 ».
    ''' </summary>
    Public Shared ReadOnly CB_ORIGINE_LOT As Date = New Date(1999, 1, 1)

    ''' <summary>Longueur du numéro de lot, en caractères.</summary>
    Public Const CB_LONGUEUR_LOT As Integer = 4

#End Region

End Class
