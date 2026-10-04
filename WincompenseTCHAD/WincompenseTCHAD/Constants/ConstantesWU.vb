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
    '''
    ''' Il porte en outre la ligne de mouvement des AGENCES PROPRES, sous le libellé
    ''' « VIREMENTS INTERBANCAIRES ÉMISES » sur la pièce manuelle de la banque : une agence
    ''' propre n'ayant pas de compte de compensation, son mouvement passe d'une agence à une
    ''' autre à l'intérieur de la banque.
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

    ''' <summary>
    ''' Période d'une pièce qui ne couvre qu'une journée : « DU 09 09 2026 ».
    '''
    ''' Écrite sans « AU » : « DU 09 AU 09 09 2026 » serait exact mais se lit comme une
    ''' faute de frappe, et un narratif dont le comptable doute est un narratif qu'il
    ''' vient faire vérifier.
    ''' </summary>
    Public Const PIECE_JOURNEE_FORMAT As String = "DU {0:00} {1:00} {2}"

    ''' <summary>
    ''' Période tenant dans un seul mois : « DU 08 AU 14 09 2026 ». Le mois et l'année
    ''' ne sont écrits qu'une fois, puisqu'ils sont communs aux deux bornes.
    ''' </summary>
    Public Const PIECE_PERIODE_FORMAT As String = "DU {0:00} AU {1:00} {2:00} {3}"

    ''' <summary>
    ''' Période franchissant un mois ou une année : « DU 28/09/2026 AU 04/10/2026 ».
    ''' Les deux dates sont écrites en entier — c'est plus long, mais c'est le seul cas
    ''' où l'abrégé serait ambigu, et les compensations de fin de mois y tombent.
    ''' </summary>
    Public Const PIECE_PERIODE_LONGUE_FORMAT As String = "DU {0:dd/MM/yyyy} AU {1:dd/MM/yyyy}"
    Public Const LIB_TVA As String = "TVA COLLECTEES WESTERN UNION"
    Public Const LIB_TTA_ENVOI As String = "TTA (TAXE SUR TRANSFER DE FONDS WU)"
    ' Double espace avant "DE" : reproduit fidèlement le libellé du classeur de référence
    ' PieceComptabilsationTchad.xlsx (colonne LIBELLES, ligne TTA Réception).
    Public Const LIB_TTA_RECEPTION As String = "TTA (TAXE SUR RECEPTION  DE FONDS WU)"
    Public Const LIB_ECART_ATTENTE As String = "ECART D'ARRONDI - COMPTE INTER BANCAIRE"

    ''' <summary>
    ''' PRÉFIXE PORTÉ PAR TOUS LES LIBELLÉS DE LA PIÈCE, sur demande écrite de la banque.
    '''
    ''' Il part dans la colonne ADDLTEXT du fichier core banking, où leur système l'attend en
    ''' tête de chaque narratif. Posé une fois ici, il suit les douze libellés de la pièce sans
    ''' qu'aucun d'eux ait à le connaître : voir PieceComptableService.Narratif, qui l'ajoute à
    ''' celui qui ne le porte pas déjà.
    ''' </summary>
    Public Const LIB_PREFIXE As String = "LD"

    ''' <summary>
    ''' Gabarit du libellé de la ligne de mouvement (activité) du point de vente.
    ''' {0} est remplacé par la Designation de l'Account.
    '''
    ''' IL A CHANGÉ SUR DEMANDE ÉCRITE DE LA BANQUE. Il valait « CCS_{0} ACTIVITE WU », et
    ''' c'est ce gabarit — notre propre préfixe, et non une donnée du référentiel — qui faisait
    ''' apparaître CCS « partout dans la narrative » du core banking, comme la banque l'a
    ''' relevé : chaque ligne de mouvement de chaque point de vente le portait.
    '''
    ''' La nouvelle forme est celle qu'elle a dictée : LD WU ACTIVITE, puis le point de vente,
    ''' puis la période — cette dernière ajoutée par Narratif, comme pour tous les autres
    ''' libellés.
    '''
    ''' « AGENCE OU SOUS-AGENT » EST LA DÉSIGNATION, et non le code Account. C'est elle que la
    ''' pièce portait déjà, c'est elle qu'un comptable lit, et la demande ne nommait pas de
    ''' code. Si la banque voulait l'Account, c'est l'appel de LibelleDuMouvement qu'il faut
    ''' changer, et lui seul.
    ''' </summary>
    Public Const LIB_MOUVEMENT_ACTIVITE_FORMAT As String = "LD WU ACTIVITE {0}"

    ''' <summary>
    ''' Groupe d'un point de vente qui n'en porte aucun — une agence propre, ou un
    ''' sous-agent dont le groupe reste à renseigner.
    '''
    ''' Nommé plutôt que laissé vide : une pièce intitulée « Groupe  » se lit comme un
    ''' défaut d'affichage, « Groupe (sans groupe) » se lit comme un état des lieux.
    ''' </summary>
    Public Const PIECE_GROUPE_SANS As String = "(sans groupe)"

    ''' <summary>
    ''' Gabarit de la ligne RAISON, en bas de chaque pièce.
    '''
    ''' LA BANQUE L'A DEMANDÉE « POUR CHAQUE SOUS-AGENT », et c'est ce que le classeur produit :
    ''' une pièce par point de vente, chacune portant en bas la raison de ses écritures. La
    ''' pièce GLOBALE, qui les rassemble toutes, ne peut en nommer aucun : elle emploie le
    ''' gabarit ci-dessous.
    '''
    ''' {0} est le point de vente, {1} la période.
    ''' </summary>
    Public Const PIECE_RAISON_FORMAT As String = "LD WU ACTIVITE {0} {1}"

    ''' <summary>
    ''' Raison de la pièce GLOBALE, qui ne se rattache à aucun point de vente en particulier.
    ''' {0} est la période.
    ''' </summary>
    Public Const PIECE_RAISON_GLOBALE_FORMAT As String = "LD WU ACTIVITE {0}"

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

#Region "Écarts de change : colonnes et valeurs du rapport de règlement"

    ' CE QUE CETTE SECTION AJOUTE, ET POURQUOI ELLE EST À PART. Le calcul des écarts de change
    ' lit le rapport de règlement COLONNE PAR COLONNE, là où l'agrégation de la compensation
    ' n'en lit que six. Les noms sont donc nommés ici une fois, plutôt que recopiés dans le
    ' service : une faute de frappe sur « ClearPrincipalPAY » ne lèverait aucune erreur — la
    ' colonne serait simplement introuvable, la valeur lue 0, et l'écart égal au montant local
    ' tout entier. C'est la faute la plus coûteuse que ce calcul puisse commettre, et la seule
    ' que le compilateur ne voit pas.

    ''' <summary>
    ''' Identifiant du point de vente dans les rapports Western Union : « AHB020013 », « ADJ226820 ».
    ''' C'est la clé par laquelle le référentiel reconnaît un sous-agent ou une agence propre.
    ''' </summary>
    Public Const COLONNE_ACCOUNT As String = "Account"

    ''' <summary>Numéro de contrôle du transfert : la référence citée par la banque et par le client.</summary>
    Public Const COLONNE_MTCN As String = "MTCN"

    ''' <summary>Nature de la ligne : "T" transaction, "A" ajustement.</summary>
    Public Const COLONNE_TYPE_TRANSACTION As String = "TransactionType"

    ''' <summary>Libellé de l'ajustement porté par une ligne "A" (ex. "FOR FULL REFUND").</summary>
    Public Const COLONNE_TYPE_AJUSTEMENT As String = "AdjustmentType"

    ''' <summary>MTCN de la transaction d'origine que l'ajustement corrige, quand il le désigne.</summary>
    Public Const COLONNE_MTCN_AJUSTE As String = "AdjustmentMTCN"

    ''' <summary>Sens de l'opération : "S" envoi, "P" paiement.</summary>
    Public Const COLONNE_SENS As String = "SendPayIndicator"

    ''' <summary>Statut de règlement de la ligne : "S" réglée, "W" en attente.</summary>
    Public Const COLONNE_STATUT_REGLEMENT As String = "TxnStatus"

    ''' <summary>Code produit Western Union de la ligne (IMTR, FTSS, AVSS…).</summary>
    Public Const COLONNE_CODE_PRODUIT As String = "ProductCode"

    ''' <summary>Principal d'un ENVOI en monnaie locale : le montant encaissé au guichet.</summary>
    Public Const COLONNE_PRINCIPAL_ENVOI_LOCAL As String = "RecPrincipalREC"

    ''' <summary>Principal d'un PAIEMENT en monnaie locale : le montant décaissé au guichet.</summary>
    Public Const COLONNE_PRINCIPAL_PAYE_LOCAL As String = "ClearPrincipalPAY"

    ''' <summary>Principal de la transaction dans la devise de règlement (colonnes LOC).</summary>
    Public Const COLONNE_PRINCIPAL_DEVISE As String = "ClearPrincipalLOC"

    ''' <summary>
    ''' Part de change de la transaction dans la devise de règlement.
    '''
    ''' ELLE N'ENTRE DANS LE CALCUL QUE POUR LES ENVOIS. Pour un paiement, sa valeur absolue est
    ''' DÉJÀ comptabilisée en commission de paiement par WUReportService.CalculerReglement :
    ''' l'ajouter au montant en devise la compterait deux fois, une fois en commission et une
    ''' fois en écart de change. Voir ChangeService.
    ''' </summary>
    Public Const COLONNE_CHANGE_DEVISE As String = "ClearFXLOC"

    ''' <summary>Les colonnes sans lesquelles aucun écart de change ne peut être calculé.</summary>
    Public Shared ReadOnly ColonnesEcartsDeChange As String() = {
        COLONNE_MTCN, COLONNE_ACCOUNT, COLONNE_TYPE_TRANSACTION, COLONNE_SENS, COLONNE_DEVISE_LOC,
        COLONNE_PRINCIPAL_ENVOI_LOCAL, COLONNE_PRINCIPAL_PAYE_LOCAL,
        COLONNE_PRINCIPAL_DEVISE, COLONNE_CHANGE_DEVISE
    }

    ''' <summary>Valeur de COLONNE_TYPE_TRANSACTION désignant une transaction réelle.</summary>
    Public Const TYPE_TRANSACTION_NORMALE As String = "T"

    ''' <summary>
    ''' Valeur de COLONNE_TYPE_TRANSACTION désignant un AJUSTEMENT : remboursement, reprise,
    ''' retraitement. Ces lignes ne portent aucun change propre et sont toutes écartées.
    ''' </summary>
    Public Const TYPE_TRANSACTION_AJUSTEMENT As String = "A"

    ''' <summary>
    ''' Libellé d'ajustement signalant un remboursement TOTAL. La transaction d'origine n'a
    ''' alors produit aucun change définitif : elle est écartée avec son remboursement.
    ''' Comparaison insensible à la casse et aux espaces de bordure.
    ''' </summary>
    Public Const AJUSTEMENT_REMBOURSEMENT_TOTAL As String = "FOR FULL REFUND"

    ''' <summary>Valeur de COLONNE_SENS désignant un envoi.</summary>
    Public Const SENS_ENVOI As String = "S"

    ''' <summary>Valeur de COLONNE_SENS désignant un paiement.</summary>
    Public Const SENS_PAIEMENT As String = "P"

    ''' <summary>
    ''' Valeur de COLONNE_STATUT_REGLEMENT désignant une transaction NON ENCORE RÉGLÉE.
    ''' Leur sort est décidé par une option : voir OptionsChangeWU.InclureEnvoisEnAttente.
    ''' </summary>
    Public Const STATUT_REGLEMENT_EN_ATTENTE As String = "W"

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

    ' Le fichier chargé dans le core banking compte DOUZE colonnes, dont huit sont constantes
    ' ou déduites. Les valeurs ci-dessous sont celles de la banque : elles ne se devinent pas et
    ' ne doivent pas être modifiées sans son accord.
    '
    ' CE BLOC A ÉTÉ CORRIGÉ SUR DEMANDE ÉCRITE DE LA BANQUE, pendant sa production, et les
    ' quatre changements sont confirmés par elle :
    '
    '     BRN ............ N01  ->  S51
    '     TXNCD débit .... U24  ->  F03
    '     TXNCD crédit ... F15  ->  F57
    '     COST_CENTER .... supprimée du fichier
    '     BATCHNO ........ désormais en majuscules (voir CoreBankingService)
    '
    ' Les fichiers DÉJÀ chargés portent les anciennes valeurs. Ce n'est pas rattrapable et n'a
    ' pas à l'être : un lot chargé l'a été sous les codes en vigueur ce jour-là.

    Public Const CB_DETBSJRNL As String = "BBR"

    ''' <summary>
    ''' Agence du LOT (colonne BRN). Corrigée de N01 en S51 sur demande de la banque.
    '''
    ''' ATTENTION, ELLE N'EST PAS LA MÊME CHOSE QUE CB_AGENCE_SIEGE. BRN désigne l'agence du
    ''' LOT, portée à l'identique par toutes ses lignes ; ACBRN désigne l'agence de CHAQUE
    ''' ÉCRITURE, et vaut CB_AGENCE_SIEGE pour les comptes internes. La demande de la banque
    ''' ne nomme que BRN : CB_AGENCE_SIEGE reste donc à N01, et un fichier porte aujourd'hui
    ''' S51 en tête de chaque ligne et N01 dans sa colonne ACBRN pour les comptes du siège.
    ''' Si la banque voulait aussi changer celui-là, c'est CB_AGENCE_SIEGE qu'il faut
    ''' modifier, et lui seul.
    ''' </summary>
    Public Const CB_BRN As String = "S51"

    Public Const CB_SRCCODE As String = "ECOSOURCE"

    ''' <summary>Code transaction d'un débit. U24 jusqu'au correctif demandé par la banque.</summary>
    Public Const CB_TXNCD_DEBIT As String = "F03"

    ''' <summary>Code transaction d'un crédit. F15 jusqu'au correctif demandé par la banque.</summary>
    Public Const CB_TXNCD_CREDIT As String = "F57"

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
    ''' Longueur maximale du narratif (colonne ADDLTEXT) du fichier chargé au core banking.
    ''' CONFIRMÉE PAR LA BANQUE : 150 caractères.
    '''
    ''' Elle n'est pas là pour tronquer — on ne coupe pas un narratif comptable, un libellé
    ''' amputé à mi-mot ne dit plus de quel sous-agent il s'agit — mais pour AVERTIR avant
    ''' l'envoi. Le plus long libellé produit aujourd'hui atteint 78 caractères, période
    ''' comprise ; la marge est confortable, et ce contrôle ne devrait jamais se déclencher.
    ''' C'est précisément pour cela qu'il doit exister : le jour où un point de vente portera
    ''' une désignation à rallonge, personne ne pensera à recompter.
    ''' </summary>
    Public Const CB_NARRATIF_LONGUEUR_MAX As Integer = 150

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

    ''' <summary>
    ''' Première lettre du numéro de lot de la pièce des ÉCARTS DE CHANGE.
    '''
    ''' Elle met ce lot dans un espace distinct de celui de la compensation, qui porte sur la
    ''' MÊME journée et produirait sans cela le même numéro. Voir
    ''' CoreBankingService.NumeroDeLotDeChange, qui démontre qu'aucune collision n'est possible.
    ''' </summary>
    Public Const CB_LOT_PREFIXE_CHANGE As String = "C"

#End Region

End Class
