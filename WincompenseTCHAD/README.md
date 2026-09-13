# WincompenseTCHAD

Automatisation de la comptabilisation Western Union J+1 (Tchad) — application Windows Forms
en VB.NET (.NET Framework 4.8 / Visual Studio 2019 / SQL Server Express).

## Ouverture du projet

1. Ouvrir `WincompenseTCHAD.sln` dans Visual Studio 2019.
2. Vérifier que le Framework cible du projet est bien **.NET Framework 4.8**.
3. Adapter la chaîne de connexion SQL Server dans `WincompenseTCHAD\App.config` si nécessaire
   (par défaut : `Server=.\SQLEXPRESS;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;`).
4. Exécuter les scripts SQL du dossier `Scripts\` (dans l'ordre numéroté) sur l'instance
   `.\SQLEXPRESS` pour créer les tables `T_Pdv_SA` / `T_Pdv_EC` et les données de test.
5. Compiler et lancer (F5).

## Architecture

```
WincompenseTCHAD/
├── WincompenseTCHAD.sln
└── WincompenseTCHAD/
    ├── WincompenseTCHAD.vbproj
    ├── App.config                          ' Chaîne de connexion SQL Server
    ├── Program.vb                          ' Point d'entrée (Sub Main)
    ├── My Project/AssemblyInfo.vb
    ├── Constants/ConstantesWU.vb           ' Taux, comptes comptables, libellés, colonnes attendues
    ├── Models/
    │   ├── CalculWU.vb                     ' Classe métier par Account
    │   ├── ComptesSystemeWU.vb             ' Comptes comptables paramétrés (table SystemeWU)
    │   └── PointDeVente.vb                 ' Sous-agent (T_Pdv_SA) et agence propre (T_Pdv_EC)
    ├── Services/
    │   ├── WUFichierService.vb             ' Contrôles de sécurité : type de rapport, concordance des périodes
│   ├── WUReportService.vb              ' Lecture fichiers (ZIP ou texte), parsing, agrégation, dates
    │   ├── WURepository.vb                 ' Lecture SQL Server pour la compensation (T_Pdv_SA / T_Pdv_EC / SystemeWU)
│   ├── PdvRepository.vb                ' CRUD points de vente et groupes (écriture isolée de la lecture)
    │   ├── WUCalculationService.vb         ' Formules, répartition, arrondi
    │   └── PieceComptableService.vb        ' Grille de contrôle, pièce comptable, équilibrage, export Excel
    └── Forms/
        ├── FrmCompensationWU.vb            ' Orchestration des événements uniquement
        ├── FrmCompensationWU.Designer.vb
        ├── FrmCompensationWU.resx
        ├── FrmPieceComptable.vb            ' Affichage d'une pièce (globale ou d'un seul PDV)
        ├── FrmComptesSysteme.vb            ' Paramétrage des comptes comptables
        ├── FrmSousAgents.vb                ' Gestion des sous-agents (CRUD)
        ├── FrmGroupesStatistiques.vb       ' Gestion des groupes statistiques (CRUD)
        └── FrmAgences.vb                   ' Gestion des agences propres (CRUD)

Scripts/
├── 01_CreateTables_GWC_WINCOMPENSE_ETD.sql
├── 02_DonneesExemple.sql
└── 03_SystemeWU.sql                       ' Comptes comptables paramétrés
```

## Hypothèses métier retenues (à valider)

1. **InclureLigneReglement** : toutes les lignes du rapport de règlement sont incluses par défaut
   (y compris `TransactionType = "A"`), dès lors que l'Account est renseigné. Fonction isolée,
   volontairement simple, à affiner selon consigne métier ultérieure.
2. **Comptes INCONNU** : toujours affichés dans la grille de contrôle (jamais ignorés) ; pour la
   pièce comptable, traités par défaut comme une agence propre (100 % banque, compte courant WU),
   avec surlignage d'anomalie. *À confirmer.*
3. **Structure des écritures de la pièce comptable** : validée par rapprochement algébrique avec
   un exemple réel du classeur `PieceComptabilsationTchad.xlsx` (agence sous-agent "BOLOLO").
   Une seule ligne de mouvement net (`NetMouvement = PrincipalEnvoi+ChargeEnvoi+Taxes−PrincipalPaye`,
   Débit si positif) sur le `CompteCompense` du PDV, en contrepartie du compte courant WU pour la
   part nette bancaire ; commissions et taxes sont des lignes de crédit uniquement, sans ligne de
   débit miroir individuelle (voir commentaires détaillés dans `GenererPieceComptable`). Vérifié à
   l'unité près sur l'exemple disponible ; *le cas d'une agence propre "EC" reste à valider faute
   d'exemple de référence pour ce type de PDV.*
4. **Solde par Account** (grille de contrôle) = `PrincipalPaye − (PrincipalEnvoi + ChargeEnvoi + Taxes)`.
5. **Cohérence des dates** : la date du rapport d'activité (`txnDateLOC`) est comparée à celle du
   rapport de règlement, reconstituée depuis `SetDateLOCYear/Month/Day` (à défaut `RepDate`).
   Une divergence bloque le traitement. Si aucune date n'est exploitable d'un côté, la
   vérification est ignorée sans bloquer (avertissement affiché dans le StatusStrip).
6. **Compte absorbant l'écart d'arrondi global** : le **compte inter bancaire 381000101**,
   d'après la ligne de paramétrage de `SystemeWU` (`Cpte_attenteDEBIT` / `Cpte_attenteCREDIT`)
   et le formulaire « Comptes Systèmes WU » de la Direction Comptable, qui le désignent tous
   deux ainsi. Il remplace le compte d'attente fictif `XXXXXXXXXX` utilisé jusqu'ici faute de
   numéro connu. *Usage à confirmer : la table nomme ces colonnes « attente », le formulaire
   « compte inter bancaire » — s'il s'agit de deux comptes distincts, seul le paramétrage est
   à corriger, le code n'est pas concerné.*
7. **Export Excel** : réalisé en liaison tardive (late binding, `Option Strict Off` isolé dans
   `PieceComptableService.vb`) afin de ne pas imposer de référence COM Excel obligatoire au
   projet. Toute la logique métier fonctionne sans Excel installé.

## Chargement des rapports : archives ZIP

Western Union livre désormais ses rapports **sous forme d'archives ZIP** portant exactement le
nom du fichier texte qu'elles contiennent (`..._ACTIVITY_REPORT_....txt`, etc.). Les deux boutons
de chargement (**1. Rapport d'activité** et **2. Rapport de règlement**) acceptent donc
indifféremment :

- l'**archive ZIP** telle qu'elle est reçue — elle est décompressée **en mémoire** au moment de la
  lecture, sans création de fichier temporaire (donc rien à nettoyer ensuite) ;
- le **fichier texte** déjà décompressé, pour les rapports plus anciens ou décompressés à la main.

Détails d'implémentation (`WUReportService`) :

| Point | Traitement |
|---|---|
| Détection | Signature binaire `PK` (0x50 0x4B) en tête de fichier, **jamais** l'extension : les archives portant le nom du rapport `.txt`, l'extension n'est pas fiable |
| Entrée retenue dans l'archive | Le fichier `.txt` le plus volumineux ; à défaut, l'entrée la plus volumineuse (`ChoisirEntreeRapport`) |
| Encodage | `Encoding.Default` (Windows/ANSI), identique à la lecture d'un fichier texte |
| Archive vide | `RapportInvalideException` : « ne contient aucun fichier exploitable » |
| Archive endommagée | `InvalidDataException` convertie en `RapportInvalideException` : « illisible ou endommagée » |

Aucune étape de traitement en aval n'est modifiée : une fois les lignes obtenues, le parsing,
l'agrégation, les calculs et la génération de la pièce comptable sont strictement identiques.

## Paramétrage des comptes comptables (table SystemeWU)

Les comptes utilisés par la pièce comptable ne sont plus figés dans le code : ils sont lus au
démarrage dans la table SQL Server **`SystemeWU`** et modifiables depuis le formulaire
**« Comptes Systèmes WU »** (bouton *Paramètres des comptes…* de l'écran principal).

### Correspondance champ du formulaire → colonne de la table

| Champ du formulaire | Colonne `SystemeWU` | Valeur en service |
|---|---|---|
| Compte courant Western Union ETD | `Cpte_PositionNette` | 32100003292 |
| Compte inter bancaire | `Cpte_attenteDEBIT` / `Cpte_attenteCREDIT` | 381000101 |
| Commission sur Transfert_Ecobank | `Cpte_Produit` | 728300148 |
| Commission sur Envoi_Ecobank | `Cpte_Produit_Envoi` | 728300148 |
| Commission sur Paiement_Ecobank | `Cpte_Produit_Paiement` | 728300149 |
| Impôts et taxe sur envoi | `Tthu` | 434000147 |
| TVA | `Tob` | 434000104 |
| TTA sur envoi WU | `Cpte_Envoi` | 434000145 |
| TTA sur paiement WU | `Cpte_Paiement` | 434000159 |

Les autres colonnes de la table (`Passif`, `Actif`, `Cpte_Charge_Publicitaire`,
`Cpte_Gainde_Change`, `Cpte_Envoi_agence`, `Cpte_Paiement_agence`) ne sont **ni lues ni
écrites** : elles ne concernent pas la pièce comptable Western Union et un enregistrement ne
les altère jamais. La table étant partagée, l'`UPDATE` ne porte que sur les dix colonnes
ci-dessus.

### Règles de fonctionnement

- **Repli sans base.** Si la base est inaccessible, la table absente ou vide, les comptes
  **par défaut du code** (`ConstantesWU`) prennent le relais : l'application reste utilisable
  et se comporte exactement comme avant le paramétrage. La barre d'état l'indique.
- **Colonne vide = valeur par défaut conservée.** Une colonne NULL ou vide ne remplace pas la
  valeur par défaut : mieux vaut un compte connu qu'une écriture sans numéro de compte.
- **Ligne visée à l'enregistrement.** Celle dont la colonne `code` correspond à la ligne lue.
  Si aucun code n'est lisible, la mise à jour n'est acceptée que lorsque la table ne contient
  qu'une seule ligne : jamais question de modifier une ligne au hasard dans une table partagée.
- **Édition sur copie.** Les comptes en service ne sont remplacés qu'une fois l'écriture en
  base réussie : une modification abandonnée ne peut pas fausser une pièce comptable.
- **Contrôles de saisie.** Un compte vide bloque l'enregistrement (il produirait une écriture
  sans numéro de compte) ; un compte non exclusivement numérique déclenche un avertissement
  — typiquement une faute de frappe — mais reste enregistrable après confirmation.
- **Prise en compte immédiate.** Les nouveaux comptes s'appliquent dès l'enregistrement. Si un
  calcul est déjà affiché, l'application invite à le relancer avant de générer la pièce.

Le script `Scripts\03_SystemeWU.sql` crée la table et sa ligne de paramétrage **si elles
n'existent pas** ; il ne modifie jamais un paramétrage en place.

### Commission Transfert et Commission Envoi désormais dissociables

Le code n'utilisait qu'un seul compte pour ces deux commissions (`728300148`). La table les
porte dans deux colonnes distinctes (`Cpte_Produit` et `Cpte_Produit_Envoi`), que le formulaire
présente comme deux champs : **elles peuvent maintenant recevoir des comptes différents sans
toucher au code**. Tant que les deux colonnes portent la même valeur, la pièce comptable est
strictement identique à celle produite auparavant.

## Gestion des points de vente (sous-agents et agences)

Deux écrans, ouverts depuis les boutons **« Sous-agents… »** et **« Agences propres… »**,
permettent de créer, consulter, modifier et supprimer les points de vente sans passer par SQL
Server Management Studio.

| Écran | Table | Champs |
|---|---|---|
| Sous-agents Western Union | `T_Pdv_SA` | Account (`Code_Pdv`), Désignation, Groupe statistique, puis — hérités du groupe — Taux, Compte d'activité (`CompteCompense`), Compte de commission ; et Code agence |
| Agences propres Ecobank | `T_Pdv_EC` | Account (`Codesite`), Désignation, Code agence Voyager |

Une agence propre ne rétrocède aucune commission — la banque en conserve 100 % — d'où l'absence
de taux et de comptes de compensation/commission, contrairement aux sous-agents.

### Règles de fonctionnement

- **L'Account n'est modifiable qu'à la création.** C'est la clé sous laquelle les rapports
  Western Union désignent le point de vente : la renommer romprait le lien avec l'historique.
  Pour la changer, supprimer puis recréer la fiche.
- **Doublon entre les deux tables signalé.** Un même Account présent dans `T_Pdv_SA` et
  `T_Pdv_EC` est une incohérence : la recherche interroge `T_Pdv_SA` en premier, la fiche
  agence serait donc silencieusement ignorée. La base n'interdit pas ce doublon — l'application
  avertit avant de créer la fiche, et laisse décider.
- **Account déjà pris.** La violation de clé primaire est traduite en message clair
  (« l'Account est déjà enregistré ») au lieu d'un code d'erreur SQL Server.
- **Taux : une fraction, deux décimales.** 0,70 signifie 70 %. Toute valeur hors de
  l'intervalle [0 ; 1] est refusée : la confusion « 70 » pour « 0,70 » multiplierait par cent
  toutes les commissions rétrocédées. La colonne étant de type `DECIMAL(4,2)`, une saisie à
  plus de deux décimales est également refusée — plutôt qu'arrondie en silence par la base.
  La virgule et le point sont acceptés indifféremment comme séparateur décimal.
- **Aucun NULL écrit.** Plusieurs colonnes facultatives du point de vue de la saisie
  (`GroupeStatistique`, `codeagence`, `[CodeAgenc-Voyager]`) sont déclarées `NOT NULL` en base :
  un champ laissé vide est donc enregistré comme chaîne vide. La lecture traite indifféremment
  `NULL` et chaîne vide.
- **Suppression confirmée, et expliquée.** La confirmation rappelle la conséquence réelle : les
  rapports portant cet Account apparaîtront en `INCONNU` et leur pièce comptable utilisera le
  compte courant WU au lieu du compte de compensation.
- **Groupe statistique : unité de paramétrage.** Le compte d'activité, le compte de commission
  et le taux appartiennent au GROUPE ; le sous-agent en hérite et ne peut pas les redéfinir
  pour lui seul (voir ci-dessous).
- **Recherche.** Le champ de recherche filtre sur l'Account ou la désignation ; il s'applique à
  la touche Entrée ou au bouton *Actualiser*, pas à chaque caractère frappé. Les caractères
  génériques de SQL (`%`, `_`, `[`) y sont neutralisés : chercher « % » cherche bien un
  pourcentage.
- **Calcul invalidé.** Après un passage dans l'un de ces écrans, si un calcul est déjà affiché,
  la barre d'état invite à le relancer : le paramétrage a pu changer.

### Groupe statistique

Le compte d'activité, le compte de commission et le taux appartiennent au GROUPE, pas au
sous-agent : voir la section « Groupes statistiques » ci-dessous.

Le chemin de comptabilisation quotidienne ne peut jamais écrire dans ces tables : la lecture
reste dans `WURepository`, l'écriture est isolée dans `PdvRepository`.

## Groupes statistiques (table T_GroupeStatistique)

**Règle métier.** Un groupe statistique porte **un** compte d'activité, **un** compte de
commission et **un** taux. Tout Account appartient à un seul groupe et en hérite ces trois
valeurs. Un compte d'activité ou de commission n'appartient qu'à un seul groupe.

Le compte d'activité est la colonne `CompteCompense` de `T_Pdv_SA` : c'est lui qui porte la
ligne de mouvement « CCS_… **ACTIVITE** WU » de la pièce comptable.

Ces trois règles ne sont plus seulement vérifiées par l'application : la table les rend
**impossibles à violer**, y compris par une écriture SQL directe.

| Contrainte | Ce qu'elle garantit |
|---|---|
| `Groupe` clé primaire | Un libellé de groupe unique |
| Index unique sur `CompteActivite` | Un compte d'activité n'appartient qu'à un groupe |
| Index unique sur `CompteCommission` | Un compte de commission n'appartient qu'à un groupe |

### Migration

Le script `Scripts\04_GroupeStatistique.sql` crée la table et y reprend les groupes déjà
présents dans `T_Pdv_SA`. **Il ne modifie rien tant que les données ne sont pas saines** : il
commence par un diagnostic en trois volets, et la migration ne s'exécute que si aucune anomalie
bloquante n'est trouvée.

| Diagnostic | Bloquant | Pourquoi |
|---|---|---|
| Groupes dont les sous-agents portent des valeurs divergentes | oui | Le groupe ne peut avoir qu'un seul jeu de valeurs : laquelle retenir ? |
| Un même compte utilisé par plusieurs groupes | oui | Violerait l'index unique |
| Sous-agents sans groupe | non | Listés pour information : ils n'hériteront d'aucune valeur |

En cas de blocage, les lignes fautives sont affichées, aucune donnée n'est touchée, et le script
se relance à volonté une fois les corrections faites. La pose d'une clé étrangère
`T_Pdv_SA.GroupeStatistique → T_GroupeStatistique.Groupe` est proposée en fin de script, **non
exécutée** : elle échouerait tant qu'il reste des sous-agents sans groupe.

### Les colonnes de T_Pdv_SA restent le miroir du groupe

`CompteCompense`, `CompteCommission` et `Taux` sont **conservées dans `T_Pdv_SA` et tenues
synchronisées** avec leur groupe. Deux raisons :

- la **comptabilisation quotidienne continue de les lire** — elle ne dépend donc pas de la
  nouvelle table, et une migration non jouée ne peut pas interrompre la production ;
- d'autres applications de la banque peuvent lire `T_Pdv_SA` : leur comportement est inchangé.

Toute modification d'un groupe est aussitôt reportée sur ses sous-agents. Une dérive éventuelle
(écriture directe en base, synchronisation interrompue) est comptée par la colonne
**Désynchronisés** de la grille, la ligne apparaît en rose, et le bouton **« Synchroniser les
sous-agents »** la corrige.

### Écran « Groupes... »

Liste, création, modification, suppression, plus la synchronisation ci-dessus.

- **Le libellé n'est modifiable qu'à la création** : c'est la clé sous laquelle les sous-agents
  se rattachent au groupe.
- **Modifier un groupe est annoncé avant, pas découvert après** : la confirmation chiffre les
  sous-agents concernés et affiche chaque valeur sous la forme `ancienne -> nouvelle`.
- **Suppression refusée** tant que des sous-agents y sont rattachés : ils perdraient leurs
  comptes et leur taux sans que rien ne le signale.
- **Compte déjà pris.** La violation d'index unique est traduite en message nommant *lequel* des
  deux comptes appartient déjà à un autre groupe.

### Écran « Sous-agents... »

Le groupe se choisit dans la liste déroulante ; les trois champs hérités se remplissent
aussitôt et restent **toujours en lecture seule** (fond grisé) — ils appartiennent au groupe.
Le bouton **« … »** à côté de la liste ouvre l'écran des groupes, en création et pré-rempli si
le libellé saisi est inconnu.

- **Le groupe est obligatoire** : sans lui, le sous-agent n'a ni compte ni taux.
- **Un groupe ne se crée pas d'un simple nom.** Un libellé inconnu n'est plus accepté à la
  volée : il faut ses trois valeurs, donc passer par l'écran des groupes, qui s'ouvre pré-rempli.
- **Fiche ayant dérivé.** Une fiche antérieure ne portant pas les valeurs de son groupe est
  affichée telle quelle, avec un avertissement — jamais réalignée en silence. L'enregistrement
  adopte les valeurs du groupe, après une confirmation qui montre chaque changement.

## Contrôles de sécurité sur les fichiers chargés

Deux erreurs de manipulation fausseraient silencieusement toute la comptabilisation : croiser
deux rapports de **périodes différentes**, ou charger un rapport **d'activité à la place d'un
rapport de règlement** (ou l'inverse). Chacune est bloquée par **deux barrières indépendantes**.

Principe retenu : un contrôle fondé sur le **nom** du fichier ne bloque que s'il constate une
contradiction *certaine*. Un nom non standard, dont on ne peut rien déduire, n'interrompt jamais
le traitement — il serait inacceptable qu'un simple renommage empêche de comptabiliser la journée.
Le contrôle fondé sur le **contenu**, lui, est systématique et fait foi.

### 1. Les deux rapports doivent couvrir la même période

| Barrière | Moment | Source | Effet |
|---|---|---|---|
| Nom du fichier | dès le clic sur le bouton | période lue dans le nom de l'archive | choix explicite : abandonner le fichier, **ou** changer de journée (l'autre rapport est alors retiré) |
| Contenu | au clic sur « Afficher » | dates lues **dans** les rapports (`ValiderCoherenceDates`) | traitement **bloqué** |

Deux nomenclatures de période sont reconnues (`WUFichierService.ExtrairePeriode`) :

| Nom du fichier | Période retenue |
|---|---|
| `RSP_TD383_ACTIVITY_REPORT_BY_ACCOUNT_20260530_20260530_202606031151.zip` | 30/05/2026 |
| `RSP_TD383_ACTIVITY_REPORT_BY_ACCOUNT_20260501_20260531_...zip` | du 01/05/2026 au 31/05/2026 |
| `Rapport d'activité par Site (N° d'opérateur) du 02 Jan 2021.zip` | 02/01/2021 |
| `rapports_du_jour.zip` | non lisible → contrôle reporté sur le contenu |

Le contrôle sur la période ne refuse pas sèchement le fichier : cela enfermerait l'utilisateur,
qui ne pourrait plus jamais passer d'une journée à une autre une fois les deux rapports chargés.
Une boîte de dialogue lui laisse donc le choix entre abandonner le fichier et changer de journée
de traitement — auquel cas l'autre rapport, devenu hors période, est retiré et doit être rechargé.
**Dans les deux cas, deux rapports de périodes différentes ne peuvent jamais être chargés ensemble.**

Seuls les groupes de **8 chiffres exactement** sont retenus comme dates : l'horodatage d'édition
du rapport (`202606031151`, 12 chiffres) est ainsi écarté et ne peut pas être pris pour une date
de période. Les mois nommés sont acceptés en français comme en anglais, abrégés ou complets ; un
libellé ambigu comme `jui` (juin ? juillet ?) est volontairement **non** reconnu — sur une donnée
comptable, ne rien conclure vaut mieux que deviner.

### 2. Un rapport ne peut pas être chargé à la place de l'autre

| Barrière | Moment | Source | Effet |
|---|---|---|---|
| Nom du fichier | dès le clic sur le bouton | mots-clés `ACTIVITY`/`activité` et `SETTLEMENT`/`règlement` | fichier **refusé** |
| Contenu | au clic sur « Afficher » | colonnes réellement présentes (`VerifierTypeRapport`) | traitement **bloqué** |

La seconde barrière est celle qui compte : **un fichier renommé franchit la première, jamais la
seconde**. Elle repose sur des colonnes discriminantes, présentes dans tout rapport d'un type et
dans aucun rapport de l'autre — vérifié sur les rapports des 02/01/2021 et 30/05/2026, soit les
deux formats Western Union :

| Signature | Colonnes |
|---|---|
| Rapport d'activité | `TaxesREC`, `TaxesPAY`, `PayPrincipalPAY`, `txnDateLOC` |
| Rapport de règlement | `TransactionType`, `PayCountry`, `ClearChargesLOC`, `ClearFXLOC`, `SetDateLOCYear` |

Deux colonnes suffisent pour conclure (`MIN_COLONNES_SIGNATURE`) : la détection résiste ainsi à la
disparition d'une colonne lors d'une future évolution du format, sans risque de confusion puisque
aucune de ces colonnes n'existe dans le rapport de l'autre type. Sur les quatre rapports réels
disponibles, la séparation est nette : score 4/0 pour les rapports d'activité, 0/5 pour les
rapports de règlement, dans les deux formats.

Si l'archive a été renommée et que son nom ne dit plus rien, le nom du **rapport contenu dans
l'archive** est analysé à son tour : il porte, lui, la nomenclature Western Union d'origine.

## Compatibilité avec les formats de rapport Western Union

Western Union a refondu le format de ses rapports (comparaison faite entre un rapport du
02/01/2021 et un rapport du 30/05/2026). **L'application traite indifféremment les deux
formats** : aucune colonne utilisée par les calculs n'a disparu, et les différences de contenu
sont absorbées automatiquement.

| Élément | Ancien format | Nouveau format | Traitement |
|---|---|---|---|
| `PayCountry` (pays local) | `TCHAD` | `CHAD` | Les deux graphies sont acceptées (`ConstantesWU.PaysPaiementLocal`) |
| Devise des montants LOC | `EUR` | `XAF` | Conversion pilotée par `LOCCurrencyCode` : ×655,957 si EUR, aucune conversion si XAF (`WUReportService.ObtenirFacteurConversion`) |
| Séparateur décimal | virgule (`1493,91`) | point (`-250000.0000`) | Détecté par position du symbole, jamais par une culture supposée (`ToDecimalSafe`) |
| Casse des colonnes | `Status`, `txnDateLOC` | `STATUS`, `TXNDATELOC` | `DataTable.Columns` est insensible à la casse |
| Colonne `STATUS` | absente | `S` / `C` / `W` | Les lignes annulées (`C`) sont exclues de l'agrégation (`InclureLigneActivite`) ; absence de la colonne = comportement historique |
| Format de date | `20210102` | `20260530` | `yyyyMMdd` analysé explicitement (`ConstantesWU.FormatsDateRapport`) |
| Date du rapport de règlement | colonnes `SetDateLOCYear/Month/Day` | idem | Reconstituée depuis ces colonnes (`ObtenirDateReglement`), ce qui rend enfin opérationnelle la validation croisée de la section 16 |
| Colonnes ajoutées | — | `Tax1REC`, `Tax2REC`, `Tax3REC`, `TotalChargesPAY`, `Assigned_Account_Name` | Ignorées par le lecteur (aucun impact) |

**Non-régression vérifiée** : relu par le code actuel, le rapport du 02/01/2021 restitue très
exactement les valeurs de référence du cahier des charges pour `AHB020200` (PrincipalEnvoi
5 966 388 ; PrincipalPaye 7 522 117,45 ; ChargeEnvoi 340 000 ; Taxes 82 075 ; CommissionPaiement
98 104,93).

### Les taxes détaillées confirment les taux du cahier des charges

Le nouveau rapport d'activité fournit le détail des taxes (`Tax1REC`, `Tax2REC`, `Tax3REC`,
dont la somme vaut exactement `TaxesREC`), là où le code les recalcule par application des taux.
Le rapprochement, effectué ligne à ligne sur le fichier du 30/05/2026 (95 lignes d'envoi, hors
annulations), **valide les formules du cahier des charges** :

| Taxe fournie | Formule du code | Lignes conformes | Écart total |
|---|---|---|---|
| `Tax1REC` | `ChargeEnvoi × 19,25 %` (TVA) | **95 / 95** | +20,50 FCFA |
| `Tax3REC` | `PrincipalEnvoi × 0,2 %` (TTA Envoi) | **95 / 95** | −1,79 FCFA |
| `Tax2REC` | `Taxes − TVA − TTAEnvoi` (Solde de taxes) | — | −18,71 FCFA |

Les écarts résiduels (une vingtaine de francs sur ~460 000 FCFA de taxes) ne sont que les
arrondis à l'unité pratiqués par Western Union. **Aucune modification des formules n'est donc
justifiée** : les remplacer par les valeurs fournies modifierait les règles du cahier des
charges pour un gain nul.

## Points restant à confirmer

- Structure des écritures pour une **agence propre "EC"** dans la pièce comptable (aucun exemple
  de référence de ce type disponible à ce jour ; seul un exemple sous-agent "SA" a pu être validé).
- Mode d'authentification SQL Server réel en production (actuellement : Windows intégré).
- Usage exact du compte inter bancaire 381000101 pour l'écart d'arrondi global (voir hypothèse 6).
- Faut-il alimenter les listes déroulantes du formulaire de paramétrage avec le plan comptable
  complet ? Elles ne proposent aujourd'hui que le compte paramétré et le compte par défaut, la
  saisie libre restant possible.
- Sous-agents restant sans groupe statistique après la migration : quel groupe leur attribuer ?
  Tant qu'il en subsiste, la clé étrangère proposée en fin de `Scripts\04_GroupeStatistique.sql`
  ne peut pas être posée.
- Règle définitive de traitement des lignes `TransactionType = "A"` du rapport de règlement.
- Traitement définitif souhaité des Accounts `INCONNU` dans la pièce comptable.

## Test de référence (section 17)

Les scripts `Scripts\02_DonneesExemple.sql` paramètrent :

- `AHB020200` comme sous-agent (`T_Pdv_SA`, `Taux = 0.70`) ;
- `AHB020057` comme agence propre Ecobank (`T_Pdv_EC`).

Pour valider le parsing/l'agrégation, charger un rapport d'activité contenant `AHB020200` à la
date du 02/01/2021 et vérifier dans `dgvControle` que les agrégats retrouvent approximativement
les valeurs de référence indiquées dans le cahier des charges (Principal Envoi ≈ 5 966 388,
Principal Payé ≈ 7 522 117,45, Commission Envoi = 69 700, Commission Paiement ≈ 98 105, etc.).
Pour `AHB020057`, vérifier qu'aucune commission sous-agent n'est générée (100 % à la banque).
