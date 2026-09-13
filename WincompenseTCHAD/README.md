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
    │   ├── PointDeVente.vb                 ' Sous-agent (T_Pdv_SA) et agence propre (T_Pdv_EC)
    │   ├── LigneHistoriqueWU.vb            ' Une journée comptabilisée pour un point de vente
    │   └── TransactionWU.vb                ' Une transaction identifiée par son MTCN
    ├── Services/
    │   ├── WUFichierService.vb             ' Contrôles de sécurité : type de rapport, concordance des périodes
│   ├── WUReportService.vb              ' Lecture fichiers (ZIP ou texte), parsing, agrégation, dates
    │   ├── WURepository.vb                 ' Lecture SQL Server pour la compensation (T_Pdv_SA / T_Pdv_EC / SystemeWU)
│   ├── PdvRepository.vb                ' CRUD points de vente et groupes (écriture isolée de la lecture)
│   ├── ExcelExportService.vb           ' Export générique de tableaux vers Excel ou PDF (titre, en-têtes, impression)
│   ├── HistoriqueRepository.vb         ' Historique des journées comptabilisées (T_HistoriqueWU)
│   ├── RapportActiviteService.vb       ' Construction des quatre états du rapport d'activité
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
        ├── FrmSousAgentsParGroupe.vb       ' Liste des sous-agents par groupe (consultation)
        ├── FrmRapportActivite.vb           ' Rapport d'activité sur une période (4 états)
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
présents dans `T_Pdv_SA`. Il est **partiel et répétable** : il migre tous les groupes sains et
laisse de côté ceux qui demandent un arbitrage, en disant lesquels et pourquoi.

| Groupe | Repris ? | Motif |
|---|---|---|
| Sous-agents portant tous les mêmes valeurs, comptes non revendiqués ailleurs | oui | — |
| Sous-agents aux valeurs divergentes | non | Laquelle retenir pour le groupe ? |
| Un de ses comptes déjà utilisé par un autre groupe | non | Violerait l'index unique |
| Déjà enregistré lors d'une exécution précédente | non | Jamais retouché |

**Un groupe non repris n'est pas un groupe perdu** : l'application continue de le proposer, avec
les valeurs lues dans `T_Pdv_SA`, permet d'y rattacher de nouveaux sous-agents, et de
l'enregistrer d'un clic après arbitrage. Rien n'est bloqué en attendant — voir « groupes
hérités » ci-dessous. Le script se relance autant de fois que nécessaire.

La pose d'une clé étrangère `T_Pdv_SA.GroupeStatistique → T_GroupeStatistique.Groupe` est
proposée en fin de script, **non exécutée** : elle échouerait tant qu'il reste des sous-agents
sans groupe.

### Groupes hérités : l'application fonctionne avant, pendant et après la migration

Un groupe encore porté par les seuls sous-agents de `T_Pdv_SA`, sans ligne dans
`T_GroupeStatistique`, est dit **hérité**. Il est listé partout comme les autres, avec les
valeurs lues dans `T_Pdv_SA`, et se comporte normalement :

- dans l'écran **Sous-agents**, il se choisit dans la liste déroulante et transmet ses valeurs ;
  au moment d'enregistrer, l'application propose de le reprendre dans la table — un refus
  n'empêche jamais le rattachement ;
- dans l'écran **Groupes**, il apparaît **en jaune** avec la colonne *Enregistré* à faux ;
  le bouton *Enregistrer* l'insère dans la table plutôt que de le modifier.

Si `T_GroupeStatistique` n'existe pas encore du tout, la lecture se rabat automatiquement sur
les groupes hérités : **l'application reste pleinement utilisable avant toute migration**.

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

### Bouton « Liste par groupe… » — restitution du paramétrage

Depuis l'écran des sous-agents, ce bouton ouvre un état de **consultation seule** organisé en
deux grilles :

| Grille | Contenu |
|---|---|
| Récapitulatif (haut) | Un ligne par groupe : libellé, nombre de sous-agents, compte d'activité, compte de commission, taux, état |
| Détail (bas) | Les sous-agents, tous groupes confondus ou ceux du seul groupe choisi |

Cliquer un groupe du récapitulatif — ou le choisir dans la liste déroulante — restreint le
détail à ce groupe. En affichage complet, la liste est triée par groupe puis par Account, et la
teinte de fond change **à chaque changement de groupe** : les blocs se distinguent sans qu'il
faille de ligne de séparation.

Deux partis pris :

- **Les sous-agents sans groupe ne sont pas omis** : ils forment une ligne `(sans groupe
  statistique)` en fin de liste, sans compte ni taux. Ce sont précisément ceux qui n'héritent de
  rien, et qu'il faut voir.
- **La colonne « État » reprend les mêmes termes que l'écran des groupes** : *Hérité — pas encore
  enregistré*, ou *N sous-agent(s) désynchronisé(s)*. Un paramétrage à régulariser se repère donc
  depuis cet état, sans changer d'écran.

La fenêtre s'ouvre en mode non modal : elle peut rester affichée pendant la saisie d'une fiche.
Les données sont lues une fois à l'ouverture — le bouton *Actualiser* les relit après une
modification.

Le bouton **« Exporter vers Excel… »** produit un classeur reprenant l'état **tel qu'il est
affiché, filtre compris** — ce qui est imprimé est ce qui a été vu :

| Élément | Contenu |
|---|---|
| Bandeau de titre | « ECOBANK TCHAD — SOUS-AGENTS PAR GROUPE STATISTIQUE », fond bleu, texte blanc |
| Sous-titres | Groupe affiché (ou « Tous les groupes »), effectifs totaux, date et heure d'édition |
| Tableau 1 | Récapitulatif par groupe |
| Tableau 2 | Détail des sous-agents, avec filtre automatique |

**Quand un groupe est retenu, l'état ne porte que sur lui** : le récapitulatif se restreint à sa
seule ligne, et trois éléments passent en **jaune** — la ligne de sous-titre qui le nomme, sa
ligne du récapitulatif, et le titre du tableau de détail (« Sous-agents du groupe « X » (N) »).
Le détail ne contenant alors que ce groupe, ses lignes ne sont pas surlignées. En affichage
« tous les groupes », le récapitulatif les reprend tous et rien n'est mis en exergue.

Mise en forme : en-têtes en gras sur fond bleuté, bordures sur toutes les cellules, taux au
format pourcentage, colonnes ajustées. Mise en page d'impression : **paysage, ajusté à la
largeur d'une page**, bandeau de titre répété en haut de chaque page, pied de page numéroté
(`Page 1 / 3`) et horodaté. Le classeur est enregistré à l'emplacement choisi puis **ouvert
dans Excel**.

Les valeurs numériques partent en tant que **nombres** et non en texte : elles restent
calculables et triables dans Excel. Le nom de fichier proposé porte le groupe et l'horodatage
(`SousAgents_RESEAU_20260913_1432.xlsx`), de sorte que deux extractions ne s'écrasent pas.

L'export est assuré par `ExcelExportService`, générique : il ne connaît que des `DataTable`,
produit indifféremment un classeur Excel ou un PDF, et peut donc servir à d'autres états. Comme l'export de la pièce comptable, il fonctionne en
**liaison tardive** — aucune référence COM Excel n'est imposée au projet, dont le numéro de
version diffère d'un poste à l'autre. Excel absent du poste, l'application continue de
fonctionner : seul l'export le signale, au moment où il est demandé.

## Rapport d'activité sur une période

Bouton **« Rapport d'activité… »** de l'écran principal. Cinq états, en cinq onglets, tous
exportables **en un seul document PDF**.

| Onglet | Contenu |
|---|---|
| 1. Synthèse | Volumes, envois, paiements, commissions et taxes, plus la **répartition entre sous-agents, agences propres et Accounts non paramétrés** |
| 2. Jour par jour | Une ligne par journée comptabilisée : c'est la page qui fait ressortir un jour anormal |
| 3. Par point de vente | **Sous-agents et agences propres séparés**, chacun avec son sous-total, classés par principal envoyé décroissant |
| 4. Par groupe statistique | Une ligne par groupe, les points de vente sans groupe formant une ligne distincte |
| 5. Évolution des commissions | Jour par jour : les trois commissions, leur total, la variation par rapport à la veille et le cumul de la période |
| 6. Transactions (MTCN) | Le détail transaction par transaction : date, Account, groupe, MTCN, sens, statut et montant |

Les **annulations** figurent dans toutes les pages de détail — et plus seulement en synthèse —
sous une colonne dédiée : sans elle, il était impossible de savoir *qui* annule, alors que
c'est précisément ce que l'on veut suivre.

**Filtre par groupe statistique.** Une liste déroulante restreint les cinq états à un seul
groupe ; l'export porte alors sur ce périmètre, le nom du groupe figure en jaune dans le
classeur et dans le nom de fichier proposé. Les groupes proposés sont ceux **présents dans
l'historique de la période**, et non ceux du paramétrage courant : un groupe supprimé depuis
reste consultable sur les journées où il existait.

### Pourquoi un PDF et non un classeur Excel

Le rapport est un état édité, destiné à circuler : le figer évite qu'il soit retouché après
coup, volontairement ou non. Le PDF ne s'ouvre pas dans un tableur et ne se modifie pas au fil
de l'eau.

Il est produit **par Excel**, à partir d'un classeur **invisible et jamais enregistré** : la
mise en page d'impression déjà en place (paysage, ajusté à la largeur d'une page, bandeau de
titre répété en haut de chaque page, pied de page numéroté et horodaté) est reprise telle
quelle, et aucun fichier intermédiaire ne subsiste sur le disque. Le document s'ouvre ensuite
dans le lecteur PDF du poste.

> **Ce que le PDF ne fait pas.** Il n'est pas infalsifiable : un PDF reste modifiable avec
> l'outil adéquat. Pour une valeur probante, il faudrait le signer électroniquement — un
> dispositif qui relève de la banque, pas de cette application.

Excel reste nécessaire sur le poste, non pour ouvrir le résultat mais pour le mettre en page.
Son absence est signalée au moment de l'export ; le rapport reste consultable à l'écran.

La **liste des sous-agents par groupe** conserve son export Excel : c'est une liste de travail,
que l'on trie et filtre, non un état à figer.

Trois précisions sur ces états :

- les **Accounts non paramétrés** forment une catégorie à part, ni sous-agents ni agences
  propres : les ranger avec les secondes reviendrait à affirmer ce que précisément on ignore ;
- la **variation** de la page 5 est laissée vide pour la première journée et lorsque la veille
  est à zéro — une variation depuis zéro n'a pas de sens, et afficher 100 % induirait en erreur ;
- une nature de point de vente absente de la période n'apparaît pas : pas de ligne à zéro.

Les quatre états sont bâtis sur **la même lecture**, agrégée différemment : leurs totaux sont
donc nécessairement identiques d'une page à l'autre. Le rapprochement entre pages est un
contrôle de cohérence, pas une coïncidence.

### Suivi des MTCN

`T_HistoriqueWU` agrège la journée par point de vente : elle ne peut pas porter le MTCN, qui
identifie **une** transaction. La table **`T_HistoriqueMTCN`** (script
`Scripts\06_HistoriqueMTCN.sql`) conserve donc le détail, une ligne par transaction d'envoi ou
de paiement, et permet de retrouver et de justifier une opération précise.

- Alimentée **dans la même transaction** que l'agrégat : les deux tables ne peuvent pas diverger.
- Les **annulations y figurent**, signalées par leur statut et surlignées en rose à l'écran :
  c'est le plus souvent une transaction annulée que l'on cherche.
- Le filtre par groupe statistique est appliqué **par la base**, plutôt que de rapatrier toute
  la période pour la trier ensuite.
- Volume : de l'ordre de 400 transactions par jour, soit environ 100 000 lignes par an.
- Pas de clé primaire sur `(DateActivite, Account, MTCN, Sens)` : rien ne garantit qu'un MTCN
  ne puisse pas apparaître deux fois le même jour pour le même point de vente — un ajustement
  en produirait un — et une contrainte trop stricte ferait échouer l'historisation d'une journée
  par ailleurs valable. L'unicité technique est assurée par une colonne d'identité.

Dans le PDF, le détail est inclus **après confirmation** au-delà de 2 000 transactions : sur un
mois complet, l'ajouter sans le dire produirait des dizaines de pages que personne n'attendait.
Il reste alors consultable à l'écran.

### Agences propres restées sans activité

L'état par point de vente reprend **toutes** les agences propres du paramétrage, y compris
celles qui n'ont rien fait sur la période — elles apparaissent à zéro, et leur nombre est
rappelé dans le sous-total. Une agence sans activité est une information de gestion ; l'omettre
reviendrait à la rendre invisible au moment même où elle mérite d'être regardée.

Ce complément ne s'applique pas lorsqu'un groupe statistique est sélectionné : une agence propre
n'appartient à aucun groupe, ceux-ci ne concernant que les sous-agents.

### La source : l'historique des journées comptabilisées

L'application traite une journée à la fois et ne conservait rien. La table
**`T_HistoriqueWU`** (script `Scripts\05_HistoriqueWU.sql`) enregistre désormais le résultat de
chaque journée — une ligne par point de vente — et sert de source à ces rapports.

- **Alimentée à la génération de la pièce comptable**, jamais à un simple affichage : seule une
  journée réellement comptabilisée entre dans l'historique.
- **Restitue ce qui a été comptabilisé**, et non un recalcul a posteriori qui dépendrait du
  paramétrage du jour. L'identification du point de vente — désignation, groupe, type — est
  recopiée telle qu'elle était ce jour-là : changer un point de vente de groupe ne réécrit pas
  le passé.
- **Écriture répétable** : une journée regénérée remplace intégralement ses lignes. Suppression
  et réinsertion dans une même transaction, pour qu'un incident laisse la journée dans son état
  antérieur plutôt qu'amputée.
- **Un échec d'historisation n'annule jamais la pièce comptable** : elle est déjà générée et
  équilibrée. L'utilisateur est averti que la journée manquera aux rapports tant qu'elle n'aura
  pas été regénérée.

L'historique ne couvre que les journées comptabilisées **après** la mise en service. Pour les
périodes antérieures, il faut regénérer les journées concernées à partir de leurs rapports.

### Volumes

Le nombre de transactions d'envoi et de paiement est désormais compté à l'agrégation
(`ActiviteAgregat`), ainsi que le nombre de transactions **annulées** — exclues des montants,
mais comptées : une journée riche en annulations mérite d'être regardée.

### Ce qui ne figure pas dans ces rapports

La répartition banque / sous-agent des commissions. Elle dépend du taux du groupe **au moment
de l'édition**, et non au moment des opérations : un rapport rétroactif deviendrait faux dès
qu'un taux change. L'historique conserve les commissions totales, qui elles ne bougent pas.

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
