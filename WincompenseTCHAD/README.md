# WincompenseTCHAD

Automatisation de la comptabilisation Western Union J+1 (Tchad) — application Windows Forms
en VB.NET (.NET Framework 4.8 / Visual Studio 2019 / SQL Server Express).

## Ouverture du projet

> **Le projet s'appelle `WincompenseTCHAD`, l'exécutable produit s'appelle `Wincompense.exe`.**
> Ce n'est pas une incohérence : `AssemblyName` nomme le fichier livré, `RootNamespace` nomme les
> types à l'intérieur. La banque ne voit que le premier, et son nom devait être neutre du pays.

1. Ouvrir `WincompenseTCHAD.sln` dans Visual Studio 2019.
2. Vérifier que le Framework cible du projet est bien **.NET Framework 4.8**.
3. Adapter la chaîne de connexion SQL Server dans `WincompenseTCHAD\App.config` si nécessaire
   (par défaut : `Server=.\SQLEXPRESS;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;`).
4. Exécuter **tous** les scripts SQL du dossier `Scripts\`, dans l'ordre numéroté, sur
   l'instance `.\SQLEXPRESS`. Les scripts `07` et `08` ne sont pas optionnels : sans le `07`,
   la table des utilisateurs et les colonnes de traçabilité n'existent pas, et l'application ne
   peut ni identifier personne ni écrire.
5. Compiler et lancer (F5). Au premier lancement, l'application propose de créer le premier
   compte administrateur : c'est lui qui créera ensuite les autres.

## L'application : une fenêtre MDI

`FrmPrincipal` est le conteneur MDI et le point d'accès unique aux écrans. Chacun s'ouvre en
fenêtre fille, ce qui permet d'en consulter plusieurs à la fois — comparer un rapport d'activité
et le paramétrage d'un sous-agent, par exemple — là où des boîtes de dialogue modales
l'interdisaient.

| Menu | Écrans |
|---|---|
| **Compensation** | Traitement de la compense, Rapport d'activité |
| **Paramétrage** | Sous-agents, Agences propres, Groupes statistiques, Autorisations du référentiel, Comptes systèmes |
| **Sécurité** | Mon mot de passe, Utilisateurs et connexions |
| **Fenêtres** | Cascade, mosaïques, fermeture de toutes les fenêtres, et la liste des fenêtres ouvertes |

Les entrées qu'un rôle n'a pas le droit d'ouvrir **ne sont pas affichées** — voir
« Utilisateurs, rôles et traçabilité » plus bas.

**Aucun écran n'est ouvert d'office.** Une fois la connexion validée, l'application affiche sa
zone de travail — le nom du logiciel et son numéro de version en filigrane — et c'est l'utilisateur
qui choisit par où commencer. Ouvrir le traitement de la compense d'emblée imposait cet écran au
commercial, qui n'y a pas accès, et faisait attendre l'agent de compense les jours où il venait
seulement consulter un rapport.

### L'écran de traitement dit son propre enchaînement

Les boutons sont numérotés et rangés dans l'ordre où on les actionne — charger l'activité,
charger le règlement, calculer, générer la pièce, produire le fichier — de sorte que la fenêtre
décrit l'enchaînement sans qu'il faille l'avoir appris.

Chaque chargement de rapport se termine par un **message explicite** : le fichier retenu, sa
période, et ce qu'il reste à faire. Le libellé s'inscrivait jusqu'ici en petit dans la barre
d'état, où un agent pouvait le manquer et croire avoir chargé alors que la boîte de dialogue
avait été refermée sans sélection.

Après le calcul, une **alerte liste les Accounts non paramétrés**, en distinguant deux
populations que tout sépare : ceux qui ne seront pas comptabilisés, et ceux qui le seront malgré
une donnée manquante.

**Un Account non paramétré n'est pas comptabilisé.** Il n'apparaît ni dans la pièce comptable,
ni dans le fichier destiné au core banking. Trois cas, et trois seulement : l'Account est absent
de `T_Pdv_SA` comme de `T_Pdv_EC` ; c'est un sous-agent sans compte de compensation ; c'est un
sous-agent sans compte de commission. Une agence propre connue reste toujours comptabilisée —
ses écritures vont sur le compte courant WU par règle métier, et non faute de mieux.

Auparavant, le mouvement d'un Account inconnu était posé sur le compte courant WU de la banque,
c'est-à-dire sur un compte qui n'est pas le sien : l'écart se corrigeait ensuite à la main,
écriture par écriture, une fois la pièce chargée. Le cas du sous-agent sans compte de commission
était pire encore : sa rétrocession n'était posée nulle part alors que sa contrepartie restait
comptée, et la pièce partait en déséquilibre du montant de cette commission — assez pour dépasser
le seuil de mille francs et bloquer la génération sans en dire la raison.

Avant de générer, l'écran **annonce les Accounts écartés, leur motif et le montant que cela
représente**, et demande confirmation — la réponse par défaut étant « non ». L'activité écartée
reste à régulariser : mieux vaut créer le point de vente manquant et relancer le calcul.

Dans la grille de contrôle, **l'écart d'arrondi remonte en quatrième colonne**. Il se trouvait en
avant-dernière position sur vingt-huit, donc hors de l'écran — alors que c'est lui qu'on vient
vérifier. L'Account reste figé à gauche pendant le défilement horizontal, sans quoi on ne sait
plus de quel point de vente on lit les montants.

**Un écran déjà ouvert n'est jamais dupliqué** : il est ramené au premier plan, et rétabli s'il
était réduit. Sans cette règle, dix clics sur un menu produiraient dix copies de la même liste,
chacune avec ses propres données, et l'utilisateur ne saurait plus laquelle fait foi.

Les boutons de paramétrage qui encombraient l'écran de traitement ont disparu : le menu les
porte désormais, et les deux libellés de fichiers chargés occupent toute la largeur. Restent
seuls sur cet écran les boutons de son propre enchaînement — charger, calculer, générer.

La barre d'état rappelle en permanence **qui est connecté** et **le serveur et la base**
auxquels l'application est reliée : une confusion entre l'environnement de test et la
production se voit immédiatement.

**La pièce comptable se regarde avant de sortir.** Elle n'est plus envoyée directement dans
Excel : elle s'affiche d'abord dans sa fenêtre, avec ses totaux et son écart, et un bouton
« Exporter vers Excel » écrit le classeur à l'emplacement choisi puis l'ouvre aussitôt. C'est
exactement la marche du fichier destiné au core banking : ce qui engage la comptabilité de la
banque doit pouvoir être lu avant d'être signé.

**Les fenêtres s'ouvrent centrées.** Une fenêtre fille est placée au centre de la zone de
travail MDI, et rétrécie d'abord si elle y est trop grande — ouverte en haut à gauche à sa
taille de conception, elle débordait sur les postes à petit écran et ses boutons du bas
devenaient inaccessibles. Les boîtes de dialogue se centrent sur la fenêtre qui les ouvre.

Deux ouvertures restent volontairement **modales**, parce qu'elles appartiennent à un
enchaînement et non à la navigation : la création d'un groupe depuis la fiche d'un sous-agent
(bouton « … ») et l'affichage d'une pièce comptable.

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
    │   ├── TransactionWU.vb                ' Une transaction identifiée par son MTCN
    │   ├── UtilisateurWU.vb                ' Un compte utilisateur, son rôle et ses droits
    │   ├── DemandeWU.vb                    ' Une écriture proposée sur le référentiel
    │   └── LigneCoreBankingWU.vb           ' Une ligne du fichier d'interface (13 colonnes)
    ├── Services/
    │   ├── WUFichierService.vb             ' Contrôles de sécurité : type de rapport, concordance des périodes
│   ├── WUReportService.vb              ' Lecture fichiers (ZIP ou texte), parsing, agrégation, dates
    │   ├── WURepository.vb                 ' Lecture SQL Server pour la compensation (T_Pdv_SA / T_Pdv_EC / SystemeWU)
│   ├── PdvRepository.vb                ' CRUD points de vente et groupes (écriture isolée de la lecture)
│   ├── ExcelExportService.vb           ' Export générique de tableaux vers Excel ou PDF (titre, en-têtes, impression)
│   ├── HistoriqueRepository.vb         ' Historique des journées comptabilisées (T_HistoriqueWU)
│   ├── RapportActiviteService.vb       ' Construction des quatre états du rapport d'activité
    │   ├── WUCalculationService.vb         ' Formules, répartition, arrondi
    │   ├── PieceComptableService.vb        ' Grille de contrôle, pièce comptable, équilibrage, export Excel
    │   ├── MotDePasseService.vb            ' Empreintes PBKDF2, robustesse, mots de passe provisoires
    │   ├── SessionWU.vb                    ' Utilisateur connecté et droits : point d'accès unique
    │   ├── UtilisateurRepository.vb        ' Comptes (T_UtilisateurWU) et journal (T_ConnexionWU)
    │   ├── DemandeRepository.vb            ' File des demandes : dépôt, autorisation, rejet
    │   ├── CoreBankingService.vb           ' Fichier d'interface : construction, numéro de lot
    │   ├── CalendrierWU.vb                 ' Jours ouvrés : week-ends et jours fériés
    │   ├── ConfigurationWU.vb              ' Où est le serveur : partage réseau, copie locale, secours
    │   ├── MiseAJourWU.vb                  ' Annonce aux postes qu'une version plus récente est publiée
    │   ├── DiagnosticSqlWU.vb              ' Traduit les refus de SQL Server en consigne exécutable
    │   └── SecretWU.vb                     ' Chiffre le mot de passe SQL pour ce poste (DPAPI)
    └── Forms/
        ├── FrmPrincipal.vb                 ' Fenêtre MDI : menus et ouverture des écrans
        ├── FrmCompensationWU.vb            ' Orchestration des événements uniquement
        ├── FrmCompensationWU.Designer.vb
        ├── FrmCompensationWU.resx
        ├── FrmPieceComptable.vb            ' Affichage d'une pièce (globale ou d'un seul PDV)
        ├── FrmComptesSysteme.vb            ' Paramétrage des comptes comptables
        ├── FrmSousAgents.vb                ' Gestion des sous-agents (CRUD)
        ├── FrmGroupesStatistiques.vb       ' Gestion des groupes statistiques (CRUD)
        ├── FrmSousAgentsParGroupe.vb       ' Liste des sous-agents par groupe (consultation)
        ├── FrmRapportActivite.vb           ' Rapport d'activité sur une période (4 états)
        ├── FrmAgences.vb                   ' Gestion des agences propres (CRUD)
        ├── FrmConnexion.vb                 ' Écran de connexion, amorçage du premier administrateur
        ├── FrmChangerMotDePasse.vb         ' Changement de mot de passe (imposé ou volontaire)
        ├── FrmDiagnostic.vb                ' Affiche un diagnostic long, avec bouton Copier
        ├── FrmUtilisateurEdition.vb        ' Création et modification d'un compte
        ├── FrmUtilisateurs.vb              ' Liste des comptes et journal des connexions
        ├── FrmDemandes.vb                  ' Autorisations du référentiel (inputer / authorizer)
        ├── FrmFichierCoreBanking.vb        ' Consultation du fichier d'interface avant export
        └── FrmParametresConnexion.vb       ' Changement de serveur, pour ce poste ou pour toute la banque

Scripts/
├── 00_InstallationComplete.sql            ' LES DIX SCRIPTS EN UN SEUL, à donner à l'informatique
├── 01_CreateTables_GWC_WINCOMPENSE_ETD.sql
├── 02_DonneesExemple.sql
├── 03_SystemeWU.sql                       ' Comptes comptables paramétrés
├── 04_GroupeStatistique.sql               ' Table des groupes + migration depuis T_Pdv_SA
├── 05_HistoriqueWU.sql                    ' Historique des journées comptabilisées
├── 06_HistoriqueMTCN.sql                  ' Détail des transactions, MTCN par MTCN
├── 07_Utilisateurs.sql                    ' Utilisateurs, journal des connexions, traçabilité
├── 08_RolesSQLServer.sql                  ' Rôles de base de données wu_compense / wu_commercial / wu_admin
├── 09_Demandes.sql                        ' Double regard : file des demandes et fonction des utilisateurs
├── 10_JoursFeries.sql                     ' Jours fériés : contrôle de la date de valeur
└── 11_AccesUtilisateurs.sql               ' Rattachement des comptes Windows aux rôles

Installation/
├── Wincompense.iss                        ' Script Inno Setup : produit Wincompense_Setup.exe
├── connexion.config.modele                ' Modèle du fichier à poser sur le partage réseau
├── Configurer-Connexion.ps1               ' Changement de serveur en ligne de commande
├── version.txt.modele                     ' Modèle du fichier annonçant la version publiée
├── LISEZMOI-Installation.md               ' Déploiement, changement de serveur, dépannage
└── PLAN-DEPLOIEMENT.md                    ' Les huit phases du déploiement, à cocher
```

## Installation sur les postes, et changement de serveur

La banque change souvent de serveur. Tout le dispositif est bâti autour de cette contrainte :
**changer de serveur ne doit pas envoyer un technicien dans les bureaux.**

### La connexion ne vit pas avec l'application

La chaîne vivait dans `Wincompense.exe.config`, à côté de l'exécutable. Trois défauts qui se
cumulaient : le fichier est dans `Program Files`, donc protégé ; il est propre à chaque poste ; et
une réinstallation l'écrase.

Elle est désormais cherchée dans un ordre où le premier trouvé l'emporte :

| Rang | Emplacement | À quoi il sert |
|---|---|---|
| 0 | **Saisie pour cette session** | Dépannage à l'installation — **rien n'est écrit** |
| 1 | Variable d'environnement `WINCOMPENSE_CONNEXION` | Dépannage, poste de test |
| 2 | **Fichier partagé** désigné à l'installation | **La source de vérité** |
| 3 | `%PROGRAMDATA%\Wincompense\wincompense.config` | Copie locale, rafraîchie à chaque lecture réussie du partage |
| 4 | `App.config` | Poste de développement |
| 5 | Valeur compilée `.\SQLEXPRESS` | Dernier recours |

**Changer de serveur, c'est modifier une ligne dans un fichier.** Les postes la prennent au
démarrage suivant.

Le rang 3 n'est pas un doublon : c'est lui qui fait travailler le poste le matin où le partage est
injoignable. Sans lui, une coupure du serveur de fichiers arrêterait la compense de toute la
banque. La copie est rafraîchie pendant que le partage répond, jamais quand il ne répond plus.

L'authentification est celle de **Windows**. Aucun mot de passe ne circule — et c'est précisément
ce qui autorise à poser la configuration sur un partage lisible par tous. Les droits d'accès à la
base restent donnés par `Scripts\08_RolesSQLServer.sql`.

### Trois façons de changer de serveur

1. **Menu Sécurité → Connexion à la base de données…**, réservé à l'administrateur Wincompense.
   L'écran teste la connexion avant d'enregistrer, dit **d'où vient** la chaîne en service, et
   fait confirmer tout réglage appliqué à l'ensemble des postes.

   La case **« Conserver ce réglage sur ce poste » est décochée par défaut**, et ce n'est pas
   un détail. Le geste ordinaire est un dépannage : sans la case, le réglage vaut le temps de
   la session, rien n'est écrit, et le poste continue de suivre la chaîne publiée avec
   l'application. Un réglage conservé, lui, l'emporte sur cette chaîne — **définitivement** :
   le poste cesse de suivre les republications, ce que l'écran annonce sans détour avant
   d'écrire. Le déploiement ClickOnce rend cette distinction décisive, puisque la chaîne y est
   publiée avec l'application et qu'un réglage local rendrait le poste sourd à toute
   correction.
2. **Le Bloc-notes** : le fichier partagé est une suite de lignes `CLE=VALEUR`. Ni XML, ni
   registre — l'informatique doit pouvoir agir sans outil et sans casser une balise.
3. **`Configurer-Connexion.ps1`**, pour une migration faite hors des heures de bureau.

Quand la banque fournit non pas un nom de serveur mais une **chaîne de connexion complète** —
chiffrement imposé, partenaire de secours, port particulier —, elle se colle telle quelle : case
« Employer une chaîne de connexion complète » dans l'écran, ou clé `CHAINE=` dans le fichier.
Elle prime alors sur le serveur et la base, qui se grisent : deux descriptions du même serveur
dont une seule compterait rendraient le fichier trompeur.

Une chaîne portant un mot de passe n'est **jamais** propagée sur le partage : il y serait lisible
en clair par tous les utilisateurs, puisque le partage doit être lisible par tous pour que le
dispositif fonctionne. L'application le refuse, et renvoie vers l'authentification Windows.

Les trois testent ou font tester la connexion avant d'écrire : un serveur mal orthographié propagé
à toute la banque arrêterait tout le monde, et se corrigerait depuis un poste qui ne se connecte
plus.

### Le programme d'installation

L'assistant d'installation accepte indifféremment **un nom de serveur ou une chaîne de
connexion complète**, dans le même champ : la banque fournit généralement la seconde. La
distinction se fait sur le signe `=`, qu'une chaîne porte toujours et qu'un nom de serveur ne
porte jamais — il n'y a donc rien à choisir, et rien à expliquer au technicien qui installe.

`Installation\Wincompense.iss` produit `Wincompense_Setup.exe` avec **Inno Setup 6.3 ou
supérieur**. Attention au point de licence : depuis la version 7, Inno Setup n'est plus gratuit
pour un usage commercial, et une banque l'est — la **6.4.x**, gratuite pour tout usage, reste le
choix le plus simple, et le script compile à l'identique. Il vérifie .NET Framework 4.8, copie l'application et les scripts SQL, **demande le
serveur et le chemin du partage**, écrit la configuration, crée le fichier partagé s'il n'existe
pas encore, et pose les raccourcis.

Deux points méritent d'être signalés :

- le dossier `%PROGRAMDATA%\Wincompense` est créé avec le droit de modification pour les
  utilisateurs. Sans cela, l'application ne pourrait pas rafraîchir sa copie locale, et
  l'administrateur ne pourrait pas changer de serveur depuis l'écran prévu pour cela ;
- le fichier partagé n'est créé que s'il **n'existe pas**. Sur le deuxième poste installé, il
  porte déjà le réglage de la banque : l'écraser avec la saisie d'un technicien ferait basculer
  tout le monde par accident.

La désinstallation laisse `%PROGRAMDATA%\Wincompense` en place, pour qu'une réinstallation
retrouve le serveur sans ressaisie, et ne touche jamais au fichier partagé, qui appartient à la
banque et non au poste.

### Un administrateur ne se crée jamais dans une base non désignée

Au tout premier démarrage, l'application propose de créer le premier administrateur. Un compte
créé dans la mauvaise base est une faute silencieuse : tout semble avoir fonctionné, et personne
ne revient jamais voir cette base.

La proposition nomme donc **le serveur, la base et la provenance de la chaîne**, et offre une
troisième réponse — *régler d'abord le serveur* — plutôt que le seul choix entre créer et
renoncer. Le technicien qui installe peut ainsi désigner la bonne base, créer le compte, et
laisser le poste revenir à la chaîne publiée au démarrage suivant.

### Aucun poste n'ignore qu'une version plus récente existe

Les postes ne se mettent pas à jour seuls : l'informatique garde la main sur le moment, et
pendant la marche en parallèle deux versions différentes fausseraient la comparaison avec la
pièce manuelle. Mais rien n'avertissait un agent qu'il travaillait sur une version dépassée —
une correction livrée un lundi pouvait rester ignorée d'un poste pendant des semaines, et deux
agents produire des pièces différentes à partir des mêmes rapports.

L'application lit donc `version.txt` sur le partage au démarrage — à côté de `connexion.config`,
donc sans réglage supplémentaire — et se compare à lui. En retard, elle l'inscrit **dans sa barre
d'état**, où un clic ouvre le dossier du programme d'installation, et le dit **une seule fois par
version** : répété chaque matin, le message serait fermé sans être lu, et le suivant le serait
aussi.

La comparaison porte sur les nombres et non sur le texte — `1.10.0.0` est postérieure à
`1.9.0.0`, alors qu'elle la précède alphabétiquement.

Tant que ce fichier n'existe pas, rien ne s'affiche et rien n'échoue. Le dispositif est un
confort, jamais une condition de démarrage — comme la copie locale de la connexion.

Le détail — prérequis, droits à poser sur le partage, publication d'une version, dépannage — est
dans `Installation\LISEZMOI-Installation.md`.

## Hypothèses métier retenues (à valider)

1. **InclureLigneReglement** : toutes les lignes du rapport de règlement sont incluses par défaut
   (y compris `TransactionType = "A"`), dès lors que l'Account est renseigné. Fonction isolée,
   volontairement simple, à affiner selon consigne métier ultérieure.
2. **Comptes INCONNU** : toujours affichés dans la grille de contrôle et repris dans l'historique
   d'activité, mais **jamais comptabilisés** — voir « Un Account non paramétré n'est pas
   comptabilisé » plus haut. Leur activité reste à régulariser, et l'écran la chiffre avant de
   générer la pièce.
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
| 3. Par point de vente | **Sous-agents et agences propres séparés**, chacun avec son sous-total, classés par principal envoyé décroissant. **Chaque Account se déroule** sur le détail de ses transactions, MTCN par MTCN |
| 4. Par groupe statistique | Une ligne par groupe, les points de vente sans groupe formant une ligne distincte |
| 5. Évolution des commissions | Jour par jour : les trois commissions, leur total, la variation par rapport à la veille et le cumul de la période |

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
- une nature de point de vente absente de la période n'apparaît pas : pas de ligne à zéro ;
- l'historique enregistre **toute l'activité de la journée**, y compris celle des Accounts qui
  n'ont pas été comptabilisés : le rapport dit ce qui s'est passé, la pièce dit ce qui a été
  comptabilisé. Les deux totaux peuvent donc différer, et c'est précisément ce qui permet de
  chiffrer ce qui reste à régulariser.

Les quatre états sont bâtis sur **la même lecture**, agrégée différemment : leurs totaux sont
donc nécessairement identiques d'une page à l'autre. Le rapprochement entre pages est un
contrôle de cohérence, pas une coïncidence.

### Le déroulé d'un point de vente : ses transactions, MTCN par MTCN

Dans l'état **par point de vente**, un Account porteur de transactions s'ouvre d'un clic sur le
`+` de sa première colonne : ses envois et ses paiements apparaissent alors sous lui, chacun
avec sa date, son MTCN, son sens et son statut. Le même mécanisme vaut pour les agences propres.

Une transaction est **une unité de volume** : son montant alimente la colonne d'envoi ou de
paiement selon son sens, et son compteur la colonne correspondante. Les cumuls du point de vente
au-dessus sont donc exactement la somme des lignes en dessous — **le déroulé justifie le total,
il ne se contente pas de l'accompagner**. Une transaction annulée ne pèse dans aucun montant :
elle ne compte que comme annulation, exactement comme dans l'agrégat de la journée, et sa ligne
est surlignée en rose.

Tout est replié à l'ouverture : un état de gestion se lit d'abord au niveau des points de vente.
Le tri des colonnes est désactivé sur cette page — il détacherait une transaction de son point
de vente.

> **Un point de vente ne se déroule pas ?** C'est qu'aucun détail n'est enregistré pour ces
> journées : leur agrégat est bien là, mais les MTCN qui le composent n'ont jamais été écrits.
> Cause la plus fréquente : ces journées ont été comptabilisées **avant** la mise en place du
> suivi des MTCN. L'application le dit explicitement à l'ouverture du rapport, et la barre
> d'état indique en permanence le nombre de transactions détaillées. Pour l'obtenir : exécuter
> `Scripts\06_HistoriqueMTCN.sql`, puis recharger les rapports de ces journées et regénérer
> leur pièce comptable — une journée regénérée remplace proprement la précédente.

`T_HistoriqueWU` agrège la journée par point de vente : elle ne peut pas porter le MTCN, qui
identifie **une** transaction. La table **`T_HistoriqueMTCN`** (script
`Scripts\06_HistoriqueMTCN.sql`) conserve donc le détail, une ligne par transaction.

- Alimentée **dans la même transaction** que l'agrégat : les deux tables ne peuvent pas diverger.
- Le filtre par groupe statistique est appliqué **par la base**, plutôt que de rapatrier toute
  la période pour la trier ensuite.
- Volume : de l'ordre de 400 transactions par jour, soit environ 100 000 lignes par an.
- Pas de clé primaire sur `(DateActivite, Account, MTCN, Sens)` : rien ne garantit qu'un MTCN
  ne puisse pas apparaître deux fois le même jour pour le même point de vente — un ajustement
  en produirait un — et une contrainte trop stricte ferait échouer l'historisation d'une journée
  par ailleurs valable. L'unicité technique est assurée par une colonne d'identité.

Dans le PDF, le détail est inclus sous chaque point de vente **après confirmation** au-delà de
2 000 transactions : sur un mois complet, l'ajouter sans le dire produirait des dizaines de pages
que personne n'attendait. Répondre « Non » produit l'état au niveau des points de vente
seulement ; le détail reste consultable à l'écran.

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

## Fichier d'interface vers le core banking

La pièce comptable reste l'objet de contrôle, lisible par un comptable. Le fichier décrit ici en
dérive : c'est la forme sous laquelle le core banking accepte d'**impacter réellement** les
comptes des sous-agents et les comptes internes de la banque.

Bouton **« 5. Fichier core banking… »** sur l'écran de traitement, actif une fois la pièce
générée : ce qui est chargé doit être exactement ce que le comptable a vu et validé à l'écran.

**Le fichier est présenté avant d'être écrit.** Un écran de consultation montre les treize
colonnes telles qu'elles seront produites — pas une version arrangée pour la lecture — avec le
nombre de lignes, la date de valeur et le contrôle d'équilibre sous les yeux. L'export n'écrit
rien tant qu'il n'est pas demandé, et **refuse de s'activer sur un fichier déséquilibré**. Une
fois produit, le classeur s'ouvre aussitôt, comme la pièce comptable.

### Les treize colonnes

| # | Colonne | Valeur |
|---|---|---|
| 1 | `DETBSJRNL` | `BBR` |
| 2 | `BRN` | `N01` |
| 3 | `BATCHNO` | numéro de lot, 4 caractères |
| 4 | `SRCCODE` | `ECOSOURCE` |
| 5 | `AMOUNT` | montant, **toujours positif**, entier FCFA |
| 6 | `ACNO` | compte mouvementé |
| 7 | `DRCR` | `D` ou `C` |
| 8 | `ACBRN` | voir ci-dessous |
| 9 | `TXNCD` | `U24` au débit, `F15` au crédit |
| 10 | `VALDT` | jour de la compense, c'est-à-dire la date du jour, au format `jj/mm/aaaa` |
| 11 | `INSTR_NO` | vide |
| 12 | `ADDLTEXT` | libellé de l'écriture |
| 13 | `COST_CENTER` | `10000` |

Une ligne de pièce portant un débit **ou** un crédit donne une ligne de fichier ; le sens part
dans `DRCR` et le montant devient positif.

### `ACBRN` : trois cas, dans cet ordre

1. Les comptes **379100319** et **379200585** sont rattachés d'office à `N01`, quel que soit le
   point de vente qui les a mouvementés.
2. Sinon, le **code agence du point de vente** : `codeagence` pour un sous-agent,
   `CodeAgenc-Voyager` pour une agence propre, selon l'Account — exactement ce que `WURepository`
   lit déjà.
3. À défaut, `N01`. Ce cas couvre la ligne d'écart d'arrondi, qui n'appartient à aucun point de
   vente, et un point de vente dont le code agence manque encore dans la base. Laisser la colonne
   vide ferait rejeter le fichier entier.

Pour que la règle 2 soit possible, la pièce comptable porte désormais une colonne `CodeAgence`.
Elle ne s'affiche pas à l'écran et ne change aucun montant : sans elle, rien ne permettait de
savoir à quel point de vente appartient une ligne.

### Date de valeur : le jour de la compense

`VALDT` porte **le jour où la compense est passée**, c'est-à-dire la date du jour. Les écritures
sont datées du moment où elles impactent réellement les comptes, et non d'une date déduite de la
journée d'activité. Une journée rattrapée trois jours plus tard prend donc la date du rattrapage
— celle que le relevé de compte montrera.

Dans le cas nominal, la compense du jour J se passe le lendemain : `VALDT` vaut J+1 sans qu'il
ait fallu le calculer.

La date n'est plus déduite, mais elle reste **contrôlée**. Une écriture datée d'un jour chômé est
rejetée par le core banking, ou repoussée d'office sans que personne ne le sache : si la compense
est passée un samedi, un dimanche ou un jour férié, l'application le signale, indique le prochain
jour ouvré et demande confirmation avant de produire le fichier. Elle ne décale jamais la date
d'elle-même — c'est la banque qui décide.

Les samedis et dimanches se déduisent du calendrier. Les jours fériés, non : ils changent chaque
année, et les fêtes musulmanes suivent le calendrier lunaire. Ils sont donc tenus dans la table
`T_JourFerieWU`, créée par `Scripts\10_JoursFeries.sql` et complétée par la banque sans
recompiler l'application.

**Le script sème les fêtes à date fixe et les lundis de Pâques jusqu'en 2030 — 35 dates — et
rien d'autre.** Les fêtes musulmanes sont annoncées chaque année et ne se calculent pas d'avance :
elles doivent être ajoutées à mesure. Tant qu'une année n'a aucun jour férié enregistré,
l'application le signale avant de produire le fichier et demande confirmation : une année vide
laisserait passer un jour férié sans que rien ne le signale.

La fenêtre d'aperçu dit en une ligne le rapport entre la journée traitée et la date portée —
« Compense du lendemain (J+1) », « Compense passée 3 jours après la journée traitée » — et
prévient lorsque le rapport chargé porte une journée postérieure à aujourd'hui, ce qui trahit un
fichier qui n'est pas le bon.

### Le numéro de lot n'est pas tiré au hasard

C'est délibéré, et c'est le point le plus important de ce format.

Le numéro de lot est **le seul élément par lequel le core banking peut reconnaître qu'on lui
présente deux fois la même journée**. Un tirage aléatoire produirait deux numéros différents pour
un même fichier réexporté, et les comptes seraient impactés en double sans que rien ne le
signale.

Il est donc dérivé de la date : le nombre de jours écoulés depuis une origine fixe, écrit en
base 36 sur quatre caractères. Deux propriétés en découlent — **une journée donne toujours le
même numéro**, et **deux journées n'en partagent jamais un**, pendant plus de quatre mille ans.
L'origine est calée pour que les numéros aient aujourd'hui la forme de ceux de la banque : le
31 mai 2026 donne `07q4`.

C'est la **journée d'activité** qui le détermine, et non la date de valeur. Celle-ci étant le
jour de la compense, rattraper le lundi les journées du vendredi, du samedi et du dimanche leur
donnerait la même date de valeur. Dérivé de cette date, le numéro aurait été identique pour ces
trois journées, et le core banking les aurait prises pour trois chargements du même fichier.

Le nom du fichier porte pour la même raison la journée d'activité : `WU_CORE_<aaaammjj>_<lot>.xlsx`.

Aucune table n'a été ajoutée : le numéro se recalcule, il n'a pas à être conservé.

### Un fichier déséquilibré ne sort pas

L'équilibre est revérifié au moment de produire le fichier, alors même que la pièce a déjà été
équilibrée à sa génération. Ce fichier impacte des comptes réels et rien ne garantit qu'il soit
produit dans la foulée : chargé déséquilibré, il déséquilibrerait la comptabilité de la banque.

### Écriture du classeur

Ligne d'en-tête puis données, sans titre ni ligne vide : le lecteur est un automate, et un
décalage de colonne ferait rejeter le fichier.

Toutes les colonnes sont écrites en **texte**, sauf `AMOUNT` écrit en nombre. Sans cela Excel
réinterprète ce qu'il croit reconnaître : un numéro de compte perdrait ses zéros de tête, et un
numéro de lot comme `0741` deviendrait le nombre 741 quand `07p1` resterait du texte.

La banque n'impose aucun nom de fichier ; celui retenu rattache sans ambiguïté un fichier
retrouvé dans un dossier à la journée qu'il comptabilise.

## Double regard sur le référentiel

Toute écriture sur les **sous-agents**, les **agences propres** et les **groupes statistiques**
passe par deux personnes : un **inputer** la saisit, un **authorizer** l'autorise. Tant qu'elle
n'est pas autorisée, elle n'existe pour personne — ni pour la comptabilisation quotidienne, ni
pour les applications tierces qui lisent ces tables.

Le but n'est pas d'ajouter une étape, c'est qu'une seule personne ne puisse pas, seule, changer
une donnée qui détermine des écritures comptables. Un taux passé de 90 % à 99 %, ou un compte de
commission remplacé, se voit dans une pièce comptable le lendemain — mais il est trop tard.

### Les dix écritures couvertes

| Objet | Opérations |
|---|---|
| Sous-agents (`T_Pdv_SA`) | Créer · Modifier · Supprimer |
| Agences propres (`T_Pdv_EC`) | Créer · Modifier · Supprimer |
| Groupes statistiques (`T_GroupeStatistique`) | Créer · Modifier · Supprimer |
| Groupes → sous-agents | **Synchroniser** |

La synchronisation est la plus sensible : un seul clic réécrit le compte d'activité, le compte de
commission et le taux de **tous** les sous-agents du groupe. Sans elle, le dispositif ne
protégerait rien — il suffirait de passer par le groupe.

### Rôle et fonction

Le **rôle** dit le domaine, la **fonction** dit le pouvoir.

| | Inputer | Authorizer | Aucune |
|---|---|---|---|
| Ouvrir les écrans du référentiel | ✔ | ✔ | ✔ (consultation) |
| Saisir une création, modification, suppression | ✔ | | |
| Autoriser ou rejeter | | ✔ | |

La fonction ne se propose qu'aux rôles `COMMERCIAL` et `ADMIN` : un agent de la compense n'a
aucun accès au référentiel, lui en attribuer une n'aurait aucun effet.

### Les six règles

1. **Personne ne décide de sa propre saisie** — y compris un administrateur. C'est tout le
   dispositif. La règle est appliquée par `DemandeRepository`, **et redoublée par une contrainte
   `CHECK` de la base** : un `UPDATE` fait à la main dans Management Studio ne peut pas davantage
   la contourner.
2. **Une seule demande en attente par objet.** Deux demandes contradictoires sur le même
   sous-agent s'appliqueraient sinon dans l'ordre où l'authorizer les traite. Garanti par un
   index unique filtré.
3. **Le rejet est motivé**, et le motif est lu par celui qui a saisi.
4. **Rien n'est effacé.** Les demandes autorisées et rejetées restent dans la file : c'est la
   piste d'audit — qui a proposé quoi, qui a décidé, quand.
5. **L'écriture et la décision tiennent dans une seule transaction.** Une demande marquée
   autorisée alors que l'écriture a échoué laisserait croire que la donnée est en base.
6. **Les données déjà présentes sont réputées autorisées.** Pas de validation rétroactive.

### Pourquoi une file séparée, et non une colonne « Statut »

Un commentaire de `PdvRepository` rappelle que les colonnes de `T_Pdv_SA` sont lues par la
comptabilisation quotidienne **et par d'autres applications**. Une ligne non autorisée qui
séjournerait dans cette table serait vue par ces applications, qui n'ont aucune raison de
connaître le nouveau statut : le contrôle serait contourné sans que personne n'y touche.

Avec la file `T_DemandeWU`, les trois tables du référentiel ne contiennent **que** de la donnée
autorisée. Aucun lecteur, interne ou externe, n'a été modifié.

Une seule table couvre les trois objets, qui partagent presque tous leurs champs. Les vues
`V_Demande_SousAgent`, `V_Demande_Agence` et `V_Demande_Groupe` rendent les noms de colonnes de
chaque table cible, pour qu'un auditeur lise `CompteCompense` et non un nom générique.

### L'écran « Autorisations du référentiel »

Menu **Paramétrage**. La grille du haut liste les demandes en attente ; celle du bas montre,
champ par champ, **la valeur actuelle et la valeur demandée**, les lignes qui changent étant
mises en évidence. Autoriser sans voir ce qu'on autorise ne serait qu'un clic de plus.

L'entrée de menu porte le **nombre de demandes en attente** : sans ce rappel, une demande peut
dormir une semaine parce que personne ne sait qu'elle existe.

Un second onglet conserve les demandes décidées, avec leur auteur, leur décideur et le motif des
rejets.

### Cohérence du groupe et de son miroir

Enregistrer un groupe réalignait aussitôt ses sous-agents, `T_Pdv_SA` étant le miroir que lit la
comptabilisation. Cette règle est conservée : **autoriser une modification de groupe réaligne ses
sous-agents dans la même transaction**. Sans cela, le groupe serait à jour et ses sous-agents
non, et la pièce comptable du lendemain utiliserait les anciens comptes sans que rien ne le
signale.

### Mise en service

Après `Scripts\09_Demandes.sql`, **aucun utilisateur n'a de fonction** — pas même
l'administrateur. Personne ne peut donc créer de sous-agent tant que les fonctions ne sont pas
attribuées, dans l'écran « Utilisateurs et connexions ».

C'est voulu : la première décision à prendre est qui saisit et qui autorise. Il faut au moins un
inputer et au moins un authorizer, et **ce ne peut pas être la même personne**.

Prévoyez **plusieurs authorizers** : avec un seul, une semaine d'absence bloque toute création de
point de vente.

## Utilisateurs, rôles et traçabilité

L'application **exige une identification** avant d'afficher le moindre écran : `Program.Main`
ouvre `FrmConnexion`, et ne construit la fenêtre MDI qu'une fois la session ouverte. Tant que
personne n'est connecté, `SessionWU` refuse **tous** les droits — une erreur d'enchaînement ne
peut donc pas ouvrir l'application sans identification.

### Mot de passe applicatif, et non session Windows

Les postes sont nominatifs, mais un agent qui laisse sa session ouverte ne doit pas pour autant
laisser l'accès à la compense. Chaque utilisateur a donc son propre mot de passe applicatif.
Le compte Windows et le nom de la machine sont néanmoins relevés et journalisés : un identifiant
utilisé depuis un poste qui n'est pas le sien se repère ainsi.

### Les trois rôles

| | Agent de la compense | Commercial | Administrateur |
|---|---|---|---|
| Charger les rapports, calculer, générer et historiser les pièces | ✔ | | ✔ |
| Sous-agents, agences propres, groupes statistiques | | ✔ | ✔ |
| Comptes comptables de la pièce (table `SystemeWU`) | | | ✔ |
| Comptes utilisateurs et journal des connexions | | | ✔ |
| Rapports d'activité | ✔ | ✔ | ✔ |

**L'agent de compense n'a aucun accès au paramétrage, pas même en lecture** : décision de la
banque. C'est ce paramétrage qui détermine les écritures ; le laisser modifiable par celui qui
les génère supprimerait le contrôle croisé recherché.

Les entrées de menu correspondantes sont **masquées et non grisées** : une entrée grisée invite
à demander le droit, alors que la question est tranchée. Le contrôle est refait à l'ouverture de
chaque écran — un écran qui ne compte que sur le menu pour être protégé ne l'est pas.

### Les mots de passe ne sont jamais enregistrés

Seule une empreinte **PBKDF2** l'est, avec son sel (16 octets) et son nombre d'itérations
(100 000). Un mot de passe perdu ne se retrouve donc pas : il se réinitialise, ce qui est le
comportement attendu. Le nombre d'itérations est conservé **avec chaque empreinte** plutôt que
fixé une fois pour toutes, afin de pouvoir être relevé quand les machines seront plus rapides
sans invalider les comptes existants — chacun se vérifie avec le sien.

PBKDF2 est retenu parce qu'il est fourni par le framework (`Rfc2898DeriveBytes`) et ne demande
aucune bibliothèque supplémentaire, contrainte posée dès l'origine du projet.

Un mot de passe doit comporter au moins 8 caractères, dont une majuscule, une minuscule et un
chiffre, et ne pas contenir l'identifiant. La règle est volontairement courte : empiler les
contraintes pousse les utilisateurs à noter leur mot de passe, ce qui dégrade la sécurité au
lieu de l'améliorer.

### Réinitialisation : un mot de passe provisoire, tiré au hasard

L'administrateur ne saisit jamais le mot de passe de quelqu'un d'autre : il demande une
réinitialisation, l'application **tire** un mot de passe provisoire de 12 caractères et
l'affiche une seule fois. Un administrateur qui choisirait lui-même ces mots de passe finirait
par leur donner toujours le même. Les caractères ambigus (`O` et `0`, `I`, `l` et `1`) en sont
écartés, puisqu'il sera lu puis recopié à la main, et le tirage vient du générateur
cryptographique : un mot de passe prévisible, même provisoire, laisse la porte ouverte jusqu'à
son changement.

Le titulaire devra en choisir un autre **avant d'accéder à l'application** : le changement est
imposé à la connexion suivante, pas renvoyé à plus tard.

### Verrouillage après échecs

Au bout de **5 échecs consécutifs**, le compte est verrouillé et seul un administrateur peut le
débloquer. Le message affiché ne distingue jamais « identifiant inconnu » de « mot de passe
erroné » : il ne doit pas renseigner sur l'existence d'un compte.

### Le dernier administrateur ne peut pas se retirer ses droits

Changer le rôle ou désactiver un compte est refusé s'il ne reste **aucun autre administrateur
actif et utilisable**. Sans cette règle, la base se retrouverait sans personne pour créer ou
débloquer un compte, et il faudrait rouvrir SQL Server à la main pour s'en sortir.

### Premier démarrage

Le script `07_Utilisateurs.sql` ne crée **aucun compte**, et surtout pas un compte à mot de
passe connu : une empreinte figée dans un fichier versionné est un mot de passe publié, que
personne ne pense ensuite à changer.

C'est l'application qui s'en charge : si la table ne contient aucun administrateur actif dont
l'empreinte est exploitable, l'écran de connexion propose de créer ce premier compte et demande
son mot de passe, haché comme tous les autres. C'est le seul moment où un compte se crée sans
que personne ne soit connecté.

La longueur de l'empreinte et du sel est contrôlée, et pas seulement leur présence : une ligne
posée à la main dans SQL Server avec une empreinte de fortune ouvrirait un compte administrateur
inutilisable, dont l'existence empêcherait pourtant la création du vrai.

### Journal des connexions (`T_ConnexionWU`)

**Toute** tentative est journalisée, réussie ou non, avec sa date, le poste, le compte Windows
et le motif du refus. Un refus non enregistré serait précisément celui qu'un auditeur
chercherait. Y figurent aussi les changements de mot de passe, les réinitialisations et les
déverrouillages.

Un échec d'écriture au journal **n'empêche jamais la connexion** : le journal sert à l'audit,
pas au contrôle d'accès, et priver un agent de son outil parce qu'une trace n'a pas pu s'écrire
serait disproportionné.

L'écran « Utilisateurs et connexions » en présente les 100, 500 ou 2 000 dernières lignes.

### Traçabilité des écritures

Le script `07_Utilisateurs.sql` ajoute, **sans rien détruire**, les colonnes `CreePar`,
`DateCreation`, `ModifiePar` et `DateModification` à `T_Pdv_SA`, `T_Pdv_EC`,
`T_GroupeStatistique` et `SystemeWU`, ainsi que `ComptabilisePar` et `DateComptabilisation` à
`T_HistoriqueWU` et `T_HistoriqueMTCN`. Les lignes existantes restent à `NULL` : on ne réécrit
pas un passé que l'on ne connaît pas.

L'application renseigne ensuite ces colonnes à chaque écriture — création et modification d'un
sous-agent, d'une agence, d'un groupe, des comptes comptables, et comptabilisation d'une
journée. Chaque écriture porte donc le nom de son auteur.

> **Le script 07 est obligatoire.** Sans lui, ces colonnes n'existent pas et les écritures
> échouent : exécutez-le avant de lancer cette version.

### Rôles SQL Server (`08_RolesSQLServer.sql`)

Les droits applicatifs ci-dessus sont doublés de trois rôles de base de données —
`wu_compense`, `wu_commercial`, `wu_admin` — qui s'appliquent aux connexions SQL Server
elles-mêmes, y compris à quelqu'un qui contournerait l'application avec SQL Server Management
Studio. Le script est rejouable et se termine par un état des droits réellement accordés.

**Aucun rôle ne reçoit `DELETE` sur le journal des connexions** : un journal que ses propres
utilisateurs peuvent effacer ne prouve rien. L'historique, lui, accepte `DELETE` pour le rôle de
compense, puisque rejouer une journée suppose d'effacer la version précédente.

Le rattachement des comptes Windows aux rôles ne peut pas être écrit d'avance : il dépend de
votre domaine. C'est l'objet de `11_AccesUtilisateurs.sql`, le seul script dont une section doit
être modifiée avant exécution.

### Authentification SQL Server : où vit le mot de passe

La banque peut fournir un **compte SQL Server et son mot de passe** plutôt qu'une
authentification Windows. La chaîne de connexion porte alors un secret, et un secret ne
s'écrit pas n'importe où : le fichier partagé est lisible par tous les utilisateurs de
l'application — c'est ce qui lui permet de valoir pour tout le monde.

La chaîne est donc **coupée en deux**, et les deux moitiés ne voyagent pas ensemble :

| Ce qui est dit | Où c'est écrit | Portée |
|---|---|---|
| **Où** est le serveur — serveur, base, nom du compte | `CHAINE=` sur le partage | Toute la banque |
| **Comment** s'y annoncer — le mot de passe | `MOTDEPASSE=` dans le fichier local, chiffré | Ce poste seulement |

Les mêler rendrait impossible de changer de serveur pour tout le monde sans diffuser un
secret à tout le monde. Séparées, la banque change de serveur dans un fichier, et chaque
poste continue de s'annoncer avec ce qu'il tient de lui-même.

Le chiffrement est celui de Windows (**DPAPI**, portée machine, `SecretWU`). Ce qu'il apporte,
et ce qu'il n'apporte pas, mérite d'être dit dans les deux sens :

- **protection réelle** — le mot de passe n'apparaît plus en clair dans `wincompense.config`,
  et le fichier copié sur une autre machine ne donne rien ;
- **ce que ça ne fait pas** — sur ce poste, un programme lancé par un utilisateur local peut
  redemander le déchiffrement à Windows. La portée machine est imposée par le fait que le
  fichier est commun à tous les comptes du poste.

Autrement dit : cela ferme la lecture accidentelle et le vol de fichier, pas l'accès
administrateur à la machine. **L'authentification Windows reste préférable quand la banque
l'accepte**, puisqu'aucun secret n'est alors conservé nulle part.

Trois détails qui se paient cher s'ils sont oubliés :

1. Un poste qui n'a **jamais reçu** le mot de passe lit bien le partage, mais ne se connecte
   pas. La propagation déplace le serveur, pas le secret — l'écran le dit avant d'enregistrer.
2. Le programme d'installation, lui, ne sait pas chiffrer pour Windows. Il écrit
   `MOTDEPASSE_CLAIR=` dans `%PROGRAMDATA%`, que l'application reprend, chiffre et **efface**
   au premier démarrage (`NettoyerLeFichierLocal`). La fenêtre d'exposition se limite à
   l'intervalle entre l'installation et le premier lancement.
3. L'écran de réglage n'affiche plus le mot de passe — il n'est plus dans `CHAINE`. Le bouton
   **Tester** le remet néanmoins (`ChaineEssayable`) : un test qui échouerait là où
   l'application réussit serait le pire des verdicts, celui qui envoie chercher une panne
   ailleurs.

#### Un compte unique partagé supprime le second verrou

Si la banque fournit **un seul** compte SQL pour tout le service, ce compte doit porter la
réunion des droits de tous les postes — donc `wu_admin`. Les trois rôles cessent alors de
distinguer quoi que ce soit au niveau SQL Server : c'est l'application qui garde seule la
distinction entre l'agent de compense, le commercial et l'administrateur, et le verrou qui
s'opposait à une connexion faite **hors** de l'application disparaît. Un compte SQL par agent,
si la banque l'accepte, rend aux trois rôles leur utilité.

### Quand SQL Server refuse le compte (`DiagnosticSqlWU`)

« Login failed for user 'ETD\wincompense' » est exact, et inexploitable : c'est de l'anglais, ça
ne dit pas si le tort est au serveur, à la base ou au compte, et surtout ça ne dit pas quoi
demander à l'informatique. L'agent devant son écran ne peut rien en faire.

`DiagnosticSqlWU` traduit ces refus en une consigne exécutable. Il ne décide de rien et n'écrit
nulle part : il rédige.

| Erreur | Ce qui manque | Ce que le diagnostic propose |
|---|---|---|
| 18456 | Le *login*, au niveau du serveur | `CREATE LOGIN … FROM WINDOWS`, puis l'utilisateur et le rôle |
| 4060, 916 | L'*utilisateur*, au niveau de la base | `CREATE USER … FOR LOGIN …` |
| 229, 230, 262, 297 | Le *rôle* | `ALTER ROLE wu_compense ADD MEMBER …` |
| 4064 | Une base par défaut valable | `ALTER LOGIN … WITH DEFAULT_DATABASE = …` |
| 18452 | Le domaine du poste n'est pas approuvé | Poste et serveur dans des domaines différents |
| −1, 2, 40, 53, 1231, 10060, 10061, 11001 | Le serveur n'a pas répondu | Nom, instance, port 1433, service arrêté |

Trois niveaux, qu'on confond facilement : le **login** ouvre la porte du bâtiment, l'**utilisateur**
celle du bureau, le **rôle** dit ce qu'on a le droit d'y faire. Il faut les trois. Un login sans
utilisateur donne 4060 ; un utilisateur sans rôle donne « SELECT permission was denied » à la
première lecture.

Le nom du compte est lu dans la chaîne de connexion et dans l'identité Windows du poste — **pas**
dans le message d'erreur, qui est traduit dans la langue du serveur et dont le découpage
varierait avec elle.

Le diagnostic s'affiche dans `FrmDiagnostic`, avec un bouton **Copier** : le texte porte le T-SQL
à exécuter, et c'est l'agent qui le transmettra. Le lui faire recopier à la main reviendrait à
lui faire inventer un nom de compte.

Quand `Expliquer` ne reconnaît pas l'erreur, elle rend une chaîne vide et l'appelant garde son
propre message : mieux vaut un message technique qu'un message vague.

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

## Règles tranchées par la banque

- **Agence propre (EC).** La structure de sa pièce est **identique à celle d'un sous-agent** :
  même enchaînement d'écritures, mêmes libellés, mêmes sens. Les trois lignes de commission
  sous-agent n'y figurent pas parce que le taux vaut zéro et qu'une écriture à zéro n'est jamais
  posée — ce n'est pas un traitement à part, c'est la même règle appliquée à 0 %, toute la
  commission revenant à la banque. Sa ligne de mouvement va sur le **compte courant WU** : une
  agence propre n'a pas de compte de compensation dans les livres de la banque.
- **Accounts non paramétrés.** Ils ne sont **pas comptabilisés** — voir plus haut.

## Points restant à confirmer

- Mode d'authentification SQL Server réel en production (actuellement : Windows intégré), et
  rattachement des comptes aux rôles `wu_compense` / `wu_commercial` / `wu_admin`. Les noms de
  domaine de la banque n'étant pas connus d'ici, `Scripts\11_AccesUtilisateurs.sql` porte la
  liste à compléter — c'est le seul script à modifier avant exécution.
- Durée de vie d'un mot de passe : aucune expiration périodique n'est imposée aujourd'hui.
  Faut-il en ajouter une, et à quelle échéance ?
- Les comptes comptables (`SystemeWU`) restent hors du double regard, sur décision de la banque.
  Ce sont pourtant les neuf comptes qui déterminent toute la pièce comptable, et seul
  l'administrateur y touche — seul.
- Format de `VALDT` dans le fichier core banking : `jj/mm/aaaa` a été retenu, faute d'indication
  contraire. À confirmer auprès de l'équipe du core banking avant le premier chargement réel.
- Les comptes 379100319 et 379200585 sont écrits dans `ConstantesWU` et non paramétrés dans
  `SystemeWU` : la banque les a donnés tels quels, et ils décrivent une règle d'aiguillage propre
  au format, non un paramétrage comptable. À basculer en paramètre si le plan comptable bouge.
- Usage exact du compte inter bancaire 381000101 pour l'écart d'arrondi global (voir hypothèse 6).
- Faut-il alimenter les listes déroulantes du formulaire de paramétrage avec le plan comptable
  complet ? Elles ne proposent aujourd'hui que le compte paramétré et le compte par défaut, la
  saisie libre restant possible.
- Sous-agents restant sans groupe statistique après la migration : quel groupe leur attribuer ?
  Tant qu'il en subsiste, la clé étrangère proposée en fin de `Scripts\04_GroupeStatistique.sql`
  ne peut pas être posée.
- Règle définitive de traitement des lignes `TransactionType = "A"` du rapport de règlement.

## Test de référence (section 17)

Les scripts `Scripts\02_DonneesExemple.sql` paramètrent :

- `AHB020200` comme sous-agent (`T_Pdv_SA`, `Taux = 0.70`) ;
- `AHB020057` comme agence propre Ecobank (`T_Pdv_EC`).

Pour valider le parsing/l'agrégation, charger un rapport d'activité contenant `AHB020200` à la
date du 02/01/2021 et vérifier dans `dgvControle` que les agrégats retrouvent approximativement
les valeurs de référence indiquées dans le cahier des charges (Principal Envoi ≈ 5 966 388,
Principal Payé ≈ 7 522 117,45, Commission Envoi = 69 700, Commission Paiement ≈ 98 105, etc.).
Pour `AHB020057`, vérifier qu'aucune commission sous-agent n'est générée (100 % à la banque).
