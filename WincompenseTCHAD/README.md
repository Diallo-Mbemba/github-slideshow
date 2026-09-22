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
    │   ├── SecretWU.vb                     ' Chiffre le mot de passe SQL pour ce poste (DPAPI)
    │   ├── PreparationBaseWU.vb            ' Constate l'accès du compte, le répare ou rédige le script
    │   ├── PieceExcelWU.vb                 ' Le formulaire de pièce de la banque, une feuille par point de vente
    │   └── PieceRepository.vb              ' Conserve et relit les pièces produites (T_PieceWU)
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
        ├── FrmPiecesArchivees.vb           ' Consultation des pièces déjà produites
        ├── FrmRapportSousAgents.vb         ' Rapport d'activité — sous-agents (hérite de FrmRapportActivite)
        ├── FrmRapportAgences.vb            ' Rapport d'activité — agences propres (idem)
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
├── 11_AccesUtilisateurs.sql               ' Rattachement des comptes Windows aux rôles
├── 12_AccesCompteApplicatif.sql           ' LE SCRIPT À REMETTRE À LA BANQUE : un compte, un rôle
└── 13_PiecesComptables.sql                ' Table T_PieceWU : les pièces conservées

Installation/
├── Wincompense.iss                        ' Script Inno Setup : produit Wincompense_Setup.exe
├── Construire-Setup.cmd                   ' Fabrique le setup : trouve ISCC, compile, ouvre Sortie\
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

**Ce qui se propage n'est pas ce qui authentifie.** La résolution se fait en deux temps, et
c'est le cœur du dispositif : `ResoudreLaSource()` dit **où est le serveur** — cela peut venir du
partage, donc valoir pour toute la banque — puis `AppliquerLeMotDePasse()` dit **comment s'y
annoncer**, et cela n'appartient qu'à la machine. Les mêler rendrait impossible de changer de
serveur pour tout le monde sans diffuser un secret à tout le monde.

En authentification **Windows**, aucun mot de passe n'existe et tout tient dans le fichier
partagé. Avec le **compte SQL Server** que la banque fournit — `etdwincompense` —, le mot de
passe est retiré de la chaîne **avant toute écriture**, puis chiffré par Windows (DPAPI, portée
`LocalMachine`, donc valable pour tous les utilisateurs du poste et pour ce poste seul) sous la
clé `MOTDEPASSE` du fichier local. Le partage ne reçoit que le serveur, la base et le **nom** du
compte. Il reste ainsi lisible par tous, ce qui est la condition même de son utilité.

Les droits d'accès à la base restent donnés par `Scripts\08_RolesSQLServer.sql`.

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
   Une seule ligne de mouvement net
   (`NetMouvement = PrincipalEnvoi + ChargeEnvoi + Taxes − PrincipalPaye + TTAReception`,
   Débit si positif — voir « La TTA sur réception entre dans le versement » plus bas) sur le `CompteCompense` du PDV, en contrepartie du compte courant WU pour la
   part nette bancaire ; commissions et taxes sont des lignes de crédit uniquement, sans ligne de
   débit miroir individuelle (voir commentaires détaillés dans `GenererPieceComptable`). Vérifié à
   l'unité près sur l'exemple disponible ; *le cas d'une agence propre "EC" reste à valider faute
   d'exemple de référence pour ce type de PDV.*
4. **TTA sur paiement** : due par un sous-agent, **pas par une agence propre**. **Confirmé par la banque** — ce n'est plus une hypothèse. Elle entre au surplus dans le versement du sous-agent : voir « La TTA sur réception entre dans le versement ».
5. **Solde par Account** (grille de contrôle) = l'opposé de `NetMouvement`, soit
   `PrincipalPaye − (PrincipalEnvoi + ChargeEnvoi + Taxes + TTAReception)`. Il vaut toujours
   l'opposé du mouvement porté sur la pièce : deux chiffres pour la même chose, et l'agent ne
   saurait plus lequel croire.
6. **Cohérence des dates** : la date du rapport d'activité (`txnDateLOC`) est comparée à celle du
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

#### La TTA sur paiement ne concerne que les sous-agents

Les deux taxes sur les flux sont symétriques dans leur calcul — 0,2 % du principal — mais pas
dans leur application :

| | TTA sur envoi (`434000145`) | TTA sur paiement (`434000159`) |
|---|---|---|
| Sous-agent (`SA`) | oui | oui |
| Agence propre (`EC`) | oui | **non — la taxe n'est pas due** |
| Non paramétré (`INCONNU`) | oui | oui *(voir plus bas)* |

**Elle n'est pas calculée, et non pas calculée puis omise de la pièce.** La différence n'est
pas de forme : calculée, elle apparaîtrait dans la grille de contrôle, dans le rapport
d'activité et dans l'historique comme une taxe que la banque devrait, alors qu'elle ne la doit
pas.

La règle tient donc en une condition, **à un seul endroit** : `WUCalculationService.AppliquerFormules`.
Tout ce qui suit en découle sans rien répéter :

| | Conséquence |
|---|---|
| La ligne de la pièce | Disparaît d'elle-même : `AjouterLigneSiNonNul` écarte un montant nul |
| Le compte courant WU | Augmente d'autant — la banque garde la somme |
| L'écart d'arrondi | Reste juste : il part de la même valeur |
| Grille de contrôle, historique, rapport, fichier core banking | Suivent, tous construits sur cette valeur |

Vérifié en simulation sur une agence propre de 45 M d'envois et 38 M de paiements : la ligne
de 76 000 F disparaît, le compte courant WU passe de 8 372 250 à 8 448 250 F, et la pièce
reste équilibrée à zéro. La pièce d'un sous-agent, elle, ne bouge pas d'un franc.

**`INCONNU` n'est pas traité comme `EC` ici**, contrairement à `RepartirCommissions`. Un Account
non paramétré n'est pas une agence propre : il n'est rien encore. Il n'entre de toute façon pas
dans la pièce ; la taxe reste donc calculée, et la grille de contrôle montre ce qu'il faudrait
payer s'il s'avérait être un sous-agent. Mettre à zéro sur une supposition serait affirmer plus
qu'on ne sait.

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

### Deux rapports d'activité, parce que deux populations (`PorteeRapportWU`)

Le rapport est éclaté en **deux fenêtres** : une pour les sous-agents, une pour le réseau
propre. Ce n'est pas une affaire de présentation — les deux populations n'ont pas les mêmes
axes d'analyse :

| | Sous-agents | Agences propres |
|---|---|---|
| Groupe statistique | Oui, c'est leur axe | **Aucun** — `T_Pdv_EC` n'a pas la colonne |
| Taux de rétrocession | 0,60 à 0,95 | **Aucun** |
| Partage de commission | Oui | **Aucun** — la banque garde 100 % |
| Regroupement par agence | Non | **Oui, c'est leur axe** |

Les mêler obligeait l'onglet « par groupe » à ranger toutes les agences dans une ligne vide,
et celui des commissions à additionner une part qui se partage avec une part qui ne se partage
pas. **Séparer n'ajoute pas une vue : cela en retire une fausse.**

`FrmRapportSousAgents` et `FrmRapportAgences` **héritent** de `FrmRapportActivite` et ne font
que poser leur population. Dupliquer huit cents lignes pour changer un filtre aurait garanti
que les deux divergent au premier correctif.

L'onglet « par groupe statistique » et celui de performance **ne coexistent jamais** : celui
qui n'a pas d'objet est retiré, pas laissé vide. Les onglets se retirent d'ailleurs et ne se
masquent pas — `Visible` n'a aucun effet sur un `TabPage`.

#### La performance des agences, à deux niveaux

**Une agence a plusieurs Accounts.** Dans `T_Pdv_EC`, l'Account est la clé primaire mais le
code agence ne porte **aucune contrainte d'unicité** : un même guichet peut tenir plusieurs
points Western Union. Un classement à un seul niveau comparerait donc des agences à des
fractions d'agences.

```
Rang  Agence                 Account     Désignation     Envois    Principal   Commissions   Part
1     001 — AGENCE SIEGE                 2 Account(s)       42   18 250 000     1 240 500  59,3 %
                             TD0001234                      28   12 100 000       820 300  39,2 %
                             TD0009876                      14    6 150 000       420 200  20,1 %
2     004 — AGENCE MOUNDOU               1 Account(s)       31   11 800 000       790 100  37,8 %
```

Trois choix, et leurs raisons :

1. **Le classement se fait sur les commissions**, décroissantes : c'est ce que la banque gagne,
   et la seule colonne qui réponde à « quelle agence rapporte le plus ». Les volumes et les
   montants restent affichés à côté — une agence peut faire du volume sans marge, et c'est
   précisément ce qu'un classement doit laisser voir.
2. **La part** est celle de l'agence dans les commissions de tout le réseau propre. Sans elle,
   une liste de performance n'est qu'une liste : on voit qui est en tête, pas de combien.
3. **Les agences sans activité figurent, à zéro.** Constater qu'une agence n'a rien fait de la
   période est un résultat — et souvent celui qu'on cherchait.

Vérifié en simulation : les parts somment à 100 % aux deux niveaux, une agence sans activité
apparaît bien, et un Account absent du référentiel tombe dans une ligne
« (agence non rattachée) » plutôt que d'être perdu en silence — c'est un signal de
paramétrage, pas un déchet.

**Une limite à connaître.** `T_HistoriqueWU` ne porte pas le code agence : le rattachement
d'un Account à son agence se fait au référentiel **d'aujourd'hui**. Un Account qui changerait
d'agence emporterait tout son passé avec lui. C'est acceptable pour un rapport de gestion — ce
n'est pas un justificatif, et le rattachement ne bouge quasiment jamais — et ça ne l'était pas
pour la pièce comptable, qui est conservée telle quelle. Le jour où cela deviendrait gênant,
une colonne `CodeAgence` dans `T_HistoriqueWU` figerait le passé.

### Consulter une pièce déjà produite (`PieceRepository`, `FrmPiecesArchivees`)

La pièce de chaque journée est **conservée ligne à ligne** dans `T_PieceWU`, et la consulter
c'est **relire ce qui a été écrit** — rien n'est recalculé.

**Pourquoi conserver plutôt que recalculer.** L'historique garde les volumes, les montants et
les totaux de commissions. Il ne garde pas le taux du sous-agent ce jour-là, ni ses comptes de
compensation et de commission, ni la ligne d'écart posée sur le compte inter bancaire.
Reconstituer une pièce ancienne avec le paramétrage d'aujourd'hui réécrirait le passé : un
sous-agent passé de 70 % à 60 % ferait apparaître une pièce qui n'a jamais été visée ni
signée. Une pièce comptable est un justificatif ; « à peu près la même » n'a pas de sens
devant un inspecteur.

**Écrite dans la MÊME transaction que l'historique.** Deux documents d'une même journée ne
doivent jamais diverger : si l'un échoue, aucun des deux n'est écrit.

**Une journée à la fois, jamais une plage.** Les rapports d'activité se consultent sur une
période parce qu'ils cumulent. Une pièce, non : c'est un justificatif daté, rattaché aux deux
rapports Western Union d'UNE journée. Les additionner sur une plage produirait un document qui
ne correspond à aucun téléchargement de la plateforme — donc à rien de vérifiable.

L'écran ne fait qu'aiguiller, et c'est délibéré :

| Bouton | Ce qui s'ouvre |
|---|---|
| **Ouvrir la pièce…** | `FrmPieceComptable`, d'où elle s'exporte au formulaire de la banque |
| **Fichier core banking…** | `FrmFichierCoreBanking`, après confirmation de la date de valeur |

Ce sont les mêmes écrans que le jour de la compense : l'agent y retrouve ses repères, et il
n'y a qu'une façon d'afficher une pièce dans toute l'application.

**Le fichier core banking n'est pas stocké**, et n'a pas à l'être : il dérive entièrement de la
pièce, par les mêmes règles. Le conserver serait garder deux fois la même chose, avec le
risque que les deux copies divergent. La **date de valeur**, elle, se redemande : c'est le jour
où les écritures sont réellement passées, donc aujourd'hui — un fichier rejoué aujourd'hui
porte la date d'aujourd'hui, comme le montrera le relevé de compte.

Deux limites, dites franchement :

1. **Les journées comptabilisées avant la mise en service** de cette conservation n'ont pas de
   pièce. L'écran le dit au lieu d'en inventer une.
2. **Le classeur exporté depuis l'archive porte la pièce globale seule**, sans les onglets par
   point de vente : la liste des `CalculWU` n'est pas conservée. Mieux vaut une pièce fidèle
   sans ses détails qu'un détail reconstitué au paramétrage d'aujourd'hui.

### L'export de la pièce : le formulaire de la banque (`PieceExcelWU`)

La pièce comptable ne s'exporte plus en tableau à quatre colonnes. Elle sort dans **le
formulaire que la comptabilité vise et signe** — celui du modèle fourni par la banque
(`classe_bis.xlsx`) — et le classeur porte **une feuille par point de vente** en plus de la
pièce globale.

**Pourquoi une feuille par point de vente.** La pièce globale équilibre la journée entière.
Mais c'est point de vente par point de vente que la comptabilité contrôle, que le sous-agent
conteste, et que l'inspection remonte. Extraire ces pièces à la main d'un tableau de trois
cents lignes est un travail de recopie, et la recopie se trompe.

**Pourquoi ce formulaire.** Le modèle n'est pas une présentation : c'est le document que le
guichet accepte ou refuse. Une pièce qui ne lui ressemble pas se fait renvoyer, quelles que
soient ses écritures.

Chaque feuille reprend le modèle dans l'ordre :

| Zone | Contenu |
|---|---|
| En-tête | `ECOBANK TCHAD`, `VERIFICATION PIECE COMPTABLE`, la case de contrôle en D6 |
| Identification | `DATE :` (la journée comptabilisée), `DE :`, `POUR :`, `AGENCE:`, et la ligne `Agence  001: …` |
| Tableau | `N° DE COMPTES` / `LIBELLES` / `MONTANTS`, bloc **DEBIT :** puis bloc **CREDIT :** |
| `RAISON :` | Compensation Western Union, désignation du point de vente, journée |
| Cartouches | `SIGNATURES REQUISES`/`FCU`, `INITIE PAR` / `CONTRÔLE PAR` / `APPROUVE PAR`, `ECRITURE`/`OPS`, `PASSEE PAR` / `AUTORISEE PAR`, date et numéro de séquence |

L'en-tête porte quatre valeurs que l'application établit seule :

| Champ | Ce qu'il porte |
|---|---|
| `DATE :` | La **journée comptabilisée** |
| `DE :` | Le service, puis **la personne connectée** — trois agents se relaient sur la compense, et c'est à l'un d'eux que le comptable renverra la pièce |
| `AGENCE:` | **L'agence de rattachement du point de vente**, retrouvée par son code dans le référentiel des agences ; à défaut son code, à défaut l'agence par défaut |
| Numéro (D11) | **Le numéro de lot de la journée**, un tiret, le rang de la pièce : `07q3-002` |

Le numéro de lot est celui-là même que porte le fichier destiné au core banking : quatre
caractères tirés de la date, donc identiques d'une exécution à l'autre pour une même journée,
et différents d'une journée à la suivante. Une pièce et l'écriture qu'elle justifie se
retrouvent ainsi l'une par l'autre — ce qu'un compteur repartant de 1 chaque matin ne
permettrait pas : deux pièces de deux journées porteraient le même numéro 2.

Le référentiel des agences est lu **une fois pour tout le classeur**, et indexé à la fois sur
l'Account et sur le code agence Voyager : une lecture par feuille ferait cinquante
allers-retours vers SQL Server pour une information qui ne bouge pas pendant l'export. Une
erreur SQL rend un référentiel vide plutôt que de faire échouer l'export — une pièce qui
porte un code d'agence au lieu de son nom reste une pièce juste ; une pièce qu'on n'a pas pu
produire, non.

Quatre écarts avec le modèle, tous délibérés :

1. **Le tableau grandit.** Le modèle réserve vingt-cinq lignes parce qu'il est fait pour être
   rempli à la main. Une pièce générée en a autant que la journée en produit : les blocs
   s'étendent, et tout ce qui suit descend d'autant.
2. **La date est celle de la journée comptabilisée**, pas `=TODAY()`. Le modèle porte la
   date du jour parce qu'il est vierge ; une pièce rejouée trois jours plus tard doit rester
   datée du jour qu'elle comptabilise.
3. **La case de contrôle D6 est une formule sur les deux plages**, là où le modèle porte une
   soustraction écrite ligne par ligne. Elle dit la même chose et reste vraie quand la pièce
   change — c'est une case de contrôle.
4. **`DE :` et `POUR :` sont renseignés.** Le modèle les laisse vides ; une pièce produite par
   l'application sait toujours d'où elle vient, et de la part de qui.

`FCU` et `OPS` sont repris tels quels : ce sont les codes de service de la banque, et ils ne
se déduisent de rien.

#### Une pièce tient sur une page, et c'est ce qui commande la géométrie

Le modèle est dessiné pour être **rempli à la main** : colonnes de 24 à 114 caractères, corps
de 20 à 28 points, lignes de 33. Il faut cela pour écrire au stylo dans une case.

Mesuré, ce modèle fait **1 241 points de large et 997 de haut** pour une pièce de six
écritures. Une page A4 en offre 487 sur 734. Il en découle deux impasses, et une seule
sortie :

| | Ce qu'on obtient |
|---|---|
| Imprimé à l'échelle | **Deux pages et demie** — et une pièce comptable en morceaux ne se signe pas |
| Foré sur une page | **39 %** — les libellés à 4 points, illisibles |
| Géométrie ramenée | **519 sur 711 points** — une page à 100 %, et 81 % au-delà de vingt écritures |

Les tailles retenues gardent les **proportions** du modèle — les intitulés plus gros que le
corps, Arial Black pour les uns et Century Schoolbook pour les autres — ramenées à ce qui
tient sur une page. La structure, les libellés et l'enchaînement des cartouches ne bougent
pas : c'est cela que le guichet reconnaît, pas le corps de la police. Tout est groupé dans
la région *Géométrie du formulaire* de `PieceExcelWU`, une constante par mesure.

Trois réglages d'impression font le reste :

1. **La zone d'impression** s'arrête à la dernière ligne écrite. Sans elle, Excel décide
   lui-même de ce qu'il imprime, et une cellule touchée par mégarde ajoute une page blanche.
2. **Une seule page en largeur, toujours.** Une pièce coupée verticalement oblige à raccorder
   les montants à leur libellé au scotch.
3. **Une seule page en hauteur pour une pièce de point de vente.** La pièce globale, elle,
   s'étale — à trois cents écritures, la forcer sur une page la réduirait à un timbre-poste.
   Son en-tête se répète en haut de chaque page : sans cela, la page 2 arrive sans date, sans
   agence et sans nom de colonne, une colonne de chiffres dont on ne sait plus ce qu'ils sont.

#### Filtrer : sous-agents seulement, ou agences propres seulement

La fenêtre d'aperçu porte une liste **« Pièces individuelles »** à trois choix, chacun suivi
du nombre d'onglets qu'il produira :

| Choix | Ce qu'on obtient |
|---|---|
| Tous les points de vente | La pièce globale, puis tous les onglets |
| Sous-agents seulement | La pièce globale, puis les onglets des `SA` |
| Agences propres seulement | La pièce globale, puis les onglets des `EC` |

Le nom de fichier proposé suit : `PieceWU_20260530_SA.xlsx` à côté de `PieceWU_20260530.xlsx`.
Un dossier d'exports reste ainsi lisible sans ouvrir les classeurs.

**Le filtre ne touche QUE les onglets individuels. La première feuille reste la pièce
globale**, entière. Ce n'est pas un oubli :

1. C'est **le** document comptable de la journée. La scinder en deux donnerait deux pièces
   dont aucune ne serait celle que la comptabilité attend.
2. L'écart d'arrondi de la journée est absorbé par le compte inter bancaire, **sur la pièce
   globale et sur elle seule** — la règle de la banque interdit d'appliquer le compte
   d'attente point de vente par point de vente. Une pièce globale filtrée sortirait donc
   déséquilibrée de quelques francs, sans rien pour les porter.

La liste ne s'affiche pas quand la fenêtre présente la pièce d'un seul point de vente : un
filtre y serait un choix entre une chose et elle-même.

**Les pièces individuelles sont regénérées**, et non découpées dans la pièce globale : c'est
`GenererPieceComptable` qui produit les deux, sur un `CalculWU` unique pour la seconde. Les
mêmes écritures, aux mêmes comptes, par construction. Un Account non comptabilisé n'a pas
d'onglet — il n'a pas d'écriture non plus, et un onglet vide laisserait croire à une pièce
à zéro.

### Préparer l'accès du compte (`PreparationBaseWU`)

L'agent colle la chaîne de la banque et voudrait que tout suive. Le bouton **« Préparer la
base… »** de l'écran de connexion fait ce qu'il est possible de faire, et dit le reste.

Il procède en trois temps, et aucun n'est supposé — tout est lu sur le serveur :

| | Temps | Ce qui se passe |
|---|---|---|
| 1 | **Constater** | Le compte entre-t-il ? La base existe-t-elle ? Les tables ? Les trois rôles ? Le compte est-il membre de l'un d'eux ? |
| 2 | **Agir** | Si le compte a les droits — `CREATE USER`, `ALTER ROLE`, après confirmation explicite |
| 3 | **Rédiger** | Sinon, le T-SQL exact, noms réels substitués, avec un bouton **Copier** |

L'interrogation se fait sur `master`, jamais sur la base visée : une base qui refuse le
compte refuse aussi la connexion, et l'on ne saurait alors rien dire du tout.

**Le temps 3 n'est pas un pis-aller, c'est le cas normal.** Créer un accès à SQL Server
demande des droits d'administration du serveur, que le compte applicatif n'a presque jamais
— et c'est heureux : une application comptable qui pourrait se donner des droits à
elle-même n'aurait plus de contrôle d'accès du tout. Il y a là un cercle que rien ne rompt :
pour créer l'accès, il faut déjà l'avoir. Mieux vaut donc un script juste et prêt à
exécuter qu'une tentative qui échoue en laissant l'agent deviner.

Le même script existe en fichier, pour qui préfère l'envoyer sans ouvrir l'application :
`Scripts\12_AccesCompteApplicatif.sql`, une seule ligne à modifier.

Et pour une informatique qui préfère l'interface au T-SQL, la même chose clic par clic :
`Installation\PROCEDURE-SSMS-Acces.md`. Tout y tient dans une seule fenêtre de SSMS — la page
*User Mapping* des propriétés du login crée l'utilisateur de base **et** lui donne son rôle.

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

### La TVA est perçue à 19,25 %, mais seuls 18 points vont au compte de TVA

Le taux ci-dessus est bien celui que Western Union perçoit : `Tax1REC` vaut `ChargeEnvoi ×
19,25 %` sur 95 lignes sur 95. Ce taux n'a pas changé et ne doit pas changer.

Ce que la pièce manuelle de la banque a révélé, sur l'agence propre `AHB020013` (Ecobank AGP
Siège, semaine du 08 au 14/09/2026), c'est que cette TVA perçue **ne va pas en totalité sur le
compte de TVA**. La banque n'y porte que **18 points** ; le **1,25 point** restant n'est pas de
la TVA au sens comptable, c'est une taxe additionnelle, et il suit le sort du solde de taxes.

Sur cette semaine, les frais d'envoi valaient 522 000 F — valeur recoupée par deux lignes que la
banque ne conteste pas : la commission sur envoi (107 010 = 20,5 %) et la TTA sur envoi
(15 788 = 0,2 % de 7 894 000 de nominal envoyé).

| Destination des 100 485 F de TVA perçue (19,25 % de 522 000) | Compte | Montant |
|---|---|---:|
| TVA portée au compte de TVA — 18 % des frais | `434000104` | 93 960 |
| Part restée dans le solde de taxes, ventilée à 75 % | `434000147` | 4 894 |
| Part restée dans le solde de taxes, ventilée à 25 % | `728300148` | 1 631 |
| **Total** | | **100 485** |

**Rien n'est perdu ni ajouté** : la somme des trois comptes vaut exactement 19,25 % des frais
d'envoi, et le total de la pièce — donc la contrepartie sur le compte courant Western Union —
est rigoureusement inchangé. Seule la répartition entre trois comptes change.

C'est ce qui explique les deux seules différences que la banque nous avait signalées, et
pourquoi elle ne signalait pas la ligne qui en était la cause :

| Ligne | Avant | Après | Pièce de la banque |
|---|---:|---:|---:|
| `434000104` TVA | 100 485 | **93 960** | 93 960 |
| `728300148` Commission sur transfert | 7 719 | **9 350** | 9 350 |
| `434000147` Impôts et taxes sur envoi | 23 157 | **28 050** | 28 050 |
| Total commissions et taxes | 443 540 | 443 540 | 443 540 |

Le mécanisme tient en une phrase : le solde de taxes est un **résidu**
(`Taxes − TVA − TTAEnvoi`). Tout franc qui n'est pas retranché au titre de la TVA y demeure, et
s'y répartit 25 % / 75 % comme le reste. C'est pourquoi la formule du solde n'a pas eu à être
touchée : seule la valeur retranchée a changé.

Dans le code, `TAUX_TVA` reste à `0.1925D` et demeure la source ; la part portée au compte de
TVA en est **dérivée** (`TAUX_TVA_COMPTE_TVA = TAUX_TVA - TAUX_TVA_HORS_COMPTE_TVA`), de sorte
qu'une révision future du taux de TVA se propage d'elle-même sans qu'aucune valeur en dur ne
puisse diverger.

**Effet sur les autres points de vente.** La correction est une fonction des seuls frais
d'envoi, identique pour une agence propre et pour un sous-agent : la ligne de TVA baisse de
1,25 % des frais, la commission sur transfert monte de 0,3125 % et les impôts et taxes sur envoi
de 0,9375 %. Le total de chaque pièce est inchangé.

**Point resté ouvert avec la banque.** Les 1 631 F ventilés à 25 % atterrissent sur
`728300148 « Commission sur transfert »`, qui est un compte de produit et non un compte de
taxe. C'est bien ce que fait la pièce manuelle de la banque, que l'application reproduit
désormais à l'identique ; il vaut néanmoins la peine de le faire confirmer par leur
comptabilité.

## Barre de progression des exports

Un export Excel ne dit rien pendant qu'il travaille. Sur une journée chargée, écrire la pièce
comptable d'une trentaine de points de vente prend une bonne dizaine de secondes pendant
lesquelles la fenêtre ne se repeint plus : l'utilisateur ne sait pas si l'application calcule ou
si elle est bloquée, et il reclique. **Tous les processus d'exportation affichent désormais une
fenêtre de progression flottante** qui nomme l'étape en cours.

### Le contrat : `Services\ProgressionWU.vb`

Le compteur d'avancement ne connaît pas l'interface graphique. Il expose deux évènements et
quatre méthodes, rien d'autre :

| Membre | Rôle |
|---|---|
| `Commencer(total)` | Annonce le nombre d'étapes prévues. `total <= 0` = avancement inconnu |
| `Etape(libelle)` | Change le texte affiché sans faire avancer la barre |
| `Avancer(libelle)` | `Etape` puis `Avancer` — le cas courant |
| `Avancer()` | Avance d'un cran, sans jamais dépasser `Total` |
| `Terminer()` | Porte la barre à son maximum |
| `EtapeChangee` / `AvancementChange` | Les deux évènements auxquels une fenêtre s'abonne |

Les services d'export (`PieceExcelWU`, `PieceComptableService`, `ExcelExportService`) reçoivent
ce compteur en paramètre **optionnel** :

```vb
Public Shared Sub Ecrire(..., Optional progression As ProgressionWU = Nothing)
```

`Nothing` est un cas normal, pas une erreur : appelé sans progression, un service exporte
exactement comme avant. Aucun service n'a de dépendance vers `System.Windows.Forms` du fait de
cette fonctionnalité, et les tests d'export restent possibles sans écran.

### La fenêtre : `Forms\FrmProgression.vb`

`FrmProgression.Ouvrir(proprietaire, titre)` construit la fenêtre, s'abonne au compteur et
l'affiche **immédiatement**. Pas de temporisation « n'afficher qu'au bout d'une seconde » : le
seul cas qui justifie la fenêtre est précisément celui d'une étape longue et unique, que
l'affichage différé manquerait.

Elle n'a ni bouton, ni `ControlBox`, ni menu système. **Il n'y a pas de bouton Annuler**, et
c'est un choix technique et non un oubli : Excel Interop est un composant **STA**, qui ne
supporte pas d'être piloté depuis un thread d'arrière-plan. L'export s'exécute donc sur le thread
de l'interface, et le seul point où la fenêtre peut se repeindre est l'intérieur de la boucle
d'export. Un bouton Annuler ne recevrait jamais son clic. Le repeint se fait par
`Control.Refresh()` et **non** par `Application.DoEvents()`, qui rouvrirait la file de messages
et permettrait à l'utilisateur de relancer l'export pendant qu'il tourne.

`Fermer()` est idempotent : elle est appelée en première instruction de chaque `Catch` — pour que
la fenêtre disparaisse **avant** le `MessageBox` d'erreur, et non derrière lui — puis de nouveau
dans le `Finally`.

### Ce que chaque export annonce

| Opération | Étapes fixes | Étapes variables |
|---|---|---|
| Export de la pièce comptable | 4 | 1 par point de vente détaillé |
| Export Excel par blocs | 3 | 1 par bloc écrit |
| Export PDF | 3 | 1 par feuille |
| Fichier core banking | 3 | — |
| Lecture des rapports Western Union | 8 | — |

Le budget d'étapes est vérifié : si une opération promet huit étapes et n'en annonce que six, la
barre s'arrête aux trois quarts et ment à l'utilisateur. Un script de simulation rejoue les
annonces de chaque opération, helpers compris, et contrôle que chaque barre atteint bien son
total.

La lecture des rapports Western Union (`FrmCompensationWU`) suivait jusqu'ici des pourcentages
écrits en dur dans la barre de la barre d'état, entrecoupés d'appels à `Application.DoEvents()`.
Les deux ont disparu : la barre d'état s'abonne maintenant à `AvancementChange`, donc à la même
source que la fenêtre flottante, et les deux affichages ne peuvent plus diverger.

## Annuler une comptabilisation

Une journée peut avoir été comptabilisée à tort : rapport Western Union vide, rapport d'une
autre journée, chargement fait deux fois. Il faut pouvoir la retirer des rapports.

### Corriger et retirer ne sont pas la même chose

**Corriger une journée ne demande rien de nouveau.** Recomptabiliser la même date écrase
intégralement la précédente : `HistoriqueRepository.EnregistrerJournee` commence par un
`DELETE` sur `T_HistoriqueWU`, `T_HistoriqueMTCN` et `T_PieceWU`, puis réinsère — le tout dans
une seule transaction. Jamais de doublon, jamais de cumul, jamais une journée à moitié
remplacée. Rechargez les bons rapports, relancez, c'est fini.

**Retirer une journée, en revanche, n'avait pas de réponse** : pour l'effacer, il aurait fallu
en comptabiliser une autre à sa place, ce qui est impossible quand la bonne réponse est « il ne
s'est rien passé ce jour-là ».

### Annuler n'est pas supprimer

Les lignes ne sont pas effacées, elles sont **déplacées** vers `T_HistoriqueAnnuleWU`,
`T_HistoriqueMTCNAnnuleWU` et `T_PieceAnnuleeWU`, sous un identifiant d'annulation qui porte le
motif, le commentaire et les deux signatures (`T_AnnulationWU`). Un `DELETE` effacerait la
preuve au moment précis où l'on en a besoin : c'est quand une journée est annulée que
l'auditeur veut savoir qui, quand et pourquoi.

**Pourquoi des tables séparées, et non une colonne `Annulee`.** `T_HistoriqueWU` a pour clé
primaire `(DateActivite, Account)` et `T_PieceWU` `(DateActivite, Ligne)`. Une journée annulée
qui resterait sur place entrerait en collision avec la journée recomptabilisée qui la remplace,
et il faudrait ajouter l'annulation à chaque clé primaire — donc à chaque lecture, donc aux
rapports d'activité, qui n'ont aucune raison de connaître ce nouveau statut. Avec des tables
d'archive, les tables vivantes ne contiennent que des journées en vigueur : **les rapports
cessent de compter une journée annulée sans qu'on touche à une seule de leurs requêtes.**

Une même date peut être annulée plusieurs fois (comptabilisée, annulée, recomptabilisée,
annulée de nouveau). C'est `IdAnnulation`, et non la date, qui identifie une archive.

### Deux personnes, comme pour le référentiel

L'annulation emprunte la file du double regard (`T_DemandeWU`), avec
`TypeObjet = 'COMPTABILISATION'` et `Operation = 'ANNULATION'`. L'authorizer garde un seul
écran ; la contrainte `CK_T_DemandeWU_PasSoiMeme` s'applique telle quelle ; et l'index unique
sur `(TypeObjet, Cle)` interdit deux demandes simultanées sur la même journée.

Retirer une journée de la comptabilité est plus grave que modifier un taux de sous-agent — et
le taux exige déjà deux personnes.

**Conséquence sur les droits SQL.** Le rôle `wu_compense` n'avait aucun accès à `T_DemandeWU`,
pas même en lecture, parce que la file ne portait que du paramétrage. Elle porte maintenant
aussi les annulations, qui appartiennent à la compense : `wu_compense` reçoit donc
`SELECT, INSERT, UPDATE`. Ce qui protège le référentiel n'est pas l'absence de ce droit, mais
la fonction INPUTER / AUTHORIZER et la contrainte de la base, qu'un `UPDATE` fait à la main ne
contourne pas davantage.

L'application applique le retrait **dans la transaction de la décision** : en-tête d'annulation,
trois déplacements et décision sur la demande ne font qu'un. Une journée à moitié archivée
serait comptée deux fois par les rapports, ou pas du tout.

### Ce que l'annulation ne fait pas : extourner

Si le fichier destiné au core banking a été injecté, les écritures sont dans les livres de la
banque. **Wincompense ne peut pas les en retirer.** L'écran de demande pose la question, la
réponse est conservée (`CoreBankingInjecte`), et elle est rappelée à l'authorizer au moment
d'autoriser. L'extourne se demande en comptabilité, séparément.

Pour que l'avertissement ne repose pas sur la seule mémoire de l'agent, **la production du
fichier est désormais tracée** (`T_FichierCoreBankingWU` : journée, date de valeur, numéro de
lot, nom du fichier, totaux, auteur, horodatage). Le fichier lui-même n'est toujours pas
conservé — il dérive entièrement de la pièce, le garder serait garder deux fois la même chose.

La trace ne dit pas que le fichier a été **injecté** : l'application écrit un classeur, elle ne
voit pas ce que le core banking en fait. Et « aucune production connue » n'est pas « jamais
produit » : les fichiers sortis avant la mise en service de la table ne sont pas enregistrés.
Les écrans le disent dans ces termes.

### Le parcours à l'écran

| Étape | Où | Qui |
|---|---|---|
| Sélectionner la journée, bouton « Annuler cette comptabilisation… » | Pièces comptables conservées | inputer |
| Voir ce qui sera retiré, répondre sur le core banking, choisir le motif | `FrmAnnulerComptabilisation` | inputer |
| Relire le contenu ACTUEL de la journée, autoriser ou rejeter | Demandes en attente | authorizer |

Le motif est obligatoire, pris dans une liste courte (rapport vide, rapport erroné, rapport
d'une autre journée, double comptabilisation, autre). « Autre » exige une explication écrite :
il n'explique rien par lui-même.

Le détail affiché à l'authorizer **relit la journée au moment de la décision**, et non au moment
du dépôt : entre les deux, elle a pu être recomptabilisée. Ce sont les chiffres affichés là qui
partiront en archive. Si la journée n'a plus rien de comptabilisé, l'écran le dit et
l'autorisation est refusée — une archive vide ferait croire à un traitement.

### La journée annulée reste visible

Elle figure dans la liste des pièces conservées, grisée et barrée, avec la colonne **État** qui
porte « ANNULÉE » et le motif. Sa pièce s'ouvre et s'exporte comme une autre : un justificatif
se relit, il n'engage rien.

**Le fichier core banking, lui, ne se reconstruit pas depuis une journée annulée** : ce fichier
passe des écritures, et repasser celles d'une journée retirée serait exactement l'inverse de ce
qu'on a voulu. Le bouton s'éteint.

### Avertissement avant de remplacer une journée

Le remplacement d'une journée déjà comptabilisée était silencieux. Il ne l'est plus : avant de
générer, `FrmCompensationWU` annonce les trois choses qui peuvent le rendre fâcheux —

- la journée a déjà été comptabilisée, par qui et à quelle heure ;
- son fichier core banking est peut-être déjà parti, auquel cas la nouvelle version ferait
  double emploi avec des écritures déjà passées ;
- une annulation de cette journée attend d'être autorisée, et l'authorizer retirerait alors la
  **nouvelle** version en croyant retirer l'ancienne.

Rien de tout cela n'interdit de continuer — corriger une journée est la façon normale de
réparer une comptabilisation fausse. L'écran informe, il ne décide pas, mais il propose « Non »
par défaut.

### Scripts à exécuter

| Script | Ce qu'il crée |
|---|---|
| `Scripts\14_AnnulationComptabilisation.sql` | Les quatre tables d'annulation, les deux colonnes ajoutées à `T_DemandeWU`, les contraintes rouvertes, la vue `V_JourneesAnnulees` |
| `Scripts\15_FichierCoreBanking.sql` | `T_FichierCoreBankingWU` |

`00_InstallationComplete.sql` les contient déjà pour une installation neuve. Sur une base
existante, **le script 14 est obligatoire avant toute nouvelle demande** : il ajoute deux
colonnes à `T_DemandeWU`, et tant qu'elles manquent, plus aucune demande — annulation comme
référentiel — ne peut être déposée ni décidée. Les écrans le disent et nomment le script.

**Aucun rôle ne reçoit `DELETE` sur les archives ni sur la trace des fichiers.** Une archive que
ses propres utilisateurs peuvent effacer ne prouve rien.

## Paramétrage : fichier de secours

*Menu Paramétrage › Paramétrage : fichier de secours… — administrateur seul.*

### Ce que ce fichier n'est pas

**Une sauvegarde de la base.** L'application ne peut pas en faire une, et ce n'est pas un choix
de conception mais quatre obstacles :

- `BACKUP DATABASE` n'est pas un droit de base de données mais un rôle serveur. Le compte
  applicatif est membre de `wu_admin` **dans la base** et n'a rien au niveau du serveur — lui
  donner le droit de sauvegarder et de restaurer, ce serait lui donner le droit d'écraser.
- Le fichier de sauvegarde atterrit sur le disque du **serveur**, sous le compte de service SQL
  Server, et non sur le poste de l'agent.
- Une sauvegarde qui ne se déclenche que si quelqu'un ouvre l'application n'est pas une
  sauvegarde : un lundi férié, une semaine de congé, et il n'y a rien.
- On ne restaure pas une base à laquelle on est connecté, et restaurer efface tout ce qui a été
  fait depuis. Ce n'est pas un bouton dans une application comptable.

**La vraie sauvegarde se fait sur le serveur, par l'équipe qui le tient** : complète chaque
nuit après la compensation, journal toutes les 15 à 30 minutes si la base est en mode
`FULL`, `RESTORE VERIFYONLY` après chaque passage, copie hors du serveur, et une restauration
d'essai par trimestre. Une sauvegarde qu'on n'a jamais restaurée n'est pas une sauvegarde.

### Ce qu'il est

De quoi remonter une installation **neuve** sans resaisir à la main des centaines de
sous-agents : les comptes comptables, les groupes statistiques, les sous-agents et les agences
propres. C'est peu, et c'est exactement ce qu'une application de poste peut promettre sans
mentir.

### Le format : une archive de fichiers texte

`Parametrage-GWC_WINCOMPENSE_ETD-20260920-1830.zip`, qui contient :

| Fichier | Contenu |
|---|---|
| `manifeste.csv` | version du format, base, serveur, date, auteur, empreinte du contenu |
| `LISEZ-MOI.txt` | ce que le fichier est, ce qu'il n'est pas, comment le recharger |
| `comptes-systeme.csv` | les neuf comptes, une ligne chacun, avec leur signification |
| `groupes-statistiques.csv` | groupe, compte d'activité, compte de commission, taux |
| `sous-agents.csv` | les sept colonnes de `T_Pdv_SA` |
| `agences-propres.csv` | les trois colonnes de `T_Pdv_EC` |
| `utilisateurs-pour-information.csv` | identifiant, nom, rôle, fonction, actif |

**Du CSV, et non un classeur.** Le fichier doit pouvoir être relu sur le poste d'une
installation qui commence, donc sans qu'on puisse parier sur la présence d'Excel — Interop est
exclu ici comme il l'est de toute la logique métier. Un CSV en point-virgule et en UTF-8 **avec
BOM** s'ouvre néanmoins d'un double-clic dans Excel, en français, colonnes séparées : on garde
la lisibilité sans la dépendance.

Deux précautions valent d'être notées. Les valeurs qui contiennent un point-virgule, un
guillemet ou un saut de ligne sont protégées, et le découpage à la lecture est écrit caractère
par caractère — un `Split` couperait en deux une désignation contenant un point-virgule, et
personne ne s'en apercevrait avant que le sous-agent correspondant soit mal créé. Les nombres,
eux, sont écrits en **culture invariante** : un taux écrit `0,70` sur un poste français et relu
`0.70` sur un autre vaudrait soixante-dix fois trop.

**L'empreinte SHA-256** des quatre fichiers de données figure au manifeste. Elle ne protège de
rien — qui modifie un fichier peut recalculer son empreinte — mais elle répond à une question :
*ce fichier est-il celui qui est sorti de l'application ?* Retoucher un taux avant de recharger
est parfois la seule façon de s'en sortir ; l'ignorer ne l'est jamais. L'écran affiche
« CONTENU MODIFIÉ depuis l'export » en rouge, et le redit dans la confirmation.

### Les utilisateurs ne se rechargent pas

Leur liste est exportée — identifiant, nom, rôle, fonction — pour qu'on sache **qui** recréer.
Aucun mot de passe n'en sort, **pas même sous forme d'empreinte** : un fichier qui circule ne
porte pas les identifiants d'une banque. Les recharger supposerait soit d'emporter ces
empreintes, soit d'inventer des mots de passe provisoires et de les écrire dans le même
fichier. Ni l'un ni l'autre. L'administrateur de la nouvelle installation est créé au premier
démarrage, comme aujourd'hui.

### Le chargement ne remplace jamais

**C'est la règle qui tient tout le reste.** Une clé déjà présente est laissée telle quelle et
comptée « déjà présente ». Le chargement ne peut donc pas détruire un référentiel : au pire, il
ne fait rien.

C'est aussi ce qui le dispense du double regard. Le contrôle à deux personnes protège les
**modifications** d'un paramétrage en service ; ici, là où quelque chose existe, l'import
s'abstient — il n'y a rien à protéger. Toute modification ultérieure repasse par la file des
demandes, comme aujourd'hui.

**L'import n'a pas son propre SQL** : il écrit par `PdvRepository.AppliquerSousAgent`,
`AppliquerAgence` et `AppliquerGroupe`, c'est-à-dire par le chemin qu'emprunte déjà une demande
autorisée. Ce qui est vrai de l'un est vrai de l'autre.

**La seule exception : les comptes comptables.** `SystemeWU` n'est pas une liste d'objets mais
une ligne de réglages, que le script d'installation crée d'emblée avec des valeurs par défaut.
S'en tenir à « ne jamais remplacer » reviendrait à ne jamais charger les comptes de la banque —
exactement ce qu'on est venu chercher. Ils sont donc remplaçables, mais jamais en silence :
l'écran montre l'**ancienne et la nouvelle valeur ligne à ligne**, en rouge là où elles
diffèrent, et l'administrateur doit cocher. Le code de la ligne locale est conservé : celui du
fichier désigne la ligne de la base d'origine, et viser une ligne inexistante ne toucherait
rien.

### Ce que l'analyse refuse, et pourquoi elle le fait avant d'écrire

| Cas | Décision |
|---|---|
| Taux hors de `[0 ; 1]` — le classique `70` pour `0,70` | refusé |
| Désignation ou compte obligatoire vide | refusé |
| Compte d'activité ou de commission déjà porté par un autre groupe | refusé |
| Account déjà enregistré de l'autre côté (sous-agent ↔ agence propre) | refusé |
| Sous-agent sans groupe, ou dont le groupe n'existe pas | créé, avec une remarque |

Les deux index uniques de `T_GroupeStatistique` et la règle « un Account n'est pas des deux
côtés » sont vérifiés **avant** la transaction. Les heurter en plein chargement ferait échouer
l'ensemble, là où les annoncer permet de corriger le fichier et de recommencer.

L'analyse est **refaite au moment de cliquer sur Charger** : entre son affichage et ce clic,
quelqu'un d'autre a pu créer un sous-agent.

### L'ordre des écritures n'est pas indifférent

Les groupes d'abord — les sous-agents s'y réfèrent —, puis les agences, puis les sous-agents, le
tout dans **une transaction**. Un chargement à moitié fait laisserait des sous-agents rattachés
à des groupes absents, et personne ne saurait dire où il s'est arrêté.

Les comptes comptables sont écrits **en dernier, hors de cette transaction**. Si le référentiel
échoue, les comptes n'auront pas bougé ; si les comptes échouent après, le référentiel est en
place et les neuf comptes se retapent sur un écran, là où trois cents sous-agents ne se retapent
pas.

### Un garde-fou de plus

Si le fichier vient d'une **autre base** que celle à laquelle on est connecté, l'écran le dit et
demande confirmation. C'est normal pour une installation neuve ; ce l'est beaucoup moins si les
deux bases appartiennent à des entités différentes — le référentiel chargé serait plausible et
faux, et rien ne le signalerait ensuite.

## Commissions encaissées par la banque

*Menu Compensation › Commissions encaissées par la banque.*

### Pourquoi cet écran n'existait pas

L'onglet « Évolution des commissions » des rapports d'activité montre la commission
**générée par l'activité** — celle que Western Union verse. Sur un sous-agent à 70 %, les sept
dixièmes affichés ne sont pas à la banque : ils lui sont rétrocédés. **Aucun écran ne disait
donc ce que la banque gagne.**

### Deux sources, deux règles

| Population | Ce que la banque garde |
|---|---|
| Sous-agents | la part non rétrocédée, soit `commission × (1 − taux)` |
| Agences propres | la totalité — leur taux vaut zéro |

La somme des deux est ce qu'elle a réellement encaissé sur la période. L'état les présente
côte à côte, par nature de commission (envoi, paiement, transfert) puis jour par jour, avec la
part rétrocédée en regard : **part banque + part sous-agents = commission totale**, vérifiable
d'un coup d'œil et recoupable avec l'onglet existant.

### Rien n'est recalculé

La part de la banque est celle qui a été **calculée le jour même**, conservée avec l'historique
(`CommissionEnvoiBanque`, `CommissionPaiementBanque`, `CommissionTransfertBanque`), ainsi que le
taux appliqué ce jour-là (`TauxSA`).

La recalculer avec les taux d'aujourd'hui ferait varier **rétroactivement** ce que la banque a
gagné le mois dernier : un sous-agent passé de 70 % à 60 % changerait le passé. C'est
exactement ce que l'application refuse déjà de faire pour les pièces comptables, et pour la
même raison.

### Les comptes concernés

Les trois comptes de commission de la banque (`Cpte_Produit`, `Cpte_Produit_Envoi`,
`Cpte_Produit_Paiement`) **reçoivent les deux populations indifféremment**. La séparation
demandée est donc **analytique** : elle se lit dans l'état, pas sur le relevé de compte. C'est
le choix retenu par la banque ; faire apparaître la séparation en comptabilité supposerait des
comptes distincts pour les agences et modifierait la pièce.

L'état affiche ces comptes avec ce que les **pièces conservées** y ont réellement porté sur la
période : nombre d'écritures, débit, crédit, net crédité. Les comptes sont **dédoublonnés** —
par défaut, le transfert et l'envoi partagent le même compte, et les compter deux fois
doublerait le total.

### Le contrôle croisé

Le total de l'état (issu de l'historique) est confronté au total porté sur ces comptes (issu
des pièces). Les deux doivent coïncider ; l'écran le dit en vert quand c'est le cas, en rouge
sinon, avec les deux explications possibles : les comptes de commission ont changé pendant la
période, ou une journée n'a pas de pièce conservée. La tolérance est d'un franc par ligne
d'historique — l'historique porte deux décimales, la pièce est arrondie à l'unité.

Les journées **annulées** n'entrent dans aucun des deux totaux : leur historique et leur pièce
sont partis en archive, ce qui est précisément ce qu'on attend d'une journée retirée.

### Ce que l'état ne sait pas, il le dit

Les journées comptabilisées **avant** l'exécution de `Scripts\16_CommissionsBanque.sql` n'ont
pas de répartition conservée. Pour elles :

- la part des **agences propres reste exacte** — leur taux vaut zéro par construction, et le
  type du point de vente est conservé depuis toujours ;
- la part des **sous-agents est inconnue**, et non nulle.

L'écran affiche alors, en rouge, combien de journées sont concernées, de quand à quand, et à
partir de quelle date la répartition est disponible. Les lignes correspondantes sont en rouge
dans le tableau jour par jour, avec une colonne « Répartition » qui vaut NON. **Un zéro et une
donnée absente ne sont pas la même chose**, et une case vide qui se lit comme un zéro est la
pire façon de se tromper. L'avertissement suit l'état jusque dans le PDF exporté : un
chiffre incomplet exporté sans sa réserve deviendrait un chiffre tout court.

### L'export se fait en PDF

Et non en classeur, à la différence des rapports d'activité : cet état se signe, se classe et
se transmet. Un tableur invite à retoucher les chiffres, et un chiffre retouché dans le fichier
qu'on présente n'est plus celui de la banque.

Le PDF reprend les trois tableaux — répartition, comptes crédités, jour par jour — précédés de
la période, de la mention que rien n'est recalculé, de l'avertissement sur les journées non
documentées s'il y en a, du contrôle croisé, et de la date d'édition avec son auteur. Cette
dernière n'est pas une politesse : deux éditions d'une même période peuvent différer si une
journée a été annulée entre-temps.

Mise en page en paysage, ajustée à une page de large — le tableau jour par jour porte sept
colonnes. La conversion passe par Excel, comme tous les PDF de l'application.

### Une fenêtre à part

L'état n'est pas un onglet des rapports d'activité : ceux-ci sont filtrés par population — une
fenêtre pour les sous-agents, une autre pour les agences propres — et un total des deux
n'aurait sa place dans ni l'une ni l'autre.

### Compatibilité avec une base non mise à jour

Les quatre colonnes sont récentes. Trois traitements les nomment, et les trois vérifient
d'abord qu'elles existent (`WURepository.ColonneExiste`) :

- la **lecture** de l'historique — sans quoi tous les rapports échoueraient pour une
  information qui n'en est qu'une parmi d'autres ;
- l'**écriture** de la comptabilisation quotidienne — c'est l'opération du jour, elle ne
  s'arrête pas pour une colonne manquante ;
- le **déplacement** d'une journée annulée vers l'archive — une journée resterait en place
  faute d'une colonne qu'elle n'a jamais eue.

`Scripts\16_CommissionsBanque.sql` ajoute les colonnes à `T_HistoriqueWU` **et** à
`T_HistoriqueAnnuleWU`, et crée la vue d'audit `V_CommissionsBanque`.

## Le bordereau de fin de journée, et son visa

### Ce que la pièce comptable ne prouve pas

La pièce porte déjà les quatre cartouches de la banque — *Initié, Contrôlé, Approuvé*, puis
*Écriture passée et autorisée*. Les **écritures** sont donc couvertes par une signature.

Mais qui signe la pièce signe **ce qui y figure**. Cinq choses n'y figurent pas :

- **ce qui n'a PAS été comptabilisé** : les Accounts non paramétrés sont écartés, et
  n'apparaissent donc nulle part sur la pièce, par construction. Un contrôleur qui ne voit que
  la pièce ne peut pas savoir qu'une agence a travaillé ce jour-là sans que ses opérations
  soient passées ;
- **quels rapports Western Union ont servi** — si une journée est contestée dans six mois, rien
  ne prouvait que c'est bien le rapport de ce jour-là qui avait été traité ;
- **l'écart d'arrondi**, noyé au milieu des lignes de la pièce ;
- **le fichier destiné au core banking**, qui est pourtant le point de non-retour ;
- **que la journée est complète** : volumes, points de vente, concordance.

Le bordereau porte ces cinq choses. Il atteste de la **façon dont la journée a été faite**, là
où la pièce atteste de ce qui a été écrit.

### Ses huit blocs

| Bloc | Contenu |
|---|---|
| 1. Identification | journée, date de valeur, numéro de lot, qui a comptabilisé et quand, état du visa |
| 2. Rapports traités | nom des deux fichiers Western Union et leur **empreinte SHA-256** |
| 3. Ce qui a été traité | points de vente, dont sous-agents et agences propres, **Accounts écartés**, envois, paiements, annulations |
| 4. Montants | principal envoyé et payé, charges, taxes, commission totale, **dont part de la banque** et part rétrocédée |
| 5. La pièce | total débit, total crédit, **équilibre**, et **l'écart d'arrondi isolé avec son compte** |
| 6. **Accounts non comptabilisés** | Account, désignation, volumes, principal et commission générée — ou « AUCUN », écrit en toutes lettres |
| 7. Core banking | fichier produit, par qui, quand, sous quel lot |
| 8. Visa et signatures | trois cartouches hauts : *Établi par*, *Vérifié par*, *Approuvé par* |

Le **bloc 6 est celui qui justifie la signature**. Les sept autres décrivent ce qui a été fait ;
celui-là décrit ce qui ne l'a pas été, et c'est la seule information qu'un supérieur ne peut
obtenir ailleurs. Quand il est vide, il le dit — une case vide ne se lit pas, « AUCUN » se lit.

**L'empreinte des rapports** ne protège de rien : qui remplace un fichier peut recalculer la
sienne. Elle répond à une question — *le fichier que vous me montrez est-il celui qui a été
traité ce jour-là ?* Sans elle, la question n'a pas de réponse. Elle est calculée au moment de
la comptabilisation, pendant que les fichiers sont encore sous la main.

### Le visa, et pourquoi il existe

Modifier un taux de sous-agent exige **deux personnes**. Comptabiliser une journée entière n'en
exigeait qu'**une**. C'était un déséquilibre curieux : le référentiel était mieux gardé que
l'écriture qu'il produit.

Le visa corrige cela. Un utilisateur ayant la fonction **authorizer** — la même que pour le
référentiel et les annulations — relit la journée et la marque visée. **Celui qui l'a
comptabilisée ne peut pas la viser**, et la contrainte `CK_T_TraitementWU_PasSoiMeme` le refuse
aussi bien qu'un `UPDATE` fait à la main dans Management Studio.

La confirmation répète les chiffres au lieu de demander « êtes-vous sûr ? » : le nombre de
points de vente, les deux totaux, l'équilibre, et **le nombre d'Accounts non comptabilisés**
s'il y en a. On est toujours sûr ; on ne relit pas toujours.

**Recomptabiliser une journée efface son visa.** Le visa atteste d'un traitement précis ;
refaire la journée en produit un autre, et laisser le visa en place ferait croire qu'un
supérieur a vu des chiffres qu'il n'a jamais vus. Il faut viser de nouveau.

### Quand il se produit

**Automatiquement**, à la fin de la comptabilisation : il s'ouvre après la pièce — l'agent
regarde d'abord ce qu'il a produit, puis ce qu'il doit faire signer. L'inverse ferait signer
avant d'avoir vu.

**Et à la demande**, depuis l'écran des pièces conservées, bouton « Bordereau de la journée… ».
C'est là que le chef de service vient relire une journée et la viser. Le bouton reste actif sur
une journée **annulée** : comprendre pourquoi elle a été retirée suppose de pouvoir relire
comment elle avait été traitée.

### Il se reconstitue

Tout vient de la base : l'en-tête de traitement (`T_TraitementWU`, écrit dans la **même
transaction** que l'historique et la pièce), l'historique, la pièce conservée et la trace du
fichier core banking. Rien n'est recalculé avec le paramétrage d'aujourd'hui.

Une journée annulée emporte son en-tête en archive (`T_TraitementAnnuleWU`), **visa compris** :
c'est ce que l'on relira pour comprendre qui avait vu quoi.

### Pour les journées antérieures

Celles comptabilisées avant l'exécution de `Scripts\17_BordereauJournee.sql` n'ont pas
d'en-tête. Leur bordereau s'affiche quand même : volumes, montants et totaux de la pièce se
relisent. Mais le **nom des rapports Western Union** n'était conservé nulle part — l'écran écrit
« NON CONSERVÉ pour cette journée » plutôt que de laisser des cases vides qui se liraient comme
une absence de rapport. Et ces journées **ne peuvent pas être visées** : on ne vise pas un
traitement dont on n'a pas la trace.

### Compatibilité avec une base non mise à jour

L'écriture de l'en-tête et son déplacement en archive vérifient tous deux que les tables
existent. **La comptabilisation du jour ne s'arrête pas** parce qu'un script d'évolution n'a pas
été joué : c'est l'opération quotidienne, elle passe quand même, sans bordereau reproductible.
De même, une annulation ne laisse pas une journée en place faute d'une table qu'elle n'a jamais
eue.

### Le visa bloque-t-il le fichier core banking ? La banque choisit

*Menu Paramétrage › Options de traitement… — administrateur seul.*

Le visa peut **constater** ou **empêcher**. Imposer l'un ou l'autre serait décider à la place
de la banque : un contrôle bloquant arrête la compense le matin où le chef de service est
absent, et un contrôle qui n'arrête rien ne contrôle pas grand-chose.

| Option | Ce qui se passe au moment de produire le fichier |
|---|---|
| **NON** *(par défaut)* | L'application **avertit** que la journée n'est pas visée — et que des Accounts n'ont pas été comptabilisés, s'il y en a — puis laisse l'agent décider. « Non » est proposé par défaut. |
| **OUI** | L'application **refuse**, et dit qui doit viser et où. |

L'écran ne se contente pas de deux boutons radio : **sous chaque choix, une phrase décrit ce
qui se passera**. Une case à cocher dont on ne voit pas la conséquence se coche au hasard. Et
activer l'option bloquante demande une confirmation qui rappelle de prévoir **plusieurs
authorizers** — avec un seul, une journée d'absence bloque la compense.

**L'option vit en base** (`T_ParametreWU`), et non sur le poste : elle décrit la façon de
travailler de la banque, pas la configuration d'un ordinateur. Un agent ne doit pas pouvoir
s'en affranchir en décochant une case chez lui. Seul l'administrateur l'écrit ; tout le monde
la lit, parce qu'elle gouverne un contrôle que l'agent subit.

**Une absence n'est jamais une interdiction.** Table absente, option jamais créée, base
injoignable : la valeur par défaut s'applique, et c'est toujours la moins bloquante. Une règle
qui s'activerait toute seule parce qu'une lecture a échoué arrêterait la compense pour une
raison que personne ne comprendrait.

**La règle est écrite une seule fois**, dans `Forms\VisaWU.vb`. Deux écrans produisent ce
fichier — celui de la compense, pour la journée qu'on vient de traiter, et celui des pièces
conservées, pour une journée ancienne — et deux copies d'une règle finissent toujours par
diverger.

Enfin, une journée **antérieure au bordereau** n'est jamais bloquée : elle n'a pas d'en-tête de
traitement, donc pas de visa possible, et refuser sa production punirait l'agent pour une
table qui n'existait pas.

### L'édition mensuelle des commissions

L'état des commissions encaissées par la banque gagne deux commandes :

- **« Mois complet »** cale la période sur le mois entier de la date de fin. Un état signé qui
  couvrirait vingt-trois jours parce que l'agent a mal cliqué serait un état faux, et rien ne le
  dirait ;
- **« Édition à signer (cartouches) »** ajoute au PDF les trois cartouches de signature et la
  mention « ÉDITION DESTINÉE À LA SIGNATURE » en exergue. Le fichier s'appelle alors
  `Commissions-banque-a-signer-202605.pdf`.

## Les icônes, la barre d'outils et l'icône de l'application

### Ce qui manquait

Les menus n'affichaient que du texte. Sur un écran de vingt entrées, l'agent lit chaque ligne
avant de trouver la sienne ; avec une image, il vise. L'application n'avait pas non plus
d'icône propre : le bureau, la barre des tâches et l'explorateur montraient l'icône blanche
par défaut de Windows, celle que portent les programmes qu'on n'a pas fini d'écrire.

### Les icônes sont dessinées, non importées

`Forms/IconesWU.vb` dessine les dix-neuf icônes au trait, avec GDI+. Il n'y a donc **aucun
fichier image à transporter** : ni dossier de PNG à côté de l'exécutable — il se perd au
premier déploiement — ni blocs binaires dans les `.resx`, qu'on ne peut ni relire ni corriger.
Le dessin est du texte, il se lit et se modifie comme le reste du code.

L'avantage décisif est ailleurs : le dessin est calculé **à la taille demandée**. Un poste en
125 % ou 150 % — courant sous Windows 10 comme sous Windows 11 — réclame des images de 20 ou
24 pixels ; une image de 16 pixels agrandie par Windows devient floue, celle-ci reste nette
parce qu'elle est retracée. Chaque dessin est composé dans un carré de 16 unités, et la mise
à l'échelle est posée une fois sur le contexte graphique.

Les couleurs sont relevées sur l'icône fournie par la banque : l'or `#FFD54C`, le vert
`#79BB00`, le bleu `#3C91BD`, avec un contour ardoise `#37474F` — et non noir, qui écrase la
couleur qu'il entoure à 16 pixels. Les menus appartiennent ainsi visiblement à la même famille
que l'icône du bureau.

Le grisé des entrées désactivées n'est pas produit ici : Windows Forms grise lui-même l'image
d'un élément désactivé. En fabriquer une seconde version doublerait le cache pour un résultat
que le système fait déjà.

### Toutes les entrées, ou aucune

Un menu où trois entrées sur dix portent une image est plus laid qu'un menu sans images : les
sept autres s'alignent sur une colonne vide et paraissent inachevées. **Toute entrée de menu
déroulant reçoit donc son icône.**

Les titres de la barre de menu — Compensation, Paramétrage, Sécurité, Fenêtres — n'en
reçoivent pas : ils ouvrent un menu, ils ne déclenchent rien, et aucune application Windows
n'y met d'image. « Quitter » fait exception : c'est une commande posée dans la barre, pas un
titre, et elle porte la sienne.

| Entrée | Icône |
|---|---|
| Traitement de la compense | balance à deux plateaux |
| Rapport d'activité — sous-agents | barres d'un graphique |
| Rapport d'activité — agences propres | immeuble |
| Pièces comptables conservées | boîte d'archives |
| Commissions encaissées par la banque | deux pièces de monnaie |
| Sous-agents | une silhouette |
| Agences propres | immeuble |
| Groupes statistiques | trois pastilles reliées |
| Autorisations du référentiel | coche dans une pastille |
| Comptes systèmes | registre à tranche bleue |
| Options de traitement | deux curseurs de réglage |
| Paramétrage : fichier de secours | chemise |
| Mon mot de passe | clé |
| Utilisateurs et connexions | deux silhouettes |
| Connexion à la base de données | cylindre à bandes |
| Cascade, mosaïques, fermer tout | fenêtres disposées, fenêtre barrée |
| Quitter | porte et flèche |

### La barre d'outils

Une barre d'outils reprend les gestes du quotidien, ceux qu'on répète chaque matin :

> Traitement du jour │ Rapport sous-agents · Rapport agences · Pièces conservées ·
> Commissions │ Autorisations

Les boutons ne décident rien : chacun reprend, un par un, la disponibilité de l'entrée de menu
qu'il double, et partage son gestionnaire — un bouton visible que le menu masque serait une
porte dérobée, et celle-là se verrait, puisqu'elle est en haut de l'écran. La barre entière
disparaît quand aucun bouton ne reste : une barre vide prendrait de la place sur la zone de
travail sans rien offrir.

Le bouton des autorisations porte une **pastille** avec le nombre de demandes en attente. Le
bouton n'affiche pas de texte : sans elle, il ne dirait rien de ce qui attend. Au-delà de 99,
la pastille affiche « 99+ », et le compte exact se lit dans le menu.

### L'icône de l'application

`Ressources/Wincompense.ico` porte les cubes isométriques fournis par la banque, en 16, 24,
32, 48 et 64 pixels. Elle est déclarée deux fois dans le projet, et les deux sont nécessaires :

- `<ApplicationIcon>` la grave dans `Wincompense.exe` — c'est elle que montrent le bureau, la
  barre des tâches et l'explorateur ;
- `<EmbeddedResource>` la rend lisible à l'exécution, pour que les fenêtres la portent — avec
  toutes ses tailles, si bien que la barre de titre et la barre des tâches prennent chacune la
  sienne, sans agrandissement.

#### Chaque fenêtre se pose l'icône elle-même

La première version la posait en **un seul endroit**, `AfficherEnfant` — l'ouverture d'une
fenêtre fille MDI. C'était trop peu : les boîtes de dialogue ne passent pas par là. Elles sont
ouvertes par `ShowDialog`, depuis quatorze endroits différents, et n'avaient donc aucune icône.

La règle est maintenant : **`IconesWU.Habiller(Me)` suit chaque `InitializeComponent()`**, dans
le constructeur de chaque formulaire. C'est une ligne de plus dans vingt-trois fichiers, et
c'était le prix à payer : Windows Forms n'offre aucun événement « une fenêtre vient de
naître », et `Application.Idle` — le seul point central possible — n'est pas garanti pendant
une boucle modale, c'est-à-dire précisément pendant l'affichage des dialogues qu'il s'agissait
de rattraper.

Une classe de base commune aurait fait la même chose en un endroit. Elle a été écartée pour une
raison d'atelier, pas de conception : **le Concepteur Windows Forms refuse d'ouvrir un
formulaire hérité tant que le projet ne compile pas**. Mettre les vingt-trois écrans derrière
cette condition, dans une application maintenue à travers l'interface de Visual Studio, coûtait
plus cher que la ligne répétée.

`verif_icone_fenetres.py` tient la règle : tout `InitializeComponent()` qui n'est pas suivi de
`IconesWU.Habiller(Me)` est signalé.

#### FixedDialog masque l'icône — c'est Windows, pas l'application

Huit fenêtres avaient l'icône et ne l'affichaient pas. `FormBorderStyle.FixedDialog` pose
`WS_EX_DLGMODALFRAME` sur la fenêtre — **le même drapeau que `ShowIcon = False`** — et Windows
ne dessine alors aucune icône dans la barre de titre.

Elles sont passées en `FixedSingle` : même comportement (fenêtre non redimensionnable), aspect
identique sous Windows 10 et Windows 11 où le cadre de dialogue ne se distingue plus, et
l'icône s'affiche. Si la banque préfère le cadre de dialogue, c'est une ligne à remettre dans
chaque `.Designer.vb`.

`FrmProgression` garde `FixedDialog` : elle n'a pas de `ControlBox`, donc pas de barre système,
donc pas d'icône à afficher — et une fenêtre d'avancement ne doit pas offrir de bouton de
fermeture.

#### Un défaut trouvé en chemin : `FrmDiagnostic` n'avait pas de constructeur

`InitializeComponent` n'y était jamais appelé. La fenêtre n'avait donc aucun contrôle, et la
première ligne d'`Afficher` — `fenetre.lblTitre.Text = …` — levait une
`NullReferenceException`. L'écran tombait exactement quand on en a besoin : pour lire le
diagnostic d'un accès SQL refusé.

VB ne rattrape pas cet oubli. Il n'ajoute l'appel implicite qu'aux classes portant l'attribut
`DesignerGenerated`, qu'aucun `.Designer.vb` de ce projet ne porte. `verif_icone_fenetres.py`
le vérifie désormais pour les vingt-quatre fenêtres.

#### Le piège du nom de ressource, et pourquoi il y a deux sources

La première version n'a porté **aucune** icône, et n'a rien signalé. La cause : le nom d'une
ressource incorporée **ne se compose pas de la même façon en VB.NET et en C#**.

| | Ce que MSBuild écrit pour `Ressources\Wincompense.ico` |
|---|---|
| C# | `WincompenseTCHAD.Ressources.Wincompense.ico` |
| **VB.NET** | `WincompenseTCHAD.Wincompense.ico` — **le dossier est ignoré** |

Le nom écrit en dur était celui de C#. `GetManifestResourceStream` a donc rendu `Nothing` —
**sans lever d'exception** : une ressource introuvable n'est pas une erreur, elle est
simplement absente. L'icône manquait partout, et rien ne le disait.

Trois parades, posées ensemble parce qu'aucune ne suffit seule :

1. `<LogicalName>Wincompense.ico</LogicalName>` dans le `.vbproj` **fixe** le nom au lieu de
   laisser MSBuild le composer ;
2. `IconesWU` cherche par la **fin** du nom, ce qui retrouve les trois écritures possibles —
   celle de VB, celle de C#, et le `LogicalName` — sans dépendre d'aucune ;
3. si la ressource manque quand même, l'icône est relue **dans l'exécutable lui-même**, où
   `ApplicationIcon` l'a gravée (`Icon.ExtractAssociatedIcon`). Cette lecture ne rend qu'une
   taille, que Windows redimensionne, mais elle ne dépend ni du nom de la ressource ni de la
   façon dont MSBuild l'a composé.

Les deux sources n'échouent pas pour les mêmes raisons : la première tient au nom de la
ressource, la seconde au fichier `.exe`. C'est précisément pour cela qu'il y en a deux.

`verif_pieges.py` porte désormais la règle : un nom de ressource incorporée écrit en dur
**avec un dossier** est signalé.

### Une réserve sur l'image fournie

L'image d'origine est une **capture d'écran de 30 × 40 pixels**, bordée d'un liseré gris et
posée sur une ombre portée. Le liseré et l'ombre ont été retirés — le fond est détouré depuis
les bords sur ce qui est *clair et sans couleur*, ce qui épargne les reflets blancs du dessin
— et le dessin utile, 28 × 30 pixels, a été recadré puis complété en carré.

**Les tailles au-delà de 32 pixels sont donc agrandies, et restent un peu douces.** Elles
conviennent au bureau et à la barre des tâches ; une image carrée nette de 256 × 256 pixels sur
fond transparent donnerait un résultat franchement meilleur en grande vignette. Le jour où la
banque la fournit, il suffit de remplacer le `.ico` : aucune ligne de code ne change.

### Ce qui a été vérifié

- `simul_icones.py` recalcule l'encombrement de **chaque forme tracée** et refuse tout trait
  qui sortirait de la grille de 16 unités — un débordement ne se verrait autrement qu'à
  l'écran, une fois livré ; il contrôle aussi que chaque valeur de l'énumération a son
  aiguillage et son tracé, que chaque entrée de menu déroulant reçoit son icône, et que chaque
  bouton de la barre a son image, son gestionnaire, son infobulle et son contrôle de droits ;
- les dix-neuf dessins ont été rendus en image, à 16, 24 et 72 pixels, et **regardés**.

## La TTA sur réception entre dans le versement du point de vente

### Ce qui n'allait pas

La pièce créditait le compte `434000159` du montant de la TTA sur réception — **sans l'avoir
encaissée nulle part**. La contrepartie tombait donc sur le compte courant Western Union, qui
se trouvait financer une taxe tchadienne qu'il ne doit pas.

Ce n'était pas un choix de politique, c'était un oubli. Il valait **19 860 F sur une seule
semaine et un seul sous-agent** — l'écart exact constaté face à la pièce manuelle de la banque
sur AHB020211, semaine du 08 au 14/09/2026.

### L'argument qui tranche : l'asymétrie est dans les données

| | À l'envoi | À la réception |
|---|---|---|
| Qui prélève la TTA ? | **Western Union** | **personne** |
| Où la trouve-t-on ? | dans `TaxesREC`, colonne `Tax3REC` | nulle part |
| Est-elle dans la caisse du point de vente ? | **oui**, et déjà dans le terme « taxes » | **non** |

Vérifié sur les données : `Tax3REC` cumule 41 444 sur la semaine, et 0,2 % du principal envoyé
vaut 41 443,376. À la réception, la plateforme n'affiche une taxe que trois fois sur cinquante-sept
paiements, et c'est le simple report de la taxe d'envoi.

**Ce que Western Union encaisse à l'envoi, la banque doit l'encaisser à la réception.** D'où le
terme qui s'ajoute :

```
NetMouvement = (PrincipalEnvoi + ChargeEnvoi + Taxes) − PrincipalPaye + TTAReception
```

### Une seule écriture de la formule

Elle vit dans **`CalculWU.NetMouvement`**, et nulle part ailleurs. `GenererPieceComptable` et
`CalculerEcartArrondi` la recopiaient chacun de leur côté ; ils l'appellent désormais. Deux
copies d'une formule finissent toujours par diverger, et celle-ci porte le montant que le point
de vente doit réellement verser.

Le compte courant Western Union suit **tout seul** : il est le solde
(`NetMouvement − toutes commissions et taxes`), donc il passe de 11 667 235 à 11 687 095 sans
qu'on y touche. Le fichier core banking, l'historique et le bordereau lisent la même valeur.

### L'agence propre ne change pas

**Règle confirmée par la banque : une agence propre ne retient pas de TTA sur paiement.** Sa
`TTAReception` vaut zéro — posé une seule fois, dans `AppliquerFormules` — et le nouveau terme
est alors sans effet. Le versement d'une agence propre reste inchangé.

### Le passé ne bouge pas, et rien n'a été prévu pour ça

Aucune date d'effet, aucun paramètre : **les journées déjà comptabilisées conservent leurs
montants par construction.** Une pièce archivée est relue telle quelle dans `T_PieceWU`
(`Compte, Libelle, Debit, Credit`), jamais recalculée ; les rapports et le bordereau cumulent
les colonnes stockées de `T_HistoriqueWU`, qui portent déjà `TTAReception`. `GenererPieceComptable`
n'est appelée que sur la journée en cours.

La seule façon de recalculer une journée ancienne est de **l'annuler puis de la recomptabiliser**,
ce qui est précisément ce que l'on veut dire quand on la refait : la nouvelle règle s'applique,
et l'archive de l'annulation conserve l'ancienne.

### Ce qui reste à dire aux sous-agents

« Prélevée sur le nominal payé » se lit de deux façons, et **la comptabilité est la même dans
les deux cas**, mais pas la caisse :

- le sous-agent **retient** la taxe au guichet : le bénéficiaire reçoit le nominal diminué de
  0,2 %, la caisse garde la taxe, et le versement correspond à l'espèce détenue ;
- le sous-agent **verse le nominal entier** : sa caisse est courte de la taxe, qu'il retrouve
  sur sa commission.

La banque doit dire laquelle des deux s'applique, sinon les caisses ne tomberont jamais juste
au guichet.

### Ce qui a été vérifié

`simul_tta_reception.py` rejoue la chaîne complète — `AppliquerFormules`, `RepartirCommissions`,
les lignes de `GenererPieceComptable`, l'arrondi au franc ligne par ligne — sur quatre situations :

- **sous-agent à 60 %** : la pièce tombe **exactement** sur celle saisie par la banque, douze
  lignes sur douze, débit 12 398 369 et compte courant 11 687 095 ;
- **sous-agent à 70 %** : la pièce tombe **exactement** sur les colonnes de la feuille de la
  banque (6 616 / 75 153 / 22 805 / 15 437 / 53 212 / 175 357) ;
- **agence propre** : aucune ligne de TTA sur réception, débit inchangé à 12 378 509 ;
- **l'écart** entre l'ancienne et la nouvelle règle vaut la TTA sur réception, au débit comme
  sur le compte courant : 19 860 des deux côtés.

> Un franc d'écart d'arrondi subsiste au taux de 60 % : c'est le résidu normal, absorbé par le
> compte d'attente (`VerifierEquilibrePiece`). La banque, elle, l'a absorbé dans son compte
> courant en saisissant 30 408 là où 40 % de 76 017,06 donne 30 407.

## Cinq corrections de forme sur la pièce

Elles viennent toutes du rapprochement ligne à ligne avec la pièce manuelle de la banque
(AHB020211, semaine du 08 au 14/09/2026). Aucune ne touche un montant ; toutes touchent ce que
le vérificateur voit.

### 1. La pièce dit la période qu'elle couvre

La date d'activité était déduite de la **première ligne lisible** du fichier, et rien ne
vérifiait les suivantes. Un rapport hebdomadaire — c'est ainsi que la banque liquide, *« DU 08
AU 14/09 »* — était donc agrégé en entier puis étiqueté d'une seule journée : celle de sa
première ligne. Les montants étaient justes, l'intitulé faux, et **rien ne le disait**.

- `WUReportService.JourneesDuRapport` inventorie désormais **toutes** les journées du rapport ;
- une seule fonction lit une date (`EssayerLireDate`), pour que l'inventaire et la lecture d'une
  date ne puissent pas comprendre deux choses différentes du même fichier ;
- l'écran **annonce** le périmètre et le fait confirmer, **avec « Non » par défaut**, en nommant
  la journée sous laquelle tout sera enregistré et en avertissant du double comptage si l'on
  charge ensuite une journée déjà comprise dans la période ;
- la pièce s'intitule alors **« activité du 08/09/2026 au 14/09/2026 »**.

Le traitement n'est **pas** refusé : la banque liquide par semaine, et lui interdire de charger
sa semaine reviendrait à lui interdire de travailler. La conséquence est dite, l'agent tranche.

> **Ce qui reste à décider.** L'historisation se fait toujours sous **une** date — la première
> journée du rapport. Tant que la banque charge une semaine par semaine, cela fonctionne ; le
> jour où elle voudra une clé de période, c'est le modèle de `T_HistoriqueWU` qu'il faudra
> ouvrir, pas un libellé.

### 2. Les numéros de compte sont du texte

`32100005296` était écrit comme un **nombre**. La cellule s'alignait à droite comme un montant,
devenait sommable, un compte commençant par zéro aurait perdu son zéro — et surtout, dès que la
colonne était rétrécie, Excel affichait **`3,21E+10`**. C'est exactement ce que la banque a vu en
collant notre pièce à côté de la sienne.

Le format `"@"` est posé sur la colonne des comptes **avant** l'écriture des valeurs. Le poser
après ne servirait à rien : la conversion a déjà eu lieu.

### 3. La ligne de total

La pièce de la banque porte le total des crédits ; la nôtre ne le portait pas. C'est lui qui
permet de conclure d'un coup d'œil, sans additionner douze lignes. C'est une **formule**
(`=SUM(...)`) et non une valeur figée : elle se recalcule si quelqu'un corrige une ligne dans le
classeur, au lieu d'afficher un total qui ne correspondrait plus à rien.

### 4. Un seul ordre de lignes

Le bloc Ecobank allait *transfert / envoi / paiement*, le bloc sous-agent *transfert / paiement /
envoi*. Deux ordres différents **dans la même pièce**, et un troisième chez la banque. Les trois
blocs suivent désormais **transfert, paiement, envoi** — celui de la banque. Dans une comparaison
ligne à ligne, un ordre différent fait perdre du temps au vérificateur et lui fait passer des
écarts.

### 5. Le libellé n'est plus doublé

Le gabarit préfixe la désignation par `CCS_`, comme le classeur de référence. Mais la plupart des
désignations commencent **déjà** par « CCS », et la pièce portait
*« CCS_CCS NGARTA RUE DE 40M ACTIVITE WU »*. Le préfixe n'est désormais posé que lorsqu'il
manque : `CCS NGARTA RUE DE 40M ACTIVITE WU`, tandis que `BOLOLO` garde son
`CCS_BOLOLO ACTIVITE WU` d'origine.

### Ce qui a été vérifié

`verif_piece_forme.py` tient les cinq règles — dix-neuf contrôles sur le code : la fonction
d'inventaire des journées, l'unicité de la lecture de date, la confirmation par défaut sur
« Non », le libellé de période, le format texte posé **avant** la valeur, la ligne de total en
formule et dans le quadrillage, l'ordre des trois lignes dans les deux blocs, et le préfixe
conditionnel — avec les deux cas limites, `CCS NGARTA` et `BOLOLO`.

## Le taux de rétrocession est borné des deux côtés

### Où il vit, et ce que la base en garde

```
T_Pdv_SA.Taux  →  CalculWU.TauxSA  →  RepartirCommissions  →  T_HistoriqueWU.TauxSA
  (le réglage)      (la journée)        (le partage)            (la trace, figée)
```

**Chaque journée comptabilisée conserve le taux qui lui a été appliqué** : changer le réglage
ne réécrit rien. Trois requêtes répondent à « quel taux, depuis quand, décidé par qui » :

```sql
-- ce que l'application applique aujourd'hui
SELECT Code_Pdv, Designationagence, Taux, DateModification, ModifiePar
FROM   T_Pdv_SA WHERE Code_Pdv = 'AHB020211';

-- qui l'a changé, quand, et sur quel double regard
SELECT IdDemande, Operation, Statut, Taux, SaisiPar, DateSaisie, DecidePar, DateDecision
FROM   T_DemandeWU WHERE TypeObjet = 'SOUS_AGENT' AND Cle = 'AHB020211' ORDER BY DateSaisie DESC;

-- la décisive : le taux réellement appliqué, journée par journée
SELECT DateActivite, TauxSA FROM T_HistoriqueWU
WHERE  Account = 'AHB020211' ORDER BY DateActivite DESC;
```

La colonne `TauxSA` de l'historique n'existe que depuis `Scripts\16_CommissionsBanque.sql` :
les journées antérieures à cette migration ne la portent pas.

> **La base dit ce qui est appliqué, pas ce qui est juste.** `T_Pdv_SA.Taux` est une saisie,
> pas une autorité. La seule autorité est le contrat qui lie la banque à son sous-agent.

### Pourquoi une borne, et pourquoi en base

Le taux est une **fraction** : `0,70` vaut 70 %. Saisir « 70 » multiplierait par cent toutes les
commissions rétrocédées — le sous-agent recevrait 7000 %, la part de la banque deviendrait
massivement négative — et **la pièce s'équilibrerait quand même**, puisque le compte courant
Western Union est calculé par différence et absorbe n'importe quoi.

**L'application le refuse déjà** : `PointDeVenteSA.Anomalies` et `GroupeStatistiqueWU.Anomalies`
bornent le taux à `[0 ; 1]`, contrôlent les deux décimales de `DECIMAL(4,2)`, et bloquent
l'enregistrement — sur l'écran des sous-agents, sur celui des groupes, et au chargement d'un
fichier de paramétrage.

**Trois portes restaient ouvertes**, et la contrainte les ferme d'un coup :

1. un `UPDATE` direct, par l'informatique ou une autre application ;
2. une ligne posée à la main dans `T_DemandeWU` — `AppliquerSousAgent` écrit le taux de la
   demande dans `T_Pdv_SA` **sans le revalider** : le contrôle a eu lieu chez le demandeur,
   pas chez celui qui autorise ;
3. une reprise de données, une restauration, un script de migration.

`Scripts\19_BornerLeTaux.sql` pose `CK_T_Pdv_SA_Taux`, `CK_T_GroupeStatistique_Taux` et
`CK_T_DemandeWU_Taux`. Elles valent **quelle que soit la version de l'application** installée
sur les postes.

### Le script ne corrige jamais une donnée

S'il trouve des valeurs hors bornes, il les **nomme** et s'arrête sans rien poser : corriger un
taux est une décision métier, pas un effet de bord de script. Il ne contient ni `UPDATE` ni
`DELETE`. Relancez-le après correction — il est rejouable.

Deux nuances assumées : `NULL` est admis sur `T_DemandeWU` (une suppression, ou une demande
portant sur une agence propre, n'a pas de taux), et la contrainte y est posée `WITH NOCHECK`
— les demandes déjà décidées sont des archives, et l'histoire ne se réécrit pas.

### Ce qui a été vérifié

`verif_taux.py` contrôle que **les deux côtés disent la même borne** : si l'une changeait sans
l'autre, la première saisie limite passerait d'un côté et serait rejetée de l'autre, avec un
message que personne ne comprendrait. Il vérifie aussi que le script ne corrige aucune donnée,
qu'il est rejouable, et que l'autorisation d'une demande ne revalide effectivement pas le taux
— ce qui est la raison d'être de la contrainte.

## Règles tranchées par la banque

- **La TTA sur réception est supportée par le sous-agent**, et s'ajoute donc à son versement.
  La plateforme ne la prélève pas — contrairement à la TTA sur envoi, qui est déjà dans
  `TaxesREC` —, la banque la recouvre et la reverse au Trésor. **Une agence propre ne retient
  pas de TTA sur paiement** : sa TTA sur réception reste nulle, et son versement est inchangé.

- **Agence propre (EC).** La structure de sa pièce est **identique à celle d'un sous-agent** :
  même enchaînement d'écritures, mêmes libellés, mêmes sens. Les trois lignes de commission
  sous-agent n'y figurent pas parce que le taux vaut zéro et qu'une écriture à zéro n'est jamais
  posée — ce n'est pas un traitement à part, c'est la même règle appliquée à 0 %, toute la
  commission revenant à la banque. Sa ligne de mouvement va sur le **compte courant WU** : une
  agence propre n'a pas de compte de compensation dans les livres de la banque.
- **Accounts non paramétrés.** Ils ne sont **pas comptabilisés** — voir plus haut.

- **Chemin d'installation.** La sécurité de la banque autorise un **fichier à un emplacement**,
  pas une application. Le chemin autorisé est
  `C:\Program Files\Default Company Name\SetupWincompense\Wincompense.exe` — le Program Files
  **natif**, sans `(x86)`. Les deux noms sans signification viennent de l'ancien projet
  d'installation Visual Studio, et c'est pour cela qu'il ne faut pas les rendre présentables :
  les changer obligerait la banque à refaire son autorisation. `Installation\Wincompense.iss`
  le reproduit avec `{autopf}` **et** `ArchitecturesInstallIn64BitMode=x64compatible` — les deux
  lignes tiennent ensemble, `{autopf}` seul donnant `Program Files (x86)` sur un poste 64 bits.

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
- Comptes distincts pour les commissions des agences propres : la banque a retenu la
  séparation ANALYTIQUE, les trois comptes de commission restant communs aux deux
  populations. Les colonnes `Cpte_Envoi_agence` et `Cpte_Paiement_agence` de `SystemeWU`
  restent donc inutilisées. Si la séparation doit un jour apparaître en comptabilité, c'est
  la pièce qu'il faudra modifier, et la banque devra fournir ces comptes.
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
- Annulation d'une comptabilisation : les fonctions INPUTER / AUTHORIZER du référentiel
  servent telles quelles. Si la banque veut que le retrait d'une journée soit décidé par
  d'autres personnes que le paramétrage des points de vente, il faudra une seconde paire de
  fonctions — une colonne de plus sur `T_UtilisateurWU`, et rien d'autre à changer.
- Faut-il purger les archives d'annulation au bout d'un certain temps, et lequel ? Aucune
  purge n'est prévue aujourd'hui, et aucun rôle n'a le droit d'effacer.
- Le visa d'une journée utilise la fonction AUTHORIZER, la même que le référentiel et les
  annulations. Si la banque veut que le chef de service de la compense soit distinct de celui
  qui autorise le paramétrage, il faudra une seconde paire de fonctions — une colonne de plus
  sur `T_UtilisateurWU`, et rien d'autre à changer.
- Sauvegarde de la base : qui tient le serveur SQL, à quelle périodicité les sauvegardes
  tournent-elles, et où sont-elles recopiées ? Le fichier de secours du paramétrage ne
  remplace rien de cela — il le complète.
- Faut-il afficher dans l'application la date de la dernière sauvegarde réussie, lue dans
  `msdb` ? C'est le jour où le travail de sauvegarde s'arrête en silence que cela sert.
  Proposé, non retenu pour l'instant.
- Le fichier de secours doit-il aussi porter les jours fériés (`T_JourFerieWU`) ? Ils sont
  regénérés pour 2026-2030 par `Scripts\10_JoursFeries.sql` sur une installation neuve,
  d'où leur absence ; les corrections apportées par la banque, elles, ne s'exportent pas.

## Test de référence (section 17)

Les scripts `Scripts\02_DonneesExemple.sql` paramètrent :

- `AHB020200` comme sous-agent (`T_Pdv_SA`, `Taux = 0.70`) ;
- `AHB020057` comme agence propre Ecobank (`T_Pdv_EC`).

Pour valider le parsing/l'agrégation, charger un rapport d'activité contenant `AHB020200` à la
date du 02/01/2021 et vérifier dans `dgvControle` que les agrégats retrouvent approximativement
les valeurs de référence indiquées dans le cahier des charges (Principal Envoi ≈ 5 966 388,
Principal Payé ≈ 7 522 117,45, Commission Envoi = 69 700, Commission Paiement ≈ 98 105, etc.).
Pour `AHB020057`, vérifier qu'aucune commission sous-agent n'est générée (100 % à la banque).
