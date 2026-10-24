---
titre: "Wincompense TCHAD — Manuel d'installation"
sous_titre: "Procédure pour les administrateurs, étape par étape"
version: "Application 1.1.0.0"
---

# Wincompense TCHAD — Manuel d'installation

**Procédure pour les administrateurs, étape par étape.**

| | |
|---|---|
| Application | Wincompense TCHAD — compensation Western Union, J+1 |
| Version couverte | 1.1.0.0 |
| Base de données | `GWC_WINCOMPENSE_ETD`, SQL Server |
| Destinataires | Informatique de la banque, administrateur Wincompense |

Ce manuel se lit **dans l'ordre**. Chaque étape se termine par une vérification à
**constater** — pas seulement à exécuter. Une étape dont la vérification n'a pas été vue
n'est pas faite : elle est supposée faite, ce qui n'est pas la même chose et se paie deux
semaines plus tard.

Les trois intervenants cités sont **INFO** (informatique de la banque), **COMPTA**
(comptabilité et compense) et **PROJET** (l'équipe qui a construit l'application).

> Ce manuel remplace, pour la conduite de l'installation, la lecture croisée de
> `LISEZMOI-Installation.md`, `PLAN-DEPLOIEMENT.md` et `PROCEDURE-SSMS-Acces.md`. Ces
> trois documents restent dans `Installation\` : le premier pour le détail technique du
> déploiement et du changement de serveur, le deuxième pour le suivi du projet phase par
> phase, le troisième pour être remis tel quel à l'informatique.

---

## 1. Avant de commencer

### 1.1 Ce qu'il faut qu'on vous ait donné

Sans ces cinq éléments, l'installation s'arrête au milieu. Les réunir avant de toucher à
un serveur fait gagner une journée.

| | Élément | Qui le fournit |
|---|---|---|
| 1 | Le nom exact de l'**instance SQL Server** de production — par exemple `SRV-SQL01` ou `SRV-SQL01\SQLEXPRESS` | INFO |
| 2 | Le **compte de connexion** de l'application : un compte SQL Server avec son mot de passe, **ou** les groupes Active Directory à autoriser | INFO |
| 3 | Le **chemin du partage réseau** qui portera la configuration — par exemple `\\SRV-FICHIERS\Wincompense` | INFO |
| 4 | Les **onze comptes comptables** du paramétrage, validés un par un | COMPTA |
| 5 | Le **référentiel des points de vente** : sous-agents, agences propres, groupes statistiques | COMPTA |

### 1.2 Les prérequis des postes

| Prérequis | Remarque |
|---|---|
| Windows 7 SP1 ou plus récent | C'est le plancher du .NET Framework 4.8 |
| .NET Framework 4.8 | Vérifié par l'installateur ; présent d'office depuis Windows 10 1903 |
| Microsoft Excel | Pour les exports et les éditions. L'application fonctionne sans, mais n'exporte plus |
| Accès réseau au serveur SQL | Port 1433, ou le port de l'instance nommée |

Aucun client SQL Server n'est à installer sur les postes : le pilote fait partie du
framework.

### 1.3 Les décisions qui bloquent la mise en production

| | Point | Qui tranche |
|---|---|---|
| ☐ | Le **compte comptable qui reçoit l'écart d'arrondi global** de la pièce | COMPTA |
| ☐ | Les **groupes Active Directory** à rattacher aux rôles `wu_compense`, `wu_commercial`, `wu_admin` | INFO |
| ☐ | Le **serveur SQL de production** et le **chemin du partage** | INFO |
| ☐ | Le format de `VALDT` dans le fichier core banking (`jj/mm/aaaa` retenu), à confirmer par un chargement de test | INFO |

> **L'écart d'arrondi est en cours d'arbitrage.** L'application affecte aujourd'hui le
> reliquat Débit − Crédit de la pièce au **compte inter bancaire** `381000101`, au débit ou
> au crédit selon son signe. La banque a demandé un **compte dédié** : tant qu'elle n'a pas
> donné son numéro, le comportement reste celui décrit ici, et il est correct.

---

## 2. Étape 1 — La base de données

**Un seul script à exécuter. Il est rejouable.**

### 2.1 Exécuter le script

1. Ouvrir **SQL Server Management Studio**, connecté au serveur de **production**.
2. Vérifier en haut de l'écran que c'est bien le bon serveur. Cette seconde de lecture est
   la seule protection contre une exécution sur la base de test — ou l'inverse.
3. Ouvrir `Scripts\00_InstallationComplete.sql`.
4. **À la première installation seulement**, vérifier une ligne, dans la **PARTIE 5** en bas
   du fichier :

   ```sql
   DECLARE @compteApplicatif SYSNAME = N'etdwincompense';
   ```

   C'est le nom du compte SQL Server avec lequel l'application se connecte. S'il porte un
   autre nom chez la banque, **remplacer celui-là, et c'est tout**.

5. Exécuter le script entier (**F5**).
6. Lire l'onglet **Messages**, puis le **compte rendu de la partie 6**.

### 2.2 Ce que le script installe

| | Objets |
|---|---|
| Base | `GWC_WINCOMPENSE_ETD`, créée si elle n'existe pas |
| Tables | **26**, avec leurs index |
| Vues | **6** |
| Rôles | **3** — `wu_admin`, `wu_compense`, `wu_commercial`, et leurs droits |
| Paramétrage de départ | les comptes comptables, les produits de transfert, les bornes du taux, les libellés de la pièce, les **dix clés** des options et du format core banking, les jours fériés à date fixe |

**Il est rejouable, et c'est sa raison d'être.** Chaque objet est protégé par un contrôle
d'existence : il crée ce qui manque, complète ce qui est incomplet, et ne touche à rien
d'autre. **Aucune donnée n'est écrasée** — ni un compte comptable saisi par la Direction
Comptable, ni un libellé, ni une option. Le relancer sur une base déjà installée est sans
risque, et c'est la façon normale de rattraper une base en retard.

> **À exécuter après chaque livraison de l'application.** Une livraison qui ajoute un écran
> ajoute souvent une colonne ou une clé de paramétrage : le script les pose, une par une, sur
> une base déjà en service. C'est la partie 4 *decies* qui s'en charge, et c'est elle qui
> évite le « le script est passé sans erreur » suivi d'un échec une semaine plus tard.

### 2.3 Ce que le script ne fait pas

- **Aucun sous-agent, aucune agence** : le référentiel se charge depuis l'application
  (étape 7).
- **Aucun compte utilisateur** : le premier administrateur se crée au premier démarrage
  (étape 6).
- **Aucun login SQL Server** : il faudrait un mot de passe, qui n'a pas sa place dans un
  fichier qui circule. Le login est créé par la banque (étape 2). S'il manque, la partie 5
  s'arrête proprement et **affiche la commande `CREATE LOGIN` à passer**.
- **Ni le mode de récupération, ni la sauvegarde.**

### 2.4 La sauvegarde, tout de suite

| | À faire | Qui |
|---|---|---|
| ☐ | Vérifier le **mode de récupération** de la base | INFO |
| ☐ | Si le mode est `FULL` : mettre en place une sauvegarde régulière du **journal de transactions** | INFO |
| ☐ | Mettre en place la **sauvegarde complète quotidienne** | INFO |

> Une base en mode `FULL` sans sauvegarde du journal finit par remplir le disque. Ce n'est
> pas un risque théorique : c'est un arrêt de production, un matin, sans prévenir.

### 2.5 Les vingt-quatre autres scripts

Le dossier `Scripts\Documentation\` porte **24 fichiers**. Ils **expliquent**, ils ne
s'exécutent pas : chacun ouvre sur la règle métier qu'il installe — pourquoi ce compte et
pas un autre, pourquoi cette contrainte, ce que l'application fait si la table manque. Tout
ce qu'ils créent est repris par `00`.

Deux exceptions, et deux seulement :

| Fichier | Quand l'exécuter |
|---|---|
| `Scripts\11_AccesUtilisateurs.sql` | **Authentification Windows uniquement** : il porte les comptes ou groupes Active Directory de la banque, que `00` ne peut pas deviner. À compléter, puis exécuter — **une fois**, à l'installation, jamais à une livraison |
| `Scripts\Documentation\12_AccesCompteApplicatif.sql` | Pour rattacher **plusieurs** comptes, un par agent, avec des rôles différents : une exécution par compte |

**Vérification à constater :** le compte rendu de la partie 6 affiche la liste des tables et
des rôles, la rubrique « Reste à faire » et la chaîne de connexion à employer.

---

## 3. Étape 2 — L'accès du compte applicatif

Deux cas, et un seul s'applique chez vous.

### 3.1 Si l'application se connecte par un compte SQL Server

C'est le cas chez Ecobank Tchad, avec le compte `etdwincompense`.

La **partie 5** du script `00` a déjà fait le travail, à une condition : que le **login**
existe au niveau du serveur. S'il n'existait pas, le script l'a dit et a affiché la commande
à passer. Trois ordres, dans cet ordre :

```sql
-- 1. le login, au niveau du serveur
USE [master];
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'etdwincompense')
    CREATE LOGIN [etdwincompense] WITH PASSWORD = N'<le mot de passe>',
                                       CHECK_POLICY = OFF;
GO
-- 2. l'utilisateur dans la base
USE [GWC_WINCOMPENSE_ETD];
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'etdwincompense')
    ALTER USER [etdwincompense] WITH LOGIN = [etdwincompense];
ELSE
    CREATE USER [etdwincompense] FOR LOGIN [etdwincompense];
GO
-- 3. le rôle de l'application
ALTER ROLE [wu_compense] ADD MEMBER [etdwincompense];
GO
```

> **Vérifier le mode d'authentification du serveur.** Une instance neuve n'accepte que
> Windows : le login existerait alors sans pouvoir servir. Clic droit sur le nom du serveur →
> *Propriétés* → *Sécurité* → **Mode d'authentification SQL Server et Windows**. Si la case
> doit être cochée, **le service SQL Server doit être redémarré** pour que le changement
> prenne effet.

Pour la même opération **sans script**, clic par clic dans SSMS, remettre à l'informatique
le document `Installation\PROCEDURE-SSMS-Acces.md` : deux minutes, dans une seule fenêtre.

### 3.2 Si les postes se connectent avec leur compte Windows

1. Ouvrir `Scripts\11_AccesUtilisateurs.sql`.
2. Compléter sa **section 1** avec les comptes ou groupes Active Directory réels.
3. L'exécuter.

Aucun mot de passe ne circule alors, et c'est précisément ce qui permet de poser la
configuration sur un partage lisible par tous.

**Vérification à constater :** depuis le poste, l'application ouvre la base et la barre
d'état affiche le serveur et la base en service.

---

## 4. Étape 3 — Le partage réseau

### 4.1 Pourquoi la connexion ne vit pas avec l'application

La banque change de serveur. Tant que la chaîne de connexion vivait dans
`Wincompense.exe.config`, en changer imposait une tournée dans les bureaux : ce fichier est
dans `Program Files`, donc protégé ; il est propre à chaque poste ; et une réinstallation
l'écrase.

Elle est donc cherchée ailleurs, dans un ordre où **le premier trouvé l'emporte** :

| Rang | Emplacement | À quoi il sert |
|---|---|---|
| 0 | Saisie pour la session en cours | Essai, sans rien conserver |
| 1 | Variable d'environnement `WINCOMPENSE_CONNEXION` | Dépannage, poste de test |
| **2** | **Fichier partagé** désigné à l'installation | **La source de vérité** |
| **3** | `%PROGRAMDATA%\Wincompense\wincompense.config` | Copie locale, rafraîchie à chaque lecture réussie du partage |
| 4 | `App.config` de l'application | Poste de développement |
| 5 | Valeur compilée `.\SQLEXPRESS` | Dernier recours |

Le rang 3 n'est pas un doublon : c'est lui qui fait travailler le poste le matin où le
partage est injoignable. Sans lui, une coupure réseau arrêterait la compense de toute la
banque.

### 4.2 Les quatre gestes

1. Créer un dossier sur le serveur de fichiers — par exemple `\\SRV-FICHIERS\Wincompense`.
2. Y copier `Installation\connexion.config.modele` sous le nom **`connexion.config`**.
3. Renseigner la ligne `SERVEUR=`.
4. Poser les droits :

| Qui | Droit |
|---|---|
| Utilisateurs de Wincompense | **Lecture** |
| Informatique, administrateur Wincompense | **Lecture et écriture** |

> **Qui peut écrire dans ce fichier commande la connexion de toute la banque.** Ce n'est pas
> une formalité de sécurité : c'est le point le plus sensible de toute l'installation.

### 4.3 Si la banque fournit une chaîne de connexion complète

Il arrive que la banque ne donne pas un nom de serveur mais une chaîne toute faite, avec des
mots-clés que `SERVEUR` et `BASE` ne savent pas exprimer : chiffrement imposé (`Encrypt`),
partenaire de secours (`Failover Partner`), groupe de disponibilité, port particulier.

Coller la chaîne après `CHAINE=`, **sur une seule ligne**, et mettre en commentaire les
lignes `SERVEUR` / `BASE` / `DELAI` :

```
# SERVEUR=...
# BASE=...
CHAINE=Server=SRV-SQL01;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;Encrypt=True;
```

`CHAINE` **prime** sur `SERVEUR`, `BASE` et `DELAI`. Renseigner les deux ferait coexister
deux descriptions du même serveur, dont une seule compte.

> **Aucun mot de passe dans le fichier partagé.** Il est lisible par tous les utilisateurs —
> c'est ce qui permet à un changement de serveur de valoir pour tout le monde. Une chaîne
> portant `User ID` / `Password` n'est pas refusée : l'application **retire le mot de passe
> avant toute écriture**, le chiffre par Windows pour cette machine, et n'écrit sur le partage
> que le serveur, la base et le **nom** du compte.
>
> **Conséquence :** le serveur se propage, le mot de passe non. Chaque poste doit le recevoir
> une fois (étape 5).

**Vérification à constater :** `connexion.config` existe sur le partage, un utilisateur
ordinaire le lit, et il ne contient aucun mot de passe.

---

## 5. Étape 4 — Construire le programme d'installation

À faire **une fois par version**, sur le poste de développement. *(PROJET)*

### 5.1 Quelle version d'Inno Setup

**Inno Setup 7 n'est plus gratuit pour un usage commercial.** Une banque qui s'en sert pour
outiller son activité est dans ce cadre. Deux issues, l'une et l'autre valables :

| Choix | Ce que cela implique |
|---|---|
| **Inno Setup 6.4.x** *(recommandé)* | Gratuit pour tout usage, commercial compris. Le script fonctionne dès la 6.3 |
| **Inno Setup 7 avec licence** | Acheter la licence auprès de jrsoftware. Rien à changer au script |

Le script ne dépend d'aucune nouveauté de la version 7.

### 5.2 Les quatre étapes

1. **Compiler en Release.** C'est l'oubli le plus fréquent, et il arrête tout. Dans Visual
   Studio, la liste déroulante affiche *Debug* par défaut : choisir **Release**, puis
   *Générer* → *Générer la solution*. Vérifier que
   `WincompenseTCHAD\WincompenseTCHAD\bin\Release\Wincompense.exe` existe et porte la date du
   jour.
2. **Installer Inno Setup**, selon le choix ci-dessus.
3. **Double-cliquer `Installation\Construire-Setup.cmd`.** Il cherche `ISCC.exe`, vérifie que
   le Release existe, compile, et ouvre le dossier du résultat. En cas de manque, il dit
   lequel et ce qu'il faut faire.
4. Le programme d'installation apparaît dans `Installation\Sortie\`, sous le nom
   `Wincompense_Setup_1.1.0.0.exe`.

### 5.3 Le numéro de version ne se recopie pas

Le script **lit la version dans l'exécutable compilé**. Pour publier une nouvelle version,
il n'y a donc **qu'un seul endroit à changer** : `My Project\AssemblyInfo.vb`
(`AssemblyVersion` **et** `AssemblyFileVersion`), puis recompiler en Release.

> **Ne jamais changer `AppId`** dans `Wincompense.iss` : c'est par lui que Windows reconnaît
> une mise à jour plutôt qu'un second produit installé côte à côte.

### 5.4 Ce que porte le fichier produit

Le setup n'est pas un binaire anonyme : il porte l'**icône de l'application** et, dans
*Propriétés → Détails*, son éditeur, son produit, sa description et sa version. C'est ce que
la sécurité de la banque regarde en premier sur un exécutable reçu par courriel.

**Ce qui manque encore, et qui ne dépend pas de nous : la signature numérique.** Un setup non
signé déclenche l'avertissement « Éditeur inconnu » de Windows SmartScreen. Y remédier demande
un **certificat de signature de code**, qui appartient à la banque. Le cas échéant, une seule
directive `SignTool` à ajouter, et rien d'autre à changer.

---

## 6. Étape 5 — Installer un poste

### 6.1 L'assistant, écran par écran

Lancer `Wincompense_Setup.exe` **en tant qu'administrateur**.

| | Écran | Ce qui s'y passe |
|---|---|---|
| 1 | Élévation Windows | L'installation écrit dans `Program Files` : les droits administrateur sont demandés |
| 2 | *(le cas échéant)* | Si .NET Framework 4.8 manque, un avertissement s'affiche **avant tout** |
| 3 | Bienvenue | — |
| 4 | Dossier de destination | `C:\Program Files\Default Company Name\SetupWincompense` — **le chemin autorisé par la banque, à ne pas changer** |
| 5 | Dossier du menu Démarrer | `Wincompense TCHAD` |
| 6 | Tâches supplémentaires | Raccourci sur le Bureau, à cocher |
| 7 | **Connexion à la base** | Le serveur **ou** la chaîne complète, et le chemin du fichier partagé |
| 8 | Prêt à installer | Récapitulatif |
| 9 | Installation | Quelques secondes |
| 10 | Terminé | Case « Lancer Wincompense TCHAD » |

À l'écran 7, les deux formes sont acceptées **dans le même champ** — l'assistant les
distingue seul, sur la présence d'un signe `=` :

| Ce que la banque vous donne | Ce que vous saisissez | Ce qui est écrit |
|---|---|---|
| Un nom de serveur | `SRV-SQL01\SQLEXPRESS` | `SERVEUR=` et `BASE=` |
| Une chaîne de connexion | `Server=SRV-SQL01;Database=…;Integrated Security=True;` | `CHAINE=` |

L'assistant refuse une chaîne qui n'indique aucun serveur, et refuse de poser sur le partage
une chaîne contenant un mot de passe.

Sur le **deuxième poste et les suivants**, le fichier partagé existe déjà : il n'est pas
écrasé. C'est voulu — la saisie d'un technicien ne doit pas faire basculer toute la banque
par accident.

### 6.2 Le chemin et le nom de l'exécutable sont imposés par la sécurité

La sécurité de la banque n'autorise pas *une application* : elle autorise **un fichier à un
emplacement précis**.

```
C:\Program Files\Default Company Name\SetupWincompense\Wincompense.exe
```

> L'Explorateur Windows en français affiche `C:\Programmes\Default Company Name\…`. C'est le
> **même dossier** : Windows ne traduit que son nom à l'écran. Pour lire le vrai chemin,
> cliquer dans la barre d'adresse.

`Default Company Name` et `SetupWincompense` ne veulent rien dire : ce sont les valeurs par
défaut de l'ancien projet d'installation Visual Studio. **C'est précisément pour cela qu'il
ne faut pas y toucher** : les rendre présentables obligerait la banque à refaire son
autorisation.

Ce n'est **pas** le dossier en `(x86)` — la banque l'a confirmé, et `Wincompense.iss` est
déjà réglé ainsi. Ne pas basculer `#define RacineProgrammes` sur `{autopf32}`.

### 6.3 Ce qui est posé sur le poste

| Emplacement | Contenu |
|---|---|
| `…\SetupWincompense\` | L'application, plus `Scripts\` (le script `00`, le `11`, et les 24 fichiers de documentation) et `Installation\` |
| `C:\ProgramData\Wincompense\wincompense.config` | La configuration du poste |
| Menu Démarrer, Bureau | Les raccourcis |
| Panneau de configuration | L'entrée de désinstallation |
| Le partage réseau | `connexion.config`, **uniquement s'il n'existe pas déjà** |

Les fichiers de débogage (`*.pdb`, `*.xml`) ne partent pas en production.

Le dossier `…\SetupWincompense\Logos\` est créé **même vide** : l'application y cherche
`WU.png`, `RIA.png` — un fichier par produit — et, quand il manque, dessine les **initiales
du produit sur un carré de sa couleur**. Format attendu : PNG carré, 128 pixels suffisent.

### 6.4 Pour vingt postes : l'installation silencieuse

```
Wincompense_Setup.exe /VERYSILENT /SUPPRESSMSGBOXES ^
    /SERVEUR="SRV-SQL01\SQLEXPRESS" ^
    /PARTAGE="\\SRV-FICHIERS\Wincompense\connexion.config"
```

> **Ces deux paramètres sont obligatoires sur un poste NEUF, et à proscrire sur une mise à
> jour.** Sur un poste neuf, sans eux, personne ne saisit rien et le poste repart sur
> `.\SQLEXPRESS` — sur aucun serveur. Sur une mise à jour, l'installateur relit ce que le
> poste porte déjà et le réécrit à l'identique : passer `/SERVEUR` par habitude écraserait un
> réglage correct sur tout le lot.
>
> Les contrôles de saisie ne s'appliquent dans aucun des deux cas : c'est l'informatique qui
> répond de ce qu'elle passe en paramètre.

Pour éprouver la commande avant de la diffuser, remplacer `/VERYSILENT` par `/SILENT` : la
barre de progression s'affiche, et les erreurs restent visibles.

Le détail — les commutateurs un par un, l'essai en trois temps, les quatre pièges qui font échouer un lot, la règle de détection pour SCCM et le script de démarrage type — est dans **`PROCEDURE-DEPLOIEMENT-SILENCIEUX.pdf`**, à remettre à l'informatique.

### 6.5 Donner le mot de passe SQL à un poste

Nécessaire **seulement** en compte SQL Server, au premier déploiement ou après un changement
de mot de passe. Un simple changement de serveur ne demande rien.

| | Façon | Quand |
|---|---|---|
| 1 | *Sécurité → Connexion à la base de données…*, cocher « Employer une chaîne de connexion complète », coller la chaîne **avec** son mot de passe, **tester**, enregistrer **sans** cocher « TOUS les postes » | Un poste, à la main |
| 2 | Ajouter `MOTDEPASSE_CLAIR=le-mot-de-passe` à `%PROGRAMDATA%\Wincompense\wincompense.config` | Déploiement scripté |
| 3 | Réinstaller le poste : l'assistant demande la chaîne | Poste neuf |

Dans le cas 2, l'application **chiffre le mot de passe au démarrage suivant et efface la
ligne en clair** : la fenêtre d'exposition se referme au premier lancement.

> Le chiffrement est lié à la machine : recopier `wincompense.config` d'un poste sur un autre
> ne transporte pas le mot de passe.

### 6.6 Sur un poste portant déjà l'ancienne version

L'installation écrase l'ancienne au même endroit et **supprime** `WincompenseTCHAD.exe`, son
`.config` et son `.pdb`.

Si l'ancienne version avait été posée par un projet d'installation Visual Studio, elle a
aussi une entrée dans *Programmes et fonctionnalités* : **la désinstaller d'abord**. Sinon,
le jour où quelqu'un désinstallera cette vieille entrée, elle emportera les fichiers de la
nouvelle version au passage.

**Vérification à constater :** l'application a atterri dans
`…\Default Company Name\SetupWincompense\Wincompense.exe`, elle démarre sans blocage de la
sécurité, et la barre d'état affiche le bon serveur et la bonne base.

---

## 7. Étape 6 — Le premier démarrage

### 7.1 Le premier administrateur

Au tout premier démarrage, la base ne contient aucun administrateur : personne ne pourrait
se connecter, ni créer personne. L'application le détecte et **propose de créer ce premier
administrateur**, en nommant la base visée — lire ce nom avant de répondre *Oui* est le seul
moyen de ne pas créer un administrateur dans la base de test.

### 7.2 Le choix du produit

Après l'identification, la **fenêtre de choix du produit** s'affiche : **Western Union** y
est *Disponible*, **Ria** *En attente*. Ouvrir Western Union — le titre de la fenêtre
devient alors `Wincompense TCHAD — Western Union`.

Le bouton **Changer de produit**, à droite de la barre de menus, ramène au choix **sans
redemander le mot de passe**.

### 7.3 Les rôles et les fonctions

| Rôle | Ce qu'il peut |
|---|---|
| `COMPENSE` | Traiter la compense, lire les rapports |
| `COMMERCIAL` | Les rapports et le référentiel, sans la compense |
| `ADMIN` | Tout, plus les comptes systèmes, les utilisateurs et la connexion à la base |

Deux **fonctions** s'y ajoutent, pour le double regard sur le référentiel :

| Fonction | Rôle dans le double regard |
|---|---|
| `INPUTER` | Saisit une création, une modification, une suppression |
| `AUTHORIZER` | L'autorise, ou la refuse |

> **L'inputer et l'authorizer doivent être deux personnes distinctes : la base le refuse
> autrement.** C'est le principe même du double regard, et il n'est pas contournable depuis
> l'application.

| | À faire | Qui |
|---|---|---|
| ☐ | Créer les agents de compense, le commercial, l'**inputer** et l'**authorizer** | Administrateur |
| ☐ | Vérifier que chacun change son mot de passe à la première connexion | Administrateur |
| ☐ | Éprouver le double regard : créer un sous-agent fictif en tant qu'inputer, l'autoriser en tant qu'authorizer, puis le supprimer | COMPTA |
| ☐ | Vérifier qu'un inputer **ne peut pas** autoriser sa propre demande | COMPTA |

---

## 8. Étape 7 — Le paramétrage, dans l'ordre

Tout se fait **depuis l'application**, menu **Paramétrage**. L'ordre ci-dessous n'est pas
indifférent : les comptes commandent la pièce, le référentiel commande ce qui est
comptabilisé.

### 8.1 Les onze comptes systèmes — *Paramétrage → Comptes systèmes*

| | Compte | Repli du code |
|---|---|---|
| 1 | Compte courant / position nette WU | `32100003292` |
| 2 | Compte inter bancaire | `381000101` |
| 3 | Commission sur transfert, part banque | `728300148` |
| 4 | Commission sur envoi, part banque | `728300148` |
| 5 | Commission sur paiement, part banque | `728300149` |
| 6 | Impôts et taxe sur envoi | `434000147` |
| 7 | TVA collectée | `434000104` |
| 8 | TTA sur envoi | `434000145` |
| 9 | TTA sur réception | `434000159` |
| 10 | **Gain de change** | *aucun* — le script amorce `371100102` |
| 11 | **Perte de change** | *aucun* — le script amorce `671100102` |

La colonne de droite n'est **pas** ce que l'application emploie : c'est la valeur sur
laquelle elle retombe quand la base est injoignable. Ce qu'elle emploie est ce que porte
`SystemeWU`, et c'est cet écran qui le montre.

> **Ces onze comptes déterminent toute la pièce comptable.** Une erreur ici ne se voit pas à
> l'écran : elle se voit sur les comptes, après chargement. Les faire valider un par un par
> la Direction Comptable.
>
> Les deux comptes de change n'ont **aucune valeur de repli dans le code** : vides, ils
> bloquent la pièce de change — et ne bloquent qu'elle. C'est voulu : un compte de change
> deviné comptabiliserait des centaines de milliers de francs sur un numéro que personne
> n'a confirmé.

### 8.2 Les options de traitement — *Paramétrage → Options de traitement…*

| Option | Ce qu'elle commande |
|---|---|
| Visa avant le fichier core banking | OUI : le fichier ne peut pas être produit tant que la journée n'est pas visée. NON : l'application avertit seulement |
| Envois en attente dans les écarts de change | OUI : les envois de statut `W` entrent dans le calcul. NON : ils en sont écartés |

### 8.3 Les libellés de la pièce — *Paramétrage → Narrative comptable…*

Deux choses s'y règlent : le **modèle** de la narrative du point de vente — la phrase que
porte la colonne `ADDLTEXT` du fichier core banking, avec ses repères `{AGENCE}`,
`{ACCOUNT}`, `{CODE_AGENCE}`, `{PERIODE}` — et le **mode** des libellés de la pièce :
`PAR_NATURE` (un libellé par nature de mouvement) ou `GLOBAL` (le même pour les douze lignes
d'un point de vente).

`PAR_NATURE` est le mode par défaut, et le mode de repli en cas de doute.

### 8.4 Le format du fichier core banking — *Paramétrage → Format du fichier core banking…*

Réservé à l'administrateur. **Sept valeurs** y sont confiées à la banque : `DETBSJRNL`,
`BRN`, `SRCCODE`, les deux `TXNCD` (débit et crédit), l'agence du siège, et la liste des
comptes rattachés au siège.

Un **aperçu d'une ligne de fichier** reste affiché en permanence, construit par le vrai
constructeur de ligne. Un onglet porte l'**historique des modifications**.

> Une base où ces clés n'existent pas produit **exactement le fichier d'aujourd'hui** :
> chaque valeur retombe sur la constante du code. Le paramétrage n'est donc pas une condition
> de démarrage.

### 8.5 Les produits de transfert — *Paramétrage → Produits de transfert…*

Western Union est *Disponible*, Ria *En attente*. Un produit se retire du service ; **aucun
bouton ne le supprime** — un produit supprimé emporterait l'historique qui s'y rattache.

### 8.6 Les jours fériés

Le script `00` sème les fêtes **à date fixe** et les lundis de Pâques. Les **fêtes
musulmanes ne se calculent pas d'avance** : elles sont à compléter, pour l'année en cours et
la suivante. *(COMPTA, chaque année)*

### 8.7 Le référentiel des points de vente

C'est la phase **la plus longue** du déploiement, et la plus souvent sous-estimée.

| | À faire | Où |
|---|---|---|
| ☐ | Charger les **sous-agents** : code, désignation, taux, **compte de compensation**, **compte de commission**, code agence | *Paramétrage → Sous-agents* |
| ☐ | Charger les **agences propres** : code site, désignation, code agence Voyager | *Paramétrage → Agences propres* |
| ☐ | Créer les **groupes statistiques** et y rattacher les sous-agents | *Paramétrage → Groupes statistiques* |
| ☐ | **Contrôle** : aucun sous-agent actif sans compte de compensation ni compte de commission | — |
| ☐ | **Contrôle** : aucun sous-agent sans groupe statistique, ou liste assumée de ceux qui restent | — |

> **Le contrôle des comptes n'est pas une formalité.** Un sous-agent auquel il manque l'un des
> deux comptes **n'est pas comptabilisé** : son activité n'apparaît ni dans la pièce ni dans
> le fichier core banking, et reste à régulariser à la main. L'application le signale et
> chiffre le montant écarté, mais mieux vaut ne jamais en arriver là.

Pour relire le réseau une fois chargé : *Paramétrage → Liste du réseau (impression)* édite
tous les points de vente, **rangés par groupe statistique**, un bloc par section, avec la
date de création et son auteur quand l'historique les connaît.

**Vérification à constater :** la liste du réseau imprimée, relue et signée par COMPTA.

---

## 9. Étape 8 — Sauvegarder le paramétrage

*Paramétrage → Paramétrage : fichier de secours…*

Le paramétrage représente des jours de saisie. Cet écran l'**exporte dans une archive**, et
sait la **relire** pour remonter une base.

L'archive porte neuf entrées :

| Entrée | Contenu |
|---|---|
| `manifeste.csv` | Version du format, base, serveur, date, auteur, **empreinte du contenu** |
| `LISEZ-MOI.txt` | Ce que l'archive contient, en clair |
| `comptes-systeme.csv` | Les onze comptes |
| `groupes-statistiques.csv` | Les groupes et leurs comptes |
| `sous-agents.csv` | Le référentiel des sous-agents |
| `agences-propres.csv` | Le référentiel des agences |
| `options-de-traitement.csv` | Les options et les sept codes du format core banking |
| `narratives-par-nature.csv` | Les libellés, nature par nature |
| `utilisateurs-pour-information.csv` | La liste des comptes, **jamais rechargée** |

À la relecture, l'écran annonce ce qui **diffère** de la base avant d'écrire quoi que ce
soit, et l'**empreinte** dit si le fichier a été retouché depuis son export.

> **Le format est en version 2.** Une archive produite par une version plus récente que
> l'application est **refusée** plutôt que lue à moitié : une v1 lisant une v2 perdrait
> silencieusement les options et les libellés.

| | À faire | Quand |
|---|---|---|
| ☐ | Exporter le paramétrage et **ranger l'archive hors du poste** | Après l'étape 7, puis après chaque modification des comptes |

---

## 10. Étape 9 — La recette, avant la mise en service

C'est ici que le déploiement se joue. **Elle ne se contourne pas.**

| | Étape | Qui |
|---|---|---|
| ☐ | Choisir une **journée réelle déjà comptabilisée à la main**, et retrouver sa pièce manuelle | COMPTA |
| ☐ | Rejouer cette journée dans l'application, de l'étape 1 à l'étape 4 | COMPTA |
| ☐ | Comparer la pièce produite et la pièce manuelle, **ligne à ligne** | COMPTA |
| ☐ | Vérifier l'**écart d'arrondi global** et le compte qui le reçoit | COMPTA |
| ☐ | Vérifier la liste des Accounts écartés, et que chacun s'explique | COMPTA |
| ☐ | Produire le fichier core banking : ses colonnes, son numéro de lot, sa date de valeur | COMPTA |
| ☐ | **Charger ce fichier dans l'environnement de TEST du core banking** | INFO |
| ☐ | Faire confirmer le format de `VALDT` par le retour de ce chargement | INFO |
| ☐ | Rejouer une journée **atypique** : un lundi rattrapant le week-end | COMPTA |
| ☐ | Rejouer une journée comportant des **annulations** | COMPTA |
| ☐ | Rejouer une journée comportant un **Account absent du référentiel** | COMPTA |
| ☐ | Éditer le **rapport d'activité** sur une période connue et le confronter aux chiffres de la banque | COMPTA |

> **Jamais dans le core banking de production à ce stade.** Le fichier impacte réellement les
> comptes ; c'est en test qu'on découvre un format inattendu, pas après.

**Go / no-go.** Si la pièce diffère de la pièce manuelle sur autre chose qu'un arrondi au
franc, on ne met pas en service : on corrige, et on recommence la recette entière.

Vient ensuite la **marche en parallèle** — deux à quatre semaines pendant lesquelles la
compense est faite à la main **et** par l'application, seule la pièce manuelle étant chargée
dans le core banking. Le détail de cette phase et de la bascule est dans
`PLAN-DEPLOIEMENT.md`, phases 6 et 7.

---

## Annexe A — Changer de serveur

C'est le geste que toute cette architecture existe pour rendre simple. **Quatre façons, par
ordre de préférence.**

**a) Depuis l'application** — menu *Sécurité → Connexion à la base de données…*, réservé à
l'administrateur :

1. Saisir le nouveau serveur.
2. **Tester la connexion.** L'écran dit quelle base répond, et sous quel nom.
3. Cocher **« Appliquer ce réglage à TOUS les postes »**.
4. Enregistrer. Les autres postes prennent le nouveau serveur à leur prochain démarrage.

**b) Dans le Bloc-notes** — ouvrir `connexion.config` sur le partage, modifier `SERVEUR=`,
enregistrer. Rien d'autre.

**c) En ligne de commande** :

```powershell
.\Configurer-Connexion.ps1 -Serveur SRV-SQL02 `
                           -Partage \\SRV-FICHIERS\Wincompense\connexion.config
```

Le script teste la connexion **avant** d'écrire. `-Force` passe outre, pour une bascule
préparée alors que le nouveau serveur n'est pas encore en ligne.

**d) Avec une chaîne complète** — cocher « Employer une chaîne de connexion complète »,
coller, tester, enregistrer. Le serveur, la base et le délai se grisent : ils ne comptent
plus.

### Ce qui reste à faire côté base

Changer de serveur ne déplace pas les données. Sur le nouveau serveur :

1. Restaurer `GWC_WINCOMPENSE_ETD`, **ou** la recréer avec
   `Scripts\00_InstallationComplete.sql`.
2. En authentification Windows, rejouer `Scripts\11_AccesUtilisateurs.sql` complété. **Un
   serveur restauré garde ses utilisateurs de base mais perd ses logins** : c'est la cause la
   plus fréquente du « Login failed for user » après une migration.
3. En compte SQL Server, recréer le login **et le rattacher** à l'utilisateur de base resté
   orphelin — les trois ordres de la section 3.1.

---

## Annexe B — Publier une mise à jour

Les postes ne se mettent pas à jour tout seuls : l'informatique garde la main sur le moment.
En revanche, **aucun poste ne doit ignorer qu'une version plus récente existe** : au
démarrage, l'application lit `version.txt` sur le partage et se compare à lui.

L'agent en retard voit une mention rouge dans la **barre d'état** — un clic ouvre le dossier
du setup, le fichier déjà sélectionné — et **un message, une seule fois par version**.
Répété chaque matin, il serait fermé sans être lu, et le suivant, celui qui compte vraiment,
le serait aussi. Rien n'est bloquant.

| | Étape |
|---|---|
| 1 | Incrémenter la version dans `My Project\AssemblyInfo.vb` (`AssemblyVersion` **et** `AssemblyFileVersion`) |
| 2 | Compiler en **Release**, puis produire le setup |
| 3 | **Rejouer `Scripts\00_InstallationComplete.sql`** sur la base de production |
| 4 | Déposer le setup sur le partage, dans un sous-dossier `Setup\` |
| 5 | **Alors seulement**, mettre à jour `version.txt` |

> **L'ordre des étapes 4 et 5 n'est pas indifférent** : annoncer une version avant d'avoir
> déposé son programme d'installation envoie les agents chercher un fichier qui n'existe pas.
>
> **Et l'étape 3 ne s'oublie pas** : une livraison qui ajoute un écran ajoute souvent une
> colonne. L'application démarrerait, et échouerait sur cet écran-là.

Tant que `version.txt` n'existe pas, rien ne s'affiche et rien n'échoue : le dispositif est
un confort, jamais une condition de démarrage.

---

## Annexe C — Vérifier ce qu'un poste emploie réellement

Trois moyens, du plus simple au plus détaillé :

- la **barre d'état** affiche en permanence le serveur et la base en service ;
- l'écran **Connexion à la base de données** indique en plus **d'où vient** la chaîne :
  « fichier partagé », « copie locale (partage injoignable) », etc. ;
- en ligne de commande : `.\Configurer-Connexion.ps1 -Afficher`.

La distinction compte : un administrateur qui croit lire le partage alors qu'il lit une copie
locale périmée modifierait un fichier sans aucun effet.

**Deux conséquences de l'ordre des rangs (section 4.1), qui surprennent :**

- **Tant qu'un réglage de session est actif, cet écran ne dit pas ce que le poste
  emploiera.** La ligne du haut affiche alors *« saisie pour cette session, non conservée sur
  le poste »*. Pour voir la vraie origine : fermer et rouvrir l'application.
- **Un réglage conservé sur le poste ne prime pas sur le fichier partagé** — c'est l'inverse.
  Pour qu'un réglage local gouverne vraiment : soit cocher « Appliquer à TOUS les postes »,
  soit **vider le champ du fichier partagé** pour que ce poste cesse de le suivre.

---

## Annexe D — Dépannage

| Symptôme | Cause probable | Ce qu'il faut faire |
|---|---|---|
| « copie locale (… injoignable) » | Le partage ne répond pas | Vérifier le serveur de fichiers et les droits de lecture. Le poste travaille sur la dernière chaîne connue : ce n'est pas urgent |
| Le poste garde l'ancien serveur | L'application n'a pas été redémarrée, ou le réglage n'a pas été propagé | Fermer et rouvrir. Vérifier que « TOUS les postes » était cochée |
| « valeur par défaut (aucune configuration trouvée) » | Poste jamais installé, ou `wincompense.config` supprimé | Relancer l'installation, ou régler par l'écran de connexion |
| Écriture refusée sur le partage | Droits insuffisants | L'écriture est réservée à l'informatique. Rien n'a été modifié |
| « Le serveur est introuvable » alors que la barre affiche `SRV-SQL01` | `SRV-SQL01` est le **nom d'exemple** de la documentation | Remplacer `SERVEUR=` par le nom réel |
| Le nom du serveur est bon, la connexion échoue | Instance nommée sans *SQL Server Browser*, TCP/IP désactivé, ou pare-feu | Démarrer *SQL Server Browser* ; activer TCP/IP ; ouvrir le port 1433 |
| « Login failed for user » après une bascule | Le mot de passe SQL ne se propage pas par le partage | Donner le mot de passe à ce poste — section 6.5 |
| L'application démarre mais aucune donnée n'apparaît | Bonne connexion, mauvaise base | Tester la connexion : l'écran affiche le nom de la base qui répond |
| Les exports Excel échouent | Excel absent du poste | Installer Excel. Les calculs et les écrans restent utilisables |
| Un écran annonce qu'une table ou une colonne manque | Le script `00` n'a pas été rejoué après la livraison | Le rejouer : il ne crée que ce qui manque |

### Trouver le nom exact du serveur

Dans SSMS : `SELECT @@SERVERNAME`. Sans SSMS : *Services* Windows → chercher **SQL Server**.

| Ce que le service affiche | Ce qu'il faut écrire dans `SERVEUR=` |
|---|---|
| `SQL Server (MSSQLSERVER)` | `NOM-DU-SERVEUR` — instance par défaut |
| `SQL Server (SQLEXPRESS)` | `NOM-DU-SERVEUR\SQLEXPRESS` |
| La base est sur le poste lui-même | `.\SQLEXPRESS`, ou `.` |

### Si la connexion échoue au démarrage

Le bouton **« Serveur… »** apparaît alors sur l'écran de connexion. Il ouvre le réglage du
serveur **sans authentification** — impossible par définition quand la base ne répond pas —
et retente la lecture dès la fermeture. Il reste masqué tant que tout va bien.

---

## Annexe E — Désinstallation

Panneau de configuration → *Programmes et fonctionnalités* → **Wincompense TCHAD**.

La désinstallation **pose la question** : *« Supprimer aussi la configuration de ce
poste ? »*, et la réponse par défaut est **Non**.

| Réponse | Effet |
|---|---|
| **Non** | `%PROGRAMDATA%\Wincompense` est conservé. Une réinstallation retrouve le serveur, le partage et le mot de passe SQL sans ressaisie. C'est la réponse d'une désinstallation de dépannage |
| **Oui** | Le dossier est supprimé. C'est la réponse d'un retrait définitif |

En désinstallation silencieuse, aucune question n'est possible : le dossier est conservé.

> Ce dossier porte le **mot de passe du compte SQL**, chiffré par Windows pour cette machine.
> L'effacer sans demander ferait payer le geste le plus banal — désinstaller pour réinstaller
> — par une ressaisie que le technicien n'a pas toujours sous la main.

Le fichier partagé, lui, n'est **jamais** touché : il appartient à la banque, pas au poste.

---

## Annexe F — Récapitulatif à cocher

| | Étape | Qui |
|---|---|---|
| ☐ | Les cinq éléments de la section 1.1 sont réunis | INFO + COMPTA |
| ☐ | `00_InstallationComplete.sql` exécuté, compte rendu lu | INFO |
| ☐ | Mode de récupération vérifié, sauvegardes en place | INFO |
| ☐ | Le compte applicatif ouvre la base (login, utilisateur, rôle) | INFO |
| ☐ | `11_AccesUtilisateurs.sql` exécuté — **authentification Windows seulement** | INFO |
| ☐ | Partage créé, `connexion.config` posé, droits posés | INFO |
| ☐ | Setup produit en **Release** avec Inno Setup 6.4.x | PROJET |
| ☐ | Poste pilote installé **au bon chemin**, application démarrée | INFO |
| ☐ | Installation silencieuse éprouvée avec `/SILENT` sur un second poste | INFO |
| ☐ | Premier administrateur créé, puis les agents, l'inputer et l'authorizer | INFO + COMPTA |
| ☐ | Double regard éprouvé, y compris le refus de l'auto-autorisation | COMPTA |
| ☐ | Onze comptes systèmes saisis et validés un par un | COMPTA |
| ☐ | Options, narratives, format core banking, produits relus | Administrateur |
| ☐ | Fêtes musulmanes complétées pour deux années | COMPTA |
| ☐ | Référentiel chargé, les deux contrôles passés | COMPTA |
| ☐ | Liste du réseau imprimée et relue | COMPTA |
| ☐ | Paramétrage exporté, archive rangée hors du poste | Administrateur |
| ☐ | Recette passée, fichier chargé dans le core banking de **test** | COMPTA + INFO |
| ☐ | Marche en parallèle lancée | COMPTA |

---

## Annexe G — Ce que ce manuel ne décide pas

| Point | En attente de |
|---|---|
| Le **compte dédié à l'écart d'arrondi**, et s'il en faut un au débit et un au crédit | La Direction Comptable |
| Le **format exact des sept codes** du fichier core banking — longueur, jeu de caractères, casse | Le core banking |
| Le **certificat de signature de code**, qui ferait taire l'avertissement « Éditeur inconnu » | La banque |
| Les **groupes Active Directory** à rattacher aux trois rôles SQL Server | L'informatique |
| L'**expiration périodique des mots de passe** : oui / non, à quelle échéance | La sécurité |

Aucun de ces cinq points n'empêche l'installation : chacun a aujourd'hui un comportement
par défaut, décrit à sa place dans ce manuel.
