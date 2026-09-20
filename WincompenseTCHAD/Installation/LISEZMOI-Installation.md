# Wincompense TCHAD — installation et changement de serveur

Ce dossier contient tout ce qu'il faut pour déployer l'application sur les postes,
et pour changer de serveur ensuite sans y revenir.

| Fichier | Rôle |
|---|---|
| `Wincompense.iss` | Script Inno Setup produisant `Wincompense_Setup.exe` |
| `connexion.config.modele` | Modèle du fichier à poser sur le partage réseau |
| `Configurer-Connexion.ps1` | Changement de serveur en ligne de commande |
| `version.txt.modele` | Modèle du fichier annonçant la version publiée |
| `LISEZMOI-Installation.md` | Ce document — installation technique et changement de serveur |
| `PLAN-DEPLOIEMENT.md` | Les huit phases du déploiement à la banque, à cocher point par point |
| `PROCEDURE-SSMS-Acces.md` | À remettre à la banque : donner accès au compte, clic par clic dans SSMS |

---

## 1. Le principe : la connexion ne vit pas avec l'application

La banque change souvent de serveur. Tant que la chaîne de connexion vivait dans
`Wincompense.exe.config`, en changer imposait une tournée dans les bureaux :
ce fichier est dans `Program Files`, donc protégé ; il est propre à chaque poste ;
et une réinstallation l'écrase.

Elle est donc cherchée ailleurs, dans un ordre où **le premier trouvé l'emporte** :

| Rang | Emplacement | À quoi il sert |
|---|---|---|
| 1 | Variable d'environnement `WINCOMPENSE_CONNEXION` | Dépannage, poste de test |
| 2 | **Fichier partagé** désigné à l'installation | **La source de vérité** |
| 3 | `%PROGRAMDATA%\Wincompense\wincompense.config` | Copie locale, rafraîchie à chaque lecture réussie du partage |
| 4 | `App.config` de l'application | Poste de développement |
| 5 | Valeur compilée `.\SQLEXPRESS` | Dernier recours |

Le rang 3 n'est pas un doublon. C'est lui qui fait travailler le poste le matin où
le partage est injoignable — sans quoi une coupure réseau arrêterait la compense de
toute la banque.

**Si l'authentification est celle de Windows**, aucun mot de passe ne circule, et c'est
précisément ce qui permet de poser la configuration sur un partage lisible par tous.

**Si la banque fournit un compte SQL Server et son mot de passe**, la chaîne est coupée en
deux : le partage reçoit le serveur, la base et le **nom** du compte ; le mot de passe reste
sur le poste, chiffré par Windows. Changer de serveur vaut toujours pour toute la banque —
mais **chaque poste doit recevoir le mot de passe une fois**, à son installation. Sans quoi il
affichera « Login failed for user ».
Les droits d'accès à la base sont donnés par `Scripts\08_RolesSQLServer.sql` (les rôles) et
`Scripts\11_AccesUtilisateurs.sql` (les comptes).

---

## 2. Construire le programme d'installation

À faire une fois par version, sur le poste de développement.

### Quelle version d'Inno Setup — à trancher par la banque

**Inno Setup 7 n'est plus gratuit pour un usage commercial.** Sa fenêtre affiche
« Non-commercial use only » et invite à acheter une licence. Une banque qui s'en sert
pour outiller son activité est dans le cadre commercial.

Deux issues, l'une et l'autre valables :

| Choix | Ce que cela implique |
|---|---|
| **Inno Setup 6.4.x** *(recommandé)* | Gratuit pour tout usage, commercial compris. Le script ci-joint fonctionne dès la 6.3. Versions antérieures dans les archives : <https://jrsoftware.org/isdl.php> |
| **Inno Setup 7 avec licence** | Acheter la licence auprès de jrsoftware. Rien à changer au script |

Le script ne dépend d'aucune nouveauté de la version 7 : il compile à l'identique
sous 6.3 et suivantes.

### Les quatre étapes

1. **Compiler en Release.** C'est l'oubli le plus fréquent, et il arrête tout.
   Dans Visual Studio, la liste déroulante de la barre d'outils affiche *Debug* par
   défaut : choisir **Release**, puis *Générer* → *Générer la solution*.
   Vérifier ensuite que `WincompenseTCHAD\WincompenseTCHAD\bin\Release\Wincompense.exe`
   existe et porte la date du jour.

   Sans cela, la compilation du script s'arrête sur un message explicite disant
   exactement cela.

2. **Installer Inno Setup**, selon le choix ci-dessus.

3. Ouvrir `Installation\Wincompense.iss`, puis **Build → Compile** (Ctrl+F9).

4. Le programme d'installation apparaît dans `Installation\Sortie\`.

Pour compiler depuis un autre emplacement — serveur de construction, dossier
déplacé — le chemin du dossier Release se passe en ligne de commande :

```
ISCC.exe /DDossierRelease="C:\chemin\vers\bin\Release" Wincompense.iss
```

Pour une nouvelle version, changer `VersionApplication` en tête du script. **Ne jamais
changer `AppId`** : c'est par lui que Windows reconnaît une mise à jour plutôt qu'un
second produit installé côte à côte.

---

## 3. Préparer le partage réseau — une seule fois

1. Créer un dossier sur le serveur de fichiers, par exemple `\\SRV-FICHIERS\Wincompense`.
2. Y copier `connexion.config.modele` sous le nom **`connexion.config`**.
3. Renseigner la ligne `SERVEUR=`.
4. Poser les droits :

| Qui | Droit |
|---|---|
| Utilisateurs de Wincompense | **Lecture** |
| Informatique, administrateur Wincompense | **Lecture et écriture** |

Ce point n'est pas une formalité : qui peut écrire dans ce fichier commande la
connexion de toute la banque.

---

## 4. Installer un poste

Lancer `Wincompense_Setup.exe` **en tant qu'administrateur**. L'assistant demande :

- **le serveur SQL Server, ou la chaîne de connexion complète** — les deux formes sont acceptées
  dans le même champ, l'assistant les distingue seul :

  | Ce que la banque vous donne | Ce que vous saisissez | Ce qui est écrit |
  |---|---|---|
  | Un nom de serveur | `SRV-SQL01\SQLEXPRESS` | `SERVEUR=` et `BASE=` |
  | Une chaîne de connexion | `Server=SRV-SQL01;Database=…;Integrated Security=True;` | `CHAINE=` |

  La distinction se fait sur le signe `=` : une chaîne de connexion en porte toujours au moins
  un, un nom de serveur jamais. Vous n'avez donc rien à choisir.

- le **chemin du fichier partagé** — par exemple `\\SRV-FICHIERS\Wincompense\connexion.config`.

L'assistant refuse une chaîne qui n'indique aucun serveur, et refuse de poser sur le partage une
chaîne contenant un mot de passe — il serait lisible en clair par tous les utilisateurs.

Il vérifie ensuite .NET Framework 4.8, copie l'application, écrit la configuration
locale, crée le fichier partagé **s'il n'existe pas encore**, et pose les raccourcis.

Sur le deuxième poste et les suivants, le fichier partagé existe déjà : il n'est pas
écrasé. C'est voulu — la saisie d'un technicien ne doit pas faire basculer toute la
banque par accident.

### Ce que l'assistant enchaîne, dans l'ordre

| | Écran | Ce qui s'y passe |
|---|---|---|
| 1 | Élévation Windows | L'installation écrit dans `Program Files` : les droits administrateur sont demandés |
| 2 | *(le cas échéant)* | Si .NET Framework 4.8 manque, un avertissement s'affiche **avant tout** |
| 3 | Bienvenue | — |
| 4 | Dossier de destination | `C:\Program Files\Default Company Name\SetupWincompense` — **le chemin autorisé par la banque, à ne pas changer** |
| 5 | Dossier du menu Démarrer | `Wincompense TCHAD` |
| 6 | Tâches supplémentaires | Raccourci sur le Bureau, à cocher |
| 7 | **Connexion à la base** | Serveur **ou** chaîne de connexion, et chemin du fichier partagé |
| 8 | Prêt à installer | Récapitulatif |
| 9 | Installation | Quelques secondes |
| 10 | Terminé | Case « Lancer Wincompense TCHAD » |

### Ce qui est posé sur le poste

| Emplacement | Contenu |
|---|---|
| `C:\Program Files\Default Company Name\SetupWincompense\` | L'application, plus `Scripts\` (les douze scripts SQL) et `Installation\` |
| `C:\ProgramData\Wincompense\wincompense.config` | La configuration, **modifiable par les utilisateurs** |
| Menu Démarrer, Bureau | Les raccourcis |
| Panneau de configuration | L'entrée de désinstallation |
| Le partage réseau | `connexion.config`, **uniquement s'il n'existe pas déjà** |

Les fichiers de débogage (`*.pdb`, `*.xml`, `*.vshost.*`) ne partent pas en production.

### Déploiement en masse — installation silencieuse

Pour vingt postes, on n'installe pas à la main. L'assistant accepte la configuration en ligne
de commande, ce qui permet de le lancer par stratégie de groupe ou par SCCM :

```
Wincompense_Setup.exe /VERYSILENT /SUPPRESSMSGBOXES ^
    /SERVEUR="SRV-SQL01\SQLEXPRESS" ^
    /PARTAGE="\\SRV-FICHIERS\Wincompense\connexion.config"
```

`/SERVEUR` accepte indifféremment un nom de serveur ou une **chaîne de connexion complète**,
comme le champ de l'assistant.

> **Sans ces paramètres, une installation silencieuse ne configure rien** : la page ne s'affiche
> pas, personne ne saisit rien, et le poste repart sur `.\SQLEXPRESS` — c'est-à-dire sur aucun
> serveur. Les contrôles de saisie ne s'appliquent pas non plus en mode silencieux : c'est
> l'informatique qui répond de ce qu'elle passe en paramètre.

Pour éprouver la commande avant de la diffuser, remplacer `/VERYSILENT` par `/SILENT` : la barre
de progression s'affiche, et les messages d'erreur éventuels restent visibles.

### Le chemin et le nom de l'exécutable sont imposés par la sécurité

La sécurité de la banque n'autorise pas *une application* : elle autorise **un fichier
à un emplacement précis**. Celui qui a été autorisé, et qui fonctionnait, est :

```
C:\Program Files\Default Company Name\SetupWincompense\Wincompense.exe
```

> L'Explorateur Windows en français l'affiche `C:\Programmes\Default Company Name\SetupWincompense`.
> C'est le **même dossier** : Windows ne traduit que son nom à l'écran, pas sur le disque.
> Pour lire le vrai chemin, cliquer dans la barre d'adresse de l'Explorateur : elle
> affiche alors `C:\Program Files\...`.

Ce chemin vient de l'ancien déploiement, fait avec un **projet d'installation Visual
Studio**. `Default Company Name` et `SetupWincompense` sont les valeurs que ce type de
projet met par défaut quand on ne renseigne ni l'éditeur ni le nom du produit. Elles ne
veulent rien dire — et c'est précisément pour cela qu'il ne faut pas y toucher : les
rendre plus présentables obligerait la banque à refaire son autorisation.

> **Au passage :** ce dossier n'est pas un dossier ClickOnce. ClickOnce installe sous
> `%LOCALAPPDATA%\Apps\2.0\`, avec des noms de dossiers illisibles. Ce qui avait été
> autorisé était donc bien un installateur classique — ce que `Wincompense.iss` reproduit.

Deux règles en découlent, et elles ne se négocient pas depuis le code :

| | Règle | Où c'est tenu |
|---|---|---|
| 1 | L'exécutable s'appelle `Wincompense.exe` | `WincompenseTCHAD.vbproj`, balise `AssemblyName` |
| 2 | Il s'installe dans `...\Default Company Name\SetupWincompense` | `Wincompense.iss`, `DefaultDirName` |

Si l'assistant d'installation propose un autre dossier, **ne pas valider** : il pose la
question et avertit, mais c'est l'opérateur qui tranche.

#### Ce n'est pas le dossier en (x86) — c'est confirmé

La question s'est posée : un projet d'installation Visual Studio en 32 bits aurait
déposé ses fichiers dans `C:\Program Files (x86)\`, affiché `C:\Programmes (x86)\`.
**La banque a confirmé le chemin sans `(x86)`** :

```
C:\Program Files\Default Company Name\SetupWincompense
```

`Wincompense.iss` est déjà réglé ainsi et **ne doit pas être touché** : ne pas basculer
`#define RacineProgrammes` sur `{autopf32}`.

> **Une subtilité qui se paie cher.** `{autopf}` seul ne désigne pas `C:\Program Files`.
> Sur un Windows 64 bits, Inno Setup le résout en `Program Files (x86)` tant que
> l'installation n'est pas en mode 64 bits. C'est la ligne
> `ArchitecturesInstallIn64BitMode=x64compatible`, dans la section `[Setup]`, qui fait
> tomber le chemin au bon endroit. **Les deux lignes tiennent ensemble** : retirer l'une
> sans l'autre déplace l'installation hors du chemin autorisé, et l'application cesse de
> démarrer sur les postes.

Cela reste à vérifier une fois, sur le premier poste installé — non pour décider, mais
pour constater : le tableau de recette du plan de déploiement porte la ligne.

#### Sur un poste portant déjà la version `WincompenseTCHAD.exe`

L'installation écrase l'ancienne au même endroit et **supprime** `WincompenseTCHAD.exe`,
son `.config` et son `.pdb` : laisser dans un dossier surveillé un binaire que la sécurité
n'autorise pas ne rendrait service à personne, et un utilisateur finirait par le lancer par
habitude.

Si l'ancienne version avait été posée par un projet d'installation Visual Studio, elle a
aussi une entrée dans *Programmes et fonctionnalités*. **La désinstaller d'abord**, avant
de lancer `Wincompense_Setup.exe` : sinon, le jour où quelqu'un désinstallera cette vieille
entrée, elle emportera les fichiers de la nouvelle version au passage.

`C:\ProgramData\Wincompense` n'est touché ni par l'une ni par l'autre désinstallation :
le serveur n'est jamais à ressaisir.

### Prérequis des postes

| Prérequis | Remarque |
|---|---|
| Windows 7 SP1 ou plus récent | Plancher du .NET Framework 4.8 |
| .NET Framework 4.8 | Vérifié par l'installateur ; présent d'office depuis Windows 10 1903 |
| Microsoft Excel | Pour les exports et le rapport PDF. L'application fonctionne sans, mais n'exporte plus |
| Accès réseau au serveur SQL | Port 1433, ou le port de l'instance nommée |

Aucun client SQL Server n'est à installer : le pilote fait partie du framework.

---

## 5. Changer de serveur

C'est le geste que toute cette architecture existe pour rendre simple. **Trois façons,
par ordre de préférence.**

### a) Depuis l'application — la voie normale

Menu **Sécurité → Connexion à la base de données…**, réservé à l'administrateur
Wincompense.

1. Saisir le nouveau serveur.
2. **Tester la connexion.** L'écran dit quelle base répond et sous quel nom la
   connexion est ouverte.
3. Cocher **« Appliquer ce réglage à TOUS les postes »**.
4. Enregistrer.

Les autres postes prennent le nouveau serveur à leur prochain démarrage.

### b) Dans le Bloc-notes

Ouvrir `\\SRV-FICHIERS\Wincompense\connexion.config`, modifier la ligne `SERVEUR=`,
enregistrer. Rien d'autre.

### c) En ligne de commande

```powershell
.\Configurer-Connexion.ps1 -Serveur SRV-SQL02 `
                           -Partage \\SRV-FICHIERS\Wincompense\connexion.config
```

Le script teste la connexion **avant** d'écrire. `-Force` passe outre, pour une
bascule préparée alors que le nouveau serveur n'est pas encore en ligne.

### d) Si la banque fournit une chaîne de connexion complète

Il arrive que la banque ne donne pas un nom de serveur mais une chaîne toute faite, avec des
mots-clés que `SERVEUR` et `BASE` ne savent pas exprimer : chiffrement imposé (`Encrypt`),
partenaire de secours (`Failover Partner`), groupe de disponibilité (`MultiSubnetFailover`),
port particulier, nom d'application.

**Dans l'application** — menu *Sécurité → Connexion à la base de données…* : cocher
**« Employer une chaîne de connexion complète »**, coller la chaîne, **tester**, puis
enregistrer. Le serveur, la base et le délai se grisent : ils ne comptent plus.

**Dans le fichier** — coller la chaîne après `CHAINE=`, **sur une seule ligne**, et mettre en
commentaire les lignes `SERVEUR` / `BASE` / `DELAI` :

```
# SERVEUR=...
# BASE=...
CHAINE=Server=SRV-SQL01;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;Encrypt=True;
```

`CHAINE` **prime** sur `SERVEUR`, `BASE` et `DELAI`. Renseigner les deux ferait coexister deux
descriptions du même serveur, dont une seule compte.

> **Aucun mot de passe dans le fichier partagé.** Il est lisible par tous les utilisateurs de
> l'application — c'est ce qui permet à un changement de serveur de valoir pour tout le monde ;
> un mot de passe y serait donc lisible en clair par tous.
>
> Une chaîne portant `User ID` / `Password` n'est **pas refusée** : la banque en fournit une.
> L'application **retire le mot de passe avant toute écriture**, le chiffre par Windows pour
> cette machine (DPAPI, portée `LocalMachine` : valable pour tous les utilisateurs du poste,
> illisible ailleurs), et n'écrit sur le partage que le serveur, la base et le **nom** du compte.
>
> **Conséquence à ne pas manquer :** le serveur se propage, le mot de passe non. Un poste qui ne
> l'a jamais reçu verra « Login failed for user » après la bascule. L'écran l'annonce avant
> d'enregistrer. Voir *« Donner le mot de passe aux autres postes »* ci-dessous.

### Donner le mot de passe aux autres postes

Nécessaire **seulement** si la banque a changé le mot de passe du compte SQL, ou lors du premier
déploiement en compte SQL. Un simple changement de serveur, à mot de passe inchangé, ne demande
rien : chaque poste applique celui qu'il détient déjà.

Trois façons, sur le poste concerné :

1. **Depuis l'application** — *Sécurité → Connexion à la base de données…*, cocher « Employer une
   chaîne de connexion complète », coller la chaîne **avec** son mot de passe, **tester**, puis
   enregistrer **sans** cocher « Appliquer à TOUS les postes ». Le mot de passe est chiffré pour
   ce poste ; le serveur, lui, continue de venir du partage — le poste ne devient pas sourd aux
   changements suivants.

2. **Une ligne dans le fichier local**, pour un déploiement scripté. Ajouter à
   `%PROGRAMDATA%\Wincompense\wincompense.config` :

   ```
   MOTDEPASSE_CLAIR=le-mot-de-passe
   ```

   Au démarrage suivant, l'application le chiffre sous `MOTDEPASSE` et **efface la ligne en
   clair**. La fenêtre d'exposition se referme au premier lancement.

3. **Réinstaller le poste** : l'assistant demande la chaîne et écrit `MOTDEPASSE_CLAIR`, repris
   par le même mécanisme.

> Le chiffrement est lié à la machine : recopier `wincompense.config` d'un poste sur un autre ne
> transporte pas le mot de passe, il n'y sera pas déchiffrable.

### Ce qui reste à faire côté base

Changer de serveur ne déplace pas les données. Sur le nouveau serveur, il faut :

1. restaurer la base `GWC_WINCOMPENSE_ETD`, ou la recréer avec
   **`Scripts\00_InstallationComplete.sql`** — un seul script, rejouable, qui monte les dix
   tables, les trois rôles et leurs droits, puis rend compte de ce qu'il a fait ;
2. exécuter **`Scripts\11_AccesUtilisateurs.sql`**, après y avoir mis vos comptes ou groupes
   Active Directory, pour redonner leurs droits aux utilisateurs Windows sur le nouveau
   serveur. **Un serveur restauré garde ses utilisateurs de base mais perd ses logins :**
   c'est la cause la plus fréquente du message « Login failed for user » après une migration ;
3. **si l'application se connecte par un compte SQL Server** — c'est le cas avec
   `etdwincompense` —, recréer ce login sur le nouveau serveur **et le rattacher** à
   l'utilisateur de base, que la restauration a laissé orphelin. Aucun script du projet ne le
   fait : les scripts livrés ne connaissent que les comptes Windows. À exécuter par
   l'informatique, sur le nouveau serveur :

   ```sql
   -- 1. le login, au niveau du serveur
   USE [master];
   IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'etdwincompense')
       CREATE LOGIN [etdwincompense] WITH PASSWORD = N'<le mot de passe>',
                                          CHECK_POLICY = OFF;
   GO
   -- 2. le rattachement de l'utilisateur orphelin, au niveau de la base
   USE [GWC_WINCOMPENSE_ETD];
   IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'etdwincompense')
       ALTER USER [etdwincompense] WITH LOGIN = [etdwincompense];
   ELSE
       CREATE USER [etdwincompense] FOR LOGIN [etdwincompense];
   GO
   -- 3. les rôles de l'application
   ALTER ROLE [wu_compense] ADD MEMBER [etdwincompense];
   GO
   ```

   Vérifier aussi que le serveur accepte le **mode mixte** (authentification Windows *et* SQL
   Server) : une instance neuve est en Windows seul, et le login existerait alors sans pouvoir
   servir.

---

## 5 bis. Publier une mise à jour

Les postes ne se mettent pas à jour tout seuls — c'est délibéré : l'informatique garde la main
sur le moment, et pendant la marche en parallèle deux versions différentes fausseraient la
comparaison avec la pièce manuelle. En revanche, **aucun poste ne doit ignorer qu'une version
plus récente existe**. Une correction livrée un lundi pouvait rester inconnue d'un poste pendant
des semaines, et deux agents produire des pièces différentes à partir des mêmes rapports.

### Ce que l'agent voit

Au démarrage, l'application lit `version.txt` sur le partage et se compare à lui. Si elle est en
retard :

- une mention rouge apparaît dans la **barre d'état** : « Version 1.1.0.0 disponible — cliquez
  ici ». Un clic ouvre le dossier du `setup.exe` dans l'Explorateur, le fichier déjà sélectionné ;
- un **message s'affiche une seule fois** par version. Répété chaque matin, il serait fermé sans
  être lu — et le suivant, celui qui compte vraiment, le serait aussi.

Rien n'est bloquant : l'agent peut travailler et installer plus tard.

### La marche à suivre

| | Étape |
|---|---|
| 1 | Incrémenter la version dans `My Project\AssemblyInfo.vb` (`AssemblyVersion` **et** `AssemblyFileVersion`) |
| 2 | Reporter le même numéro dans `VersionApplication`, en tête de `Wincompense.iss` |
| 3 | Compiler en **Release**, puis produire le setup |
| 4 | Déposer le `setup.exe` sur le partage, dans un sous-dossier `Setup\` |
| 5 | **Alors seulement**, mettre à jour `version.txt` |

> **L'ordre des étapes 4 et 5 n'est pas indifférent.** Annoncer une version avant d'avoir déposé
> son programme d'installation envoie les agents chercher un fichier qui n'existe pas.

`version.txt` se pose à côté de `connexion.config`, au même endroit — l'application le cherche
dans le dossier du fichier de connexion, sans réglage supplémentaire. Son modèle commenté est
`version.txt.modele`.

Tant que ce fichier n'existe pas, rien ne s'affiche et rien n'échoue : le dispositif est un
confort, jamais une condition de démarrage.

---

## 6. Vérifier ce qu'un poste emploie réellement

Trois moyens, du plus simple au plus détaillé :

- la **barre d'état** de la fenêtre principale affiche en permanence le serveur et la
  base en service ;
- l'écran **Connexion à la base de données** indique en plus **d'où vient** la chaîne :
  « fichier partagé », « copie locale (partage injoignable) », etc. ;
- en ligne de commande :

```powershell
.\Configurer-Connexion.ps1 -Afficher
```

La distinction compte : un administrateur qui croit lire le partage alors qu'il lit
une copie locale périmée modifierait un fichier sans aucun effet.

---

## 7. Dépannage

| Symptôme | Cause probable | Ce qu'il faut faire |
|---|---|---|
| « copie locale (… injoignable) » dans l'écran de connexion | Le partage ne répond pas | Vérifier le serveur de fichiers et les droits de lecture. Le poste travaille sur la dernière chaîne connue : ce n'est pas urgent |
| Le poste garde l'ancien serveur après un changement | L'application n'a pas été redémarrée, ou le réglage n'a pas été propagé | Fermer et rouvrir. Vérifier que la case « TOUS les postes » était cochée |
| « valeur par défaut (aucune configuration trouvée) » | Poste jamais installé, ou `wincompense.config` supprimé | Relancer l'installation, ou régler par l'écran de connexion |
| Écriture refusée sur le partage | Droits insuffisants | L'écriture est réservée à l'informatique et à l'administrateur. Rien n'a été modifié |
| « Le serveur est introuvable ou n'est pas accessible » alors que la barre du bas affiche `SRV-SQL01` | `SRV-SQL01` est le **nom d'exemple** de la documentation, pas votre serveur | Remplacer la ligne `SERVEUR=` par le nom réel. Voir « Trouver le nom exact du serveur » ci-dessous |
| Le nom du serveur est bon, la connexion échoue quand même | Instance nommée sans le service *SQL Server Browser*, TCP/IP désactivé, ou pare-feu | Démarrer *SQL Server Browser* ; activer TCP/IP dans *SQL Server Configuration Manager* ; ouvrir le port 1433 |
| L'application démarre mais aucune donnée n'apparaît | Bonne connexion, mauvaise base | Tester la connexion : l'écran affiche le nom de la base qui répond |
| Les exports Excel échouent | Excel absent du poste | Installer Excel. Les calculs et les écrans restent utilisables |

### Trouver le nom exact du serveur

Sur la machine qui héberge SQL Server, dans SQL Server Management Studio :

```sql
SELECT @@SERVERNAME
```

Ou, sans SSMS : *Services* Windows → chercher **SQL Server**.

| Ce que le service affiche | Ce qu'il faut écrire dans `SERVEUR=` |
|---|---|
| `SQL Server (MSSQLSERVER)` | `NOM-DU-SERVEUR` — instance par défaut, rien à ajouter |
| `SQL Server (SQLEXPRESS)` | `NOM-DU-SERVEUR\SQLEXPRESS` |
| La base est sur le poste lui-même | `.\SQLEXPRESS`, ou `.` pour une instance par défaut |

### Si la connexion échoue au démarrage

Le bouton **« Serveur... »** apparaît alors sur l'écran de connexion. Il ouvre le
réglage du serveur sans qu'il faille s'authentifier — impossible par définition quand
la base ne répond pas — et retente la lecture dès la fermeture.

Il reste masqué tant que tout va bien : régler le serveur n'est pas un geste quotidien.

---

## 8. Désinstallation

Panneau de configuration → *Programmes et fonctionnalités* → **Wincompense TCHAD**.

`%PROGRAMDATA%\Wincompense` **n'est pas supprimé** : une réinstallation retrouve ainsi
le serveur sans qu'on ait à le ressaisir. Pour repartir de zéro, supprimer ce dossier
à la main.

Le fichier partagé, lui, n'est jamais touché : il appartient à la banque, pas au poste.
