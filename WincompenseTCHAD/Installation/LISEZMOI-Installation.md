# Wincompense TCHAD — installation et changement de serveur

Ce dossier contient tout ce qu'il faut pour déployer l'application sur les postes,
et pour changer de serveur ensuite sans y revenir.

| Fichier | Rôle |
|---|---|
| `Wincompense.iss` | Script Inno Setup produisant `WincompenseTCHAD_Setup.exe` |
| `connexion.config.modele` | Modèle du fichier à poser sur le partage réseau |
| `Configurer-Connexion.ps1` | Changement de serveur en ligne de commande |
| `LISEZMOI-Installation.md` | Ce document — installation technique et changement de serveur |
| `PLAN-DEPLOIEMENT.md` | Les huit phases du déploiement à la banque, à cocher point par point |

---

## 1. Le principe : la connexion ne vit pas avec l'application

La banque change souvent de serveur. Tant que la chaîne de connexion vivait dans
`WincompenseTCHAD.exe.config`, en changer imposait une tournée dans les bureaux :
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

**L'authentification est celle de Windows.** Aucun mot de passe ne circule, et c'est
précisément ce qui permet de poser la configuration sur un partage lisible par tous.
Les droits d'accès à la base sont donnés par `Scripts\08_RolesSQLServer.sql`.

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
   Vérifier ensuite que `WincompenseTCHAD\WincompenseTCHAD\bin\Release\WincompenseTCHAD.exe`
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

Lancer `WincompenseTCHAD_Setup.exe` **en tant qu'administrateur**. L'assistant demande :

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
> un mot de passe y serait donc lisible en clair par tous. Si la chaîne fournie contient
> `User ID` / `Password`, demandez la version en authentification Windows
> (`Integrated Security=True`). L'application refuse de propager sur le partage une chaîne
> portant un mot de passe, et fait confirmer si le réglage ne vaut que pour un poste.

### Ce qui reste à faire côté base

Changer de serveur ne déplace pas les données. Sur le nouveau serveur, il faut :

1. restaurer ou recréer la base `GWC_WINCOMPENSE_ETD` — les scripts sont dans
   `Scripts\`, exécutés dans l'ordre `01_` à `10_` ;
2. rejouer `Scripts\08_RolesSQLServer.sql` pour redonner leurs droits aux
   utilisateurs Windows.

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
