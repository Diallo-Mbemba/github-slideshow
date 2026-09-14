# Wincompense TCHAD — installation et changement de serveur

Ce dossier contient tout ce qu'il faut pour déployer l'application sur les postes,
et pour changer de serveur ensuite sans y revenir.

| Fichier | Rôle |
|---|---|
| `Wincompense.iss` | Script Inno Setup produisant `WincompenseTCHAD_Setup.exe` |
| `connexion.config.modele` | Modèle du fichier à poser sur le partage réseau |
| `Configurer-Connexion.ps1` | Changement de serveur en ligne de commande |
| `LISEZMOI-Installation.md` | Ce document |

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

1. **Compiler en Release** : Visual Studio → *Générer* → *Configuration Release* →
   *Générer la solution*. Vérifier que `WincompenseTCHAD\bin\Release\WincompenseTCHAD.exe`
   existe et porte la bonne date.
2. **Installer Inno Setup 6** (gratuit) : <https://jrsoftware.org/isdl.php>
3. Ouvrir `Installation\Wincompense.iss` dans Inno Setup, puis **Build → Compile**
   (Ctrl+F9).
4. Le programme d'installation apparaît dans `Installation\Sortie\`.

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

- le **serveur SQL Server** — par exemple `SRV-SQL01` ou `SRV-SQL01\SQLEXPRESS` ;
- le **chemin du fichier partagé** — par exemple `\\SRV-FICHIERS\Wincompense\connexion.config`.

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
| L'application démarre mais aucune donnée n'apparaît | Bonne connexion, mauvaise base | Tester la connexion : l'écran affiche le nom de la base qui répond |
| Les exports Excel échouent | Excel absent du poste | Installer Excel. Les calculs et les écrans restent utilisables |

---

## 8. Désinstallation

Panneau de configuration → *Programmes et fonctionnalités* → **Wincompense TCHAD**.

`%PROGRAMDATA%\Wincompense` **n'est pas supprimé** : une réinstallation retrouve ainsi
le serveur sans qu'on ait à le ressaisir. Pour repartir de zéro, supprimer ce dossier
à la main.

Le fichier partagé, lui, n'est jamais touché : il appartient à la banque, pas au poste.
