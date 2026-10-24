# Wincompense TCHAD — Le déploiement silencieux

**Installer et mettre à jour les postes sans aller dans les bureaux.**

| | |
|---|---|
| Application | Wincompense TCHAD — version 1.1.0.0 |
| Destinataires | Informatique de la banque |
| Programme concerné | `Wincompense_Setup_1.1.0.0.exe` |

Ce document explique le **principe**, puis la **marche à suivre**, puis les **quatre pièges**
qui font échouer un déploiement de masse — chacun ayant sa cause dans le programme
d'installation lui-même, et donc sa parade.

---

## 1. Le principe

Le programme d'installation est le même pour tout le monde. Ce qui change, c'est la façon
de le lancer.

| | Ce qui se passe |
|---|---|
| **Lancé à la main** | L'assistant s'affiche : dix écrans, dont une page où le technicien saisit le serveur et le fichier partagé |
| **Lancé en silencieux** | **Aucun écran.** Les réponses viennent de la ligne de commande, ou de ce que le poste porte déjà |

Le mode silencieux n'est donc pas un mode dégradé : c'est le même travail, **avec les
réponses données d'avance**. C'est ce qui permet de le confier à une stratégie de groupe ou
à SCCM, et de mettre vingt postes à jour la nuit.

> **Ce que le mode silencieux ne fait pas :** il ne vérifie rien à votre place. Les contrôles
> de saisie de l'assistant — serveur vide, chaîne sans serveur, chemin de partage absent — ne
> s'appliquent pas, parce qu'il n'y a pas de saisie. **C'est l'informatique qui répond de ce
> qu'elle passe en paramètre.**

---

## 2. Deux cas, deux commandes — ne pas les confondre

### 2.1 Première installation d'un poste

Le poste ne porte encore aucune configuration : **les deux paramètres sont obligatoires.**

```
Wincompense_Setup_1.1.0.0.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART ^
    /SERVEUR="SRV-SQL01\SQLEXPRESS" ^
    /PARTAGE="\\SRV-FICHIERS\Wincompense\connexion.config" ^
    /LOG="\\SRV-FICHIERS\Wincompense\Journaux\%COMPUTERNAME%.log"
```

**Sans eux, le poste repart sur `.\SQLEXPRESS`** — c'est-à-dire sur aucun serveur. Il
s'installera, il démarrera, et il ne trouvera pas la base.

### 2.2 Mise à jour d'un poste déjà installé

**Aucun paramètre de connexion.**

```
Wincompense_Setup_1.1.0.0.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART ^
    /LOG="\\SRV-FICHIERS\Wincompense\Journaux\%COMPUTERNAME%.log"
```

L'installateur **relit ce que le poste porte déjà** — serveur, partage, et le mot de passe
SQL chiffré — et le réécrit à l'identique. C'est voulu : sans cela, une mise à jour lancée
sans `/SERVEUR` remettrait chaque poste sur `.\SQLEXPRESS` et lui retirerait son partage.
L'informatique croirait n'avoir changé que l'exécutable, et la banque entière perdrait sa
base un lundi matin.

> **Ne passez pas `/SERVEUR` sur une mise à jour si rien ne change.** C'est inutile, et une
> valeur erronée écraserait un réglage correct sur tous les postes du lot.

### 2.3 L'ordre dans lequel l'installateur se décide

| Rang | Ce qu'il prend | |
|---|---|---|
| 1 | Ce qui est passé en **ligne de commande** | `/SERVEUR`, `/PARTAGE` |
| 2 | Ce que le **poste porte déjà** | `CHAINE`, puis `SERVEUR` dans `wincompense.config` |
| 3 | La valeur par défaut | `.\SQLEXPRESS` — autant dire rien |

---

## 3. Les commutateurs, un par un

| Commutateur | Ce qu'il fait |
|---|---|
| `/VERYSILENT` | Aucune fenêtre, pas même la barre de progression |
| `/SILENT` | Affiche la barre de progression. **À employer pour l'essai**, jamais pour le lot |
| `/SUPPRESSMSGBOXES` | Répond à la place de l'opérateur aux boîtes de dialogue de l'installateur. **N'a d'effet qu'avec `/SILENT` ou `/VERYSILENT`** |
| `/NORESTART` | N'autorise aucun redémarrage du poste. Indispensable en déploiement de nuit |
| `/LOG="chemin"` | Écrit le journal détaillé de l'installation. **C'est lui qui vous dira pourquoi un poste a échoué** |
| `/SERVEUR="…"` | Le serveur SQL **ou** la chaîne de connexion complète — l'installateur les distingue seul, sur la présence d'un signe `=` |
| `/PARTAGE="…"` | Le chemin du fichier de connexion commun |

> **Les guillemets ne sont pas facultatifs** : un chemin UNC et un nom d'instance portent des
> caractères que l'interpréteur de commandes découperait. Le `^` en fin de ligne est la
> continuation de ligne de Windows — tout écrire sur une seule ligne revient au même.

---

## 4. Avant de diffuser : l'essai en trois temps

On ne lance jamais un lot de vingt postes sur une commande qu'on n'a pas vue fonctionner.

| | Étape | Ce qu'on regarde |
|---|---|---|
| 1 | Un poste témoin, en **`/SILENT`** | La barre de progression avance, aucune boîte de dialogue n'apparaît, l'installation se termine |
| 2 | Le même poste, désinstallé puis repris en **`/VERYSILENT`** | Le journal dit « terminé », et l'application démarre sur le **bon serveur** |
| 3 | **Trois postes** du lot réel | Les trois journaux sont identiques |

Alors seulement, le lot entier.

**Ce qu'il faut constater sur le poste témoin, et pas seulement supposer :**

- l'application est bien dans `C:\Program Files\Default Company Name\SetupWincompense\` ;
- la **barre d'état** affiche le serveur et la base attendus ;
- `C:\ProgramData\Wincompense\wincompense.config` porte le bon `SERVEUR=` et le bon
  `PARTAGE=`.

---

## 5. Les quatre pièges

### 5.1 Le poste n'a pas .NET Framework 4.8 — et le déploiement s'arrête net

L'installateur vérifie ce prérequis **avant tout**, et pose la question : « Poursuivre
malgré tout ? », la réponse par défaut étant **Non**.

Cette question-là **n'est pas supprimée par `/SUPPRESSMSGBOXES`** : les boîtes affichées par
le script d'installation ne le sont jamais. Sur un poste sans .NET 4.8, en `/VERYSILENT`,
une fenêtre s'ouvre donc **sur un poste que personne ne regarde**, et le déploiement y reste
suspendu.

> **La parade : ne jamais envoyer le setup sur un poste qui n'a pas .NET 4.8.** Posez-le en
> condition de déploiement — SCCM et Intune savent tester une clé de registre :
>
> ```
> HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full
>     Release (DWORD) >= 528040
> ```
>
> C'est exactement le test que fait l'installateur. Présent d'office depuis Windows 10 1903 ;
> le problème ne se pose que sur les postes anciens.

### 5.2 Le fichier partagé n'existe pas, ou le poste n'a pas le droit d'y écrire

À la fin de l'installation, l'installateur **crée le fichier partagé s'il n'existe pas**. Si
le dossier est injoignable ou interdit en écriture, il affiche une erreur — une boîte de
dialogue de plus, elle aussi non supprimée.

> **La parade : créer `connexion.config` une fois, à la main, avant le lot** (voir le manuel
> d'installation, section 4). Il existe alors déjà, et l'installateur n'y touche pas : c'est
> précisément la règle qui empêche la saisie d'un technicien de faire basculer toute la
> banque.

### 5.3 Le mot de passe SQL ne voyage pas

Si l'application se connecte par un **compte SQL Server**, chaque poste doit recevoir le mot
de passe **une fois**. Il est chiffré par Windows **pour cette machine** : il ne se propage
ni par le partage, ni par une copie de fichier d'un poste à l'autre.

| | Façon | À retenir |
|---|---|---|
| 1 | Déposer par script, **avant** l'installation, une ligne `MOTDEPASSE_CLAIR=…` dans `C:\ProgramData\Wincompense\wincompense.config` | L'application **le chiffre au premier démarrage et efface la ligne en clair** |
| 2 | Passer la chaîne complète, mot de passe compris, dans `/SERVEUR` | **Déconseillé** : la ligne de commande se retrouve dans les journaux de SCCM et dans le script de la GPO |

> En authentification **Windows**, rien de tout cela : aucun mot de passe ne circule.

### 5.4 L'application est ouverte pendant la mise à jour

L'installateur est réglé pour **fermer l'application** si elle tourne. Sur un poste où un
agent est en train de traiter une journée de compense, cela interrompt son travail au milieu.

> **La parade : déployer hors des heures de compense.** La compense se fait le matin, en
> J+1 ; un déploiement de nuit ne gêne personne.

---

## 6. Savoir quels postes sont à jour

### 6.1 La clé de registre — c'est elle qui sert de règle de détection

L'installation inscrit sa version dans la base de registre, sous l'identifiant du produit :

```
HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\
    {6FCC1274-03A5-5832-9B2D-CB6797179959}_is1
        DisplayVersion   = 1.1.0.0
        DisplayName      = Wincompense TCHAD
```

C'est la **règle de détection** à donner à SCCM ou à Intune : « installé si
`DisplayVersion` ≥ 1.1.0.0 ». C'est aussi ce qu'un script de démarrage doit tester pour ne
pas réinstaller un poste déjà à jour, nuit après nuit.

> **L'identifiant ne change jamais d'une version à l'autre.** C'est par lui que Windows
> reconnaît une mise à jour plutôt qu'un second produit installé à côté du premier.

### 6.2 Les journaux

Avec `/LOG="\\SRV-FICHIERS\Wincompense\Journaux\%COMPUTERNAME%.log"`, chaque poste dépose
son compte rendu dans un dossier unique. Le lendemain matin, un coup d'œil au dossier dit
combien de postes ont répondu — et lesquels manquent.

### 6.3 Les codes de retour

**`0` signifie succès.** Toute autre valeur est un échec : ne pas la traduire au jugé,
**ouvrir le journal du poste**, qui nomme la cause.

---

## 7. Par quel outil déployer

### 7.1 Stratégie de groupe (GPO)

> **Attention :** la fonction *Installation de logiciels* d'une GPO **ne prend que des
> paquets MSI**. `Wincompense_Setup.exe` n'en est pas un. Il faut donc passer par un
> **script de démarrage de l'ordinateur** — qui s'exécute avec les droits du système, ce que
> l'installation exige (`PrivilegesRequired=admin`).

Un script de démarrage, et non un script d'ouverture de session : l'agent qui ouvre sa
session n'est pas administrateur de son poste.

**Modèle à adapter** — il ne réinstalle que si la version est différente :

```bat
@echo off
set SOURCE=\\SRV-FICHIERS\Wincompense\Setup\Wincompense_Setup_1.1.0.0.exe
set ATTENDUE=1.1.0.0
set CLE=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6FCC1274-03A5-5832-9B2D-CB6797179959}_is1

for /f "tokens=3" %%v in ('reg query "%CLE%" /v DisplayVersion 2^>nul') do set POSEE=%%v
if "%POSEE%"=="%ATTENDUE%" exit /b 0

"%SOURCE%" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART ^
    /LOG="\\SRV-FICHIERS\Wincompense\Journaux\%COMPUTERNAME%.log"
exit /b %ERRORLEVEL%
```

Pour une **première** installation, ajouter `/SERVEUR` et `/PARTAGE` à la dernière commande.

### 7.2 SCCM / Intune

| Champ | Valeur |
|---|---|
| Commande d'installation | `Wincompense_Setup_1.1.0.0.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART` |
| Commande de désinstallation | `"C:\Program Files\Default Company Name\SetupWincompense\unins000.exe" /VERYSILENT` |
| Règle de détection | Clé de registre `…{6FCC1274-…}_is1`, valeur `DisplayVersion` ≥ `1.1.0.0` |
| Condition préalable | .NET Framework 4.8 — `Release` ≥ `528040` |
| Comportement d'installation | **Pour le système**, hors session utilisateur |
| Code de retour | `0` = succès |

---

## 8. Désinstaller silencieusement

```
"C:\Program Files\Default Company Name\SetupWincompense\unins000.exe" /VERYSILENT
```

En mode silencieux, la question « Supprimer aussi la configuration de ce poste ? » **n'est
pas posée**, et le dossier `C:\ProgramData\Wincompense` est **conservé** — c'est le
comportement le moins destructeur, et celui qu'on veut pour une réinstallation : le poste
retrouve son serveur et son mot de passe sans aucune ressaisie.

Le fichier partagé, lui, n'est jamais touché : il appartient à la banque, pas au poste.

---

## 9. Récapitulatif à cocher

| | À faire, dans cet ordre |
|---|---|
| ☐ | `connexion.config` existe sur le partage, et porte le bon serveur |
| ☐ | Les droits sont posés : lecture pour les utilisateurs, écriture pour l'informatique |
| ☐ | Le setup est déposé dans `Setup\` sur le partage |
| ☐ | Le dossier `Journaux\` existe et les postes peuvent y écrire |
| ☐ | La condition « .NET 4.8 » est posée dans l'outil de déploiement |
| ☐ | Essai sur un poste témoin en `/SILENT` — aucune boîte de dialogue |
| ☐ | Même poste en `/VERYSILENT` — l'application démarre sur le bon serveur |
| ☐ | Trois postes du lot réel, journaux comparés |
| ☐ | En compte SQL Server : le mot de passe est prévu, poste par poste |
| ☐ | L'heure du déploiement est **hors des heures de compense** |
| ☐ | Le lot est lancé |
| ☐ | Le lendemain : dossier des journaux relu, postes manquants repris |
