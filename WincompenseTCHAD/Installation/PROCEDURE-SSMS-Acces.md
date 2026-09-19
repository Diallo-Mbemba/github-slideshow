# Wincompense — donner accès au compte applicatif

**Document à remettre à l'informatique de la banque.**
Procédure dans SQL Server Management Studio, sans script.

Durée : environ deux minutes. Tout tient dans **une seule fenêtre** — celle des propriétés
du login, page *User Mapping*. Cette page crée l'utilisateur de base **et** lui donne son
rôle en même temps ; il n'y a pas deux opérations à enchaîner.

---

## Ce qu'il faut avoir sous la main

| | Élément | Exemple |
|---|---|---|
| 1 | Le nom de l'instance SQL Server | `SRV-SQL01` ou `SRV-SQL01\SQLEXPRESS` |
| 2 | Le nom du compte applicatif | celui communiqué avec la chaîne de connexion |
| 3 | Son mot de passe | **seulement si le login n'existe pas encore** |
| 4 | Le nom de la base | `GWC_WINCOMPENSE_ETD` |

Vous devez être connecté à SSMS avec un compte membre de **sysadmin** (ou de
*securityadmin* et *db_owner* sur cette base).

---

## Étape 0 — Vérification préalable : le mode d'authentification

*À ne faire que si le compte applicatif est un compte **SQL Server** (et non un compte
Windows du domaine).*

1. Dans l'**Explorateur d'objets**, clic droit sur le **nom du serveur**, tout en haut.
2. **Propriétés** (*Properties*).
3. Page **Sécurité** (*Security*), à gauche.
4. Sous *Authentification du serveur*, la case **Mode d'authentification SQL Server et
   Windows** (*SQL Server and Windows Authentication mode*) doit être cochée.

> Si elle ne l'est pas, la cocher, puis **redémarrer le service SQL Server** — sans ce
> redémarrage, le changement ne prend pas effet. Tant que l'instance n'accepte que Windows,
> aucun compte SQL ne pourra se connecter, quels que soient les droits qu'on lui donne.

---

## Étape 1 — Retrouver le login

1. Dans l'**Explorateur d'objets**, déplier le serveur.
2. Déplier **Sécurité** (*Security*) — celui du **serveur**, tout en haut, pas celui d'une
   base.
3. Déplier **Connexions** (*Logins*).
4. Chercher le compte applicatif dans la liste.

**S'il est là** → passer directement à l'**étape 3**.
**S'il n'est pas là** → étape 2.

---

## Étape 2 — Créer le login *(seulement s'il n'existe pas)*

1. Clic droit sur **Connexions** (*Logins*) → **Nouvelle connexion…** (*New Login…*).
2. Page **Général** (*General*) :

| Champ | Valeur |
|---|---|
| Nom de la connexion | le nom du compte applicatif |
| Type d'authentification | **Authentification SQL Server** |
| Mot de passe / Confirmer | celui fourni avec la chaîne de connexion |
| Base de données par défaut | `GWC_WINCOMPENSE_ETD` |

3. **Trois cases à décocher**, sous le mot de passe :

| Case | Pourquoi la décocher |
|---|---|
| **L'utilisateur doit changer de mot de passe à la prochaine connexion** | **Impérative.** Une application ne sait pas changer un mot de passe : elle serait refusée à chaque démarrage, sans pouvoir rien y faire |
| Le mot de passe expire (*Enforce password expiration*) | Le jour de l'expiration, tous les postes s'arrêtent en même temps, sans préavis |
| Appliquer la stratégie de mot de passe (*Enforce password policy*) | À décocher **seulement** si le mot de passe fourni ne satisfait pas la stratégie du domaine et que la banque ne souhaite pas le changer |

4. **Ne pas cliquer sur OK tout de suite** — enchaîner sur l'étape 3, qui est une autre page
   de la même fenêtre.

---

## Étape 3 — Ouvrir la base au compte et lui donner son rôle

*Depuis la fenêtre de l'étape 2, ou par clic droit sur le login existant →
**Propriétés** (*Properties*).*

1. Dans le volet de gauche, cliquer sur **Mappage de l'utilisateur** (*User Mapping*).
2. Dans le tableau du **haut**, trouver la ligne **`GWC_WINCOMPENSE_ETD`** et **cocher la
   case** de la colonne *Mapper* (*Map*), à gauche.

   > La colonne *Utilisateur* (*User*) se remplit alors toute seule avec le nom du compte.
   > C'est cela qui crée l'utilisateur dans la base — il n'y a rien d'autre à faire pour lui.

3. **La ligne `GWC_WINCOMPENSE_ETD` restant sélectionnée**, regarder le tableau du **bas**,
   intitulé *Rôles de base de données pour : GWC_WINCOMPENSE_ETD*
   (*Database role membership for: GWC_WINCOMPENSE_ETD*).
4. Y cocher **`wu_admin`**.

   > `public` reste coché : c'est le comportement normal de SQL Server, ne pas y toucher.

5. **OK**.

### Le choix du rôle

| Rôle | Quand le choisir |
|---|---|
| **`wu_admin`** | **Le cas courant** : un compte unique employé par tous les postes. Il doit porter la réunion des droits de tous les postes |
| `wu_compense` | Un compte par agent, poste de compense |
| `wu_commercial` | Un compte par agent, saisie des points de vente |

Un compte par agent vaut mieux : les trois rôles retrouvent alors leur utilité, et un accès
direct à la base — par SSMS, hors de l'application — reste borné au métier réel de chacun.
Si la banque ne fournit qu'un compte, c'est `wu_admin`.

> **Si aucun rôle `wu_…` n'apparaît dans la liste du bas**, c'est que la base n'a pas encore
> reçu les scripts d'installation. Exécuter `Scripts\00_InstallationComplete.sql`, puis
> reprendre à l'étape 3.

---

## Étape 4 — Vérifier

1. Clic droit sur le login → **Propriétés** → **Mappage de l'utilisateur**.
2. La ligne `GWC_WINCOMPENSE_ETD` est cochée, et `wu_admin` l'est aussi dans le tableau du
   bas.

C'est tout. Côté application, l'agent ouvre *Sécurité → Connexion à la base de données* et
clique sur **Tester la connexion** : la réponse attendue est

> Connexion réussie — base GWC_WINCOMPENSE_ETD, ouverte au nom de …

---

## Si cela ne marche toujours pas

L'application traduit les refus de SQL Server et indique quoi faire. Les trois messages
possibles, et ce qu'ils signifient :

| Message | Ce qui manque | Étape à reprendre |
|---|---|---|
| *Login failed for user* (18456) | Le login, ou le mot de passe, ou le mode mixte | 0 et 2 |
| *Cannot open database* (4060) | La case de la ligne `GWC_WINCOMPENSE_ETD` n'est pas cochée | 3, point 2 |
| *SELECT permission was denied* (229) | Le rôle n'est pas coché | 3, point 4 |

Trois niveaux, qu'on confond facilement : le **login** ouvre la porte du bâtiment,
l'**utilisateur de base** celle du bureau, le **rôle** dit ce qu'on a le droit d'y faire.
Il faut les trois — et la page *User Mapping* donne les deux derniers d'un seul geste.

---

## Pour qui préfère le T-SQL

La même chose, en trois ordres : `Scripts\12_AccesCompteApplicatif.sql`. Une seule ligne à
renseigner, celle du nom du compte.
