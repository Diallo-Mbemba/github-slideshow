# Wincompense TCHAD — plan de déploiement à la banque

Document de conduite, à suivre point par point. Chaque case se coche quand la vérification
associée a été **constatée**, pas seulement exécutée.

Trois intervenants : **INFO** (informatique de la banque), **COMPTA** (comptabilité et
compense), **PROJET** (l'équipe qui a construit l'application).

Les phases s'enchaînent dans l'ordre. La phase 5 porte un **go / no-go** : elle ne se
contourne pas.

---

## Phase 0 — Décisions préalables

Rien ne commence tant que ces points ne sont pas tranchés. Quatre d'entre eux bloquent la
mise en production.

| | Point | Qui tranche | Bloquant |
|---|---|---|---|
| ☐ | Format de `VALDT` dans le fichier core banking (`jj/mm/aaaa` retenu) | INFO / core banking | **oui** |
| ☐ | Usage du compte inter bancaire `381000101` pour l'écart d'arrondi global | COMPTA | **oui** |
| ☐ | Groupes Active Directory à rattacher aux rôles `wu_compense`, `wu_commercial`, `wu_admin` | INFO | **oui** |
| ☐ | Nom du serveur SQL de production et chemin du partage réseau | INFO | **oui** |
| ☐ | Durée de conservation réglementaire des historiques | Conformité | non |
| ☐ | Règle sur les lignes `TransactionType = "A"` du rapport de règlement | COMPTA | non |
| ☐ | Expiration périodique des mots de passe : oui / non, à quelle échéance | Sécurité | non |
| ☐ | Icône `.ico` de la banque | PROJET | non |

**Tranchés, pour mémoire :** la pièce d'une agence propre a la même structure que celle d'un
sous-agent, sa ligne de mouvement allant sur le compte courant WU ; les Accounts non
paramétrés ne sont pas comptabilisés.

---

## Phase 1 — La base de données

| | Étape | Qui |
|---|---|---|
| ☐ | Exécuter **`Scripts\00_InstallationComplete.sql`** sur le serveur de production | INFO |
| ☐ | Lire le **compte rendu** en fin de script : tables créées, rôles, et ce qui reste à faire | INFO |
| ☐ | Compléter la **PARTIE 5** du script avec vos groupes Active Directory, puis réexécuter cette partie | INFO |
| ☐ | Vérifier le **mode de récupération** de la base | INFO |
| ☐ | Si le mode est `FULL` : mettre en place une sauvegarde régulière du **journal de transactions** | INFO |
| ☐ | Mettre en place la **sauvegarde complète quotidienne** | INFO |
| ☐ | Alimenter `SystemeWU` : les neuf comptes comptables, validés un par un | COMPTA |
| ☐ | Alimenter `T_JourFerieWU` : fêtes musulmanes de l'année en cours et de la suivante | COMPTA |

> **Un seul script suffit.** `00_InstallationComplete.sql` crée la base, les dix tables, les
> trois rôles et leurs droits, puis rend compte de ce qu'il a fait. Il est rejouable : le
> relancer ne détruit rien et ne crée que ce qui manque.
>
> Il ne remplace pas les scripts `01_` à `10_`, qui restent disponibles pour une intervention
> ciblée — mais il évite l'ordre d'exécution, qui n'était pas indifférent : `09_Demandes.sql`
> ajoute la colonne `Fonction` sans laquelle la gestion des utilisateurs échoue, et les droits
> doivent être accordés en dernier, une fois toutes les tables créées.

> **Les neuf comptes de `SystemeWU` déterminent toute la pièce comptable.** Une erreur ici ne
> se voit pas à l'écran : elle se voit sur les comptes, après chargement.

---

## Phase 2 — Le référentiel des points de vente

| | Étape | Qui |
|---|---|---|
| ☐ | Charger `T_Pdv_SA` : code, désignation, taux, **compte de compensation**, **compte de commission**, code agence | COMPTA |
| ☐ | Charger `T_Pdv_EC` : code site, désignation, code agence Voyager | COMPTA |
| ☐ | Créer les groupes statistiques et y rattacher les sous-agents | COMPTA |
| ☐ | **Contrôle** : aucun sous-agent actif sans compte de compensation ni compte de commission | COMPTA |
| ☐ | **Contrôle** : aucun sous-agent sans groupe statistique, ou liste assumée de ceux qui restent | COMPTA |

> **Le contrôle des comptes n'est pas une formalité.** Un sous-agent auquel il manque l'un des
> deux comptes **n'est pas comptabilisé** : son activité n'apparaît ni dans la pièce, ni dans le
> fichier core banking, et reste à régulariser à la main. L'application le signale et chiffre le
> montant écarté, mais mieux vaut ne jamais en arriver là.

C'est la phase la plus longue du déploiement, et la plus souvent sous-estimée.

---

## Phase 3 — Le partage réseau et le poste pilote

| | Étape | Qui |
|---|---|---|
| ☐ | Créer le dossier partagé, par exemple `\\SRV-FICHIERS\Wincompense` | INFO |
| ☐ | Y poser `connexion.config` et renseigner la ligne `SERVEUR=` | INFO |
| ☐ | Droits : **lecture** pour les utilisateurs, **écriture** pour l'informatique seule | INFO |
| ☐ | Déposer le `setup.exe` dans `Setup\` sur le partage, puis y poser `version.txt` | INFO |
| ☐ | Compiler la solution en **Release** | PROJET |
| ☐ | Produire `Wincompense_Setup.exe` avec Inno Setup **6.4.x** | PROJET |
| ☐ | Installer sur **un seul poste**, celui de l'agent de compense principal | INFO |
| ☐ | Éprouver la commande d'installation silencieuse sur un second poste, avec `/SILENT` avant `/VERYSILENT` | INFO |
| ☐ | Vérifier que la barre d'état affiche le bon serveur et la bonne base | COMPTA |

> Qui peut écrire dans `connexion.config` commande la connexion de **tous** les postes.

Le détail de ces étapes est dans `LISEZMOI-Installation.md`.

---

## Phase 4 — Les comptes utilisateurs

| | Étape | Qui |
|---|---|---|
| ☐ | Premier démarrage : créer le premier administrateur | INFO + COMPTA |
| ☐ | Créer les agents de compense, le commercial, l'**inputer** et l'**authorizer** | Administrateur |
| ☐ | Vérifier que chacun change son mot de passe à la première connexion | Administrateur |
| ☐ | Éprouver le double regard : créer un sous-agent fictif en tant qu'inputer, l'autoriser en tant qu'authorizer, puis le supprimer | COMPTA |
| ☐ | Vérifier qu'un inputer ne peut pas autoriser sa propre demande | COMPTA |

> L'inputer et l'authorizer doivent être **deux personnes distinctes** : la base le refuse
> autrement. C'est le principe même du double regard.

---

## Phase 5 — Recette — go / no-go

C'est ici que le déploiement se joue.

| | Étape | Qui |
|---|---|---|
| ☐ | Choisir une **journée réelle déjà comptabilisée à la main**, et retrouver sa pièce manuelle | COMPTA |
| ☐ | Rejouer cette journée dans l'application, de l'étape 1 à l'étape 4 | COMPTA |
| ☐ | Comparer la pièce produite et la pièce manuelle, **ligne à ligne** | COMPTA |
| ☐ | Vérifier l'écart d'arrondi global et son affectation au compte d'attente | COMPTA |
| ☐ | Vérifier la liste des Accounts écartés, et que chacun s'explique | COMPTA |
| ☐ | Produire le fichier core banking : treize colonnes, numéro de lot, date de valeur | COMPTA |
| ☐ | **Charger ce fichier dans l'environnement de TEST du core banking** | INFO |
| ☐ | Faire confirmer le format de `VALDT` par le retour du chargement de test | INFO |
| ☐ | Rejouer une journée **atypique** : un lundi rattrapant le week-end | COMPTA |
| ☐ | Rejouer une journée comportant des annulations | COMPTA |
| ☐ | Rejouer une journée comportant un Account absent du référentiel | COMPTA |
| ☐ | Éditer le rapport d'activité sur une période connue et le confronter aux chiffres de la banque | COMPTA |

> **Jamais dans le core banking de production à ce stade.** Le fichier impacte réellement les
> comptes ; c'est en test qu'on découvre un format inattendu, pas après.

**Go / no-go.** Si la pièce diffère de la pièce manuelle sur autre chose qu'un arrondi au
FCFA, on ne passe pas à la phase 6. On corrige, et on recommence la recette entière.

---

## Phase 6 — Marche en parallèle

Deux à quatre semaines. La compense est faite **à la main et par l'application**, chaque jour.

| | Étape | Qui |
|---|---|---|
| ☐ | Chaque jour : produire les deux pièces et les comparer | COMPTA |
| ☐ | Ne charger dans le core banking que la **pièce manuelle** | COMPTA |
| ☐ | Tenir un **journal des écarts** : date, nature, cause, correction | COMPTA |
| ☐ | Expliquer chaque écart avant la bascule — aucun ne reste « à voir plus tard » | COMPTA + PROJET |

C'est la phase que les projets suppriment quand le calendrier presse. C'est aussi celle qui
évite les incidents comptables.

---

## Phase 7 — Bascule

| | Étape | Qui |
|---|---|---|
| ☐ | Constater **dix journées consécutives sans écart inexpliqué** | COMPTA |
| ☐ | Décision formelle de bascule, écrite et datée | COMPTA + INFO |
| ☐ | Installer l'application sur tous les postes concernés | INFO |
| ☐ | Former les agents : l'enchaînement numéroté 1 → 5, la lecture des alertes, le double regard | PROJET |
| ☐ | **Accompagner la première journée en production**, sur place | PROJET |
| ☐ | Première pièce chargée en production dans le core banking, sous surveillance | COMPTA + INFO |

---

## Phase 8 — Après la bascule

| | Étape | Qui | Échéance |
|---|---|---|---|
| ☐ | Vérifier la sauvegarde **par une restauration réelle** sur un serveur de test | INFO | première semaine |
| ☐ | Écrire la **procédure de secours** : que fait l'agent si l'application ne démarre pas à 8 h | PROJET + COMPTA | première semaine |
| ☐ | Conserver la procédure manuelle utilisable | COMPTA | six mois |
| ☐ | Désigner qui appelle qui, et sous quel délai | INFO | première semaine |
| ☐ | À chaque nouvelle version : déposer le setup **puis** mettre à jour `version.txt` | INFO | à chaque livraison |
| ☐ | Compléter `T_JourFerieWU` des fêtes musulmanes de l'année suivante | COMPTA | chaque année |
| ☐ | Vérifier la taille de la base et la croissance du journal de transactions | INFO | chaque trimestre |

> Une sauvegarde qui n'a jamais été restaurée n'est pas une sauvegarde : c'est une supposition.

---

## Ce qui reste du ressort du projet

| | Point |
|---|---|
| ☐ | Icône `.ico` de la banque, à intégrer au projet et au programme d'installation |
| ☐ | Rattachement des rôles SQL Server aux groupes AD, à écrire une fois le domaine connu |
| ☐ | Décision sur l'expiration des mots de passe, à implémenter le cas échéant |
| ☐ | Comptes `379100319` et `379200585` : à basculer de `ConstantesWU` vers `SystemeWU` si le plan comptable bouge |
