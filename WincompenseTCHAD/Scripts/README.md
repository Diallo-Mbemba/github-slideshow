# Ce dossier ne contient que ce qui s'exécute

```
Scripts\
├── 00_InstallationComplete.sql    ← LE SEUL À EXÉCUTER
├── 11_AccesUtilisateurs.sql       ← uniquement en authentification Windows
└── Documentation\                 ← les 22 autres : ils expliquent, ils ne s'exécutent pas
```

---

## `00_InstallationComplete.sql`

C'est le seul. **À exécuter après chaque livraison de l'application.**

Il est **rejouable** : il crée ce qui manque, complète ce qui est incomplet, et ne touche à
rien d'autre. Aucune donnée n'est écrasée — ni un compte comptable saisi par la Direction
Comptable, ni un libellé, ni une option. Le relancer sur une base déjà installée est sans
risque, et c'est la façon normale de rattraper une base en retard.

Après l'exécution, lire l'onglet **Messages**, puis le **compte rendu de la partie 6** : il
dit ce qui a été fait et ce qui reste à faire.

---

## À la première installation : une ligne à vérifier, dans `00`

Dans la **partie 5**, en bas du fichier :

```sql
DECLARE @compteApplicatif SYSNAME = N'etdwincompense';
```

C'est le nom du compte SQL Server avec lequel l'application se connecte. S'il porte un
autre nom chez la banque, **on remplace celui-là, et c'est tout** : la partie 5 crée
l'utilisateur dans la base et lui accorde son rôle.

Elle ne crée pas le **login**, et ne le peut pas : il faudrait son mot de passe, qui n'a
pas sa place dans un fichier qui circule. Le login est créé par la banque. S'il manque, la
partie 5 s'arrête proprement et affiche la commande `CREATE LOGIN` à passer.

---

## Le seul cas où un deuxième fichier est nécessaire

**L'authentification Windows.** Si les postes se connectent avec leur compte de domaine au
lieu d'un compte SQL Server, il faut en plus `11_AccesUtilisateurs.sql` : il porte les
comptes ou groupes **Active Directory** de la banque, que `00` ne peut pas deviner. À
compléter, puis exécuter — **une fois**, à l'installation, jamais à une livraison.

C'est la seule raison pour laquelle il n'est pas rangé avec la documentation.

---

## En résumé

| Situation | À exécuter |
|---|---|
| Chaque livraison | `00` |
| Première installation, compte SQL Server | `00` (vérifier la ligne `@compteApplicatif`) |
| Première installation, authentification Windows | `00`, puis `11` complété |

Jamais trois.

---

## Ce que `00` ne fait pas

- **Aucun sous-agent, aucune agence.** Le référentiel se charge depuis l'application.
  Les données d'exemple de `Documentation\02_DonneesExemple.sql` sont volontairement
  exclues : elles n'ont leur place que sur un environnement de test.
- **Aucun compte utilisateur.** Le premier administrateur se crée au premier démarrage de
  l'application, qui demande confirmation de la base visée.
- **Aucun login SQL Server.** Il faudrait un mot de passe, qui n'a pas sa place dans un
  fichier qui circule.
- **Ni le mode de récupération, ni la sauvegarde.** À vérifier séparément : une base en
  mode `FULL` sans sauvegarde du journal finit par remplir le disque.

---

## `Documentation\` — pourquoi ces fichiers restent là

Chacun porte, en tête, l'explication de la règle métier qu'il installe — pourquoi ce
compte et pas un autre, pourquoi cette contrainte, ce que l'application fait si la table
manque. C'est cette explication qui a de la valeur, pas l'ordre `CREATE`. Les supprimer
pour « faire propre » reviendrait à jeter la seule documentation écrite du schéma.

**Ils ne s'exécutent pas.** Un seul d'entre eux peut encore l'être, dans un cas précis :
`Documentation\12_AccesCompteApplicatif.sql`, qui est la partie 5 de `00` sous forme
autonome, et qui sert à rattacher **plusieurs** comptes — un par agent, avec des rôles
différents (`wu_compense`, `wu_commercial`), une exécution par compte.

Si l'un d'eux crée un objet que `00` ne reprend pas, la banque devrait exécuter deux
fichiers au lieu d'un : c'est exactement ce que le contrôleur `verif_installation.py`
vérifie avant chaque livraison, objet par objet.
