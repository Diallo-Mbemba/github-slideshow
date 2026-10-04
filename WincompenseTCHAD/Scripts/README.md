# Un seul script à exécuter

## `00_InstallationComplete.sql`

C'est le seul. Tous les autres fichiers numérotés de ce dossier sont la **documentation**
de ce qu'il fait, découpée par sujet : ils expliquent, ils ne s'exécutent pas.

**À exécuter après chaque livraison de l'application.** Il est **rejouable** : il crée ce
qui manque, complète ce qui est incomplet, et ne touche à rien d'autre. Aucune donnée n'est
écrasée — ni un compte comptable saisi par la Direction Comptable, ni un libellé, ni une
option. Le relancer sur une base déjà installée est sans risque.

Après l'exécution, lire l'onglet **Messages**, puis le **compte rendu de la partie 6** : il
dit ce qui a été fait et ce qui reste à faire.

---

## Les deux exceptions, à l'installation seulement

Ils demandent une valeur que `00` ne peut pas deviner, et qu'il ne faut pas inventer.
Ce sont des gestes d'installation, faits **une fois** :

| Fichier | À compléter avant de l'exécuter |
|---|---|
| `11_AccesUtilisateurs.sql` | les comptes ou groupes **Active Directory** de la banque |
| `12_AccesCompteApplicatif.sql` | le nom du **compte SQL Server** fourni par la banque |

---

## Ce que `00` ne fait pas

- **Aucun sous-agent, aucune agence.** Le référentiel se charge depuis l'application.
  Les données d'exemple de `02_DonneesExemple.sql` sont volontairement exclues : elles
  n'ont leur place que sur un environnement de test.
- **Aucun compte utilisateur.** Le premier administrateur se crée au premier démarrage de
  l'application, qui demande confirmation de la base visée.
- **Aucun login SQL Server.** Il faudrait un mot de passe, qui n'a pas sa place dans un
  fichier qui circule.
- **Ni le mode de récupération, ni la sauvegarde.** À vérifier séparément : une base en
  mode `FULL` sans sauvegarde du journal finit par remplir le disque.

---

## Pourquoi les autres fichiers restent là

Chacun porte, en tête, l'explication de la règle métier qu'il installe — pourquoi ce
compte et pas un autre, pourquoi cette contrainte, ce que l'application fait si la table
manque. C'est cette explication qui a de la valeur, pas l'ordre `CREATE`. Les supprimer
pour « faire propre » reviendrait à jeter la seule documentation écrite du schéma.

Si l'un d'eux crée un objet que `00` ne reprend pas, la banque devrait exécuter deux
fichiers au lieu d'un : c'est exactement ce que le contrôleur `verif_installation.py`
vérifie avant chaque livraison, objet par objet.
