# WincompenseTCHAD

Automatisation de la comptabilisation Western Union J+1 (Tchad) — application Windows Forms
en VB.NET (.NET Framework 4.8 / Visual Studio 2019 / SQL Server Express).

## Ouverture du projet

1. Ouvrir `WincompenseTCHAD.sln` dans Visual Studio 2019.
2. Vérifier que le Framework cible du projet est bien **.NET Framework 4.8**.
3. Adapter la chaîne de connexion SQL Server dans `WincompenseTCHAD\App.config` si nécessaire
   (par défaut : `Server=.\SQLEXPRESS;Database=GWC_WINCOMPENSE;Integrated Security=True;`).
4. Exécuter les scripts SQL du dossier `Scripts\` (dans l'ordre numéroté) sur l'instance
   `.\SQLEXPRESS` pour créer les tables `T_Pdv_SA` / `T_Pdv_EC` et les données de test.
5. Compiler et lancer (F5).

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
    ├── Models/CalculWU.vb                  ' Classe métier par Account
    ├── Services/
    │   ├── WUReportService.vb              ' Lecture fichiers, parsing, agrégation, dates
    │   ├── WURepository.vb                 ' Accès SQL Server (T_Pdv_SA / T_Pdv_EC)
    │   ├── WUCalculationService.vb         ' Formules, répartition, arrondi
    │   └── PieceComptableService.vb        ' Grille de contrôle, pièce comptable, équilibrage, export Excel
    └── Forms/
        ├── FrmCompensationWU.vb            ' Orchestration des événements uniquement
        ├── FrmCompensationWU.Designer.vb
        └── FrmCompensationWU.resx

Scripts/
├── 01_CreateTables_GWC_WINCOMPENSE.sql
└── 02_DonneesExemple.sql
```

## Hypothèses métier retenues (à valider)

1. **InclureLigneReglement** : toutes les lignes du rapport de règlement sont incluses par défaut
   (y compris `TransactionType = "A"`), dès lors que l'Account est renseigné. Fonction isolée,
   volontairement simple, à affiner selon consigne métier ultérieure.
2. **Comptes INCONNU** : toujours affichés dans la grille de contrôle (jamais ignorés) ; pour la
   pièce comptable, traités par défaut comme une agence propre (100 % banque, compte courant WU),
   avec surlignage d'anomalie. *À confirmer.*
3. **Structure des écritures de la pièce comptable** : reconstruite selon une logique de partie
   double auto-cohérente (voir commentaires dans `PieceComptableService.GenererPieceComptable`),
   faute d'accès au classeur `PieceComptabilsationTchad.xlsx`. *À valider contre le modèle réel.*
4. **Solde par Account** (grille de contrôle) = `PrincipalPaye − (PrincipalEnvoi + ChargeEnvoi + Taxes)`.
5. **Colonne de date côté règlement** : si absente, la vérification de cohérence de date n'est pas
   bloquante (avertissement affiché dans le StatusStrip).
6. **Compte d'attente `XXXXXXXXXX`** : valeur littérale provisoire (`ConstantesWU.CPT_ATTENTE`),
   à remplacer par le numéro de compte réel avant mise en production.
7. **Export Excel** : réalisé en liaison tardive (late binding, `Option Strict Off` isolé dans
   `PieceComptableService.vb`) afin de ne pas imposer de référence COM Excel obligatoire au
   projet. Toute la logique métier fonctionne sans Excel installé.

## Points restant à confirmer

- Détail exact des écritures du modèle `PieceComptabilsationTchad.xlsx`.
- Nom exact de la colonne de date dans le rapport de règlement.
- Mode d'authentification SQL Server réel en production (actuellement : Windows intégré).
- Règle définitive de traitement des lignes `TransactionType = "A"` du rapport de règlement.
- Traitement définitif souhaité des Accounts `INCONNU` dans la pièce comptable.

## Test de référence (section 17)

Les scripts `Scripts\02_DonneesExemple.sql` paramètrent :

- `AHB020200` comme sous-agent (`T_Pdv_SA`, `Taux = 0.70`) ;
- `AHB020057` comme agence propre Ecobank (`T_Pdv_EC`).

Pour valider le parsing/l'agrégation, charger un rapport d'activité contenant `AHB020200` à la
date du 02/01/2021 et vérifier dans `dgvControle` que les agrégats retrouvent approximativement
les valeurs de référence indiquées dans le cahier des charges (Principal Envoi ≈ 5 966 388,
Principal Payé ≈ 7 522 117,45, Commission Envoi = 69 700, Commission Paiement ≈ 98 105, etc.).
Pour `AHB020057`, vérifier qu'aucune commission sous-agent n'est générée (100 % à la banque).
