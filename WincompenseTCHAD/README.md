# WincompenseTCHAD

Automatisation de la comptabilisation Western Union J+1 (Tchad) — application Windows Forms
en VB.NET (.NET Framework 4.8 / Visual Studio 2019 / SQL Server Express).

## Ouverture du projet

1. Ouvrir `WincompenseTCHAD.sln` dans Visual Studio 2019.
2. Vérifier que le Framework cible du projet est bien **.NET Framework 4.8**.
3. Adapter la chaîne de connexion SQL Server dans `WincompenseTCHAD\App.config` si nécessaire
   (par défaut : `Server=.\SQLEXPRESS;Database=GWC_WINCOMPENSE_ETD;Integrated Security=True;`).
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
├── 01_CreateTables_GWC_WINCOMPENSE_ETD.sql
└── 02_DonneesExemple.sql
```

## Hypothèses métier retenues (à valider)

1. **InclureLigneReglement** : toutes les lignes du rapport de règlement sont incluses par défaut
   (y compris `TransactionType = "A"`), dès lors que l'Account est renseigné. Fonction isolée,
   volontairement simple, à affiner selon consigne métier ultérieure.
2. **Comptes INCONNU** : toujours affichés dans la grille de contrôle (jamais ignorés) ; pour la
   pièce comptable, traités par défaut comme une agence propre (100 % banque, compte courant WU),
   avec surlignage d'anomalie. *À confirmer.*
3. **Structure des écritures de la pièce comptable** : validée par rapprochement algébrique avec
   un exemple réel du classeur `PieceComptabilsationTchad.xlsx` (agence sous-agent "BOLOLO").
   Une seule ligne de mouvement net (`NetMouvement = PrincipalEnvoi+ChargeEnvoi+Taxes−PrincipalPaye`,
   Débit si positif) sur le `CompteCompense` du PDV, en contrepartie du compte courant WU pour la
   part nette bancaire ; commissions et taxes sont des lignes de crédit uniquement, sans ligne de
   débit miroir individuelle (voir commentaires détaillés dans `GenererPieceComptable`). Vérifié à
   l'unité près sur l'exemple disponible ; *le cas d'une agence propre "EC" reste à valider faute
   d'exemple de référence pour ce type de PDV.*
4. **Solde par Account** (grille de contrôle) = `PrincipalPaye − (PrincipalEnvoi + ChargeEnvoi + Taxes)`.
5. **Cohérence des dates** : la date du rapport d'activité (`txnDateLOC`) est comparée à celle du
   rapport de règlement, reconstituée depuis `SetDateLOCYear/Month/Day` (à défaut `RepDate`).
   Une divergence bloque le traitement. Si aucune date n'est exploitable d'un côté, la
   vérification est ignorée sans bloquer (avertissement affiché dans le StatusStrip).
6. **Compte d'attente `XXXXXXXXXX`** : valeur littérale provisoire (`ConstantesWU.CPT_ATTENTE`),
   à remplacer par le numéro de compte réel avant mise en production.
7. **Export Excel** : réalisé en liaison tardive (late binding, `Option Strict Off` isolé dans
   `PieceComptableService.vb`) afin de ne pas imposer de référence COM Excel obligatoire au
   projet. Toute la logique métier fonctionne sans Excel installé.

## Compatibilité avec les formats de rapport Western Union

Western Union a refondu le format de ses rapports (comparaison faite entre un rapport du
02/01/2021 et un rapport du 30/05/2026). **L'application traite indifféremment les deux
formats** : aucune colonne utilisée par les calculs n'a disparu, et les différences de contenu
sont absorbées automatiquement.

| Élément | Ancien format | Nouveau format | Traitement |
|---|---|---|---|
| `PayCountry` (pays local) | `TCHAD` | `CHAD` | Les deux graphies sont acceptées (`ConstantesWU.PaysPaiementLocal`) |
| Devise des montants LOC | `EUR` | `XAF` | Conversion pilotée par `LOCCurrencyCode` : ×655,957 si EUR, aucune conversion si XAF (`WUReportService.ObtenirFacteurConversion`) |
| Séparateur décimal | virgule (`1493,91`) | point (`-250000.0000`) | Détecté par position du symbole, jamais par une culture supposée (`ToDecimalSafe`) |
| Casse des colonnes | `Status`, `txnDateLOC` | `STATUS`, `TXNDATELOC` | `DataTable.Columns` est insensible à la casse |
| Colonne `STATUS` | absente | `S` / `C` / `W` | Les lignes annulées (`C`) sont exclues de l'agrégation (`InclureLigneActivite`) ; absence de la colonne = comportement historique |
| Format de date | `20210102` | `20260530` | `yyyyMMdd` analysé explicitement (`ConstantesWU.FormatsDateRapport`) |
| Date du rapport de règlement | colonnes `SetDateLOCYear/Month/Day` | idem | Reconstituée depuis ces colonnes (`ObtenirDateReglement`), ce qui rend enfin opérationnelle la validation croisée de la section 16 |
| Colonnes ajoutées | — | `Tax1REC`, `Tax2REC`, `Tax3REC`, `TotalChargesPAY`, `Assigned_Account_Name` | Ignorées par le lecteur (aucun impact) |

**Non-régression vérifiée** : relu par le code actuel, le rapport du 02/01/2021 restitue très
exactement les valeurs de référence du cahier des charges pour `AHB020200` (PrincipalEnvoi
5 966 388 ; PrincipalPaye 7 522 117,45 ; ChargeEnvoi 340 000 ; Taxes 82 075 ; CommissionPaiement
98 104,93).

### Les taxes détaillées confirment les taux du cahier des charges

Le nouveau rapport d'activité fournit le détail des taxes (`Tax1REC`, `Tax2REC`, `Tax3REC`,
dont la somme vaut exactement `TaxesREC`), là où le code les recalcule par application des taux.
Le rapprochement, effectué ligne à ligne sur le fichier du 30/05/2026 (95 lignes d'envoi, hors
annulations), **valide les formules du cahier des charges** :

| Taxe fournie | Formule du code | Lignes conformes | Écart total |
|---|---|---|---|
| `Tax1REC` | `ChargeEnvoi × 19,25 %` (TVA) | **95 / 95** | +20,50 FCFA |
| `Tax3REC` | `PrincipalEnvoi × 0,2 %` (TTA Envoi) | **95 / 95** | −1,79 FCFA |
| `Tax2REC` | `Taxes − TVA − TTAEnvoi` (Solde de taxes) | — | −18,71 FCFA |

Les écarts résiduels (une vingtaine de francs sur ~460 000 FCFA de taxes) ne sont que les
arrondis à l'unité pratiqués par Western Union. **Aucune modification des formules n'est donc
justifiée** : les remplacer par les valeurs fournies modifierait les règles du cahier des
charges pour un gain nul.

## Points restant à confirmer

- Structure des écritures pour une **agence propre "EC"** dans la pièce comptable (aucun exemple
  de référence de ce type disponible à ce jour ; seul un exemple sous-agent "SA" a pu être validé).
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
