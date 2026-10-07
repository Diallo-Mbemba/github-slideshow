/*
    =========================================================================
    Narrative : les libellés par nature de mouvement, et le journal
    =========================================================================

    DEUX TABLES, ET UNE LIGNE D'OPTION.

      1. T_NarrativeNatureWU    : le libellé de chaque nature de mouvement.
      2. T_JournalParametreWU   : qui a changé quoi, quand, et ce qu'il y avait avant.
      3. L'option NARRATIVE_MODE, dans T_ParametreWU : laquelle des deux façons s'applique
         À LA PIÈCE — le fichier core banking, lui, ne dépend pas de ce mode.

    LE MODE, ET CE QU'IL CHANGE

      PAR_NATURE  — chaque ligne de la PIÈCE porte le libellé de sa nature : mouvement,
                    contrepartie, commissions, taxes, écart d'arrondi. C'est la forme de la
                    pièce manuelle de la banque, et c'est le DÉFAUT.
      GLOBAL      — les douze lignes d'un point de vente portent LE MÊME libellé, celui du
                    modèle global (option NARRATIVE_MODELE).

    LE FICHIER CORE BANKING NE DÉPEND PAS DE CE MODE. Sa colonne ADDLTEXT porte TOUJOURS la
    narrative unique du point de vente, celle du modèle global — rectificatif de la banque du
    06/10/2026. La pièce conservée transporte donc les deux textes : son libellé par nature
    dans la colonne Libelle, et la narrative du point de vente dans la colonne Narratif.

    BASCULER NE DEMANDE AUCUNE SAISIE. Une nature ABSENTE de T_NarrativeNatureWU suit SON
    LIBELLÉ HISTORIQUE, celui que le code porte en constante : la table est donc créée VIDE,
    et une base installée et laissée telle quelle rend exactement la pièce que la banque
    connaît. Elle n'écrit dans cette table que ce qu'elle veut changer.

    L'ÉCART D'ARRONDI FAIT EXCEPTION, ET DANS LES DEUX MODES

    Sa ligne est posée APRÈS la pièce, pour absorber la différence globale : elle ne se
    rattache ni à un point de vente ni à une période, les deux repères du modèle global. Lui
    appliquer ce modèle la réduirait à « LD WU ACTIVITE ». Elle garde donc son texte propre —
    « LD ECART D'ARRONDI - COMPTE INTER BANCAIRE », valeur par défaut du code — que la banque
    peut éditer comme les autres.

    LE JOURNAL NE SE RÉÉCRIT PAS

    Aucun droit d'UPDATE ni de DELETE n'est accordé dessus, à personne, pas même à wu_admin.
    Un journal qu'on peut corriger ne prouve rien. Il est écrit DANS LA MÊME TRANSACTION que
    le changement qu'il décrit : pas de changement sans trace, et pas de trace sans
    changement — la seconde serait pire, elle accuserait quelqu'un d'une modification qui n'a
    pas eu lieu.

    POURQUOI T_NarrativeNatureWU ACCEPTE LE DELETE, ELLE

    Remettre une nature sur le modèle global, c'est EFFACER sa ligne. Garder une ligne vide
    donnerait une nature personnalisée à chaîne vide, c'est-à-dire des écritures sans libellé
    le jour où le mode bascule. L'effacement est donc le geste normal de cet écran, et non une
    opération de maintenance.

    SI CE SCRIPT N'EST PAS EXÉCUTÉ, L'APPLICATION FONCTIONNE

    Le mode reste GLOBAL, le mode « par nature » est proposé grisé avec le nom de ce script,
    et le paramétrage s'enregistre sans laisser de trace — l'écran le dit. Une absence n'est
    jamais une interdiction.

    À exécuter APRÈS 18 et 22 (ou après 00), sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. Les libellés par nature de mouvement
--
--    La clé est le CODE de la nature, une chaîne, et non l'entier de l'énumération du code
--    applicatif. Un entier rendrait la table illisible ici, et surtout : renuméroter
--    l'énumération réaffecterait silencieusement les libellés saisis à d'autres lignes de la
--    pièce. Les codes attendus sont ceux de NaturesMouvementWU.Code :
--
--        MOUVEMENT                     COMMISSION_TRANSFERT_SA     TVA
--        COMPTE_COURANT                COMMISSION_PAIEMENT_SA      TTA_ENVOI
--        COMMISSION_TRANSFERT_BANQUE   COMMISSION_ENVOI_SA         TTA_RECEPTION
--        COMMISSION_PAIEMENT_BANQUE    IMPOTS_TAXE_ENVOI           ECART_ARRONDI
--        COMMISSION_ENVOI_BANQUE
--
--    Un code que l'application ne connaît pas est IGNORÉ à la lecture, et la ligne concernée
--    reprend le modèle global : une saisie à la main ne doit pas empêcher la pièce de sortir.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_NarrativeNatureWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_NarrativeNatureWU
    (
        Nature              NVARCHAR(40)    NOT NULL,

        -- Le gabarit, avec ses repères : {AGENCE}, {ACCOUNT}, {CODE_AGENCE}, {PERIODE}.
        -- Même largeur que T_ParametreWU.Valeur : c'est le même genre de texte.
        Modele              NVARCHAR(255)   NOT NULL,

        DateModification    DATETIME        NULL,
        ModifiePar          NVARCHAR(50)    NULL,

        -- Une ligne sans modèle serait une nature « personnalisée » à vide, c'est-à-dire des
        -- écritures sans libellé. L'absence de ligne est la façon de dire « modèle global ».
        CONSTRAINT CK_T_NarrativeNatureWU_Modele CHECK (LEN(LTRIM(RTRIM(Modele))) > 0),

        CONSTRAINT PK_T_NarrativeNatureWU PRIMARY KEY (Nature)
    );

    PRINT 'Table T_NarrativeNatureWU créée (vide : toutes les natures suivent le modèle global).';
END
ELSE
    PRINT 'Table T_NarrativeNatureWU déjà présente : création ignorée.';
GO

-- Droits posés IMMÉDIATEMENT après la création, et non regroupés en fin de script : une
-- erreur au milieu laisserait sinon une table sans droits, que l'application voit mais ne
-- peut pas lire. C'est la leçon du script 08.
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.T_NarrativeNatureWU TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.T_NarrativeNatureWU TO wu_commercial');
GO
-- DELETE compris, et c'est voulu : remettre une nature sur le modèle global, c'est effacer
-- sa ligne. Voir l'en-tête de ce script.
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.T_NarrativeNatureWU TO wu_admin');
GO

-- =========================================================================
-- 2. Le journal du paramétrage
--
--    Il couvre TOUT le paramétrage de narrative sous une seule clé :
--
--        NARRATIVE_MODELE        le modèle global
--        NARRATIVE_MODE          le mode
--        NATURE_<CODE>           le libellé d'une nature
--
--    Le préfixe NATURE_ évite qu'une nature nommée TVA entre en collision avec une option
--    qui s'appellerait ainsi un jour.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_JournalParametreWU') IS NULL
BEGIN
    CREATE TABLE dbo.T_JournalParametreWU
    (
        Id                  BIGINT          IDENTITY(1,1) NOT NULL,

        Cle                 NVARCHAR(60)    NOT NULL,

        -- NULL, et non chaîne vide, quand la clé n'avait pas de valeur avant : le journal
        -- distingue « il n'y avait rien » de « il y avait une chaîne vide ».
        AncienneValeur      NVARCHAR(255)   NULL,
        NouvelleValeur      NVARCHAR(255)   NULL,

        ModifiePar          NVARCHAR(50)    NULL,
        Poste               NVARCHAR(100)   NULL,

        DateModification    DATETIME        NOT NULL DEFAULT (GETDATE()),

        CONSTRAINT PK_T_JournalParametreWU PRIMARY KEY (Id)
    );

    -- La consultation se fait toujours du plus récent au plus ancien, et toujours bornée.
    CREATE INDEX IX_T_JournalParametreWU_Date
        ON dbo.T_JournalParametreWU (DateModification DESC, Id DESC);

    PRINT 'Table T_JournalParametreWU créée.';
END
ELSE
    PRINT 'Table T_JournalParametreWU déjà présente : création ignorée.';
GO

-- TOUT LE MONDE LIT, SEUL L'ADMINISTRATEUR AJOUTE, PERSONNE NE CORRIGE.
--
-- Pas de GRANT UPDATE, pas de GRANT DELETE, pour aucun rôle : un journal qu'on peut
-- réécrire ne prouve rien. Purger deux ans d'historique restera un geste du DBA, fait
-- sciemment avec ses propres droits, et non un clic dans l'application.
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_compense')
    EXEC('GRANT SELECT ON dbo.T_JournalParametreWU TO wu_compense');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_commercial')
    EXEC('GRANT SELECT ON dbo.T_JournalParametreWU TO wu_commercial');
GO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE type = 'R' AND name = N'wu_admin')
    EXEC('GRANT SELECT, INSERT ON dbo.T_JournalParametreWU TO wu_admin');
GO

-- =========================================================================
-- 3. L'option : laquelle des deux façons s'applique à la pièce
--
--    PAR_NATURE par défaut, c'est-à-dire la pièce telle que la banque la connaît. La valeur
--    existante n'est JAMAIS écrasée ici : rejouer ce script ne fait pas rebasculer la banque.
--
--    LE RECTIFICATIF DU 06/10/2026 — qui ramène de GLOBAL à PAR_NATURE les bases où personne
--    n'a touché au paramètre — est dans Scripts\00_InstallationComplete.sql, le seul script
--    qui s'exécute. Il ne figure pas ici : ce fichier documente la création des objets.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_ParametreWU') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'NARRATIVE_MODE')
BEGIN
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'NARRATIVE_MODE', N'PAR_NATURE',
            N'PAR_NATURE : chaque ligne de la pièce porte le libellé de sa nature de mouvement (table T_NarrativeNatureWU). GLOBAL : les douze lignes d''un point de vente portent le même libellé. Le fichier core banking porte toujours la narrative unique du point de vente, quel que soit ce mode.',
            GETDATE(), N'installation');

    PRINT 'Option NARRATIVE_MODE créée à PAR_NATURE.';
END
ELSE IF OBJECT_ID(N'dbo.T_ParametreWU') IS NULL
    PRINT 'T_ParametreWU absente : exécutez Scripts\00_InstallationComplete.sql, qui crée tout.';
ELSE
    PRINT 'Option NARRATIVE_MODE déjà présente : valeur conservée.';
GO

-- =========================================================================
-- 4. Contrôle
--
--    La première requête doit normalement ne rien rendre : une table vide signifie que
--    toutes les natures suivent le modèle global, c'est-à-dire que rien n'a changé.
-- =========================================================================
SELECT Nature, Modele, LEN(Modele) AS LongueurDuModele, ModifiePar, DateModification
FROM   dbo.T_NarrativeNatureWU
ORDER  BY Nature;
GO

SELECT TOP (20) Cle, AncienneValeur, NouvelleValeur, ModifiePar, Poste, DateModification
FROM   dbo.T_JournalParametreWU
ORDER  BY DateModification DESC, Id DESC;
GO
