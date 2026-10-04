/*
    =========================================================================
    Narrative comptable — le modèle du libellé des écritures
    =========================================================================

    CE QUE CE SCRIPT AJOUTE : UNE LIGNE. Pas de table, pas de colonne, pas de droit.

    Le texte que la banque veut lire sur chacune de ses écritures — et dans la colonne
    ADDLTEXT du fichier chargé au core banking — devient un PARAMÈTRE. Il vit dans
    T_ParametreWU, créée par le script 18, qui est faite exactement pour ça : une clé, une
    valeur, son libellé, qui l'a changée et quand. Les droits y sont déjà posés par le
    script 18 : tout le monde LIT, seul wu_admin ÉCRIT.

    POURQUOI CE TEXTE QUITTE LE CODE

    Il a changé QUATRE FOIS en quatre livraisons : « CCS_<point de vente> ACTIVITE WU », puis
    le préfixe LD, puis la forme dictée « LD WU ACTIVITE <point de vente> <période> », puis
    les majuscules. Chaque mot a coûté une modification du code, une compilation, un commit,
    un pull et un redéploiement sur les postes — pour un texte qui n'entre dans aucun calcul.
    Ce qui se LIT dans le grand livre appartient à la Direction Comptable ; ce qui s'y
    CALCULE reste au code et ne bougera pas d'ici.

    CE SCRIPT N'EST PAS INDISPENSABLE, ET C'EST VOULU

    Sans lui, l'application applique le modèle par défaut de ConstantesWU — qui est la forme
    actuelle au caractère près — et l'écran « Narrative comptable » crée la ligne au premier
    enregistrement. Une base restaurée d'avant ce script sort donc exactement les narratives
    d'aujourd'hui. Il sert à ce que la ligne soit LISIBLE dans Management Studio, avec son
    libellé explicatif, sans attendre que quelqu'un ouvre l'écran.

    LES REPÈRES RECONNUS, et il n'y en a pas d'autres :

        {AGENCE}        désignation du point de vente (agence ou sous-agent)
        {ACCOUNT}       numéro Account du point de vente
        {CODE_AGENCE}   code agence, celui qui part dans la colonne ACBRN
        {PERIODE}       période couverte, « DU 08 AU 14 09 2026 »

    Un repère inconnu est REFUSÉ par l'écran, et non recopié : « LD WU ACTIVITE {AGENCY} »
    partirait au grand livre avec ses accolades, sur douze lignes et autant de points de
    vente. Un modèle saisi ici en SQL direct échappe à ce contrôle — la phrase sortirait
    alors avec ses accolades, et le seul garde-fou restant serait la limite de cent cinquante
    caractères du core banking.

    POURQUOI « LD WU ACTIVITE {AGENCE} {PERIODE} » ET PAS AUTRE CHOSE

    C'est la forme que la banque a dictée par écrit. Elle peut la changer depuis l'écran ;
    rejouer ce script ne la remettra JAMAIS d'autorité — la valeur existante est conservée,
    comme pour toutes les options du projet.

    À exécuter APRÈS 18 (ou après 00), sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. La table doit exister : elle vient du script 18
--
--    Ce script ne la crée pas. S'il la créait, une base où 18 n'a pas été joué se
--    retrouverait avec une table à une seule option, et le jour où 18 passerait enfin il
--    trouverait la table déjà là et n'y ajouterait rien d'autre.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_ParametreWU') IS NULL
BEGIN
    RAISERROR(N'La table T_ParametreWU est absente : exécutez d''abord Scripts\18_OptionsTraitement.sql.', 16, 1);
END
GO

-- =========================================================================
-- 2. Le modèle de narrative
--
--    La valeur existante n'est JAMAIS écrasée : rejouer ce script ne défait pas ce que la
--    Direction Comptable a saisi.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_ParametreWU') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'NARRATIVE_MODELE')
BEGIN
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'NARRATIVE_MODELE', N'LD WU ACTIVITE {AGENCE} {PERIODE}',
            N'Modèle de la narrative des lignes de la pièce comptable et de la colonne ADDLTEXT du fichier core banking. Repères reconnus : {AGENCE}, {ACCOUNT}, {CODE_AGENCE}, {PERIODE}.',
            GETDATE(), N'installation');

    PRINT 'Option NARRATIVE_MODELE créée : LD WU ACTIVITE {AGENCE} {PERIODE}.';
END
ELSE IF OBJECT_ID(N'dbo.T_ParametreWU') IS NOT NULL
BEGIN
    PRINT 'Option NARRATIVE_MODELE déjà présente : valeur conservée.';
END
GO

-- =========================================================================
-- 3. Contrôle
--
--    La longueur est affichée parce qu'elle compte : le core banking refuse un narratif de
--    plus de 150 caractères, et le modèle n'est qu'une PARTIE de ce qui sera rendu — les
--    repères y ajoutent le nom du point de vente et la période. L'écran de l'application
--    mesure le résultat sur le cas le plus long du référentiel ; ici on ne voit que le
--    gabarit.
-- =========================================================================
SELECT Cle, Valeur, LEN(Valeur) AS LongueurDuModele, ModifiePar, DateModification
FROM   dbo.T_ParametreWU
ORDER  BY Cle;
GO
