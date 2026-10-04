/*
    =========================================================================
    Borner le taux de rétrocession — contraintes CHECK
    =========================================================================

    CE QUE CE SCRIPT EMPÊCHE

    Le taux est une FRACTION : 0,70 signifie 70 %. Saisir « 70 » au lieu de « 0,70 »
    multiplierait par cent toutes les commissions rétrocédées — le sous-agent recevrait
    7000 % de la commission, la part de la banque deviendrait massivement négative, et la
    pièce comptable s'équilibrerait quand même : le compte courant Western Union est calculé
    par différence, il absorbe n'importe quoi.

    L'APPLICATION LE REFUSE DÉJÀ

    PointDeVenteSA.Anomalies et GroupeStatistiqueWU.Anomalies bornent le taux à [0 ; 1] et
    bloquent l'enregistrement, sur l'écran des sous-agents, sur celui des groupes, et au
    chargement d'un fichier de paramétrage. Ce script ne remplace pas ces contrôles.

    IL RESTE TROIS PORTES QUE L'APPLICATION NE GARDE PAS

      1. un UPDATE direct, par l'informatique ou une autre application ;
      2. une ligne posée à la main dans T_DemandeWU : l'autorisation d'une demande écrit son
         taux dans T_Pdv_SA SANS le revalider — le contrôle a eu lieu chez le demandeur, pas
         chez celui qui autorise ;
      3. une reprise de données, une restauration, un script de migration.

    Une contrainte en base ferme les trois d'un coup, et elle vaut quelle que soit la version
    de l'application installée sur les postes.

    CE SCRIPT NE CORRIGE AUCUNE DONNÉE

    Si des valeurs hors bornes existent déjà, il les NOMME et s'arrête sans rien poser : la
    correction d'un taux est une décision métier, pas un effet de bord de script. Relancez-le
    après correction.

    À exécuter APRÈS 01 à 18, sur la base GWC_WINCOMPENSE_ETD. Rejouable.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

SET NOCOUNT ON;
GO

-- =========================================================================
-- 1. Inventaire de ce qui est hors bornes AVANT de poser quoi que ce soit
-- =========================================================================
DECLARE @horsBornes INT = 0;

IF OBJECT_ID(N'dbo.T_Pdv_SA') IS NOT NULL
BEGIN
    SELECT @horsBornes = @horsBornes + COUNT(*)
    FROM dbo.T_Pdv_SA
    WHERE Taux < 0 OR Taux > 1;

    IF EXISTS (SELECT 1 FROM dbo.T_Pdv_SA WHERE Taux < 0 OR Taux > 1)
    BEGIN
        PRINT 'ANOMALIE — sous-agents dont le taux est hors de [0 ; 1] :';
        SELECT Code_Pdv, Designationagence, Taux, DateModification, ModifiePar
        FROM   dbo.T_Pdv_SA
        WHERE  Taux < 0 OR Taux > 1
        ORDER  BY Code_Pdv;
    END
END

IF OBJECT_ID(N'dbo.T_GroupeStatistique') IS NOT NULL
BEGIN
    SELECT @horsBornes = @horsBornes + COUNT(*)
    FROM dbo.T_GroupeStatistique
    WHERE Taux < 0 OR Taux > 1;

    IF EXISTS (SELECT 1 FROM dbo.T_GroupeStatistique WHERE Taux < 0 OR Taux > 1)
    BEGIN
        PRINT 'ANOMALIE — groupes statistiques dont le taux est hors de [0 ; 1] :';
        SELECT Groupe, Taux
        FROM   dbo.T_GroupeStatistique
        WHERE  Taux < 0 OR Taux > 1
        ORDER  BY Groupe;
    END
END

-- T_DemandeWU porte des demandes DÉJÀ décidées : une valeur hors bornes y est de l'histoire,
-- et l'histoire ne se réécrit pas. Seules les demandes encore en attente sont comptées comme
-- bloquantes — ce sont les seules qui pourraient encore écrire dans T_Pdv_SA.
IF OBJECT_ID(N'dbo.T_DemandeWU') IS NOT NULL
BEGIN
    SELECT @horsBornes = @horsBornes + COUNT(*)
    FROM dbo.T_DemandeWU
    WHERE Statut = N'EN_ATTENTE' AND Taux IS NOT NULL AND (Taux < 0 OR Taux > 1);

    IF EXISTS (SELECT 1 FROM dbo.T_DemandeWU
               WHERE Statut = N'EN_ATTENTE' AND Taux IS NOT NULL AND (Taux < 0 OR Taux > 1))
    BEGIN
        PRINT 'ANOMALIE — demandes EN ATTENTE dont le taux est hors de [0 ; 1] :';
        SELECT IdDemande, TypeObjet, Operation, Cle, Taux, SaisiPar, DateSaisie
        FROM   dbo.T_DemandeWU
        WHERE  Statut = N'EN_ATTENTE' AND Taux IS NOT NULL AND (Taux < 0 OR Taux > 1)
        ORDER  BY DateSaisie;
    END
END

IF @horsBornes > 0
BEGIN
    PRINT '';
    PRINT '=========================================================================';
    PRINT 'ARRÊT : ' + CAST(@horsBornes AS NVARCHAR(10)) + ' valeur(s) hors de [0 ; 1].';
    PRINT 'Aucune contrainte n''a été posée. Corrigez les lignes ci-dessus — le taux est';
    PRINT 'une FRACTION : 0,70 pour 70 % — puis relancez ce script.';
    PRINT '=========================================================================';
END
ELSE
BEGIN
    PRINT 'Aucun taux hors bornes : les contraintes peuvent être posées.';
END
GO

-- =========================================================================
-- 2. T_Pdv_SA
-- =========================================================================
IF OBJECT_ID(N'dbo.T_Pdv_SA') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_Pdv_SA_Taux')
   AND NOT EXISTS (SELECT 1 FROM dbo.T_Pdv_SA WHERE Taux < 0 OR Taux > 1)
BEGIN
    ALTER TABLE dbo.T_Pdv_SA
        ADD CONSTRAINT CK_T_Pdv_SA_Taux CHECK (Taux >= 0 AND Taux <= 1);

    PRINT 'Contrainte CK_T_Pdv_SA_Taux posée.';
END
ELSE IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_Pdv_SA_Taux')
BEGIN
    PRINT 'Contrainte CK_T_Pdv_SA_Taux déjà présente.';
END
GO

-- =========================================================================
-- 3. T_GroupeStatistique
-- =========================================================================
IF OBJECT_ID(N'dbo.T_GroupeStatistique') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_GroupeStatistique_Taux')
   AND NOT EXISTS (SELECT 1 FROM dbo.T_GroupeStatistique WHERE Taux < 0 OR Taux > 1)
BEGIN
    ALTER TABLE dbo.T_GroupeStatistique
        ADD CONSTRAINT CK_T_GroupeStatistique_Taux CHECK (Taux >= 0 AND Taux <= 1);

    PRINT 'Contrainte CK_T_GroupeStatistique_Taux posée.';
END
ELSE IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_GroupeStatistique_Taux')
BEGIN
    PRINT 'Contrainte CK_T_GroupeStatistique_Taux déjà présente.';
END
GO

-- =========================================================================
-- 4. T_DemandeWU — la porte la plus discrète
--
--    L'autorisation d'une demande écrit son taux dans T_Pdv_SA sans le revalider : le
--    contrôle a eu lieu chez le demandeur. Une ligne posée à la main dans cette table
--    traverserait donc tout le double regard sans jamais être contrôlée.
--
--    NULL est admis : une demande de suppression, ou portant sur une agence propre, n'a
--    pas de taux. Une contrainte CHECK laisse passer NULL d'elle-même — c'est écrit ici
--    pour que personne ne croie à un oubli.
--
--    WITH NOCHECK : les demandes DÉJÀ décidées ne sont pas revalidées. Ce sont des
--    archives ; les refuser empêcherait de poser la contrainte sans rien protéger de plus.
-- =========================================================================
IF OBJECT_ID(N'dbo.T_DemandeWU') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_Taux')
   AND NOT EXISTS (SELECT 1 FROM dbo.T_DemandeWU
                   WHERE Statut = N'EN_ATTENTE' AND Taux IS NOT NULL AND (Taux < 0 OR Taux > 1))
BEGIN
    ALTER TABLE dbo.T_DemandeWU WITH NOCHECK
        ADD CONSTRAINT CK_T_DemandeWU_Taux CHECK (Taux IS NULL OR (Taux >= 0 AND Taux <= 1));

    PRINT 'Contrainte CK_T_DemandeWU_Taux posée (les demandes déjà décidées ne sont pas revalidées).';
END
ELSE IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = N'CK_T_DemandeWU_Taux')
BEGIN
    PRINT 'Contrainte CK_T_DemandeWU_Taux déjà présente.';
END
GO

-- =========================================================================
-- 5. Ce que le script a laissé derrière lui
-- =========================================================================
PRINT '';
PRINT 'État des contraintes de taux :';

SELECT  c.name                AS Contrainte,
        OBJECT_NAME(c.parent_object_id) AS TableVisee,
        c.definition           AS Definition,
        c.is_disabled          AS Desactivee,
        c.is_not_trusted       AS NonVerifieeSurLExistant
FROM    sys.check_constraints AS c
WHERE   c.name IN (N'CK_T_Pdv_SA_Taux', N'CK_T_GroupeStatistique_Taux', N'CK_T_DemandeWU_Taux')
ORDER   BY TableVisee;
GO
