/*
    =========================================================================
    Le format du fichier core banking : sept codes, confiés à la banque
    =========================================================================

    CE SCRIPT NE CRÉE AUCUNE TABLE. Il ajoute sept lignes à T_ParametreWU, et c'est tout.
    Il est reproduit dans Scripts\00_InstallationComplete.sql, qui est le seul à exécuter.

    POURQUOI CES CODES SONT SORTIS DU CODE

    La banque en a changé QUATRE d'un coup, et c'est écrit dans les sources :

        BRN ............ N01  ->  S51
        TXNCD débit .... U24  ->  F03
        TXNCD crédit ... F15  ->  F57
        COST_CENTER .... supprimée du fichier

    Chacun a coûté un cycle complet : modification du code, compilation, commit, pull,
    redéploiement sur tous les postes. Pour des codes que la banque connaît et que nous,
    non. Ce sont les siens ; elle les corrige désormais depuis l'écran « Format du
    fichier core banking », réservé à l'administrateur.

    L'ENJEU EST PLUS GRAND QUE CELUI DE LA NARRATIVE. Un libellé mal écrit fait une ligne
    laide dans le grand livre. Un TXNCD mal écrit fait une écriture du MAUVAIS TYPE, ou un
    lot rejeté le matin de la compense. D'où trois garde-fous : un contrôle à la saisie,
    le journal T_JournalParametreWU, et un aperçu d'une ligne de fichier avant d'enregistrer.

    LES SEPT CLÉS

        CB_DETBSJRNL        BBR                     colonne DETBSJRNL
        CB_BRN              S51                     colonne BRN, agence DU LOT
        CB_SRCCODE          ECOSOURCE               colonne SRCCODE
        CB_TXNCD_DEBIT      F03                     colonne TXNCD, lignes au débit
        CB_TXNCD_CREDIT     F57                     colonne TXNCD, lignes au crédit
        CB_AGENCE_SIEGE     N01                     ACBRN des comptes internes
        CB_COMPTES_SIEGE    379100319,379200585     les comptes concernés

    BRN ET CB_AGENCE_SIEGE NE SONT PAS LA MÊME CHOSE, et c'est l'erreur à ne pas faire.
    BRN désigne l'agence du LOT, portée à l'identique par toutes ses lignes. ACBRN désigne
    l'agence de CHAQUE écriture : celle du point de vente qui l'a produite, ou celle du
    siège pour les comptes ci-dessus et pour les lignes qui n'appartiennent à aucun point
    de vente — l'écart d'arrondi, les écritures de change. Un fichier porte aujourd'hui S51
    en tête de chaque ligne et N01 dans ACBRN pour ces comptes-là.

    LA LISTE DES COMPTES TIENT DANS UNE SEULE OPTION, séparée par des virgules. Deux
    numéros de compte sans attribut ne justifient pas une table à eux : le journal y gagne
    une ligne lisible, et l'écran l'affiche quand même en petite grille. Le jour où la
    banque voudrait y attacher un libellé ou une date d'effet, il faudra une table.

    CE QUI N'EST PAS PARAMÉTRABLE, ET NE LE SERA PAS

        DRCR        « D » / « C ». Ce n'est pas un code de la banque, c'est la sémantique
                    du débit et du crédit. Le rendre modifiable offrirait d'inverser tous
                    les sens d'un fichier en une saisie.
        AMOUNT      calculé.
        ACNO        le compte de la ligne, déjà paramétré dans SystemeWU.
        VALDT       la date de valeur.
        ADDLTEXT    la narrative, déjà paramétrée (écran « Narrative comptable »).
        INSTR_NO    toujours vide.
        BATCHNO     une règle, pas une valeur : base 36 des jours écoulés depuis une
                    origine. Son origine, elle, pourrait devenir un paramètre un jour.

    UNE OPTION ABSENTE ET UNE OPTION VIDE NE DISENT PAS LA MÊME CHOSE, pour la liste des
    comptes. Absente : une base que ce script n'a pas encore touchée, et l'application rend
    les deux comptes écrits dans son code. Vide : une banque qui n'en rattache aucun, et
    l'application respecte son choix.

    CE QUI N'EST PAS HISTORISÉ, ET QU'IL FAUT SAVOIR

    La pièce est conservée ligne par ligne : son libellé, son narratif et son code agence
    sont ceux de sa génération. Le FICHIER core banking, lui, est RECONSTRUIT à chaque
    demande, et relit ces sept codes dans le paramétrage COURANT. Rejouer une journée
    archivée après un changement produit donc un fichier différent de celui qui a été
    chargé. L'écran des pièces archivées AVERTIT avant de produire ; il ne bloque pas,
    parce que rejouer avec les codes du jour est souvent exactement ce qu'on veut.

    REJOUABLE : chaque bloc est en IF NOT EXISTS, aucune valeur saisie n'est écrasée.
    =========================================================================
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_DETBSJRNL')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_DETBSJRNL', N'BBR',
            N'Colonne DETBSJRNL du fichier core banking, identique sur toutes les lignes.',
            GETDATE(), N'installation');
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_BRN')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_BRN', N'S51',
            N'Colonne BRN : agence DU LOT, portée à l''identique par toutes ses lignes. Ce n''est pas l''agence du siège, qui s''écrit en ACBRN sur les seuls comptes internes.',
            GETDATE(), N'installation');
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_SRCCODE')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_SRCCODE', N'ECOSOURCE',
            N'Colonne SRCCODE du fichier core banking, identique sur toutes les lignes.',
            GETDATE(), N'installation');
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_TXNCD_DEBIT')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_TXNCD_DEBIT', N'F03', N'Colonne TXNCD des lignes au DÉBIT.',
            GETDATE(), N'installation');
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_TXNCD_CREDIT')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_TXNCD_CREDIT', N'F57', N'Colonne TXNCD des lignes au CRÉDIT.',
            GETDATE(), N'installation');
GO

IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_AGENCE_SIEGE')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_AGENCE_SIEGE', N'N01',
            N'Agence portée en ACBRN par les comptes du siège, et par les lignes qui ne se rattachent à aucun point de vente — écart d''arrondi, écritures de change.',
            GETDATE(), N'installation');
GO

-- LA LISTE TIENT DANS UNE SEULE OPTION, separee par des virgules. Deux numeros de compte
-- sans attribut ne justifient pas une table a eux : le journal y gagne une ligne lisible
-- (« avant : 379100319,379200585 »), et l'ecran l'affiche quand meme en petite grille.
--
-- UNE OPTION ABSENTE ET UNE OPTION VIDE NE DISENT PAS LA MEME CHOSE. Absente : une base que
-- ce script n'a pas encore touchee, et l'application rend les deux comptes du code. Vide :
-- une banque qui n'en rattache aucun, et l'application respecte son choix.
IF NOT EXISTS (SELECT 1 FROM dbo.T_ParametreWU WHERE Cle = N'CB_COMPTES_SIEGE')
    INSERT INTO dbo.T_ParametreWU (Cle, Valeur, Libelle, DateModification, ModifiePar)
    VALUES (N'CB_COMPTES_SIEGE', N'379100319,379200585',
            N'Comptes rattachés d''office à l''agence du siège en ACBRN, séparés par des virgules. Règle d''aiguillage du format core banking, et non plan comptable : aucune écriture ne leur est passée de ce fait.',
            GETDATE(), N'installation');
GO
