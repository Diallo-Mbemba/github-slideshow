/*
    Table SystemeWU : paramétrage des comptes comptables utilisés par la pièce comptable
    Western Union. Ces comptes sont lus au démarrage de l'application et modifiables depuis
    le formulaire « Comptes Systèmes WU » (bouton « Paramètres des comptes... »).

    Ce script est IDEMPOTENT : il ne crée la table et la ligne de paramétrage que si elles
    n'existent pas déjà, et ne modifie JAMAIS une ligne existante — le paramétrage en place,
    qui fait foi, ne doit pas être écrasé par un script.

    À n'exécuter que si la table SystemeWU n'existe pas encore dans l'environnement visé.
*/

USE GWC_WINCOMPENSE_ETD;
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'SystemeWU')
BEGIN
    CREATE TABLE dbo.SystemeWU
    (
        Passif                      NVARCHAR(255) NULL,
        Actif                       NVARCHAR(255) NULL,
        Tob                         NVARCHAR(255) NULL,   -- TVA collectée Western Union
        Tthu                        NVARCHAR(255) NULL,   -- Impôts et taxe sur envoi
        Cpte_PositionNette          NVARCHAR(255) NULL,   -- Compte courant Western Union ETD
        Cpte_Produit                NVARCHAR(255) NULL,   -- Commission sur Transfert_Ecobank
        Cpte_Charge_Publicitaire    NVARCHAR(255) NULL,
        Cpte_Gainde_Change          NVARCHAR(255) NULL,
        Cpte_Envoi                  NVARCHAR(255) NULL,   -- TTA sur envoi WU
        Cpte_Paiement               NVARCHAR(255) NULL,   -- TTA sur paiement WU
        Cpte_attenteDEBIT           NVARCHAR(255) NULL,   -- Compte inter bancaire (débit)
        Cpte_attenteCREDIT          NVARCHAR(255) NULL,   -- Compte inter bancaire (crédit)
        Cpte_Produit_Envoi          NVARCHAR(255) NULL,   -- Commission sur Envoi_Ecobank
        Cpte_Produit_Paiement       NVARCHAR(255) NULL,   -- Commission sur Paiement_Ecobank
        Cpte_Envoi_agence           NVARCHAR(255) NULL,
        Cpte_Paiement_agence        NVARCHAR(255) NULL,
        code                        NCHAR(10)     NULL    -- Identifie la ligne de paramétrage
    );
END
GO

/*
    Ligne de paramétrage Western Union.

    Les valeurs ci-dessous sont celles du paramétrage en service au moment de la rédaction,
    reprises du formulaire « Comptes Systèmes WU » de la Direction Comptable. Les colonnes
    non utilisées par l'application (Passif, Actif, Cpte_Charge_Publicitaire,
    Cpte_Gainde_Change, Cpte_Envoi_agence, Cpte_Paiement_agence) sont laissées à NULL :
    l'application ne les lit ni ne les écrit jamais.

    La colonne « code » doit être renseignée : sans elle, l'application refuse d'enregistrer
    dès que la table contient plusieurs lignes, ne pouvant pas déterminer laquelle mettre à
    jour sans risquer d'altérer le paramétrage d'une autre application.
*/
IF NOT EXISTS (SELECT 1 FROM dbo.SystemeWU)
BEGIN
    INSERT INTO dbo.SystemeWU
    (
        Tob, Tthu,
        Cpte_PositionNette,
        Cpte_Produit, Cpte_Produit_Envoi, Cpte_Produit_Paiement,
        Cpte_Envoi, Cpte_Paiement,
        Cpte_attenteDEBIT, Cpte_attenteCREDIT,
        code
    )
    VALUES
    (
        '434000104',    -- TVA collectée Western Union
        '434000147',    -- Impôts et taxe sur envoi
        '32100003292',  -- Compte courant Western Union ETD
        '728300148',    -- Commission sur Transfert_Ecobank
        '728300148',    -- Commission sur Envoi_Ecobank
        '728300149',    -- Commission sur Paiement_Ecobank
        '434000145',    -- TTA sur envoi WU
        '434000159',    -- TTA sur paiement WU
        '381000101',    -- Compte inter bancaire (débit)
        '381000101',    -- Compte inter bancaire (crédit)
        '123'
    );
END
GO

-- Contrôle : comptes effectivement lus par l'application.
SELECT  code                    AS [Code],
        Cpte_PositionNette      AS [Compte courant WU ETD],
        Cpte_attenteDEBIT       AS [Compte inter bancaire],
        Cpte_Produit            AS [Commission Transfert],
        Cpte_Produit_Envoi      AS [Commission Envoi],
        Cpte_Produit_Paiement   AS [Commission Paiement],
        Tthu                    AS [Impots et taxe envoi],
        Tob                     AS [TVA],
        Cpte_Envoi              AS [TTA envoi],
        Cpte_Paiement           AS [TTA paiement]
FROM    dbo.SystemeWU;
GO
