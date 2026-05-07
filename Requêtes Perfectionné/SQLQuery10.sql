-- ===========================================
-- REQUÊTE : Grand Livre (format tabulaire) - VERSION CORRIGÉE
-- ===========================================

DECLARE @Devise VARCHAR(50) = 'FCFA';
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-AGBOVILLE';
DECLARE @DateDebut DATE = '2026-02-01';
DECLARE @DateFin DATE = '2026-02-25';
DECLARE @NumeroCompte VARCHAR(50) = '10111100';

-- Récupérer les infos du compte
DECLARE @IntituleCompte VARCHAR(200);
SELECT @IntituleCompte = c.LIBELLE
FROM COMPTES c
LEFT JOIN POINTS_SERVICE ps ON c.ID_AGENCE = ps.ID
WHERE c.NUM_CPTE = @NumeroCompte AND ps.NOM = @Agence;

-- Calcul du solde initial
DECLARE @SoldeInitial FLOAT;
SELECT @SoldeInitial = ISNULL((
    SELECT SUM(h.MONTANT_OPERATION * CASE WHEN h.SENS = 'C' THEN 1 ELSE -1 END)
    FROM HDPM h
    INNER JOIN COMPTES c ON h.ID_COMPTE = c.ID
    WHERE c.NUM_CPTE = @NumeroCompte
        AND CAST(h.DATE_OPERATION AS DATE) < @DateDebut
), 0);

-- Mouvements avec solde calculé
WITH Mouvements AS (
    SELECT 
        h.DATE_OPERATION,
        h.DESCRIPTION AS LibelleOperation,  -- Au lieu de LIBELLE_OPERATION
        h.NUM_TRANSACTION AS Reference,      -- NUM_TRANSACTION comme référence
        h.NUMERO_RECU,
        CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END AS Debit,
        CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END AS Credit,
        tpo.CLE_LIBELLE AS TypeOperation,
        -- Récupérer l'utilisateur via la table OPERATIONS
        u.NOM + ' ' + ISNULL(u.PRENOM, '') AS Utilisateur,
        ROW_NUMBER() OVER (ORDER BY h.DATE_OPERATION, h.ID) AS RowNum
    FROM HDPM h
    LEFT JOIN OPERATIONS o ON h.ID_OPERATION = o.ID  -- Lien avec OPERATIONS
    LEFT JOIN UTILISATEURS u ON o.ID_UTILISATEUR = u.id  -- Utilisateur dans OPERATIONS
    LEFT JOIN TYPES_OPERATION tpo ON h.ID_TYPE_OPERATION = tpo.ID
    WHERE h.ID_COMPTE IN (SELECT ID FROM COMPTES WHERE NUM_CPTE = @NumeroCompte)
        AND CAST(h.DATE_OPERATION AS DATE) BETWEEN @DateDebut AND @DateFin
)
SELECT 
    CONVERT(VARCHAR, DATE_OPERATION, 103) AS [Date],
    LibelleOperation AS [Libellé],
    Reference AS [Référence],
    NUMERO_RECU AS [N° reçu],
    FORMAT(Debit, 'N0') AS [Débit (FCFA)],
    FORMAT(Credit, 'N0') AS [Crédit (FCFA)],
    FORMAT(@SoldeInitial + SUM(Debit - Credit) OVER (ORDER BY DATE_OPERATION, RowNum), 'N0') AS [Solde (FCFA)],
    TypeOperation AS [Type],
    Utilisateur AS [Opérateur]
FROM Mouvements
ORDER BY DATE_OPERATION, RowNum;

-- Récapitulatif
SELECT 
    'RÉCAPITULATIF' AS [ ],
    FORMAT(@SoldeInitial, 'N0') AS [Solde initial],
    FORMAT(ISNULL((SELECT SUM(Debit) FROM Mouvements), 0), 'N0') AS [Total débit],
    FORMAT(ISNULL((SELECT SUM(Credit) FROM Mouvements), 0), 'N0') AS [Total crédit],
    FORMAT(@SoldeInitial + 
           ISNULL((SELECT SUM(Debit - Credit) FROM Mouvements), 0), 'N0') AS [Solde final];