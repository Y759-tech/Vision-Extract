-- ===========================================
-- REQUÊTE : Balance comptable complète
-- DESCRIPTION : Balance avec soldes période et cumulés
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @DateDebut DATE = '2026-01-01';        -- Date début de période
DECLARE @DateFin DATE = '2026-02-25';          -- Date fin de période
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-AGBOVILLE';  -- Nom de l'agence
DECLARE @Devise VARCHAR(50) = 'FCFA';          -- Devise
DECLARE @CompteDebut VARCHAR(20) = '0';        -- Compte début (optionnel)
DECLARE @CompteFin VARCHAR(20) = '9';          -- Compte fin (optionnel)
DECLARE @TypeBalance VARCHAR(50) = 'Balance générale';  -- Type de balance
-- ============================================

WITH SoldesComptes AS (
    SELECT 
        c.ID AS CompteID,
        c.NUM_CPTE AS NumeroCompte,
        c.LIBELLE AS Intitule,
        c.TYPE_COMPTE AS TypeCompte,
        d.LIBELLE AS Devise,
        ps.NOM AS AgenceNom,
        
        -- Solde à nouveau (avant la période)
        ISNULL((
            SELECT SUM(h.MONTANT_OPERATION * CASE WHEN h.SENS = 'C' THEN 1 ELSE -1 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) < @DateDebut
        ), 0) AS SoldeANouveau,
        
        -- Mouvements de la période
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) BETWEEN @DateDebut AND @DateFin
        ), 0) AS MouvementPeriodeDebit,
        
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) BETWEEN @DateDebut AND @DateFin
        ), 0) AS MouvementPeriodeCredit,
        
        -- Mouvements de l'exercice (depuis début d'année)
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND YEAR(CAST(h.DATE_OPERATION AS DATE)) = YEAR(@DateFin)
                AND CAST(h.DATE_OPERATION AS DATE) <= @DateFin
        ), 0) AS MouvementExerciceDebit,
        
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND YEAR(CAST(h.DATE_OPERATION AS DATE)) = YEAR(@DateFin)
                AND CAST(h.DATE_OPERATION AS DATE) <= @DateFin
        ), 0) AS MouvementExerciceCredit
        
    FROM COMPTES c
    LEFT JOIN DEVISES d ON c.ID_DEVISE = d.ID
    LEFT JOIN POINTS_SERVICE ps ON c.ID_AGENCE = ps.ID
    WHERE c.ETAT = 'O'
        AND (@Agence = '' OR ps.NOM = @Agence)
        AND (@Devise = '' OR d.LIBELLE LIKE '%' + @Devise + '%')
)
SELECT 
    NumeroCompte AS [Compte],
    Intitule AS [Intitulé],
    
    -- À nouveau
    CASE WHEN SoldeANouveau > 0 THEN ABS(SoldeANouveau) ELSE 0 END AS [A nouveau Débit],
    CASE WHEN SoldeANouveau < 0 THEN ABS(SoldeANouveau) ELSE 0 END AS [A nouveau Crédit],
    
    -- Mouvements période
    MouvementPeriodeDebit AS [Mouvements période Débit],
    MouvementPeriodeCredit AS [Mouvements période Crédit],
    
    -- Mouvements exercice
    MouvementExerciceDebit AS [Mouvements exercice Débit],
    MouvementExerciceCredit AS [Mouvements exercice Crédit],
    
    -- Soldes cumulés
    (CASE WHEN SoldeANouveau > 0 THEN SoldeANouveau ELSE 0 END + MouvementExerciceDebit) AS [Solde Débit],
    (CASE WHEN SoldeANouveau < 0 THEN ABS(SoldeANouveau) ELSE 0 END + MouvementExerciceCredit) AS [Solde Crédit],
    
    -- Solde final (Débit - Crédit)
    (SoldeANouveau + MouvementExerciceDebit - MouvementExerciceCredit) AS [Solde final],
    
    -- Pour le tri par compte
    LEFT(NumeroCompte, 2) AS [Classe]
    
FROM SoldesComptes
WHERE 
    -- Filtre sur les comptes qui ont des mouvements ou un solde non nul
    (SoldeANouveau != 0 OR MouvementPeriodeDebit > 0 OR MouvementPeriodeCredit > 0)
    
    -- Filtre sur la plage de comptes si spécifiée
    AND (@CompteDebut = '0' OR CAST(LEFT(NumeroCompte + '000000', 6) AS BIGINT) >= CAST(@CompteDebut + '000000' AS BIGINT))
    AND (@CompteFin = '9' OR CAST(LEFT(NumeroCompte + '000000', 6) AS BIGINT) <= CAST(@CompteFin + '999999' AS BIGINT))

ORDER BY 
    Classe,
    NumeroCompte;