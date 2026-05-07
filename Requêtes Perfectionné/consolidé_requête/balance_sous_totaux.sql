-- ===========================================
-- REQUÊTE : Balance comptable complète avec sous-totaux
-- DESCRIPTION : Balance avec soldes période, cumulés et sous-totaux par classe
-- FORMAT : Identique à l'exemple du SIG
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @DateDebut DATE = '2026-01-01';        -- Date début de période
DECLARE @DateFin DATE = '2026-02-25';          -- Date fin de période
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-AGBOVILLE';  -- Nom de l'agence
DECLARE @Devise VARCHAR(50) = 'FCFA';          -- Devise
DECLARE @CompteDebut VARCHAR(20) = '0';        -- Compte début
DECLARE @CompteFin VARCHAR(20) = '9';          -- Compte fin
-- ============================================

WITH SoldesComptes AS (
    -- Calcul des soldes pour chaque compte
    SELECT 
        c.ID AS CompteID,
        c.NUM_CPTE AS NumeroCompte,
        c.LIBELLE AS Intitule,
        LEFT(c.NUM_CPTE, 2) AS ClasseCompte,
        d.LIBELLE AS Devise,
        ps.NOM AS AgenceNom,
        
        -- Solde à nouveau (avant la période)
        ISNULL((
            SELECT SUM(h.MONTANT_OPERATION * CASE WHEN h.SENS = 'C' THEN 1 ELSE -1 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) < @DateDebut
        ), 0) AS SoldeANouveau,
        
        -- Mouvements de la période (Débit)
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) BETWEEN @DateDebut AND @DateFin
        ), 0) AS MouvementPeriodeDebit,
        
        -- Mouvements de la période (Crédit)
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND CAST(h.DATE_OPERATION AS DATE) BETWEEN @DateDebut AND @DateFin
        ), 0) AS MouvementPeriodeCredit,
        
        -- Mouvements de l'exercice (Débit)
        ISNULL((
            SELECT SUM(CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END)
            FROM HDPM h
            WHERE h.ID_COMPTE = c.ID
                AND YEAR(CAST(h.DATE_OPERATION AS DATE)) = YEAR(@DateFin)
                AND CAST(h.DATE_OPERATION AS DATE) <= @DateFin
        ), 0) AS MouvementExerciceDebit,
        
        -- Mouvements de l'exercice (Crédit)
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
),
BalanceDetail AS (
    -- Détail des comptes avec soldes calculés
    SELECT 
        NumeroCompte,
        Intitule,
        ClasseCompte,
        
        -- À nouveau (Débit/Crédit)
        CASE WHEN SoldeANouveau > 0 THEN ABS(SoldeANouveau) ELSE 0 END AS [ANouveauDebit],
        CASE WHEN SoldeANouveau < 0 THEN ABS(SoldeANouveau) ELSE 0 END AS [ANouveauCredit],
        
        -- Mouvements période
        MouvementPeriodeDebit AS [MouvementPeriodeDebit],
        MouvementPeriodeCredit AS [MouvementPeriodeCredit],
        
        -- Mouvements exercice
        MouvementExerciceDebit AS [MouvementExerciceDebit],
        MouvementExerciceCredit AS [MouvementExerciceCredit],
        
        -- Soldes cumulés
        (CASE WHEN SoldeANouveau > 0 THEN SoldeANouveau ELSE 0 END + MouvementExerciceDebit) AS [SoldeDebit],
        (CASE WHEN SoldeANouveau < 0 THEN ABS(SoldeANouveau) ELSE 0 END + MouvementExerciceCredit) AS [SoldeCredit],
        
        -- Solde final (pour calcul)
        (SoldeANouveau + MouvementExerciceDebit - MouvementExerciceCredit) AS [SoldeFinal]
        
    FROM SoldesComptes
    WHERE 
        -- Garder uniquement les comptes avec activité
        (SoldeANouveau != 0 OR MouvementPeriodeDebit > 0 OR MouvementPeriodeCredit > 0)
        
        -- Filtre sur la plage de comptes
        AND (@CompteDebut = '0' OR CAST(LEFT(NumeroCompte + '000000', 6) AS BIGINT) >= CAST(@CompteDebut + '000000' AS BIGINT))
        AND (@CompteFin = '9' OR CAST(LEFT(NumeroCompte + '000000', 6) AS BIGINT) <= CAST(@CompteFin + '999999' AS BIGINT))
)

-- 1ère partie : Les comptes détaillés
SELECT 
    NumeroCompte AS [Compte],
    Intitule AS [Intitulé],
    CAST(ANouveauDebit AS BIGINT) AS [A nouveau Débit],
    CAST(ANouveauCredit AS BIGINT) AS [A nouveau Crédit],
    CAST(MouvementPeriodeDebit AS BIGINT) AS [Mouvements période Débit],
    CAST(MouvementPeriodeCredit AS BIGINT) AS [Mouvements période Crédit],
    CAST(MouvementExerciceDebit AS BIGINT) AS [Mouvements exercice Débit],
    CAST(MouvementExerciceCredit AS BIGINT) AS [Mouvements exercice Crédit],
    CAST(SoldeDebit AS BIGINT) AS [Solde Débit],
    CAST(SoldeCredit AS BIGINT) AS [Solde Crédit],
    ClasseCompte,
    1 AS OrdreTri  -- Les comptes d'abord
FROM BalanceDetail

UNION ALL

-- 2ème partie : Sous-totaux par classe
SELECT 
    'Sous total compte ' + ClasseCompte AS [Compte],
    '' AS [Intitulé],
    CAST(SUM(ANouveauDebit) AS BIGINT),
    CAST(SUM(ANouveauCredit) AS BIGINT),
    CAST(SUM(MouvementPeriodeDebit) AS BIGINT),
    CAST(SUM(MouvementPeriodeCredit) AS BIGINT),
    CAST(SUM(MouvementExerciceDebit) AS BIGINT),
    CAST(SUM(MouvementExerciceCredit) AS BIGINT),
    CAST(SUM(SoldeDebit) AS BIGINT),
    CAST(SUM(SoldeCredit) AS BIGINT),
    ClasseCompte,
    2 AS OrdreTri  -- Les sous-totaux ensuite
FROM BalanceDetail
GROUP BY ClasseCompte

UNION ALL

-- 3ème partie : Total général
SELECT 
    'Total général' AS [Compte],
    '' AS [Intitulé],
    CAST(SUM(ANouveauDebit) AS BIGINT),
    CAST(SUM(ANouveauCredit) AS BIGINT),
    CAST(SUM(MouvementPeriodeDebit) AS BIGINT),
    CAST(SUM(MouvementPeriodeCredit) AS BIGINT),
    CAST(SUM(MouvementExerciceDebit) AS BIGINT),
    CAST(SUM(MouvementExerciceCredit) AS BIGINT),
    CAST(SUM(SoldeDebit) AS BIGINT),
    CAST(SUM(SoldeCredit) AS BIGINT),
    'ZZ' AS ClasseCompte,  -- Pour que le total soit à la fin
    3 AS OrdreTri
FROM BalanceDetail

ORDER BY 
    ClasseCompte,
    OrdreTri,
    Compte;
	'