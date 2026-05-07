-- ===========================================
-- REQUÊTE : Parts sociales
-- DESCRIPTION : Liste des parts sociales souscrites par les clients
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @DateDebut DATE = '2025-01-01';              -- Ex: '2025-01-01' ou NULL
DECLARE @DateFin DATE = '2025-01-31';                -- Ex: '2025-12-31' ou NULL
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-BONOUA';           -- Ex: 'REMUCI : AGENCE-BONOUA' ou ''
DECLARE @Client VARCHAR(100) = '';           -- ID client ou nom ou ''
DECLARE @TypeOperation VARCHAR(50) = '';     -- 'ACHAT', 'RETRAIT' ou ''
DECLARE @MontantMin FLOAT = NULL;            -- Montant minimum ou NULL
DECLARE @MontantMax FLOAT = NULL;            -- Montant maximum ou NULL
-- ============================================

SELECT 
    ops.ID AS [ID Opération],
    ops.ID_ADHERENT AS [ID Client],
    a.NOM_ADHERENT AS [Nom client],
    a.CODE AS [Code client],
    ps.NOM AS [Agence],  -- On utilise POINTS_SERVICE.NOM au lieu de AGENCES.NOM
    ops.NOMBRE AS [Nombre de parts],
    psv.VALEUR AS [Valeur nominale],  -- Renommé pour éviter confusion avec la table
    (ops.NOMBRE * psv.VALEUR) AS [Montant total],
    o.DATE_OPERATION AS [Date opération],
    o.DATE_SAISIE AS [Date saisie],
    tpo.CLE_LIBELLE AS [Type opération],
    CASE 
        WHEN tpo.CLE_LIBELLE LIKE '%ACHAT%' OR tpo.CLE_LIBELLE LIKE '%SOUSCRIPTION%' THEN 'ACTIVES'
        WHEN tpo.CLE_LIBELLE LIKE '%RETRAIT%' OR tpo.CLE_LIBELLE LIKE '%CLOTURE%' THEN 'RETIRÉES'
        ELSE 'AUTRE'
    END AS [Statut],
    d.LIBELLE AS [Devise],  -- DEVISES.LIBELLE au lieu de NOM_DEVISE
    ops.version
FROM OPERATIONS_PART_SOC ops
INNER JOIN ADHERENTS a ON ops.ID_ADHERENT = a.ID
LEFT JOIN POINTS_SERVICE ps ON a.ID_AGENCE = ps.ID  -- Changé AGENCES → POINTS_SERVICE
LEFT JOIN PARTS_SOCIALE psv ON ops.ID_PART_SOCIALE = psv.ID  -- Renommé pour éviter confusion
LEFT JOIN OPERATIONS o ON ops.ID_OPERATION = o.ID
LEFT JOIN TYPES_OPERATION tpo ON o.ID_TYPE_OPERATION = tpo.ID
LEFT JOIN DEVISES d ON ops.ID_DEVISE = d.ID
WHERE 1=1
    -- FILTRE DATE OPÉRATION
    AND (@DateDebut IS NULL OR o.DATE_OPERATION >= @DateDebut)
    AND (@DateFin IS NULL OR o.DATE_OPERATION <= @DateFin)
    -- FILTRE AGENCE (avec LIKE pour recherche partielle)
    AND (@Agence = '' OR ps.NOM LIKE '%' + @Agence + '%')
    -- FILTRE CLIENT
    AND (
        @Client = '' 
        OR a.ID LIKE '%' + @Client + '%'
        OR a.NOM_ADHERENT LIKE '%' + @Client + '%'
        OR a.CODE LIKE '%' + @Client + '%'
    )
    -- FILTRE TYPE OPÉRATION
    AND (@TypeOperation = '' OR tpo.CLE_LIBELLE LIKE '%' + @TypeOperation + '%')
    -- FILTRE MONTANT
    AND (@MontantMin IS NULL OR (ops.NOMBRE * psv.VALEUR) >= @MontantMin)
    AND (@MontantMax IS NULL OR (ops.NOMBRE * psv.VALEUR) <= @MontantMax)
ORDER BY o.DATE_OPERATION DESC;
