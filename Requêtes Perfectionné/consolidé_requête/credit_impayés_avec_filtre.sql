-- ===========================================
-- REQUÊTE : Crédits impayés avec FILTRES PARAMÉTRABLES
-- SOURCE : /api/credits-impayes
-- DESCRIPTION : Liste tous les crédits en retard avec filtres
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-BONOUA';           -- Ex: 'REMUCI : AGENCE-BONOUA' ou '' pour toutes
DECLARE @Gestionnaire VARCHAR(100) = '';     -- Ex: 'KONAN KOFFI' ou '' pour tous
DECLARE @TypeCredit VARCHAR(100) = '';       -- Ex: 'Crédit équipement' ou '' pour tous
DECLARE @Client VARCHAR(200) = '';           -- Ex: 'KOUASSI' ou '' pour tous
DECLARE @JoursRetardMin INT = 61;          -- Ex: 30 ou NULL pour pas de minimum
DECLARE @JoursRetardMax INT = 66;          -- Ex: 90 ou NULL pour pas de maximum
DECLARE @MontantMin FLOAT = 5000000;            -- Ex: 100000 ou NULL pour pas de minimum
DECLARE @MontantMax FLOAT = NULL;            -- Ex: 1000000 ou NULL pour pas de maximum
-- ============================================

SELECT 
    num_manuel AS [N° manuel],
    nom_client + ' ' + prenoms_client AS [Client],
    mtt_pret AS [Montant accordé],
    date_premiere_echeance AS [Date première échéance],
    date_fin_echeance AS [Date échéance finale],
    DATEDIFF(day, date_fin_echeance, GETDATE()) AS [Jours retard],
    nom_agence AS [Agence],
    gestionnaire_pret AS [Gestionnaire],
    produit AS [Type de crédit],
    -- Calcul du type de défaut
    CASE 
        WHEN DATEDIFF(day, date_fin_echeance, GETDATE()) > 
             DATEDIFF(day, date_premiere_echeance, date_fin_echeance) THEN 'DÉFAUT TOTAL'
        WHEN DATEDIFF(day, date_premiere_echeance, date_fin_echeance) = 0 THEN 'DURÉE INCONNUE'
        WHEN CAST(
            (DATEDIFF(day, date_premiere_echeance, date_fin_echeance) - 
             DATEDIFF(day, date_fin_echeance, GETDATE())) * 100.0 / 
            DATEDIFF(day, date_premiere_echeance, date_fin_echeance) 
        AS DECIMAL(5,2)) < 50 THEN 'DÉFAUT PRÉCOCE'
        ELSE 'DÉFAUT TARDIF'
    END AS [Type de défaut],
    code_client AS [Code client],
    telephone AS [Téléphone],
    date_adhesion AS [Date adhésion]
FROM dbo.extra_credits_view
WHERE date_fin_echeance IS NOT NULL
    AND DATEDIFF(day, date_fin_echeance, GETDATE()) > 0
    -- FILTRE AGENCE
    AND (@Agence = '' OR nom_agence = @Agence)
    -- FILTRE GESTIONNAIRE
    AND (@Gestionnaire = '' OR gestionnaire_pret = @Gestionnaire)
    -- FILTRE TYPE DE CRÉDIT (LIKE pour recherche partielle)
    AND (@TypeCredit = '' OR produit LIKE '%' + @TypeCredit + '%')
    -- FILTRE CLIENT (recherche dans nom + prénoms)
    AND (@Client = '' OR (nom_client + ' ' + prenoms_client) LIKE '%' + @Client + '%')
    -- FILTRE JOURS DE RETARD MINIMUM
    AND (@JoursRetardMin IS NULL OR DATEDIFF(day, date_fin_echeance, GETDATE()) >= @JoursRetardMin)
    -- FILTRE JOURS DE RETARD MAXIMUM
    AND (@JoursRetardMax IS NULL OR DATEDIFF(day, date_fin_echeance, GETDATE()) <= @JoursRetardMax)
    -- FILTRE MONTANT MINIMUM
    AND (@MontantMin IS NULL OR mtt_pret >= @MontantMin)
    -- FILTRE MONTANT MAXIMUM
    AND (@MontantMax IS NULL OR mtt_pret <= @MontantMax)
ORDER BY [Jours retard] DESC;