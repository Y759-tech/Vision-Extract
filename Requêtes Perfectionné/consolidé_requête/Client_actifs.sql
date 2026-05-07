-- ===========================================
-- REQUÊTE : Clients actifs avec tous les filtres
-- SOURCE : /api/clients-actifs
-- DESCRIPTION : Liste des clients actifs avec historique crédits et épargne
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @Agence VARCHAR(100) = '';              -- Ex: 'REMUCI : AGENCE-BONOUA' ou '' pour toutes
DECLARE @TypeClient VARCHAR(50) = '';           -- 'PARTICULIER', 'ENTREPRISE', 'GROUPE', 'AUTRE' ou ''
DECLARE @Produit VARCHAR(50) = '';              -- 'EPARGNE', 'CREDIT', 'LES_DEUX', 'AUCUN' ou ''
DECLARE @Gestionnaire VARCHAR(100) = '';        -- Nom du gestionnaire ou ''
-- ============================================

WITH CreditsParClient AS (
    SELECT 
        ecv.id_client as code_client,
        COUNT(DISTINCT ecv.id_pret) as nb_credits,
        SUM(ecv.mtt_pret) as total_credits_historique,
        SUM(CASE WHEN ecv.date_solde IS NULL THEN ecv.mtt_pret ELSE 0 END) as encours_actuel,
        COUNT(DISTINCT CASE WHEN ecv.date_solde IS NULL THEN ecv.id_pret END) as nb_credits_encours,
        MAX(ecv.nom_agence) as agence,
        MAX(ecv.gestionnaire_pret) as gestionnaire,
        MAX(ecv.date_effet) as derniere_date_credit
    FROM dbo.extra_credits_view ecv
    WHERE ecv.id_client IS NOT NULL
    GROUP BY ecv.id_client
),
ComptesParClient AS (
    SELECT 
        ca.ID_ADHERENT as code_client,
        COUNT(DISTINCT ca.id) as nb_comptes,
        MAX(c.DATE_OUVERTURE) as derniere_ouverture_compte
    FROM COMPTES_ADHERENT ca
    INNER JOIN COMPTES c ON ca.id = c.ID
    GROUP BY ca.ID_ADHERENT
)
SELECT 
    a.ID AS [ID Client],
    a.NOM_ADHERENT AS [Nom Client],
    a.CODE AS [Code Client],
    a.DATE_INSCRIPTION AS [Date inscription],
    a.EST_VALIDE AS [Est valide],
    ISNULL(cp.nb_credits, 0) AS [Nb crédits total],
    ISNULL(cp.total_credits_historique, 0) AS [Total crédits historique],
    ISNULL(cp.nb_credits_encours, 0) AS [Nb crédits encours],
    ISNULL(cp.encours_actuel, 0) AS [Encours actuel],
    ISNULL(cp.agence, 'Aucune') AS [Agence],
    ISNULL(cp.gestionnaire, 'Aucun') AS [Gestionnaire],
    ISNULL(cc.nb_comptes, 0) AS [Nb comptes],
    CASE 
        WHEN cc.code_client IS NOT NULL AND cp.code_client IS NOT NULL THEN 'Les deux'
        WHEN cc.code_client IS NOT NULL THEN 'Épargne'
        WHEN cp.code_client IS NOT NULL THEN 'Crédit'
        ELSE 'Aucun'
    END AS [Produit],
    CASE a.ID_TYPE_ADHERENT
        WHEN 1 THEN 'Particulier'
        WHEN 2 THEN 'Entreprise'
        WHEN 3 THEN 'Groupe'
        ELSE 'Autre'
    END AS [Type client],
    cp.derniere_date_credit AS [Dernier crédit],
    cc.derniere_ouverture_compte AS [Dernier compte]
FROM ADHERENTS a
LEFT JOIN CreditsParClient cp ON a.ID = cp.code_client
LEFT JOIN ComptesParClient cc ON a.ID = cc.code_client
WHERE a.EST_VALIDE = 1
    -- FILTRE AGENCE
    AND (
        @Agence = '' 
        OR cp.agence LIKE '%' + @Agence + '%' 
        OR (@Agence = 'AUCUNE' AND cp.agence IS NULL)
    )
    -- FILTRE GESTIONNAIRE
    AND (
        @Gestionnaire = '' 
        OR cp.gestionnaire LIKE '%' + @Gestionnaire + '%' 
        OR (@Gestionnaire = 'Tous les gestionnaires' AND cp.gestionnaire IS NOT NULL)
    )
    -- FILTRE TYPE CLIENT
    AND (
        @TypeClient = '' 
        OR (@TypeClient = 'PARTICULIER' AND a.ID_TYPE_ADHERENT = 1)
        OR (@TypeClient = 'ENTREPRISE' AND a.ID_TYPE_ADHERENT = 2)
        OR (@TypeClient = 'GROUPE' AND a.ID_TYPE_ADHERENT = 3)
        OR (@TypeClient = 'AUTRE' AND a.ID_TYPE_ADHERENT NOT IN (1, 2, 3))
    )
    -- FILTRE PRODUIT
    AND (
        @Produit = '' 
        OR (@Produit = 'EPARGNE' AND cc.code_client IS NOT NULL)
        OR (@Produit = 'CREDIT' AND cp.code_client IS NOT NULL)
        OR (@Produit = 'LES_DEUX' AND cc.code_client IS NOT NULL AND cp.code_client IS NOT NULL)
        OR (@Produit = 'AUCUN' AND cc.code_client IS NULL AND cp.code_client IS NULL)
    )
ORDER BY ISNULL(cp.encours_actuel, 0) DESC, a.DATE_INSCRIPTION DESC;