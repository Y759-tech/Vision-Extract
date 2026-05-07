-- Clients actifs - VERSION CORRIGÉE AVEC TOUS LES FILTRES
DECLARE @Agence VARCHAR(100) = '';
DECLARE @TypeClient VARCHAR(50) = '';
DECLARE @Produit VARCHAR(50) = '';
DECLARE @DateDebut DATE = NULL;
DECLARE @DateFin DATE = NULL;

WITH CreditsParClient AS (
    SELECT 
        ecv.id_client as code_client,
        COUNT(DISTINCT ecv.id_pret) as nb_credits,
        SUM(ecv.mtt_pret) as total_credits,
        MAX(ecv.nom_agence) as agence,
        MAX(ecv.gestionnaire_pret) as gestionnaire,
        MAX(ecv.date_effet) as derniere_date_credit
    FROM dbo.extra_credits_view ecv
    WHERE ecv.id_client IS NOT NULL
      AND (@DateDebut IS NULL OR ecv.date_effet >= @DateDebut)
      AND (@DateFin IS NULL OR ecv.date_effet <= @DateFin)
    GROUP BY ecv.id_client
),
ComptesParClient AS (
    SELECT 
        ca.ID_ADHERENT as code_client,
        COUNT(DISTINCT ca.id) as nb_comptes,
        MAX(c.DATE_OUVERTURE) as derniere_ouverture_compte
    FROM COMPTES_ADHERENT ca
    INNER JOIN COMPTES c ON ca.id = c.ID
    WHERE (@DateDebut IS NULL OR c.DATE_OUVERTURE >= @DateDebut)
      AND (@DateFin IS NULL OR c.DATE_OUVERTURE <= @DateFin)
    GROUP BY ca.ID_ADHERENT
)
SELECT 
    a.ID AS [ID Client],
    a.NOM_ADHERENT AS [Nom Client],
    a.CODE AS [Code Client],
    a.DATE_INSCRIPTION AS [Date inscription],
    a.EST_VALIDE AS [Est valide],
    -- Informations sur les crédits
    ISNULL(cp.nb_credits, 0) AS [Nb crédits],
    ISNULL(cp.total_credits, 0) AS [Total crédits],
    ISNULL(cp.agence, 'Aucune') AS [Agence],
    ISNULL(cp.gestionnaire, 'Aucun') AS [Gestionnaire],
    -- Informations sur les comptes
    ISNULL(cc.nb_comptes, 0) AS [Nb comptes],
    -- Déterminer le type de produit
    CASE 
        WHEN cc.code_client IS NOT NULL AND cp.code_client IS NOT NULL THEN 'Les deux'
        WHEN cc.code_client IS NOT NULL THEN 'Épargne'
        WHEN cp.code_client IS NOT NULL THEN 'Crédit'
        ELSE 'Aucun'
    END AS [Produit],
    -- Type d'adhérent
    CASE a.ID_TYPE_ADHERENT
        WHEN 1 THEN 'Particulier'
        WHEN 2 THEN 'Entreprise'
        WHEN 3 THEN 'Groupe'
        ELSE 'Autre'
    END AS [Type client],
    -- Dates importantes
    cp.derniere_date_credit AS [Dernier crédit],
    cc.derniere_ouverture_compte AS [Dernier compte]
FROM ADHERENTS a
LEFT JOIN CreditsParClient cp ON a.ID = cp.code_client
LEFT JOIN ComptesParClient cc ON a.ID = cc.code_client
WHERE a.EST_VALIDE = 1  -- Clients valides seulement
  AND (
    @Produit = '' OR 
    (@Produit = 'EPARGNE' AND cc.code_client IS NOT NULL) OR
    (@Produit = 'CREDIT' AND cp.code_client IS NOT NULL) OR
    (@Produit = 'LES_DEUX' AND cc.code_client IS NOT NULL AND cp.code_client IS NOT NULL) OR
    (@Produit = 'AUCUN' AND cc.code_client IS NULL AND cp.code_client IS NULL)
  )
  AND (@Agence = '' OR cp.agence = @Agence OR (@Agence = 'AUCUNE' AND cp.agence IS NULL))
  AND (@TypeClient = '' OR 
       (@TypeClient = 'PARTICULIER' AND a.ID_TYPE_ADHERENT = 1) OR
       (@TypeClient = 'ENTREPRISE' AND a.ID_TYPE_ADHERENT = 2) OR
       (@TypeClient = 'GROUPE' AND a.ID_TYPE_ADHERENT = 3) OR
       (@TypeClient = 'AUTRE' AND a.ID_TYPE_ADHERENT NOT IN (1, 2, 3)))
ORDER BY ISNULL(cp.total_credits, 0) DESC, a.DATE_INSCRIPTION DESC;