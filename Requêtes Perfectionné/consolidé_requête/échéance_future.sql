-- ===========================================
-- REQUÊTE : Tombée d'échéance (Échéances futures)
-- DESCRIPTION : Liste des échéances à venir pour une période donnée
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @DateDebut DATE = '2026-03-05';       -- Date début de recherche
DECLARE @DateFin DATE = '2026-03-05';         -- Date fin de recherche
DECLARE @Agence VARCHAR(100) = '';            -- Ex: 'REMUCI : AGENCE-BONOUA' ou ''
DECLARE @Gestionnaire VARCHAR(100) = '';      -- Nom du gestionnaire ou ''
DECLARE @Client VARCHAR(100) = '';            -- ID client ou nom ou ''
DECLARE @Produit VARCHAR(100) = '';           -- Type de crédit ou ''
DECLARE @TypeCredit VARCHAR(100) = '';        -- Type de crédit détaillé
DECLARE @NumContrat VARCHAR(100) = '';        -- N° de contrat
-- ============================================

SELECT 
    -- Informations de l'échéance
    t.ID AS [ID Échéance],
    t.DATE_ECHEANCE AS [Date échéance],
    t.CAPITAL AS [Capital],
    t.INTERET AS [Intérêt],
    t.COMMISSION AS [Commission],
    -- Montant total (capital + intérêt + commissions + taxes)
    (t.CAPITAL 
     + ISNULL(t.INTERET, 0) 
     + ISNULL(t.COMMISSION, 0) 
     + ISNULL(t.CSS_COMMIS, 0)
     + ISNULL(t.CSS_INT, 0)
     + ISNULL(t.TAXE_COMMIS, 0)
     + ISNULL(t.TAXE_INT, 0)) AS [Montant total],
    t.DATE_SOLDE AS [Date solde],
    
    -- Informations du crédit
    p.NUMERO_PRET AS [N° contrat],
    ecv.num_manuel AS [N° manuel],
    ecv.date_effet AS [Date décaissement],
    ecv.mtt_pret AS [Montant crédit],
    
    -- Informations du client
    a.ID AS [ID Client],
    a.NOM_ADHERENT AS [Nom client],
    a.CODE AS [Code client],
    ecv.telephone AS [Téléphone],
    
    -- Informations de l'agence et gestionnaire
    ps.NOM AS [Agence],
    ecv.gestionnaire_pret AS [Gestionnaire],
    ecv.superviseur AS [Superviseur],
    
    -- Informations du produit
    ecv.produit AS [Produit],
    ecv.objet_fin AS [Objet fin.],
    ecv.source_fin AS [Source financement],
    ecv.type_source_fin AS [Type crédit],
    ecv.terme_credit AS [Type terme crédit],
    ecv.periodicite AS [Périodicité],
    ecv.devise AS [Devise],
    
    -- Calculs supplémentaires
    DATEDIFF(day, GETDATE(), t.DATE_ECHEANCE) AS [Jours restants],
    CASE 
        WHEN t.DATE_SOLDE IS NOT NULL THEN 'SOLDE'
        WHEN t.DATE_ECHEANCE < GETDATE() THEN 'ÉCHUE'
        WHEN t.DATE_ECHEANCE BETWEEN GETDATE() AND DATEADD(day, 7, GETDATE()) THEN 'URGENT'
        ELSE 'À VENIR'
    END AS [Statut échéance]

FROM TABAMOR t
INNER JOIN CYCLES_PRET cp ON t.ID_CYCLE_PRET = cp.ID
INNER JOIN PRETS p ON cp.ID_PRET = p.ID
LEFT JOIN extra_credits_view ecv ON p.ID = ecv.id_pret
LEFT JOIN ADHERENTS a ON ecv.id_client = a.ID
LEFT JOIN POINTS_SERVICE ps ON a.ID_AGENCE = ps.ID

WHERE 1=1
    -- Filtre sur la période d'échéance
    AND t.DATE_ECHEANCE BETWEEN @DateDebut AND @DateFin
    
    -- Filtre sur les crédits non soldés (si on veut seulement les échéances à payer)
    AND t.DATE_SOLDE IS NULL
    
    -- FILTRE AGENCE
    AND (@Agence = '' OR ps.NOM LIKE '%' + @Agence + '%')
    
    -- FILTRE GESTIONNAIRE
    AND (@Gestionnaire = '' OR ecv.gestionnaire_pret LIKE '%' + @Gestionnaire + '%')
    
    -- FILTRE CLIENT
    AND (
        @Client = '' 
        OR a.ID LIKE '%' + @Client + '%'
        OR a.NOM_ADHERENT LIKE '%' + @Client + '%'
        OR a.CODE LIKE '%' + @Client + '%'
    )
    
    -- FILTRE PRODUIT
    AND (@Produit = '' OR ecv.produit LIKE '%' + @Produit + '%')
    
    -- FILTRE TYPE CRÉDIT
    AND (@TypeCredit = '' OR ecv.type_source_fin LIKE '%' + @TypeCredit + '%')
    
    -- FILTRE NUMÉRO CONTRAT
    AND (@NumContrat = '' OR p.NUMERO_PRET LIKE '%' + @NumContrat + '%')

ORDER BY 
    t.DATE_ECHEANCE,
    ps.NOM,
    ecv.gestionnaire_pret;