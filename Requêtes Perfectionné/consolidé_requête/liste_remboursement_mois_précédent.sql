-- ===========================================
-- REQUÊTE : Remboursements du mois précédent
-- DESCRIPTION : Liste tous les remboursements effectués le mois dernier
-- ===========================================

-- ========== PARAMÈTRES À MODIFIER ==========
DECLARE @DateReference DATE = GETDATE();                    -- Date de référence (aujourd'hui)
DECLARE @Agence VARCHAR(100) = '';                          -- Ex: 'REMUCI : AGENCE-BONOUA' ou ''
DECLARE @Gestionnaire VARCHAR(100) = '';                    -- Nom du gestionnaire ou ''
DECLARE @Client VARCHAR(100) = '';                          -- ID client ou nom ou ''
DECLARE @TypeCredit VARCHAR(100) = '';                      -- Type de crédit ou ''
-- ============================================

-- Calcul du mois précédent
DECLARE @MoisPrecedent DATE = DATEADD(month, -1, @DateReference);
DECLARE @DebutMois DATE = DATEFROMPARTS(YEAR(@MoisPrecedent), MONTH(@MoisPrecedent), 1);
DECLARE @FinMois DATE = EOMONTH(@DebutMois);

SELECT 
    -- Informations du remboursement
    rc.ID AS [ID Remboursement],
    rc.ID_OPERATION_CRD AS [ID Opération],
    o.DATE_OPERATION AS [Date remboursement],
    o.DATE_SAISIE AS [Date saisie],
    
    -- Détails du remboursement
    rc.CAPITAL AS [Capital remboursé],
    rc.INTERET AS [Intérêt payé],
    rc.PENALITE AS [Pénalité],
    rc.COMMISSION AS [Commission],
    (ISNULL(rc.CAPITAL, 0) + ISNULL(rc.INTERET, 0) + 
     ISNULL(rc.PENALITE, 0) + ISNULL(rc.COMMISSION, 0)) AS [Montant total],
    
    -- Lien avec l'échéance
    rc.ID_TABAMORT AS [ID Échéance],
    t.DATE_ECHEANCE AS [Date échéance concernée],
    t.CAPITAL AS [Capital prévu],
    t.INTERET AS [Intérêt prévu],
    
    -- Informations du crédit
    p.NUMERO_PRET AS [N° contrat],
    ecv.num_manuel AS [N° manuel],
    ecv.date_effet AS [Date décaissement],
    ecv.mtt_pret AS [Montant crédit initial],
    
    -- Informations du client
    a.ID AS [ID Client],
    a.NOM_ADHERENT AS [Nom client],
    a.CODE AS [Code client],
    ecv.telephone AS [Téléphone],
    
    -- Informations de l'agence
    ps.NOM AS [Agence],
    ecv.gestionnaire_pret AS [Gestionnaire],
    ecv.produit AS [Type crédit],
    
    -- Mode de paiement (si disponible)
    o.MODE_PAIEMENT AS [Mode paiement],
    o.NUMERO_RECU AS [N° reçu],
    o.NUM_TRANSACTION AS [N° transaction]

FROM REMBOURS_CRD rc
INNER JOIN OPERATIONS_CRD oc ON rc.ID_OPERATION_CRD = oc.ID
INNER JOIN OPERATIONS o ON oc.ID_OPERATION = o.ID
LEFT JOIN TABAMOR t ON rc.ID_TABAMORT = t.ID
INNER JOIN PRETS p ON oc.ID_PRET = p.ID
LEFT JOIN extra_credits_view ecv ON p.ID = ecv.id_pret
LEFT JOIN ADHERENTS a ON ecv.id_client = a.ID
LEFT JOIN POINTS_SERVICE ps ON a.ID_AGENCE = ps.ID

WHERE 1=1
    -- Filtre sur la période (mois précédent)
    AND o.DATE_OPERATION BETWEEN @DebutMois AND @FinMois
    
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
    
    -- FILTRE TYPE CRÉDIT
    AND (@TypeCredit = '' OR ecv.produit LIKE '%' + @TypeCredit + '%')

ORDER BY 
    o.DATE_OPERATION DESC,
    ps.NOM,
    ecv.gestionnaire_pret;