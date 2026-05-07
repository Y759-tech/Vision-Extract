-- ===========================================
-- REQUÊTE : Crédits impayés
-- SOURCE : /api/credits-impayes
-- DESCRIPTION : Liste tous les crédits en retard de paiement
-- ===========================================

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
ORDER BY [Jours retard] DESC;