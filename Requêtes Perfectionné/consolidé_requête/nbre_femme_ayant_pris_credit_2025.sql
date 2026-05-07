-- Nombre de femmes ayant pris un crédit en 2025
SELECT 
    COUNT(DISTINCT ecv.id_client) AS [Nombre de femmes avec crédit],
    ISNULL(SUM(ecv.mtt_pret), 0) AS [Montant total crédits]
FROM dbo.extra_credits_view ecv
INNER JOIN dbo.extra_clients_view ecl ON ecv.id_client = ecl.id_client
WHERE YEAR(ecv.date_effet) = 2025
    AND ecl.sexe = 'F'
    AND ecv.date_effet IS NOT NULL
