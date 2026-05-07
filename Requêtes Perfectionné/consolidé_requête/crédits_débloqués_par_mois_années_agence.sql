WITH CreditsUniques AS (
    SELECT DISTINCT
        dc.ID,
        ecv.nom_agence,
        ecv.mtt_pret,
        ecv.date_effet
    FROM dbo.DOSSIERS_CREDIT dc
    INNER JOIN dbo.extra_credits_view ecv ON ecv.id_dossier = dc.ID
    WHERE dc.ETAT_DOSSIER = 'ACCORDEE'
      AND ecv.date_effet IS NOT NULL
      AND YEAR(ecv.date_effet) = 2025
      AND ecv.nom_agence = 'REMUCI : AGENCE-BONOUA'
)
SELECT 
    nom_agence AS AGENCE,
    MONTH(date_effet) as Mois,
    YEAR(date_effet) as Annee,
    COUNT(ID) as Nombre_Credits_Debloques,
    SUM(mtt_pret) as Volume_Credits_Debloques
FROM CreditsUniques
GROUP BY nom_agence, MONTH(date_effet), YEAR(date_effet)
ORDER BY Annee, Mois, AGENCE