SELECT 
    ecv.nom_agence AS AGENCE,
    MONTH(dc.DATE_DECISION) as Mois,
    YEAR(dc.DATE_DECISION) as Annee,
    COUNT(*) as Nombre_Credits,
    SUM(dc.MONTANT_ACCORDE) as Volume_Credits
FROM dbo.DOSSIERS_CREDIT dc
LEFT JOIN dbo.extra_credits_view ecv ON ecv.id_dossier = dc.ID
WHERE dc.ETAT_DOSSIER = 'ACCORDEE'
  AND dc.DATE_DECISION IS NOT NULL
  AND YEAR(dc.DATE_DECISION) = 2025
  AND ecv.nom_agence = 'REMUCI : AGENCE-BONOUA'
GROUP BY ecv.nom_agence, MONTH(dc.DATE_DECISION), YEAR(dc.DATE_DECISION)
ORDER BY Annee, Mois, AGENCE