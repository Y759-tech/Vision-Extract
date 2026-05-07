
DECLARE @Agence VARCHAR(100) = '';
DECLARE @Gestionnaire VARCHAR(100) = '';
DECLARE @DateDebut DATE = '2024-01-30';
DECLARE @DateFin DATE = '2024-01-30';

SELECT
    ecv.num_manuel AS [N° manuel],
    ecv.code_client AS [ID Client],
    ecv.nom_client + ' ' + ecv.prenoms_client AS [Client],
    ecv.date_adhesion AS [Date adhésion],
    ecv.nom_agence AS [Agence],
    ecv.gestionnaire_pret AS [Gestionnaire]
FROM dbo.extra_credits_view ecv
WHERE ecv.date_adhesion BETWEEN @DateDebut AND @DateFin
  AND (@Agence = '' OR ecv.nom_agence = @Agence)
  AND (@Gestionnaire = '' OR ecv.gestionnaire_pret = @Gestionnaire OR @Gestionnaire = 'Tous les gestionnaires')
ORDER BY ecv.date_adhesion DESC;