-- Déclaration des variables avec des valeurs fixes
DECLARE @Agence VARCHAR(100) = 'REMUCI : AGENCE-BONOUA';
DECLARE @Gestionnaire VARCHAR(100) = '';
DECLARE @DateDebut DATE = '2025-01-01';
DECLARE @DateFin DATE = '2025-01-31';

WITH CreditsUniques AS (
    SELECT DISTINCT
        dc.ID,
        dc.NUM_DOSSIER,
        ecv.num_manuel,
        ecv.nom_client + ' ' + ecv.prenoms_client AS Client,
        ecv.mtt_pret AS Montant,
        ecv.date_effet AS DateDeblocage,
        ecv.date_fin_echeance AS FinEcheance,
        ecv.nb_echeance AS NbEcheances,
        dc.NBRE_DIFFERE AS NbDifferes,
        ecv.taux AS TauxBenef,
        ecv.nom_agence AS Agence,
        ecv.gestionnaire_pret AS Gestionnaire,
        ecv.code_client AS CodeClient
    FROM dbo.DOSSIERS_CREDIT dc
    INNER JOIN dbo.extra_credits_view ecv ON ecv.id_dossier = dc.ID
    WHERE dc.ETAT_DOSSIER = 'ACCORDEE'
      AND ecv.date_effet IS NOT NULL
      AND ecv.date_effet BETWEEN @DateDebut AND @DateFin
      AND (@Agence = '' OR ecv.nom_agence = @Agence)
      AND (@Gestionnaire = '' OR ecv.gestionnaire_pret = @Gestionnaire OR @Gestionnaire = 'Tous les gestionnaires')
)
SELECT 
    NUM_DOSSIER AS [N° dossier],
    num_manuel AS [N° manuel],
    Client,
    Montant,
    DateDeblocage AS [Date déblocage],
    FinEcheance AS [Fin échéance],
    NbEcheances AS [Nb échéances],
    NbDifferes AS [Nb différés],
    TauxBenef AS [Taux bénéfice],
    Agence,
    Gestionnaire,
    CodeClient AS [Code client]
FROM CreditsUniques
ORDER BY DateDeblocage DESC;