 
DECLARE @DateDebut DATE = '2024-01-01';
DECLARE @DateFin DATE = '2025-12-31';
DECLARE @CodeAgence VARCHAR(50) = '';
DECLARE @IDClient VARCHAR(50) = '';
DECLARE @TypeCompte VARCHAR(50) = '';
DECLARE @SoldeMin FLOAT = NULL;
DECLARE @SoldeMax FLOAT = NULL;

SELECT 
    c.ID as 'Numéro Compte',
    c.NUM_CPTE as 'Référence Compte',
    c.LIBELLE as 'Libellé Compte',
    CASE 
        WHEN c.LIBELLE LIKE 'Epargne libre%' THEN 'Epargne libre'
        WHEN c.LIBELLE LIKE 'Compte à vue%' THEN 'Compte à vue'
        ELSE 'Autre'
    END as 'Type Compte',
    c.DATE_OUVERTURE as 'Date Ouverture',
    c.ETAT as 'Statut',
    a.NOM_ADHERENT as 'Nom du Client',
    a.ID as 'ID Client',
    c.ID_AGENCE as 'Code Agence',
    COALESCE(
        (SELECT SUM(MONTANT_OPERATION * CASE WHEN SENS = 'C' THEN 1 ELSE -1 END)
         FROM HDPM h 
         WHERE h.ID_COMPTE = c.ID),
        0
    ) as 'Solde Actuel'
FROM COMPTES c
LEFT JOIN ADHERENTS a ON c.ID = a.ID_COMPTE_ADHERENT
WHERE c.ETAT = 'O'
    AND c.DATE_OUVERTURE BETWEEN @DateDebut AND @DateFin
    AND (@CodeAgence = '' OR c.ID_AGENCE = @CodeAgence)
    AND (
        @IDClient = '' OR 
        a.ID = @IDClient OR 
        a.ID LIKE '%' + @IDClient
    )
    AND (
        @TypeCompte = '' OR 
        CASE 
            WHEN c.LIBELLE LIKE 'Epargne libre%' THEN 'Epargne libre'
            WHEN c.LIBELLE LIKE 'Compte à vue%' THEN 'Compte à vue'
            ELSE 'Autre'
        END = @TypeCompte
    )
    AND (
        (@SoldeMin IS NULL AND @SoldeMax IS NULL) OR
        (@SoldeMin IS NOT NULL AND @SoldeMax IS NOT NULL AND 
         COALESCE((SELECT SUM(MONTANT_OPERATION * CASE WHEN SENS = 'C' THEN 1 ELSE -1 END)
                   FROM HDPM h WHERE h.ID_COMPTE = c.ID), 0) BETWEEN @SoldeMin AND @SoldeMax) OR
        (@SoldeMin IS NOT NULL AND @SoldeMax IS NULL AND 
         COALESCE((SELECT SUM(MONTANT_OPERATION * CASE WHEN SENS = 'C' THEN 1 ELSE -1 END)
                   FROM HDPM h WHERE h.ID_COMPTE = c.ID), 0) >= @SoldeMin) OR
        (@SoldeMin IS NULL AND @SoldeMax IS NOT NULL AND 
         COALESCE((SELECT SUM(MONTANT_OPERATION * CASE WHEN SENS = 'C' THEN 1 ELSE -1 END)
                   FROM HDPM h WHERE h.ID_COMPTE = c.ID), 0) <= @SoldeMax)
    )
ORDER BY c.DATE_OUVERTURE DESC;