# planning_executor.py - Exécute les planifications
import pyodbc
import pandas as pd
from datetime import datetime
from email_sender import send_email, EMAIL_CONFIG
import io
from openpyxl.utils import get_column_letter

def get_connection():
    return pyodbc.connect(
        "Driver={SQL Server};"
        "Server=localhost\\SQLEXPRESS;"
        "Database=REMUCI_VISION;"
        "Trusted_Connection=yes;"
        "Timeout=30;",
        autocommit=True
    )

def get_rapport_data(type_extrait):
    """Récupère les données selon le type d'extraction"""
    conn = get_connection()
    
    try:
        if type_extrait == 'credits':
            query = """
            SELECT TOP 100
                ecv.num_manuel AS [N° Contrat],
                ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
                ecv.mtt_pret AS [Montant (FCFA)],
                ecv.date_effet AS [Date déblocage],
                ecv.nom_agence AS [Agence],
                ecv.gestionnaire_pret AS [Gestionnaire]
            FROM dbo.extra_credits_materialized ecv
            WHERE ecv.date_effet >= DATEADD(day, -7, GETDATE())
            ORDER BY ecv.date_effet DESC
            """
            df = pd.read_sql(query, conn)
            return df, "Crédits débloqués"
        
        elif type_extrait == 'unpaid':
            query = """
            SELECT 
                num_manuel AS [N° Contrat],
                nom_client + ' ' + ISNULL(prenoms_client, '') AS [Client],
                mtt_pret AS [Montant (FCFA)],
                date_fin_echeance AS [Date échéance],
                DATEDIFF(day, date_fin_echeance, GETDATE()) AS [Jours retard],
                nom_agence AS [Agence]
            FROM dbo.extra_credits_materialized
            WHERE date_fin_echeance < GETDATE()
              AND (date_solde IS NULL OR date_solde > date_fin_echeance)
            ORDER BY [Jours retard] DESC
            """
            df = pd.read_sql(query, conn)
            return df, "Crédits impayés"
        
        elif type_extrait == 'clients':
            query = """
            SELECT DISTINCT TOP 100
                code_client AS [Code client],
                nom_client + ' ' + ISNULL(prenoms_client, '') AS [Client],
                date_adhesion AS [Date adhésion],
                telephone AS [Téléphone],
                nom_agence AS [Agence]
            FROM dbo.extra_credits_materialized
            WHERE date_adhesion >= DATEADD(day, -30, GETDATE())
            ORDER BY date_adhesion DESC
            """
            df = pd.read_sql(query, conn)
            return df, "Nouveaux clients"
        
        elif type_extrait == 'comptes':
            query = """
            SELECT TOP 100
                NUM_CPTE AS [Numéro Compte],
                LIBELLE AS [Libellé],
                DATE_OUVERTURE AS [Date ouverture],
                (SELECT NOM_ADHERENT FROM ADHERENTS WHERE ID_COMPTE_ADHERENT = COMPTES.ID) AS [Client]
            FROM COMPTES
            WHERE ETAT = 'O'
            ORDER BY DATE_OUVERTURE DESC
            """
            df = pd.read_sql(query, conn)
            return df, "Comptes ouverts"
        
        else:
            return None, "Type inconnu"
    except Exception as e:
        print(f"Erreur: {e}")
        return None, "Erreur"
    finally:
        conn.close()


def dataframe_to_excel_bytes(df, sheet_name="Données"):
    if df is None or df.empty:
        return None
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()


def check_and_execute_plannings():
    """Vérifie les planifications à exécuter"""
    conn = get_connection()
    
    try:
        now = datetime.now()
        query = """
        SELECT * FROM dbo.PLANIFICATIONS 
        WHERE ACTIF = 1 
        AND DATE_HEURE <= ?
        AND (DERNIER_ENVOI IS NULL OR DATEDIFF(day, DERNIER_ENVOI, ?) >= 
            CASE FREQUENCE 
                WHEN 'daily' THEN 1 
                WHEN 'weekly' THEN 7 
                WHEN 'monthly' THEN 30 
                ELSE 999 
            END)
        """
        df = pd.read_sql(query, conn, params=[now, now])
        
        for _, row in df.iterrows():
            print(f"Exécution de la planification {row['ID']}: {row['TYPE_EXTRAIT']}")
            
            # Récupérer les données
            df_data, titre = get_rapport_data(row['TYPE_EXTRAIT'])
            
            if df_data is not None and not df_data.empty:
                # Créer l'email
                excel_content = dataframe_to_excel_bytes(df_data)
                
                html_body = f"""
                <html>
                <body style="font-family: Arial;">
                    <div style="background: #1a472a; color: white; padding: 20px; text-align: center;">
                        <h2>REMU-CI VisionExtract</h2>
                        <h3>Extraction planifiée: {titre}</h3>
                    </div>
                    <div style="padding: 20px;">
                        <p>Veuillez trouver ci-joint le fichier Excel demandé.</p>
                        <p>Date d'exécution: {now.strftime('%d/%m/%Y %H:%M:%S')}</p>
                        <hr>
                        <p style="font-size: 12px; color: #666;">Email généré automatiquement</p>
                    </div>
                </body>
                </html>
                """
                
                destinataires = [email.strip() for email in row['DESTINATAIRES'].split(',')]
                
                success = send_email(
                    to_emails=destinataires,
                    subject=f"[REMU-CI] {titre} - {now.strftime('%d/%m/%Y')}",
                    html_body=html_body,
                    attachments=[(f"{titre}_{now.strftime('%Y%m%d')}.xlsx", excel_content)]
                )
                
                # Mettre à jour le statut
                update_query = """
                UPDATE dbo.PLANIFICATIONS 
                SET DERNIER_ENVOI = ?, STATUT = ? 
                WHERE ID = ?
                """
                cursor = conn.cursor()
                cursor.execute(update_query, (now, 'ENVOYÉ' if success else 'ÉCHEC', row['ID']))
                conn.commit()
                
                print(f"  ✅ Envoi à {len(destinataires)} destinataires: {'Succès' if success else 'Échec'}")
            else:
                print(f"  ⚠️ Aucune donnée pour {row['TYPE_EXTRAIT']}")
    
    except Exception as e:
        print(f"Erreur: {e}")
    finally:
        conn.close()


if __name__ == "__main__":
    print("Vérification des planifications...")
    check_and_execute_plannings()