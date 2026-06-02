# email_sender.py - Service d'envoi d'emails automatisés pour Windows
import os
import sys
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import pandas as pd
import pyodbc
from io import BytesIO
from openpyxl.utils import get_column_letter

# ============================================================================
# CONFIGURATION
# ============================================================================

DB_SERVER = "localhost\\SQLEXPRESS"
DB_NAME = "REMUCI_VISION"

# Configuration email - À MODIFIER SELON VOS BESOINS
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'smtp_user': 'sollomarius@gmail.com',
    'smtp_password': 'hzka capg ysyb sluf',
    'from_email': 'sollomarius@gmail.com',
    'from_name': 'REMU-CI VisionExtract'
}

# Destinataires - À MODIFIER AVEC LES VRAIS EMAILS
RECIPIENTS = {
    'finance': ['finance@remu-ci.com', 'daf@remu-ci.com'],
    'exploitation': ['exploitation@remu-ci.com', 'directeur@remu-ci.com'],
    'credit': ['credit@remu-ci.com', 'responsable_credit@remu-ci.com'],
    'audit': ['audit@remu-ci.com', 'interne@remu-ci.com'],
    'direction': ['direction@remu-ci.com']
}

# Configuration des rapports par service
REPORTS_CONFIG = {
    'finance': ['credits_debloques', 'balance_j1', 'grand_livre', 'remboursements'],
    'exploitation': ['nouveaux_clients', 'comptes_ouverts', 'clients_actifs', 'parts_sociales'],
    'credit': ['credits_impayes', 'echeances_futures', 'credits_debloques'],
    'audit': ['balance_j1', 'credits_impayes', 'grand_livre']
}

# ============================================================================
# LOGS
# ============================================================================

LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, f'email_log_{datetime.now().strftime("%Y%m")}.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('EmailSender')

# ============================================================================
# CONNEXION BASE DE DONNÉES
# ============================================================================

def get_connection():
    try:
        conn_str = f"Driver={{SQL Server}};Server={DB_SERVER};Database={DB_NAME};Trusted_Connection=yes;Timeout=60;"
        conn = pyodbc.connect(conn_str, autocommit=True)
        logger.info("Connexion SQL Server réussie")
        return conn
    except Exception as e:
        logger.error(f"Erreur de connexion SQL: {e}")
        return None

# ============================================================================
# FONCTIONS DE DATE
# ============================================================================

def get_yesterday():
    return (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

def get_last_month_range():
    today = datetime.now()
    first_day_current = today.replace(day=1)
    last_day_previous = first_day_current - timedelta(days=1)
    first_day_previous = last_day_previous.replace(day=1)
    return first_day_previous.strftime('%Y-%m-%d'), last_day_previous.strftime('%Y-%m-%d')

def get_next_month_range():
    today = datetime.now()
    if today.month == 12:
        next_month = today.replace(year=today.year+1, month=1, day=1)
    else:
        next_month = today.replace(month=today.month+1, day=1)
    last_day_next = (next_month.replace(month=next_month.month+1, day=1) - timedelta(days=1))
    return next_month.strftime('%Y-%m-%d'), last_day_next.strftime('%Y-%m-%d')

# ============================================================================
# FONCTIONS D'EXTRACTION
# ============================================================================

def extract_credits_debloques(date_debut=None, date_fin=None):
    if not date_debut:
        date_debut = get_yesterday()
    if not date_fin:
        date_fin = get_yesterday()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            ecv.num_manuel AS [N° Contrat],
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.mtt_pret AS [Montant (FCFA)],
            ecv.date_effet AS [Date déblocage],
            ecv.date_fin_echeance AS [Date fin échéance],
            ecv.nb_echeance AS [Nb échéances],
            ecv.nom_agence AS [Agence],
            ecv.gestionnaire_pret AS [Gestionnaire],
            ecv.produit AS [Produit],
            ecv.code_client AS [Code client]
        FROM dbo.extra_credits_view ecv
        WHERE CAST(ecv.date_effet AS DATE) BETWEEN ? AND ?
        ORDER BY ecv.date_effet DESC
        """
        df = pd.read_sql(query, conn, params=[date_debut, date_fin])
        logger.info(f"Credits débloqués: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur credits_debloques: {e}")
        return None
    finally:
        conn.close()

def extract_nouveaux_clients(date_debut=None, date_fin=None):
    if not date_debut:
        date_debut = get_yesterday()
    if not date_fin:
        date_fin = get_yesterday()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT DISTINCT
            ecv.code_client AS [Code client],
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.date_adhesion AS [Date adhésion],
            ecv.telephone AS [Téléphone],
            ecv.nom_agence AS [Agence],
            ecv.sexe AS [Sexe]
        FROM dbo.extra_credits_view ecv
        WHERE CAST(ecv.date_adhesion AS DATE) BETWEEN ? AND ?
        ORDER BY ecv.date_adhesion DESC
        """
        df = pd.read_sql(query, conn, params=[date_debut, date_fin])
        logger.info(f"Nouveaux clients: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur nouveaux_clients: {e}")
        return None
    finally:
        conn.close()

def extract_comptes_ouverts(date_debut=None, date_fin=None):
    if not date_debut:
        date_debut = get_yesterday()
    if not date_fin:
        date_fin = get_yesterday()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            c.NUM_CPTE as [Numéro Compte],
            c.LIBELLE as [Libellé],
            c.DATE_OUVERTURE as [Date ouverture],
            a.NOM_ADHERENT as [Client],
            ISNULL((
                SELECT SUM(MONTANT_OPERATION * CASE WHEN SENS = 'C' THEN 1 ELSE -1 END)
                FROM HDPM h WHERE h.ID_COMPTE = c.ID
            ), 0) as [Solde actuel]
        FROM COMPTES c
        LEFT JOIN ADHERENTS a ON c.ID = a.ID_COMPTE_ADHERENT
        WHERE c.ETAT = 'O'
          AND CAST(c.DATE_OUVERTURE AS DATE) BETWEEN ? AND ?
        ORDER BY c.DATE_OUVERTURE DESC
        """
        df = pd.read_sql(query, conn, params=[date_debut, date_fin])
        logger.info(f"Comptes ouverts: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur comptes_ouverts: {e}")
        return None
    finally:
        conn.close()

def extract_credits_impayes(jours_min=1, jours_max=9999):
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            ecv.num_manuel AS [N° Contrat],
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.mtt_pret AS [Montant (FCFA)],
            ecv.date_fin_echeance AS [Date échéance],
            DATEDIFF(day, ecv.date_fin_echeance, GETDATE()) AS [Jours retard],
            ecv.nom_agence AS [Agence],
            ecv.gestionnaire_pret AS [Gestionnaire],
            ecv.produit AS [Produit],
            ecv.telephone AS [Téléphone]
        FROM dbo.extra_credits_view ecv
        WHERE ecv.date_fin_echeance < GETDATE()
          AND (ecv.date_solde IS NULL OR ecv.date_solde > ecv.date_fin_echeance)
          AND DATEDIFF(day, ecv.date_fin_echeance, GETDATE()) BETWEEN ? AND ?
        ORDER BY [Jours retard] DESC
        """
        df = pd.read_sql(query, conn, params=[jours_min, jours_max])
        logger.info(f"Crédits impayés: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur credits_impayes: {e}")
        return None
    finally:
        conn.close()

def extract_clients_actifs():
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT DISTINCT
            ecv.code_client AS [Code client],
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.telephone AS [Téléphone],
            ecv.nom_agence AS [Agence],
            ecv.gestionnaire_pret AS [Gestionnaire],
            MAX(ecv.date_effet) AS [Dernier crédit],
            COUNT(ecv.id_pret) AS [Nb crédits]
        FROM dbo.extra_credits_view ecv
        WHERE ecv.date_effet >= DATEADD(month, -3, GETDATE())
        GROUP BY ecv.code_client, ecv.nom_client, ecv.prenoms_client, ecv.telephone, ecv.nom_agence, ecv.gestionnaire_pret
        ORDER BY [Dernier crédit] DESC
        """
        df = pd.read_sql(query, conn)
        logger.info(f"Clients actifs: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur clients_actifs: {e}")
        return None
    finally:
        conn.close()

def extract_parts_sociales():
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT TOP 500
            a.NOM_ADHERENT AS [Client],
            a.CODE AS [Code],
            ops.NOMBRE AS [Nombre parts],
            (ops.NOMBRE * psv.VALEUR) AS [Montant (FCFA)],
            o.DATE_OPERATION AS [Date]
        FROM OPERATIONS_PART_SOC ops
        LEFT JOIN ADHERENTS a ON ops.ID_ADHERENT = a.ID
        LEFT JOIN PARTS_SOCIALE psv ON ops.ID_PART_SOCIALE = psv.ID
        LEFT JOIN OPERATIONS o ON ops.ID_OPERATION = o.ID
        WHERE ops.NOMBRE > 0
        ORDER BY o.DATE_OPERATION DESC
        """
        df = pd.read_sql(query, conn)
        logger.info(f"Parts sociales: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur parts_sociales: {e}")
        return None
    finally:
        conn.close()

def extract_echeances_futures():
    start_date, end_date = get_next_month_range()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.num_manuel AS [N° Contrat],
            ecv.mtt_pret AS [Montant crédit (FCFA)],
            ecv.date_fin_echeance AS [Date échéance],
            DATEDIFF(day, GETDATE(), ecv.date_fin_echeance) AS [Jours restants],
            ecv.nom_agence AS [Agence],
            ecv.gestionnaire_pret AS [Gestionnaire],
            ecv.telephone AS [Téléphone]
        FROM dbo.extra_credits_view ecv
        WHERE CAST(ecv.date_fin_echeance AS DATE) BETWEEN ? AND ?
          AND ecv.date_solde IS NULL
        ORDER BY ecv.date_fin_echeance
        """
        df = pd.read_sql(query, conn, params=[start_date, end_date])
        logger.info(f"Échéances futures: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur echeances_futures: {e}")
        return None
    finally:
        conn.close()

def extract_remboursements():
    start_date, end_date = get_last_month_range()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            ecv.nom_client + ' ' + ISNULL(ecv.prenoms_client, '') AS [Client],
            ecv.num_manuel AS [N° Contrat],
            ecv.mtt_pret AS [Montant crédit (FCFA)],
            ecv.date_solde AS [Date remboursement],
            ecv.nom_agence AS [Agence]
        FROM dbo.extra_credits_view ecv
        WHERE CAST(ecv.date_solde AS DATE) BETWEEN ? AND ?
          AND ecv.date_solde IS NOT NULL
        ORDER BY ecv.date_solde DESC
        """
        df = pd.read_sql(query, conn, params=[start_date, end_date])
        logger.info(f"Remboursements: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur remboursements: {e}")
        return None
    finally:
        conn.close()

def extract_balance_j1():
    date_balance = get_yesterday()
    
    conn = get_connection()
    if not conn:
        return None
    
    try:
        query = """
        SELECT 
            c.NUM_CPTE AS [Numéro Compte],
            c.LIBELLE AS [Intitulé],
            ISNULL(SUM(CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END), 0) AS [Débit],
            ISNULL(SUM(CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END), 0) AS [Crédit]
        FROM COMPTES c
        LEFT JOIN HDPM h ON c.ID = h.ID_COMPTE AND CAST(h.DATE_OPERATION AS DATE) <= ?
        WHERE c.ETAT = 'O'
        GROUP BY c.NUM_CPTE, c.LIBELLE
        ORDER BY c.NUM_CPTE
        """
        df = pd.read_sql(query, conn, params=[date_balance])
        
        if not df.empty:
            df['Solde Débiteur'] = df.apply(lambda x: x['Débit'] - x['Crédit'] if (x['Débit'] - x['Crédit']) > 0 else 0, axis=1)
            df['Solde Créditeur'] = df.apply(lambda x: x['Crédit'] - x['Débit'] if (x['Crédit'] - x['Débit']) > 0 else 0, axis=1)
        
        logger.info(f"Balance J-1: {len(df)} comptes")
        return df
    except Exception as e:
        logger.error(f"Erreur balance_j1: {e}")
        return None
    finally:
        conn.close()

def extract_grand_livre(compte=None):
    conn = get_connection()
    if not conn:
        return None
    
    try:
        if compte:
            compte_condition = f"AND c.NUM_CPTE = '{compte}'"
        else:
            compte_condition = ""
        
        query = f"""
        SELECT TOP 5000
            c.NUM_CPTE AS [Numéro Compte],
            c.LIBELLE AS [Libellé],
            h.DATE_OPERATION AS [Date],
            h.LIBELLE_OPERATION AS [Libellé opération],
            h.NUM_PIECE AS [N° Pièce],
            CASE WHEN h.SENS = 'D' THEN h.MONTANT_OPERATION ELSE 0 END AS [Débit],
            CASE WHEN h.SENS = 'C' THEN h.MONTANT_OPERATION ELSE 0 END AS [Crédit]
        FROM COMPTES c
        LEFT JOIN HDPM h ON c.ID = h.ID_COMPTE
        WHERE h.DATE_OPERATION IS NOT NULL
          {compte_condition}
        ORDER BY h.DATE_OPERATION DESC, c.NUM_CPTE
        """
        df = pd.read_sql(query, conn)
        logger.info(f"Grand livre: {len(df)} lignes")
        return df
    except Exception as e:
        logger.error(f"Erreur grand_livre: {e}")
        return None
    finally:
        conn.close()

# ============================================================================
# EXPORT EXCEL
# ============================================================================

def create_multi_sheet_excel(reports_data, report_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in zip(report_names, reports_data):
            if df is not None and not df.empty:
                sheet_name = name[:31]
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

# ============================================================================
# ENVOI EMAIL
# ============================================================================

def send_email(to_emails, subject, html_body, attachments=None):
    if not to_emails:
        logger.warning("Aucun destinataire spécifié")
        return False
    
    to_emails = [email for email in to_emails if email and email.strip()]
    if not to_emails:
        logger.warning("Aucun email valide")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = f"{EMAIL_CONFIG['from_name']} <{EMAIL_CONFIG['from_email']}>"
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))
        
        if attachments:
            for filename, content in attachments:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(content)
                encoders.encode_base64(part)
                from email.header import Header
                part.add_header('Content-Disposition', f'attachment; filename="{Header(filename, "utf-8").encode()}"')
                msg.attach(part)
        
        with smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port']) as server:
            server.starttls()
            server.login(EMAIL_CONFIG['smtp_user'], EMAIL_CONFIG['smtp_password'])
            server.send_message(msg)
        
        logger.info(f"Email envoyé à {len(to_emails)} destinataires")
        return True
        
    except Exception as e:
        logger.error(f"Erreur envoi email: {e}")
        return False

# ============================================================================
# FONCTION PRINCIPALE
# ============================================================================

def send_daily_reports():
    logger.info("=" * 60)
    logger.info("DÉBUT DE L'ENVOI DES RAPPORTS QUOTIDIENS")
    logger.info(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    yesterday = get_yesterday()
    try:
        yesterday_str = datetime.strptime(yesterday, '%Y-%m-%d').strftime('%d/%m/%Y')
    except:
        yesterday_str = yesterday
    
    logger.info("Extraction des données...")
    
    reports = {
        'credits_debloques': extract_credits_debloques(yesterday, yesterday),
        'nouveaux_clients': extract_nouveaux_clients(yesterday, yesterday),
        'comptes_ouverts': extract_comptes_ouverts(yesterday, yesterday),
        'credits_impayes': extract_credits_impayes(),
        'clients_actifs': extract_clients_actifs(),
        'parts_sociales': extract_parts_sociales(),
        'echeances_futures': extract_echeances_futures(),
        'remboursements': extract_remboursements(),
        'balance_j1': extract_balance_j1(),
        'grand_livre': extract_grand_livre()
    }
    
    stats = {}
    for key, df in reports.items():
        stats[key] = len(df) if df is not None else 0
    
    results = []
    
    # SERVICE FINANCE
    finance_reports = []
    finance_names = []
    for report_name in REPORTS_CONFIG['finance']:
        if report_name in reports and reports[report_name] is not None and not reports[report_name].empty:
            finance_reports.append(reports[report_name])
            finance_names.append(report_name.replace('_', ' ').title())
    
    if finance_reports:
        excel_content = create_multi_sheet_excel(finance_reports, finance_names)
        html_body = f"""
        <html>
        <head><style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background: #1a472a; color: white; padding: 20px; text-align: center; }}
            .content {{ padding: 20px; }}
            .stats {{ background: #f5f5f5; padding: 15px; border-radius: 10px; margin: 20px 0; }}
            .footer {{ font-size: 12px; color: #666; text-align: center; margin-top: 30px; }}
        </style></head>
        <body>
            <div class="header">
                <h2>REMU-CI VisionExtract</h2>
                <h3>Rapport Quotidien - Service Finance</h3>
                <p>Date des données: {yesterday_str}</p>
            </div>
            <div class="content">
                <h3>📊 Récapitulatif</h3>
                <div class="stats">
                    <ul>
                        <li>💰 Crédits débloqués: {stats.get('credits_debloques', 0)} dossiers</li>
                        <li>🏦 Nouveaux comptes: {stats.get('comptes_ouverts', 0)}</li>
                        <li>💵 Remboursements: {stats.get('remboursements', 0)} opérations</li>
                        <li>📊 Total comptes: {stats.get('balance_j1', 0)}</li>
                    </ul>
                </div>
                <p>Fichier Excel joint avec les détails.</p>
            </div>
            <div class="footer">
                <p>Email automatique - Ne pas répondre</p>
            </div>
        </body>
        </html>
        """
        success = send_email(RECIPIENTS['finance'], f"[REMU-CI] Rapport Finance - {yesterday_str}", html_body, [(f"Finance_Report_{yesterday}.xlsx", excel_content)])
        results.append(('Finance', success))
    
    # SERVICE EXPLOITATION
    exploitation_reports = []
    exploitation_names = []
    for report_name in REPORTS_CONFIG['exploitation']:
        if report_name in reports and reports[report_name] is not None and not reports[report_name].empty:
            exploitation_reports.append(reports[report_name])
            exploitation_names.append(report_name.replace('_', ' ').title())
    
    if exploitation_reports:
        excel_content = create_multi_sheet_excel(exploitation_reports, exploitation_names)
        html_body = f"""
        <html>
        <head><style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background: #1a472a; color: white; padding: 20px; text-align: center; }}
            .content {{ padding: 20px; }}
            .stats {{ background: #f5f5f5; padding: 15px; border-radius: 10px; margin: 20px 0; }}
        </style></head>
        <body>
            <div class="header">
                <h2>REMU-CI VisionExtract</h2>
                <h3>Rapport Quotidien - Exploitation</h3>
                <p>Date: {yesterday_str}</p>
            </div>
            <div class="content">
                <div class="stats">
                    <ul>
                        <li>👥 Nouveaux clients: {stats.get('nouveaux_clients', 0)}</li>
                        <li>✅ Clients actifs: {stats.get('clients_actifs', 0)}</li>
                        <li>🏦 Nouveaux comptes: {stats.get('comptes_ouverts', 0)}</li>
                        <li>📈 Parts sociales: {stats.get('parts_sociales', 0)}</li>
                    </ul>
                </div>
                <p>Fichier Excel joint.</p>
            </div>
        </body>
        </html>
        """
        success = send_email(RECIPIENTS['exploitation'], f"[REMU-CI] Rapport Exploitation - {yesterday_str}", html_body, [(f"Exploitation_Report_{yesterday}.xlsx", excel_content)])
        results.append(('Exploitation', success))
    
    # SERVICE CRÉDIT
    credit_reports = []
    credit_names = []
    for report_name in REPORTS_CONFIG['credit']:
        if report_name in reports and reports[report_name] is not None and not reports[report_name].empty:
            credit_reports.append(reports[report_name])
            credit_names.append(report_name.replace('_', ' ').title())
    
    if credit_reports:
        excel_content = create_multi_sheet_excel(credit_reports, credit_names)
        montant_impayes = 0
        if reports.get('credits_impayes') is not None and not reports['credits_impayes'].empty:
            montant_impayes = reports['credits_impayes']['Montant (FCFA)'].sum()
        
        alert_html = f"""
        <div class="alert" style="background: #ffebee; border-left: 4px solid #ef4444; padding: 15px; margin: 20px 0;">
            <strong>⚠️ ALERTE IMPAYÉS</strong><br>
            {stats.get('credits_impayes', 0)} dossiers - {montant_impayes:,.0f} FCFA
        </div>
        """ if stats.get('credits_impayes', 0) > 0 else ""
        
        html_body = f"""
        <html>
        <head><style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background: #1a472a; color: white; padding: 20px; text-align: center; }}
            .content {{ padding: 20px; }}
            .stats {{ background: #f5f5f5; padding: 15px; border-radius: 10px; margin: 20px 0; }}
        </style></head>
        <body>
            <div class="header">
                <h2>REMU-CI VisionExtract</h2>
                <h3>Rapport Quotidien - Crédit</h3>
                <p>Date: {yesterday_str}</p>
            </div>
            <div class="content">
                <div class="stats">
                    <ul>
                        <li>💰 Nouveaux crédits: {stats.get('credits_debloques', 0)}</li>
                        <li>⚠️ Crédits impayés: {stats.get('credits_impayes', 0)}</li>
                        <li>💵 Montant impayé: {montant_impayes:,.0f} FCFA</li>
                        <li>📅 Échéances à venir: {stats.get('echeances_futures', 0)}</li>
                    </ul>
                </div>
                {alert_html}
                <p>Fichier Excel joint.</p>
            </div>
        </body>
        </html>
        """
        success = send_email(RECIPIENTS['credit'], f"[REMU-CI] Rapport Crédit - {yesterday_str}", html_body, [(f"Credit_Report_{yesterday}.xlsx", excel_content)])
        results.append(('Credit', success))
    
    # SERVICE AUDIT
    audit_reports = []
    audit_names = []
    for report_name in REPORTS_CONFIG['audit']:
        if report_name in reports and reports[report_name] is not None and not reports[report_name].empty:
            audit_reports.append(reports[report_name])
            audit_names.append(report_name.replace('_', ' ').title())
    
    if audit_reports:
        excel_content = create_multi_sheet_excel(audit_reports, audit_names)
        html_body = f"""
        <html>
        <head><style>
            .header {{ background: #1a472a; color: white; padding: 20px; text-align: center; }}
            .content {{ padding: 20px; }}
        </style></head>
        <body>
            <div class="header">
                <h2>REMU-CI VisionExtract</h2>
                <h3>Rapport Quotidien - Audit</h3>
                <p>Date: {yesterday_str}</p>
            </div>
            <div class="content">
                <p>Rapports d'audit pour contrôle interne.</p>
                <p><strong>Contenu:</strong> Balance, Impayés, Grand livre, Crédits</p>
            </div>
        </body>
        </html>
        """
        success = send_email(RECIPIENTS['audit'], f"[REMU-CI] Rapport Audit - {yesterday_str}", html_body, [(f"Audit_Report_{yesterday}.xlsx", excel_content)])
        results.append(('Audit', success))
    
    logger.info("=" * 40)
    logger.info("RÉCAPITULATIF:")
    for service, success in results:
        status = "✅ SUCCÈS" if success else "❌ ÉCHEC"
        logger.info(f"   {service}: {status}")
    logger.info("=" * 40)
    
    # Sauvegarde récapitulatif
    summary_file = os.path.join(LOG_DIR, f'summary_{datetime.now().strftime("%Y%m%d")}.txt')
    with open(summary_file, 'w', encoding='utf-8') as f:
        f.write(f"=== RAPPORT VISIONEXTRACT ===\n")
        f.write(f"Date: {datetime.now()}\n")
        f.write(f"Période: {yesterday_str}\n\n")
        for service, success in results:
            f.write(f"{service}: {'SUCCÈS' if success else 'ÉCHEC'}\n")
    
    return results

# ============================================================================
# POINT D'ENTRÉE
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("REMU-CI VisionExtract - Envoi automatique")
    print(f"Démarrage: {datetime.now()}")
    print("=" * 60)
    
    send_daily_reports()
    
    print(f"\nLogs dans: {LOG_DIR}")
    print("=" * 60)