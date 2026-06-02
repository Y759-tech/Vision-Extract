import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Votre mot de passe (avec espaces)
MOT_DE_PASSE = "hzka capg ysyb sluf"

# Création du message
msg = MIMEMultipart()
msg['From'] = 'sollomarius@gmail.com'
msg['To'] = 'sollomarius@gmail.com'
msg['Subject'] = '✅ VisionExtract - Test réussi !'

corps = """
<html>
<body style="font-family: Arial, sans-serif;">
    <div style="background-color: #1a472a; color: white; padding: 20px; text-align: center;">
        <h2>REMU-CI VisionExtract</h2>
    </div>
    <div style="padding: 20px;">
        <h3>✅ Félicitations !</h3>
        <p>Votre configuration email fonctionne parfaitement.</p>
        <p>Les rapports quotidiens seront envoyés automatiquement.</p>
        <hr>
        <p style="color: #666; font-size: 12px;">Email généré automatiquement - Ne pas répondre</p>
    </div>
</body>
</html>
"""

msg.attach(MIMEText(corps, 'html'))

try:
    print("Connexion au serveur Gmail...")
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    
    print("Authentification...")
    server.login('sollomarius@gmail.com', MOT_DE_PASSE)
    
    print("Envoi de l'email...")
    server.send_message(msg)
    server.quit()
    
    print("=" * 50)
    print("✅ EMAIL ENVOYÉ AVEC SUCCÈS !")
    print("✅ Vérifiez votre boîte de réception")
    print("=" * 50)
    
except Exception as e:
    print("=" * 50)
    print(f"❌ ERREUR: {e}")
    print("=" * 50)