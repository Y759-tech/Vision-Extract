# generate_license_for_client.py - À UTILISER POUR GÉNÉRER UNE LICENCE POUR UN CLIENT
from license_manager import LicenseManager

print("=" * 60)
print("🔑 GÉNÉRATION DE LICENCE POUR CLIENT")
print("=" * 60)

# Demander le Machine ID du client
machine_id = input("Entrez le Machine ID du client : ").strip().upper()

# Demander la durée en mois
mois = input("Durée de la licence (mois) [défaut: 6] : ").strip()
if mois.isdigit():
    days = int(mois) * 30
else:
    days = 180  # 6 mois par défaut

# Générer la licence
lm = LicenseManager()
license_key, license_data = lm.generate_license_key(machine_id, expiry_days=days)

print("\n" + "=" * 60)
print("✅ LICENCE GÉNÉRÉE")
print("=" * 60)
print(f"Machine ID : {machine_id}")
print(f"Clé de licence : {license_key}")
print(f"Date d'expiration : {license_data['expiry_date']}")
print(f"Durée : {days} jours ({days//30} mois)")
print("=" * 60)
print("\n📧 Envoyez cette clé à votre client :")
print(f"   {license_key}")
print("=" * 60)