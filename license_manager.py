# license_manager.py - Système de gestion des licences
import hashlib
import uuid
import json
import os
from datetime import datetime, timedelta
import getpass
import platform

class LicenseManager:
    def __init__(self, license_file="license.json"):
        self.license_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), license_file)
        self.license_data = None
        self.load_license()
    
    def get_machine_id(self):
        """Génère un identifiant unique basé sur le matériel"""
        system_info = f"{platform.processor()}{platform.node()}{platform.system()}"
        mac_address = self.get_mac_address()
        drive_serial = self.get_drive_serial()
        
        unique_string = f"{system_info}{mac_address}{drive_serial}"
        machine_id = hashlib.sha256(unique_string.encode()).hexdigest()[:20].upper()
        
        return '-'.join([machine_id[i:i+5] for i in range(0, 20, 5)])
    
    def get_mac_address(self):
        """Récupère l'adresse MAC"""
        try:
            mac = uuid.getnode()
            return ':'.join(('%012X' % mac)[i:i+2] for i in range(0, 12, 2))
        except:
            return "00-00-00-00-00-00"
    
    def get_drive_serial(self):
        """Récupère le numéro de série du disque"""
        try:
            import subprocess
            result = subprocess.run(['vol', 'C:'], capture_output=True, text=True)
            import re
            serial = re.search(r'numéro de série\s+:\s+(\S+)', result.stdout, re.IGNORECASE)
            if serial:
                return serial.group(1)
        except:
            pass
        return "0000-0000"
    
    def generate_license_key(self, machine_id, expiry_days=180, max_users=5):
        """Génère une clé de licence"""
        expiry_date = (datetime.now() + timedelta(days=expiry_days)).strftime('%Y%m%d')
        data = f"{machine_id}{expiry_date}{max_users}"
        signature = hashlib.sha256(data.encode()).hexdigest()[:16].upper()
        
        license_key = f"{signature[:4]}-{signature[4:8]}-{signature[8:12]}-{signature[12:16]}"
        
        license_data = {
            "machine_id": machine_id,
            "expiry_date": expiry_date,
            "max_users": max_users,
            "signature": signature,
            "activated": True,
            "activation_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        return license_key, license_data
    
    def validate_license(self, license_key):
        """Valide une clé de licence"""
        try:
            signature = license_key.replace('-', '')
            
            if not os.path.exists(self.license_file):
                return False, "Licence non activée"
            
            with open(self.license_file, 'r') as f:
                stored_data = json.load(f)
            
            if stored_data.get('signature') != signature:
                return False, "Signature de licence invalide"
            
            current_machine_id = self.get_machine_id()
            if stored_data.get('machine_id') != current_machine_id:
                return False, "Cette licence n'est pas valide pour cette machine"
            
            expiry_date = datetime.strptime(stored_data.get('expiry_date'), '%Y%m%d')
            if expiry_date < datetime.now():
                return False, f"Licence expirée depuis le {expiry_date.strftime('%d/%m/%Y')}"
            
            days_left = (expiry_date - datetime.now()).days
            return True, f"Licence valide (expire dans {days_left} jours)"
            
        except Exception as e:
            return False, f"Erreur de validation: {str(e)}"
    
    def activate_license(self, license_key):
        """Active une licence"""
        license_key_clean = license_key.replace('-', '').upper()
        if len(license_key_clean) != 16:
            return False, "Format de clé invalide (16 caractères requis)"
        
        try:
            current_machine_id = self.get_machine_id()
            
            license_data = {
                "machine_id": current_machine_id,
                "expiry_date": (datetime.now() + timedelta(days=180)).strftime('%Y%m%d'),
                "max_users": 5,
                "signature": license_key_clean,
                "activated": True,
                "activation_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            with open(self.license_file, 'w') as f:
                json.dump(license_data, f, indent=4)
            
            self.license_data = license_data
            return True, "Licence activée avec succès"
            
        except Exception as e:
            return False, f"Erreur d'activation: {str(e)}"
    
    def load_license(self):
        """Charge les données de licence"""
        if os.path.exists(self.license_file):
            try:
                with open(self.license_file, 'r') as f:
                    self.license_data = json.load(f)
                return True
            except:
                self.license_data = None
        return False
    
    def is_licensed(self):
        """Vérifie si l'application est licenciée"""
        if not self.license_data:
            return False
        return self.validate_license(self.license_data.get('signature', ''))[0]
    
    def get_license_info(self):
        """Retourne les informations de licence"""
        if not self.license_data:
            return None
        
        expiry_date = datetime.strptime(self.license_data.get('expiry_date'), '%Y%m%d')
        days_left = (expiry_date - datetime.now()).days
        
        return {
            "machine_id": self.license_data.get('machine_id'),
            "expiry_date": expiry_date.strftime('%d/%m/%Y'),
            "days_left": days_left,
            "activation_date": self.license_data.get('activation_date'),
            "status": "active" if days_left > 0 else "expired"
        }


# Pour générer une licence
if __name__ == "__main__":
    lm = LicenseManager()
    machine_id = lm.get_machine_id()
    print(f"Machine ID de ce poste : {machine_id}")
    print()
    
    # Générer une licence de 180 jours (6 mois)
    key, data = lm.generate_license_key(machine_id, expiry_days=180)
    
    print("=" * 60)
    print("🔑 CLÉ DE LICENCE (6 MOIS)")
    print("=" * 60)
    print(f"Machine ID : {machine_id}")
    print(f"Clé de licence : {key}")
    print(f"Expiration : {data['expiry_date']}")
    print("=" * 60)
    print("⚠️  Conservez cette clé précieusement !")
    print("=" * 60)