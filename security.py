import re
from flask import abort, request
import html

class Security:
    """Classe de sécurité pour valider et nettoyer les entrées"""
    
    @staticmethod
    def validate_date_format(date_string):
        """Valide le format de date YYYY-MM-DD"""
        if not date_string:
            return True
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_string):
            abort(400, "Format de date invalide. Utilisez YYYY-MM-DD")
        return True
    
    @staticmethod
    def sanitize_string(input_string, max_length=100):
        """Nettoie et valide les chaînes de caractères"""
        if not input_string:
            return ""
        
        if len(input_string) > max_length:
            abort(400, f"Paramètre trop long (max {max_length} caractères)")
        
        # Échapper les caractères HTML
        sanitized = html.escape(input_string.strip())
        return sanitized
    
    @staticmethod
    def validate_numeric(value, param_name, min_val=None, max_val=None):
        """Valide les valeurs numériques"""
        if not value:
            return None
        
        try:
            num = float(value)
            if min_val is not None and num < min_val:
                abort(400, f"{param_name} doit être >= {min_val}")
            if max_val is not None and num > max_val:
                abort(400, f"{param_name} doit être <= {max_val}")
            return num
        except ValueError:
            abort(400, f"{param_name} doit être un nombre valide")
    
    @staticmethod
    def validate_agence(agence):
        """Valide le nom de l'agence"""
        valid_agences = [
            "FAÎTIERE", "REMUCI : AGENCE-BONOUA", "REMUCI : AGENCE-ABOISSO",
            "REMUCI : AGENCE-BASSAM", "REMUCI : AGENCE-AGBOVILLE",
            "REMUCI : AGENCE-TIASSALE", "REMUCI : AGENCE-DIVO",
            "REMUCI : AGENCE-ADZOPE", "REMUCI : GRAND-LAHOU",
            "REMUCI : AGENCE-DABOU"
        ]
        
        if agence and agence not in valid_agences:
            abort(400, "Agence non valide")
        return agence