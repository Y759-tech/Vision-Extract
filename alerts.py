import logging
from datetime import datetime
from typing import Dict, List, Optional

logger = logging.getLogger('remuci_alerts')

class AlertSystem:
    def __init__(self):
        self.thresholds = {
            'taux_impayes_critique': 10.0,    # > 10% = critique
            'taux_impayes_avertissement': 5.0, # > 5% = avertissement
            'retard_moyen_critique': 180,      # > 6 mois = critique
            'retard_moyen_avertissement': 90,  # > 3 mois = avertissement
            'concentration_risque': 40.0,      # > 40% sur une agence
            'dossiers_anciens': 365,           # > 1 an = ancien
        }
        
        self.alert_history = []
    
    def analyze_credits_data(self, data: Dict) -> List[Dict]:
        """Analyse les données de crédits et génère des alertes"""
        alerts = []
        
        # Analyse du taux d'impayés global
        taux_impayes = data.get('taux_impayes_global', 0)
        if taux_impayes > self.thresholds['taux_impayes_critique']:
            alerts.append({
                'niveau': 'CRITIQUE',
                'type': 'TAUX_IMPAYES',
                'titre': 'Taux d\'impayés critique',
                'message': f"Le taux d'impayés global est de {taux_impayes:.1f}% (seuil: {self.thresholds['taux_impayes_critique']}%)",
                'action': 'Revue immédiate de la politique de crédit requise',
                'timestamp': datetime.now()
            })
        elif taux_impayes > self.thresholds['taux_impayes_avertissement']:
            alerts.append({
                'niveau': 'AVERTISSEMENT',
                'type': 'TAUX_IMPAYES',
                'titre': 'Taux d\'impayés élevé',
                'message': f"Le taux d'impayés global est de {taux_impayes:.1f}%",
                'action': 'Surveillance renforcée nécessaire',
                'timestamp': datetime.now()
            })
        
        # Analyse de la concentration des risques
        agences = data.get('analyse_agences', [])
        if agences:
            total_montant = sum(agence.get('MontantTotal', 0) for agence in agences)
            if total_montant > 0:
                agence_principale = max(agences, key=lambda x: x.get('MontantTotal', 0))
                concentration = (agence_principale.get('MontantTotal', 0) / total_montant) * 100
                
                if concentration > self.thresholds['concentration_risque']:
                    alerts.append({
                        'niveau': 'AVERTISSEMENT',
                        'type': 'CONCENTRATION_RISQUE',
                        'titre': 'Risque concentré',
                        'message': f"{concentration:.1f}% du risque sur {agence_principale.get('Agence', 'N/A')}",
                        'action': 'Diversification du portefeuille recommandée',
                        'timestamp': datetime.now()
                    })
        
        # Analyse des dossiers anciens
        dossiers_anciens = data.get('dossiers_anciens_count', 0)
        if dossiers_anciens > 0:
            alerts.append({
                'niveau': 'ATTENTION',
                'type': 'DOSSIERS_ANCIENS',
                'titre': 'Dossiers anciens détectés',
                'message': f"{dossiers_anciens} dossiers ont plus d'un an de retard",
                'action': 'Plan de recouvrement urgent nécessaire',
                'timestamp': datetime.now()
            })
        
        # Journalisation des alertes
        for alert in alerts:
            logger.warning(f"ALERTE {alert['niveau']}: {alert['titre']} - {alert['message']}")
            self.alert_history.append(alert)
        
        return alerts
    
    def get_recent_alerts(self, limit: int = 10) -> List[Dict]:
        """Récupère les alertes récentes"""
        return sorted(self.alert_history, key=lambda x: x['timestamp'], reverse=True)[:limit]

# Instance globale
alert_system = AlertSystem()