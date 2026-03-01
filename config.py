import os
from dotenv import load_dotenv

load_dotenv()  # lit ton fichier .env automatiquement

# Token API Ecobalyse
API_TOKEN = os.getenv("ECOBALYSE_TOKEN")

# URL de base de l'API
BASE_URL = "https://ecobalyse.beta.gouv.fr/api"

# Endpoints
ENDPOINTS = {
    "textile_simulator": f"{BASE_URL}/textile/simulator",
    "textile_materials": f"{BASE_URL}/textile/materials",
    "textile_products":  f"{BASE_URL}/textile/products",
    "textile_countries": f"{BASE_URL}/textile/countries",
    "textile_trims":     f"{BASE_URL}/textile/trims",
    "textile_simulator_detailed": f"{BASE_URL}/textile/simulator/detailed",
}

# Pays fallback quand un pays n'est pas disponible pour une étape
FALLBACK_COUNTRY = "CN"  # CN = pays générique disponible partout dans Ecobalyse

# Nombre maximum de matières par produit dans l'Excel
MAX_MATERIALS = 5

# Impacts environnementaux à inclure dans l'output
IMPACT_LABELS = {
    "cch": "Changement climatique (kg CO2 éq.)",
    "pef": "Score PEF (µPt)",
    "ecs": "Score d'impacts (µPts)",
    "wtu": "Utilisation eau (m³)",
    "ldu": "Utilisation des sols (pt)",
    "fru": "Ressources fossiles (MJ)",
    "etf": "Écotoxicité eau douce (CTUe)",
    "acd": "Acidification (mol H+ éq.)",
    "bvi": "Biodiversité locale (BVI)",
}