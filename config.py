import os
import streamlit as st

try:
    API_TOKEN = st.secrets["ECOBALYSE_TOKEN"]
except Exception:
    from dotenv import load_dotenv
    load_dotenv()
    API_TOKEN = os.getenv("ECOBALYSE_TOKEN")

BASE_URL = "https://ecobalyse.beta.gouv.fr/api"

ENDPOINTS = {
    "textile_simulator":          f"{BASE_URL}/textile/simulator",
    "textile_simulator_detailed": f"{BASE_URL}/textile/simulator/detailed",
    "textile_materials":          f"{BASE_URL}/textile/materials",
    "textile_products":           f"{BASE_URL}/textile/products",
    "textile_countries":          f"{BASE_URL}/textile/countries",
    "textile_trims":              f"{BASE_URL}/textile/trims",
}

FALLBACK_COUNTRY = "CN"
MAX_MATERIALS    = 5

IMPACT_LABELS = {
    "ecs": "Score d'impacts global (µPts)",
    "cch": "Changement climatique (kg CO2 éq.)",
    "pef": "Score PEF (µPt)",
    "wtu": "Utilisation eau (m³)",
    "ldu": "Utilisation des sols (pt)",
    "fru": "Ressources fossiles (MJ)",
    "etf": "Écotoxicité eau douce (CTUe)",
    "acd": "Acidification (mol H+ éq.)",
}