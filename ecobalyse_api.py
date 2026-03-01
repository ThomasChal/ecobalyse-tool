import requests
from config import API_TOKEN, ENDPOINTS, FALLBACK_COUNTRY

# Headers communs à toutes les requêtes
def _get_headers():
    headers = {"Content-Type": "application/json"}
    if API_TOKEN:
        headers["Authorization"] = f"Bearer {API_TOKEN}"
    return headers


# --- FONCTIONS DE LISTE (pour remplir les onglets de référence) ---

def get_materials():
    """Récupère la liste de toutes les matières textiles disponibles."""
    response = requests.get(ENDPOINTS["textile_materials"], headers=_get_headers())
    response.raise_for_status()
    return response.json()

def get_products():
    """Récupère la liste des types de produits textiles."""
    response = requests.get(ENDPOINTS["textile_products"], headers=_get_headers())
    response.raise_for_status()
    return response.json()

def get_countries():
    """Récupère la liste des pays disponibles pour le textile."""
    response = requests.get(ENDPOINTS["textile_countries"], headers=_get_headers())
    response.raise_for_status()
    return response.json()

def get_trims():
    """Récupère la liste des accessoires disponibles."""
    response = requests.get(ENDPOINTS["textile_trims"], headers=_get_headers())
    response.raise_for_status()
    return response.json()


# --- FONCTION PRINCIPALE DE SIMULATION ---

def simulate_textile(row: dict) -> dict:
    payload = _build_payload(row)
    result = _call_simulator(payload)

    if "error" in result:
        # On tente jusqu'à 5 fois, en corrigeant un pays invalide à chaque fois
        fallback_notes = []
        for _ in range(5):
            invalid_country = _extract_invalid_country(result["error"])
            if not invalid_country:
                break
            payload, note = _apply_country_fallback(payload, invalid_country)
            fallback_notes.append(note)
            result = _call_simulator(payload)
            if "error" not in result:
                break

        result["fallback_note"] = " | ".join(fallback_notes) if fallback_notes else ""
    else:
        result["fallback_note"] = ""

    return result


def _call_simulator(payload: dict) -> dict:
    try:
        response = requests.post(
            ENDPOINTS["textile_simulator_detailed"],  # ← changement ici
            json=payload,
            headers=_get_headers()
        )
        if response.status_code == 200:
            return response.json()
        else:
            return {"error": response.json().get("error", str(response.status_code))}
    except Exception as e:
        return {"error": str(e)}


def _build_payload(row: dict) -> dict:
    """Construit le JSON à envoyer à l'API — mode réglementaire."""

    def is_empty(v):
        if v is None:
            return True
        try:
            import math
            if math.isnan(float(v)):
                return True
        except (ValueError, TypeError):
            pass
        return str(v).strip().lower() in ("", "nan", "none")

    row = {k: v for k, v in row.items() if not is_empty(v)}

    # Matières
    materials = []
    for i in range(1, 6):
        mat_id    = row.get(f"mat{i}_id")
        mat_share = row.get(f"mat{i}_share")
        if mat_id and mat_share:
            mat = {"id": str(mat_id), "share": float(mat_share)}
            if row.get(f"mat{i}_country"):
                mat["country"] = str(row[f"mat{i}_country"])
            materials.append(mat)

    payload = {
        "mass":      float(row["mass"]),
        "product":   str(row["product"]),
        "materials": materials,
    }

    # Pays des étapes
    for field in ["countrySpinning", "countryFabric", "countryDyeing", "countryMaking"]:
        if field in row:
            payload[field] = str(row[field])

    return payload


def _extract_invalid_country(error) -> str:
    """Extrait le code pays invalide depuis le message d'erreur de l'API."""
    import re
    if isinstance(error, dict):
        for msg in error.values():
            match = re.search(r"Le code pays (\w+) n'est pas utilisable", str(msg))
            if match:
                return match.group(1)
    match = re.search(r"Le code pays (\w+) n'est pas utilisable", str(error))
    return match.group(1) if match else None


def _apply_country_fallback(payload: dict, invalid_country: str) -> tuple:
    """Remplace uniquement le pays invalide dans tout le payload."""
    changed = []
    country_fields = ["countrySpinning", "countryFabric", "countryDyeing", "countryMaking"]

    for field in country_fields:
        if payload.get(field) == invalid_country:
            payload[field] = FALLBACK_COUNTRY
            changed.append(f"{field}: {invalid_country} → {FALLBACK_COUNTRY}")

    for i, mat in enumerate(payload.get("materials", [])):
        if mat.get("country") == invalid_country:
            mat["country"] = FALLBACK_COUNTRY
            changed.append(f"materials[{i}].country: {invalid_country} → {FALLBACK_COUNTRY}")

    note = f"FALLBACK: {' | '.join(changed)}" if changed else ""
    return payload, note