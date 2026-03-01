import requests
from config import API_TOKEN, ENDPOINTS, FALLBACK_COUNTRY

# Headers communs à toutes les requêtes
def _get_headers():
    headers = {"Content-Type": "application/json"}
    if API_TOKEN:
        headers["token"] = API_TOKEN
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
    """
    Prend une ligne du fichier Excel (dict), appelle l'API Ecobalyse,
    et retourne les impacts environnementaux.

    En cas d'erreur de pays, retente avec le pays fallback et le note.
    """
    payload = _build_payload(row)
    result = _call_simulator(payload)

    # Si erreur 400, on tente le fallback pays
    if "error" in result:
        payload_fallback, fallback_note = _apply_country_fallback(payload)
        result = _call_simulator(payload_fallback)
        result["fallback_note"] = fallback_note
    else:
        result["fallback_note"] = ""

    return result


def _call_simulator(payload: dict) -> dict:
    """Appelle l'endpoint /textile/simulator et retourne le résultat ou une erreur."""
    try:
        response = requests.post(
            ENDPOINTS["textile_simulator"],
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
    """Construit le JSON à envoyer à l'API à partir d'une ligne Excel."""

    # Construction du tableau de matières
    materials = []
    for i in range(1, 6):  # jusqu'à 5 matières
        mat_id = row.get(f"mat{i}_id")
        mat_share = row.get(f"mat{i}_share")
        if mat_id and mat_share:
            mat = {"id": str(mat_id), "share": float(mat_share)}
            if row.get(f"mat{i}_country"):
                mat["country"] = str(row[f"mat{i}_country"])
            if row.get(f"mat{i}_spinning"):
                mat["spinning"] = str(row[f"mat{i}_spinning"])
            materials.append(mat)

    payload = {
        "mass": float(row["mass"]),
        "product": str(row["product"]),
        "materials": materials,
    }

    # Ajout des champs optionnels seulement s'ils sont renseignés
    optional_fields = [
        "countrySpinning", "countryFabric", "countryDyeing", "countryMaking",
        "fabricProcess", "dyeingProcessType", "makingComplexity",
        "airTransportRatio", "makingWaste", "makingDeadStock",
        "fading", "upcycled", "business",
    ]
    for field in optional_fields:
        value = row.get(field)
        if value is not None and str(value).strip() != "" and str(value) != "nan":
            payload[field] = value

    return payload


def _apply_country_fallback(payload: dict) -> tuple:
    """
    Remplace tous les pays par le pays fallback et génère une note explicative.
    Retourne le payload modifié et la note.
    """
    country_fields = ["countrySpinning", "countryFabric", "countryDyeing", "countryMaking"]
    changed = []

    for field in country_fields:
        if field in payload:
            original = payload[field]
            payload[field] = FALLBACK_COUNTRY
            changed.append(f"{field}: {original} → {FALLBACK_COUNTRY}")

    note = "FALLBACK PAYS: " + " | ".join(changed) if changed else ""
    return payload, note