import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from config import IMPACT_LABELS, MAX_MATERIALS


# --- COULEURS POUR LA MISE EN FORME ---
COLOR_HEADER     = "1F4E79"  # bleu foncé
COLOR_REQUIRED   = "D9E1F2"  # bleu clair — champs obligatoires
COLOR_OPTIONAL   = "EEF3F9"  # bleu très clair — champs optionnels
COLOR_REF        = "E2EFDA"  # vert clair — onglets de référence
COLOR_OUTPUT     = "FFF2CC"  # jaune clair — résultats
COLOR_ERROR      = "FFCCCC"  # rouge clair — erreurs


# --- GÉNÉRATION DU TEMPLATE EXCEL ---

def generate_template(materials: list, products: list, countries: list) -> str:
    """
    Génère le fichier Excel template avec :
    - un onglet INPUT pré-formaté avec validations
    - des onglets de référence (matières, produits, pays)
    Retourne le chemin du fichier généré.
    """
    wb = Workbook()

    _create_input_sheet(wb, materials, products, countries)
    _create_ref_sheet(wb, "REF_MATERIALS", ["uuid", "nom"], 
                      [(m["id"], m["name"]) for m in materials])
    _create_ref_sheet(wb, "REF_PRODUCTS", ["id", "nom"],
                      [(p["id"], p["name"]) for p in products])
    _create_ref_sheet(wb, "REF_COUNTRIES", ["code", "nom"],
                      [(c["code"], c["name"]) for c in countries])

    # Supprime la feuille vide créée par défaut
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    path = "ecobalyse_template.xlsx"
    wb.save(path)
    return path


def _create_input_sheet(wb, materials, products, countries):
    ws = wb.create_sheet("TEXTILE_INPUT")

    # Construction des colonnes
    columns = [
        ("product_name",      "Nom du produit",              True),
        ("product",           "Type de produit (id)",        True),
        ("mass",              "Masse (kg)",                  True),
    ]
    for i in range(1, MAX_MATERIALS + 1):
        req = (i == 1)  # seule la première matière est obligatoire
        columns += [
            (f"mat{i}_id",      f"Matière {i} (uuid)",         req),
            (f"mat{i}_share",   f"Matière {i} part (0-1)",     req),
            (f"mat{i}_country", f"Matière {i} pays",           False),
        ]
    columns += [
        ("countrySpinning",   "Pays Filature",               False),
        ("countryFabric",     "Pays Tissage",                False),
        ("countryDyeing",     "Pays Teinture",               False),
        ("countryMaking",     "Pays Confection",             False),
        ("fabricProcess",     "Procédé tissage",             False),
        ("dyeingProcessType", "Type teinture",               False),
        ("makingComplexity",  "Complexité confection",       False),
        ("airTransportRatio", "Transport aérien (0-1)",      False),
        ("makingWaste",       "Perte confection (0-0.4)",    False),
        ("makingDeadStock",   "Stock dormant (0-0.3)",       False),
        ("fading",            "Délavage (True/False)",       False),
        ("upcycled",          "Remanufacturé (True/False)",  False),
        ("business",          "Type entreprise",             False),
    ]

    # Écriture des headers
    for col_idx, (col_id, label, required) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COLOR_HEADER)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Couleur de fond sur les 50 premières lignes selon obligatoire/optionnel
        fill_color = COLOR_REQUIRED if required else COLOR_OPTIONAL
        for row in range(2, 52):
            ws.cell(row=row, column=col_idx).fill = PatternFill("solid", fgColor=fill_color)

    ws.row_dimensions[1].height = 40
    ws.freeze_panes = "A2"  # fige la ligne de header

    # Validations de données pour les enums fixes
    _add_enum_validation(ws, columns, products, countries)


def _add_enum_validation(ws, columns, products, countries):
    """Ajoute des listes déroulantes pour les champs à valeurs fixes."""
    col_map = {col[0]: idx + 1 for idx, col in enumerate(columns)}

    # Produits
    product_ids = ",".join([p["id"] for p in products])
    _add_dropdown(ws, col_map["product"], f'"{product_ids}"')

    # Pays pour chaque étape
    country_codes = ",".join([c["code"] for c in countries])
    for field in ["countrySpinning", "countryFabric", "countryDyeing", "countryMaking"]:
        if field in col_map:
            _add_dropdown(ws, col_map[field], f'"{country_codes}"')

    # Enums fixes
    _add_dropdown(ws, col_map["fabricProcess"],
        '"knitting-mix,knitting-fully-fashioned,knitting-integral,knitting-circular,knitting-straight,weaving"')
    _add_dropdown(ws, col_map["dyeingProcessType"],
        '"average,continuous,discontinuous"')
    _add_dropdown(ws, col_map["makingComplexity"],
        '"very-high,high,medium,low,very-low,not-applicable"')
    _add_dropdown(ws, col_map["business"],
        '"small-business,large-business-with-services,large-business-without-services"')


def _add_dropdown(ws, col_idx, formula: str):
    col_letter = get_column_letter(col_idx)
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}51")


def _create_ref_sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E7D32")
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    ws.sheet_state = "visible"


# --- LECTURE DU FICHIER INPUT ---

def read_input(file) -> list:
    """Lit l'onglet TEXTILE_INPUT et retourne une liste de dicts."""
    df = pd.read_excel(file, sheet_name="TEXTILE_INPUT", dtype=str)
    df = df.dropna(how="all")  # supprime les lignes complètement vides
    return df.to_dict(orient="records")


# --- ÉCRITURE DU FICHIER OUTPUT ---

def write_output(rows_input: list, results: list) -> str:
    """
    Génère le fichier Excel de résultats.
    Chaque ligne = un produit avec ses paramètres + ses scores environnementaux.
    """
    output_rows = []

    for row, result in zip(rows_input, results):
        out = {"product_name": row.get("product_name", "")}

        if "error" in result and not result.get("impacts"):
            # Erreur non récupérée par le fallback
            out["statut"] = "ERREUR"
            out["erreur_detail"] = str(result["error"])
            out["fallback_note"] = result.get("fallback_note", "")
        else:
            out["statut"] = "OK"
            out["erreur_detail"] = ""
            out["fallback_note"] = result.get("fallback_note", "")
            impacts = result.get("impacts", {})
            for code, label in IMPACT_LABELS.items():
                out[label] = impacts.get(code, "")

        output_rows.append(out)

    df = pd.DataFrame(output_rows)
    path = "ecobalyse_results.xlsx"

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RESULTATS")
        ws = writer.sheets["RESULTATS"]

        # Mise en forme du header
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=COLOR_HEADER)

        # Coloration des lignes en erreur
        for row in ws.iter_rows(min_row=2):
            if row[1].value == "ERREUR":
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=COLOR_ERROR)
            elif row[3].value:  # fallback utilisé
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=COLOR_OUTPUT)

    return path