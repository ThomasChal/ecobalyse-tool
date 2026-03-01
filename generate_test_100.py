import pandas as pd
import random
from ecobalyse_api import get_materials, get_products, get_countries

# Récupération des vraies données depuis l'API
materials = get_materials()
products  = get_products()
countries = get_countries()

# Labels lisibles (identiques à excel_handler.py)
FABRIC_PROCESS_LABELS = {
    "knitting-mix":             "Tricotage moyen",
    "knitting-fully-fashioned": "Tricotage fully-fashioned",
    "knitting-integral":        "Tricotage intégral",
    "knitting-circular":        "Tricotage circulaire",
    "knitting-straight":        "Tricotage rectiligne",
    "weaving":                  "Tissage",
}
DYEING_LABELS = {
    "average":       "Teinture moyenne",
    "continuous":    "Teinture continue",
    "discontinuous": "Teinture discontinue",
}
COMPLEXITY_LABELS = {
    "very-high":      "Très élevée",
    "high":           "Élevée",
    "medium":         "Moyenne",
    "low":            "Faible",
    "very-low":       "Très faible",
    "not-applicable": "Non applicable",
}
BUSINESS_LABELS = {
    "small-business":                  "PME/TPE",
    "large-business-with-services":    "Grande entreprise avec services",
    "large-business-without-services": "Grande entreprise sans services",
}

# Listes de labels depuis les données API
mat_names     = [m["name"] for m in materials]
prod_names    = [p["name"] for p in products]
country_names = [c["name"] for c in countries]

fabric_labels    = list(FABRIC_PROCESS_LABELS.values())
dyeing_labels    = list(DYEING_LABELS.values())
complexity_labels = list(COMPLEXITY_LABELS.values())
business_labels  = list(BUSINESS_LABELS.values())

# Masses typiques par nom de produit
MASSES = {
    "T-shirt / Polo": (0.12, 0.25),
    "Jean":           (0.50, 0.90),
    "Pull":           (0.30, 0.70),
    "Chemise":        (0.15, 0.30),
    "Pantalon":       (0.35, 0.70),
    "Manteau":        (0.80, 1.80),
    "Chaussettes":    (0.05, 0.12),
    "Caleçon":        (0.08, 0.15),
    "Slip":           (0.06, 0.12),
    "Jupe":           (0.20, 0.45),
}

rows = []
for i in range(1, 101):
    prod_name = random.choice(prod_names)
    mass_range = MASSES.get(prod_name, (0.15, 0.50))
    mass = round(random.uniform(*mass_range), 2)

    nb_materials = random.choices([1, 2], weights=[0.6, 0.4])[0]
    mat1 = random.choice(mat_names)

    if nb_materials == 2:
        mat2 = random.choice([m for m in mat_names if m != mat1])
        share1 = round(random.uniform(0.5, 0.9), 2)
        share2 = round(1 - share1, 2)
    else:
        mat2 = ""
        share1 = 1.0
        share2 = ""

    row = {
        "Nom du produit":             f"Produit Test {i:03d} - {prod_name}",
        "Type de produit":            prod_name,
        "Masse (kg)":                 mass,
        "Matière 1":                  mat1,
        "Matière 1 part (0-1)":       share1,
        "Matière 1 pays":             random.choice(country_names),
        "Matière 1 filature":         "",
        "Matière 2":                  mat2,
        "Matière 2 part (0-1)":       share2,
        "Matière 2 pays":             random.choice(country_names) if mat2 else "",
        "Matière 2 filature":         "",
        "Matière 3":                  "",
        "Matière 3 part (0-1)":       "",
        "Matière 3 pays":             "",
        "Matière 3 filature":         "",
        "Matière 4":                  "",
        "Matière 4 part (0-1)":       "",
        "Matière 4 pays":             "",
        "Matière 4 filature":         "",
        "Matière 5":                  "",
        "Matière 5 part (0-1)":       "",
        "Matière 5 pays":             "",
        "Matière 5 filature":         "",
        "Pays Filature":              random.choice(country_names),
        "Pays Tissage":               random.choice(country_names),
        "Pays Teinture":              random.choice(country_names),
        "Pays Confection":            random.choice(country_names),
        "Procédé tissage":            random.choice(fabric_labels),
        "Type teinture":              random.choice(dyeing_labels),
        "Complexité confection":      random.choice(complexity_labels),
        "Transport aérien (0-1)":     "",
        "Perte confection (0-0.4)":   "",
        "Stock dormant (0-0.3)":      "",
        "Délavage (oui/non)":         random.choice(["oui", "non", ""]),
        "Remanufacturé (oui/non)":    "",
        "Type entreprise":            random.choice(business_labels),
    }
    rows.append(row)

df = pd.DataFrame(rows)

with pd.ExcelWriter("ecobalyse_test_100.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="SAISIE", index=False)
    ws = writer.sheets["SAISIE"]
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 28

print(f"✅ {len(rows)} produits générés dans ecobalyse_test_100.xlsx")
