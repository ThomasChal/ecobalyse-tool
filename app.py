import streamlit as st
from ecobalyse_api import (
    get_materials, get_products, get_countries,
    simulate_textile
)
from excel_handler import generate_template, read_input, write_output

# --- CONFIG DE LA PAGE ---
st.set_page_config(
    page_title="Ecobalyse Tool",
    page_icon="🌿",
    layout="wide"
)

st.title("🌿 Ecobalyse Import / Export")
st.caption("Calcul d'impacts environnementaux via l'API Ecobalyse")

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("Les données de référence sont chargées depuis l'API Ecobalyse en temps réel.")

    if st.button("🔄 Rafraîchir les données de référence"):
        st.cache_data.clear()
        st.success("Cache vidé !")

# --- CHARGEMENT DES DONNÉES DE RÉFÉRENCE (mis en cache) ---
@st.cache_data(ttl=3600)  # recharge toutes les heures
def load_reference_data():
    materials = get_materials()
    products  = get_products()
    countries = get_countries()
    return materials, products, countries

with st.spinner("Chargement des données de référence..."):
    try:
        materials, products, countries = load_reference_data()
        st.sidebar.success(f"✅ {len(materials)} matières · {len(products)} produits · {len(countries)} pays")
    except Exception as e:
        st.error(f"Impossible de contacter l'API Ecobalyse : {e}")
        st.stop()

# --- SECTION 1 : TÉLÉCHARGER LE TEMPLATE ---
st.header("1️⃣ Télécharger le template Excel")
st.markdown("Génère un fichier Excel pré-rempli avec les listes de référence et les listes déroulantes.")

if st.button("📥 Générer et télécharger le template"):
    with st.spinner("Génération du template..."):
        try:
            path = generate_template(materials, products, countries)
            with open(path, "rb") as f:
                st.download_button(
                    label="⬇️ Télécharger ecobalyse_template.xlsx",
                    data=f,
                    file_name="ecobalyse_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erreur lors de la génération : {e}")

st.divider()

# --- SECTION 2 : IMPORTER ET CALCULER ---
st.header("2️⃣ Importer un fichier Excel complété")

uploaded_file = st.file_uploader(
    "Dépose ton fichier Excel ici",
    type=["xlsx"],
    help="Le fichier doit contenir un onglet 'TEXTILE_INPUT' avec les colonnes du template."
)

if uploaded_file:
    try:
        rows = read_input(uploaded_file)
        st.success(f"✅ {len(rows)} produit(s) détecté(s) dans le fichier.")

        # Aperçu des données importées
        with st.expander("👁️ Aperçu des données importées"):
            import pandas as pd
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

        # Bouton de lancement
        if st.button(f"🚀 Lancer le calcul pour {len(rows)} produit(s)"):
            results = []
            errors  = 0

            progress = st.progress(0, text="Calcul en cours...")
            status   = st.empty()

            for i, row in enumerate(rows):
                name = row.get("product_name", f"Produit {i+1}")
                status.text(f"⏳ Traitement : {name}")

                result = simulate_textile(row)
                results.append(result)

                if "error" in result and not result.get("impacts"):
                    errors += 1

                progress.progress((i + 1) / len(rows), text=f"{i+1}/{len(rows)} produits traités")

            status.empty()
            progress.empty()

            # Résumé
            fallbacks = sum(1 for r in results if r.get("fallback_note"))
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Succès",   len(results) - errors)
            col2.metric("⚠️ Fallback pays", fallbacks)
            col3.metric("❌ Erreurs",  errors)

            # Génération et téléchargement de l'output
            output_path = write_output(rows, results)
            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇️ Télécharger les résultats",
                    data=f,
                    file_name="ecobalyse_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Affichage des erreurs détaillées si besoin
            if errors > 0:
                with st.expander("🔍 Détail des erreurs"):
                    for row, result in zip(rows, results):
                        if "error" in result and not result.get("impacts"):
                            st.error(f"**{row.get('product_name')}** — {result['error']}")

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        st.info("Vérifie que le fichier contient bien un onglet 'TEXTILE_INPUT'.")