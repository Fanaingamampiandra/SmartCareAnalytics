# app.py — point d’entrée ; sidebar, filtres et en-tête communs, contenu par page
import importlib
import streamlit as st

from utils import load_data
from pages import PAGE_MODULES

# Pages disponibles (nom affiché)
PAGES = list(PAGE_MODULES.keys())

# Données par page (chemin CSV ou None si pas encore de données)
DATA_PATHS = {
    "Logistique": "data/logistics/logistics-data-with-crise.csv",
    "Activité & Service": "data/activity-service/activity-service-data-with-crise.csv",
    "Capacité": "data/capacity/capacity-data-with-crise.csv",
    "Finance": "data/finance/finance-data-with-crise.csv",
    "Patients": "data/patients/patients-data-with-crise.csv",
    "Qualité":"data/quality/quality-data-with-crise.csv" ,
    "RH": "data/hr/hr-data-with-crise.csv",
}

st.set_page_config(page_title="PSL–CFX | Infographie (Normal vs Crise)", layout="wide")


def get_years_for_filters():
    """Années disponibles pour les filtres (à partir d’une source commune)."""
    path = DATA_PATHS.get("Logistique")
    if path:
        try:
            df = load_data(path)
            return sorted(df["ANNEE"].unique().tolist())
        except Exception:
            pass
    return list(range(2011, 2026))


def pick_value_cols(hosp: str):
    if hosp == "TOTAL":
        return "TOTAL_NORMAL", "TOTAL_CRISE"
    if hosp == "PLF":
        return "PLF_NORMAL", "PLF_CRISE"
    return "CFX_NORMAL", "CFX_CRISE"


# ---------------------------
# Sidebar : page + filtres (commun à toutes les pages)
# ---------------------------
st.sidebar.header("Navigation")
page_choice = st.sidebar.selectbox("Page", options=PAGES, label_visibility="collapsed")

years = get_years_for_filters()
st.sidebar.header("Filtres")
year_choice = st.sidebar.selectbox("Année", options=["Toutes"] + years)
mode_choice = st.sidebar.radio(
    "Mode",
    options=["Normal", "Crise"],
    index=0,
    format_func=lambda x: {
        "Normal": "Situation normale",
        "Crise": "Crise sanitaire (simulation)",
    }[x],
)
hospital_choice = st.sidebar.radio(
    "Site / Total",
    options=["TOTAL", "PLF", "CFX"],
    format_func=lambda x: {"TOTAL": "Total (PSL + CFX)", "PLF": "Pitié-Salpêtrière (PSL)", "CFX": "Charles Foix (CFX)"}[x],
)

normal_col, crise_col = pick_value_cols(hospital_choice)

# ---------------------------
# En-tête commun : titre + mode affiché
# ---------------------------
st.title(f"Infographie {page_choice} PSL–CFX")
st.caption(f"Mode affiché : **{mode_choice}**")

# ---------------------------
# Contenu : délégation au module de la page
# ---------------------------
module_name = PAGE_MODULES[page_choice]
page_module = importlib.import_module(f"pages.{module_name}")

context = {
    "data_path": DATA_PATHS.get(page_choice),
    "year_choice": year_choice,
    "mode_choice": mode_choice,
    "hospital_choice": hospital_choice,
    "normal_col": normal_col,
    "crise_col": crise_col,
    "years": years,
    "page_name": page_choice,
}
page_module.render(st, **context)

st.caption("Note : le mode crise est une simulation basée sur des coefficients (scénarios).")
