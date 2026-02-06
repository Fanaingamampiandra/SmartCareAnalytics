# app.py — point d’entrée ; sidebar, filtres et en-tête communs, contenu par page
import importlib
import streamlit as st

from utils import load_data
from pages import PAGE_MODULES

# Pages disponibles (nom affiché)
PAGES = list(PAGE_MODULES.keys())

# Données par page (chemin CSV ou None si pas encore de données)
DATA_PATHS = {
    "Logistique": "data/logistics/logistics-all.csv",
    "Activité & Service": "data/activity-service/activity-service-donnees_mensuelles_reconstituees.csv",
    "Capacité": "data/capacity/capacity-donnees_mensuelles_reconstituees.csv",
    "Finance": "data/finance/finance-donnees_mensuelles_reconstituees.csv",
    "Patients": "data/patients/patients-all.csv",
    "Qualité":"data/quality/quality-donnees_mensuelles_reconstituees.csv" ,
    "RH": "data/hr/hr-all.csv",
}

st.set_page_config(page_title="PSL–CFX | Infographie (Normal vs Crise)", layout="wide")

# Masque le menu multipage par défaut de Streamlit dans la sidebar
st.markdown(
    """
    <style>
    [data-testid="stSidebarNav"] {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def get_years_for_filters():
    """Années disponibles pour les filtres (à partir d’une source commune)."""
    all_years = set()
    for _name, path in DATA_PATHS.items():
        if not path:
            continue
        try:
            df = load_data(path)
            year_col = "ANNEE" if "ANNEE" in df.columns else ("year" if "year" in df.columns else None)
            if year_col is not None:
                all_years.update(df[year_col].astype(int).dropna().unique().tolist())
        except Exception:
            pass
    if not all_years:
        all_years = set(range(2011, 2026))
    return sorted(all_years)


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
show_forecast = st.sidebar.checkbox("Afficher prévision 2017", value=False)
# 2017 n’apparaît dans la liste qu’une fois la prévision activée
years_for_select = [y for y in years if y != 2017 or show_forecast]
year_choice = st.sidebar.selectbox("Année", options=["Toutes"] + years_for_select)
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
    "show_forecast": show_forecast,
}
page_module.render(st, **context)
