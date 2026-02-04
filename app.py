# app.py
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

st.set_page_config(page_title="PSL–CFX | Logistique (Normal vs Crise)", layout="wide")

@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["ANNEE"] = df["ANNEE"].astype(int)
    return df

DATA_PATH = "data/logistics-crise-comparaison.csv"
df = load_data(DATA_PATH)

st.title("Infographie Logistique PSL–CFX")

# ---------------------------
# Sidebar (minimal)
# ---------------------------
st.sidebar.header("Filtres")
years = sorted(df["ANNEE"].unique().tolist())
year_choice = st.sidebar.selectbox("Année", options=["Toutes"] + years)
mode_choice = st.sidebar.radio(
    "Mode",
    options=["Normal", "Crise"],
    index=0,  # Par défaut "Normal"
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

def pick_value_cols(hosp: str):
    if hosp == "TOTAL":
        return "TOTAL_NORMAL", "TOTAL_CRISE"
    if hosp == "PLF":
        return "PLF_NORMAL", "PLF_CRISE"
    return "CFX_NORMAL", "CFX_CRISE"

normal_col, crise_col = pick_value_cols(hospital_choice)

# Indication du mode choisi
st.caption(f"Mode affiché : **{mode_choice}**")

# Données filtrées année uniquement (pas de filtre unité)
dff = df.copy()
if year_choice != "Toutes":
    dff = dff[dff["ANNEE"] == int(year_choice)]

compare_cols = ["ANNEE", "INDICATEUR", "SOUS-INDICATEUR", "UNITE", normal_col, crise_col]
cmp = (
    dff[["ANNEE", "INDICATEUR", "SOUS-INDICATEUR"]]
    .drop_duplicates()
    .merge(df[compare_cols], on=["ANNEE", "INDICATEUR", "SOUS-INDICATEUR"], how="left")
)
cmp["CHANGEMENT"] = cmp[crise_col] - cmp[normal_col]
cmp["EVOLUTION_%"] = np.where(
    cmp[normal_col].fillna(0) == 0,
    np.nan,
    (cmp["CHANGEMENT"] / cmp[normal_col]) * 100,
)
cmp["TENDANCE"] = cmp["CHANGEMENT"].apply(lambda x: "⬆️" if x > 0 else ("⬇️" if x < 0 else "➜"))

# Palette années (identique à "Toutes") pour garder la même couleur par année quand une année est sélectionnée
PALETTE_ANNEES = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]

# ---------------------------
# Un onglet par indicateur ; dans chaque onglet, un graphe par unité (avec l'unité affichée)
# ---------------------------
indicateurs = sorted(cmp["INDICATEUR"].dropna().unique().tolist())
tabs = st.tabs(indicateurs)

for tab, indic in zip(tabs, indicateurs):
    with tab:
        st.subheader(indic)
        df_indic = cmp[cmp["INDICATEUR"] == indic]
        unites = sorted(df_indic["UNITE"].dropna().unique().tolist())

        for unite in unites:
            df_u = df_indic[df_indic["UNITE"] == unite]

            # Si une seule sous-catégorie pour cette unité, on met son nom dans le sous-titre ; sinon juste l'unité
            sous_labels = sorted(df_u["SOUS-INDICATEUR"].dropna().unique().tolist())
            if len(sous_labels) == 1:
                subtitle = f"{sous_labels[0]} ({unite})"
            else:
                subtitle = f"({unite})"

            # Année "Toutes" → afficher par année, une couleur par année
            by_year = year_choice == "Toutes"

            if by_year:
                # Agrégation par (SOUS-INDICATEUR, ANNEE)
                agg = (
                    df_u.groupby(["SOUS-INDICATEUR", "ANNEE"])[[normal_col, crise_col]]
                    .sum()
                    .reset_index()
                )
                agg["ANNEE"] = agg["ANNEE"].astype(int)
                col_to_show = normal_col if mode_choice == "Normal" else crise_col
                mode_label = "Situation normale" if mode_choice == "Normal" else "Crise (simulation)"
                long = agg[["SOUS-INDICATEUR", "ANNEE", col_to_show]].rename(columns={col_to_show: "Valeur"})
                bars = (
                    alt.Chart(long)
                    .mark_bar()
                    .encode(
                        x=alt.X("SOUS-INDICATEUR:N", title="Sous-indicateur"),
                        y=alt.Y("Valeur:Q", title=f"Volume ({unite})"),
                        color=alt.Color("ANNEE:O", title="Année", scale=alt.Scale(domain=years, range=PALETTE_ANNEES)),
                    )
                )
                sub = f"{subtitle} — par année ({mode_label})"
                chart = bars.properties(
                    title={"text": indic, "subtitle": sub},
                    height=450,
                )
            else:
                # Une année fixe : même couleur que pour cette année dans la vue "Toutes"
                col_to_show = normal_col if mode_choice == "Normal" else crise_col
                mode_label = "Situation normale" if mode_choice == "Normal" else "Crise (simulation)"
                year_idx = years.index(int(year_choice)) if int(year_choice) in years else 0
                color_annee = PALETTE_ANNEES[year_idx % len(PALETTE_ANNEES)]
                agg = (
                    df_u.groupby("SOUS-INDICATEUR")[[col_to_show]]
                    .sum()
                    .reset_index()
                )
                agg = agg.rename(columns={col_to_show: "Valeur"})
                bars = (
                    alt.Chart(agg)
                    .mark_bar()
                    .encode(
                        x=alt.X("SOUS-INDICATEUR:N", title="Sous-indicateur"),
                        y=alt.Y("Valeur:Q", title=f"Volume ({unite})"),
                        color=alt.value(color_annee),
                    )
                )
                chart = bars.properties(
                    title={
                        "text": indic,
                        "subtitle": f"{subtitle} — {mode_label}",
                    },
                    height=450,
                )
            st.altair_chart(chart, use_container_width=True)
            st.markdown("<div style='margin-bottom: 3.5rem;'></div>", unsafe_allow_html=True)

        # Tableau détaillé pour cet indicateur
        with st.expander("Voir le détail (tableau)"):
            out = df_indic.rename(columns={
                normal_col: "Valeur (normal)",
                crise_col: "Valeur (crise)",
                "CHANGEMENT": "Changement",
                "EVOLUTION_%": "Évolution (%)",
                "TENDANCE": "Tendance",
            })
            cols = ["ANNEE", "UNITE", "SOUS-INDICATEUR", "Valeur (normal)", "Valeur (crise)", "Changement", "Évolution (%)", "Tendance"]
            st.dataframe(out[[c for c in cols if c in out.columns]].sort_values(["UNITE", "SOUS-INDICATEUR", "ANNEE"]), use_container_width=True)

st.caption("Note : le mode crise est une simulation basée sur des coefficients (scénarios).")
