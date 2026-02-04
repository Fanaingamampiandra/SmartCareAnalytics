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

st.title("Infographie Logistique PSL–CFX — Mode Normal / Mode Crise")

# ---------------------------
# Sidebar controls (filtres)
# ---------------------------
st.sidebar.header("Filtres")

years = sorted(df["ANNEE"].unique().tolist())
year_choice = st.sidebar.selectbox("Année", options=["Toutes"] + years)

hospital_choice = st.sidebar.radio(
    "Site / Total",
    options=["TOTAL", "PLF", "CFX"],
    format_func=lambda x: {"TOTAL": "Total (PSL + CFX)", "PLF": "Pitié-Salpêtrière (PSL)", "CFX": "Charles Foix (CFX)"}[x],
)

mode_choice = st.sidebar.radio(
    "Mode",
    options=["Normal", "Crise"],
    format_func=lambda x: "Situation normale" if x == "Normal" else "Crise sanitaire (simulation)"
)

group_level = st.sidebar.selectbox(
    "Vue principale",
    options=["INDICATEUR", "SOUS-INDICATEUR"],
    format_func=lambda x: "Par grande catégorie" if x == "INDICATEUR" else "Par sous-catégorie"
)

top_n_sub = st.sidebar.slider("Nombre de sous-catégories affichées (Top)", 5, 30, 15)
show_details = st.sidebar.checkbox("Afficher le tableau détaillé", value=True)

# ---------------------------
# Helpers
# ---------------------------
def pick_value_cols(hosp: str):
    if hosp == "TOTAL":
        return "TOTAL_NORMAL", "TOTAL_CRISE"
    if hosp == "PLF":
        return "PLF_NORMAL", "PLF_CRISE"
    return "CFX_NORMAL", "CFX_CRISE"

normal_col, crise_col = pick_value_cols(hospital_choice)

# Filtrage année
dff = df.copy()
if year_choice != "Toutes":
    dff = dff[dff["ANNEE"] == int(year_choice)]

# Valeur affichée selon mode
dff["VALEUR"] = dff[normal_col] if mode_choice == "Normal" else dff[crise_col]

# Table de comparaison (normal vs crise)
# On garde aussi l'unité, supposée constante pour un couple (indicateur, sous-indicateur)
compare_cols = ["ANNEE", "INDICATEUR", "SOUS-INDICATEUR", "UNITE", normal_col, crise_col]
cmp = (
    dff[["ANNEE", "INDICATEUR", "SOUS-INDICATEUR"]]
    .drop_duplicates()
    .merge(
        df[compare_cols],
        on=["ANNEE", "INDICATEUR", "SOUS-INDICATEUR"],
        how="left",
    )
)

cmp["CHANGEMENT"] = cmp[crise_col] - cmp[normal_col]
cmp["EVOLUTION_%"] = np.where(
    cmp[normal_col].fillna(0) == 0,
    np.nan,
    (cmp["CHANGEMENT"] / cmp[normal_col]) * 100
)
cmp["TENDANCE"] = cmp["CHANGEMENT"].apply(lambda x: "⬆️" if x > 0 else ("⬇️" if x < 0 else "➜"))

# ---------------------------
# KPI row (résumé)
# ---------------------------
total_normal = cmp[normal_col].sum()
total_crise = cmp[crise_col].sum()
changement_total = total_crise - total_normal
evolution_total_pct = (changement_total / total_normal * 100) if total_normal != 0 else np.nan

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total (situation normale)", f"{total_normal:,.0f}".replace(",", " "))
c2.metric("Total (crise simulée)", f"{total_crise:,.0f}".replace(",", " "))
c3.metric("Changement (crise vs normal)", f"{changement_total:,.0f}".replace(",", " "))
c4.metric("Évolution (%)", f"{evolution_total_pct:.1f}%" if not np.isnan(evolution_total_pct) else "—")

st.divider()

# ---------------------------
# Charts
# ---------------------------
left, right = st.columns([1.2, 1])

with left:
    st.subheader("Évolution par année (normal vs crise)")
    if year_choice == "Toutes":
        by_year = df.groupby("ANNEE")[[normal_col, crise_col]].sum().reset_index()
        by_year = by_year.rename(columns={normal_col: "Situation normale", crise_col: "Crise (simulation)"})
        st.line_chart(by_year.set_index("ANNEE"))
    else:
        st.info("Choisis 'Toutes' pour voir la courbe sur plusieurs années.")

with right:
    st.subheader("Répartition de l’activité (vue principale)")
    label = "Catégorie"
    grp = dff.groupby(group_level)["VALEUR"].sum().reset_index().sort_values("VALEUR", ascending=False)
    grp = grp.rename(columns={group_level: label})
    st.bar_chart(grp.set_index(label))

st.divider()

# ---------------------------
# Focus spécifique : Déchets / Cartons
# ---------------------------
st.subheader("Focus : déchets cartons (normal vs crise)")

cartons = df[(df["INDICATEUR"] == "Déchets") & (df["SOUS-INDICATEUR"] == "Cartons")].copy()
if cartons.empty:
    st.info("Aucune donnée disponible pour le sous-indicateur 'Déchets / Cartons'.")
else:
    unite_cartons = (
        cartons["UNITE"].dropna().iloc[0]
        if "UNITE" in cartons.columns and cartons["UNITE"].notna().any()
        else ""
    )

    cartons = cartons.sort_values("ANNEE")
    cartons_chart = cartons[["ANNEE", normal_col, crise_col]].rename(
        columns={normal_col: "Situation normale", crise_col: "Crise (simulation)"}
    )

    # Courbe d'évolution
    st.line_chart(cartons_chart.set_index("ANNEE"))

    # Graphe complémentaire : barres comparant normal vs crise par année
    cartons_long = cartons_chart.melt(
        id_vars="ANNEE", var_name="Situation", value_name="Valeur"
    )
    bar = (
        alt.Chart(cartons_long)
        .mark_bar()
        .encode(
            x=alt.X("ANNEE:O", title="Année"),
            y=alt.Y(
                "Valeur:Q",
                title=f"Volume ({unite_cartons})" if unite_cartons else "Volume",
            ),
            color=alt.Color("Situation:N", title="Mode"),
            column=alt.Column("ANNEE:O", title=""),
        )
        .properties(height=200)
    )
    st.altair_chart(bar, use_container_width=True)

st.divider()

# ---------------------------
# Stats par sous-indicateur (même si vue principale = indicateur)
# ---------------------------
st.subheader("Détail par sous-catégorie (sous-indicateurs)")

sub = dff.groupby(["INDICATEUR", "SOUS-INDICATEUR"])["VALEUR"].sum().reset_index()
sub = sub.sort_values("VALEUR", ascending=False).head(top_n_sub)

# Affichage graphique + tableau
col1, col2 = st.columns([1, 1])

with col1:
    st.caption(f"Top {top_n_sub} sous-catégories — mode : {mode_choice}")
    # Pour un bar_chart simple, on fabrique un label lisible
    sub_plot = sub.copy()
    sub_plot["Poste"] = sub_plot["INDICATEUR"] + " — " + sub_plot["SOUS-INDICATEUR"]
    st.bar_chart(sub_plot.set_index("Poste")[["VALEUR"]])

with col2:
    st.caption("Chiffres (Top sous-catégories)")
    sub_table = sub.copy().rename(columns={"VALEUR": "Valeur"})
    st.dataframe(sub_table, use_container_width=True)

st.divider()

# ---------------------------
# Effet crise : ce qui change
# ---------------------------
st.subheader("Effet de la crise : ce qui change")

lvl = group_level
impact = cmp.groupby(lvl)[[normal_col, crise_col, "CHANGEMENT"]].sum().reset_index()
impact["EVOLUTION_%"] = np.where(
    impact[normal_col].fillna(0) == 0,
    np.nan,
    (impact["CHANGEMENT"] / impact[normal_col]) * 100
)
impact["TENDANCE"] = impact["CHANGEMENT"].apply(lambda x: "⬆️" if x > 0 else ("⬇️" if x < 0 else "➜"))
impact = impact.sort_values("CHANGEMENT", ascending=False)

a, b = st.columns(2)

with a:
    st.caption("Plus fortes hausses")
    st.dataframe(
        impact[[lvl, "TENDANCE", "CHANGEMENT", "EVOLUTION_%"]].head(10).rename(columns={
            lvl: "Poste",
            "CHANGEMENT": "Changement (volume)",
            "EVOLUTION_%": "Évolution (%)"
        }),
        use_container_width=True
    )

with b:
    st.caption("Plus fortes baisses")
    st.dataframe(
        impact[[lvl, "TENDANCE", "CHANGEMENT", "EVOLUTION_%"]].tail(10).sort_values("CHANGEMENT").rename(columns={
            lvl: "Poste",
            "CHANGEMENT": "Changement (volume)",
            "EVOLUTION_%": "Évolution (%)"
        }),
        use_container_width=True
    )

# ---------------------------
# Detail table
# ---------------------------
if show_details:
    st.subheader("Tableau détaillé (normal vs crise)")
    out = cmp.copy().rename(columns={
        normal_col: "Valeur (normal)",
        crise_col: "Valeur (crise)",
        "CHANGEMENT": "Changement (volume)",
        "EVOLUTION_%": "Évolution (%)",
        "TENDANCE": "Tendance"
    })
    out = out[
        [
            "ANNEE",
            "INDICATEUR",
            "SOUS-INDICATEUR",
            "UNITE",
            "Valeur (normal)",
            "Valeur (crise)",
            "Changement (volume)",
            "Évolution (%)",
            "Tendance",
        ]
    ]
    st.dataframe(out.sort_values(["ANNEE", "INDICATEUR", "SOUS-INDICATEUR"]), use_container_width=True)

st.caption("Note : le mode crise est une simulation basée sur des coefficients (scénarios).")