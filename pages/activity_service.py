#pages/activity_service.py — Activité & Service content
import streamlit as st
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from utils import load_data

COLOR_NORMAL = "#2E86AB"
COLOR_CRISE = "#A23B72"
COLOR_STABLE = "#06A77D"
COLOR_VOLATILE = "#D64550"

MONTH_NAMES = {
    1: "Jan",
    2: "Fev",
    3: "Mar",
    4: "Avr",
    5: "Mai",
    6: "Jun",
    7: "Jul",
    8: "Aou",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}

ACTIVITY_KEYWORDS = [
    "activit",
    "acte",
    "sejour",
    "sejours",
    "urgenc",
    "hospital",
    "ambulatoire",
    "mco",
    "cause",
    "admission",
]


def classify_category(indicateur):
    text = str(indicateur).lower()
    for key in ACTIVITY_KEYWORDS:
        if key in text:
            return "Activites"
    return "Services"


def render(
    st_module,
    *,
    data_path,
    year_choice,
    mode_choice,
    normal_col,  # gardé pour compatibilité
    crise_col,
    years,
    show_forecast=False,
    **kwargs,
):
    """
    Dashboard interactif pour l'analyse d'activite/service PSL-CFX.
    Affiche: graphiques, camemberts, statistiques, comparaisons Normal/Crise.
    """

    if not data_path:
        st_module.info("Aucune donnee configuree pour cette page.")
        return

    try:
        df = load_data(data_path)
    except Exception as e:
        st_module.error(f"Impossible de charger les donnees : {e}")
        return

    base_cols = {
        "year",
        "month",
        "site_code",
        "indicateur",
        "sous_indicateur",
        "unite",
        "value",
    }
    if not base_cols.issubset(df.columns):
        st_module.error(f"Colonnes de base manquantes : {base_cols - set(df.columns)}")
        return

    if "value_crise" not in df.columns:
        st_module.warning(
            "⚠️ Colonne 'value_crise' manquante. Generation automatique avec facteur 1.3x"
        )
        df["value_crise"] = df["value"] * 1.3

    if "type" not in df.columns:
        cv_by_indic = (
            df.groupby(["indicateur", "sous_indicateur"])["value"]
            .apply(lambda x: (x.std() / x.mean() * 100) if x.mean() > 0 else 0)
            .to_dict()
        )
        df["type"] = df.apply(
            lambda row: "VOLATILE"
            if cv_by_indic.get((row["indicateur"], row["sous_indicateur"]), 0) > 10
            else "STABLE",
            axis=1,
        )

    if "categorie" not in df.columns:
        df["categorie"] = df["indicateur"].apply(classify_category)

    hospital_choice = kwargs.get("hospital_choice", "TOTAL")
    if hospital_choice in ("PLF", "CFX"):
        df = df[df["site_code"] == hospital_choice]

    if year_choice != "Toutes":
        df = df[df["year"] == int(year_choice)]
    if not show_forecast:
        df = df[df["year"] != 2017]

    value_col = "value_crise" if mode_choice == "Crise" else "value"
    mode_label = "CRISE" if mode_choice == "Crise" else "NORMAL"

    st_module.markdown(
        f"### 📊 Dashboard Activite Service — Mode **{mode_label}**"
    )

    def render_section(df_section, title):
        if df_section.empty:
            st_module.info(f"Aucune donnee pour {title}.")
            return

        tabs = st_module.tabs(
            [
                "📈 Vue Globale",
                "🔍 Par Indicateur",
                "⚖️ Comparaison Normal/Crise",
                "📋 Donnees Brutes",
            ]
        )

        with tabs[0]:
            st_module.subheader(f"Resume Global — {title}")

            col1, col2, col3, col4 = st_module.columns(4)

            total_activity = df_section[value_col].sum()
            avg_monthly = df_section.groupby("month")[value_col].sum().mean()
            nb_months = df_section["month"].nunique()
            nb_indicateurs = df_section["indicateur"].nunique()

            col1.metric("📊 Total", f"{total_activity:,.0f}")
            col2.metric("📅 Moy. Mensuelle", f"{avg_monthly:,.0f}")
            col3.metric("🗓️ Mois", f"{nb_months}")
            col4.metric("📈 Indicateurs", f"{nb_indicateurs}")

            st_module.markdown("---")
            monthly_data = df_section.groupby("month")[value_col].sum().reset_index()
            monthly_data["mois_nom"] = monthly_data["month"].map(MONTH_NAMES)

            fig_monthly = px.bar(
                monthly_data,
                x="mois_nom",
                y=value_col,
                title="Evolution Mensuelle - Total",
                labels={value_col: "Volume", "mois_nom": "Mois"},
                color_discrete_sequence=[
                    COLOR_CRISE if mode_choice == "Crise" else COLOR_NORMAL
                ],
            )
            fig_monthly.update_layout(height=400, xaxis_title="Mois", yaxis_title="Activite")
            st_module.plotly_chart(fig_monthly, use_container_width=True)

            st_module.markdown("---")
            indic_data = (
                df_section.groupby("indicateur")[value_col]
                .sum()
                .reset_index()
                .sort_values(value_col, ascending=False)
            )
            fig_pie = px.pie(
                indic_data,
                values=value_col,
                names="indicateur",
                title="Repartition par Indicateur",
                hole=0.3,
            )
            st_module.plotly_chart(fig_pie, use_container_width=True)

            st_module.markdown("---")
            site_data = df_section.groupby("site_code")[value_col].sum().reset_index()
            fig_site = px.pie(
                site_data,
                values=value_col,
                names="site_code",
                title="Repartition PLF vs CFX",
                color_discrete_map={"PLF": COLOR_NORMAL, "CFX": COLOR_CRISE},
            )
            st_module.plotly_chart(fig_site, use_container_width=True)

        with tabs[1]:
            st_module.subheader("Analyse Detaillee par Indicateur")

            indicateurs = sorted(df_section["indicateur"].dropna().unique().tolist())
            selected_indic = st_module.selectbox(
                "Selectionner un indicateur", indicateurs, key=f"indic_{title}"
            )

            df_indic = df_section[df_section["indicateur"] == selected_indic]

            st_module.markdown("##### Statistiques")
            stats = (
                df_indic.groupby("sous_indicateur")[value_col]
                .agg(
                    [
                        ("Total", "sum"),
                        ("Moyenne", "mean"),
                        ("Min", "min"),
                        ("Max", "max"),
                        ("Ecart-Type", "std"),
                    ]
                )
                .round(2)
                .sort_values("Total", ascending=False)
            )
            st_module.dataframe(stats, use_container_width=True)

            st_module.markdown("---")
            monthly_indic = (
                df_indic.groupby(["month", "sous_indicateur"])[value_col]
                .sum()
                .reset_index()
            )
            if not monthly_indic.empty:
                fig_indic = px.line(
                    monthly_indic,
                    x="month",
                    y=value_col,
                    color="sous_indicateur",
                    title=f"Evolution Mensuelle — {selected_indic}",
                    markers=True,
                    labels={value_col: "Valeur", "month": "Mois"},
                )
                fig_indic.update_layout(height=400)
                st_module.plotly_chart(fig_indic, use_container_width=True)

            sous_data = (
                df_indic.groupby("sous_indicateur")[value_col]
                .sum()
                .reset_index()
            )
            fig_sous = px.pie(
                sous_data,
                values=value_col,
                names="sous_indicateur",
                title=f"Repartition — {selected_indic}",
            )
            st_module.plotly_chart(fig_sous, use_container_width=True)

        with tabs[2]:
            st_module.subheader("Comparaison Mode NORMAL vs CRISE")

            comparison = (
                df_section.groupby(["year", "month"])
                .agg({"value": "sum", "value_crise": "sum"})
                .reset_index()
            )
            comparison["month_nom"] = comparison["month"].map(MONTH_NAMES)
            comparison["ecart_pct"] = (
                (comparison["value_crise"] - comparison["value"])
                / comparison["value"].replace(0, np.nan)
                * 100
            ).round(2)

            fig_comp = make_subplots(specs=[[{"secondary_y": False}]])
            fig_comp.add_trace(
                go.Scatter(
                    x=comparison["month"],
                    y=comparison["value"],
                    name="Mode NORMAL",
                    mode="lines+markers",
                    line=dict(color=COLOR_NORMAL, width=3),
                )
            )
            fig_comp.add_trace(
                go.Scatter(
                    x=comparison["month"],
                    y=comparison["value_crise"],
                    name="Mode CRISE",
                    mode="lines+markers",
                    line=dict(color=COLOR_CRISE, width=3, dash="dash"),
                )
            )
            fig_comp.update_layout(
                title="Comparaison des Profils Mensuels",
                xaxis_title="Mois",
                yaxis_title="Activite",
                height=400,
                hovermode="x unified",
            )
            st_module.plotly_chart(fig_comp, use_container_width=True)

            st_module.markdown("---")
            fig_ecart = px.bar(
                comparison,
                x="month_nom",
                y="ecart_pct",
                title="Ecart CRISE vs NORMAL (%)",
                labels={"ecart_pct": "Ecart %", "month_nom": "Mois"},
                color="ecart_pct",
                color_continuous_scale="RdYlGn_r",
            )
            st_module.plotly_chart(fig_ecart, use_container_width=True)

            st_module.markdown("---")
            st_module.markdown("##### Tableau Comparatif")
            comp_table = comparison[
                ["month", "month_nom", "value", "value_crise", "ecart_pct"]
            ].copy()
            comp_table.columns = [
                "Mois (num)",
                "Mois",
                "Normal",
                "Crise",
                "Ecart (%)",
            ]
            st_module.dataframe(comp_table, use_container_width=True)

        with tabs[3]:
            st_module.subheader("Donnees Mensuelles Brutes")

            col1, col2 = st_module.columns(2)
            with col1:
                selected_site = st_module.multiselect(
                    "Sites",
                    df_section["site_code"].unique(),
                    default=df_section["site_code"].unique(),
                    key=f"site_{title}",
                )
            with col2:
                selected_type = st_module.multiselect(
                    "Type",
                    df_section["type"].unique(),
                    default=df_section["type"].unique(),
                    key=f"type_{title}",
                )

            df_filtered = df_section[
                (df_section["site_code"].isin(selected_site))
                & (df_section["type"].isin(selected_type))
            ].copy()

            pivot_data = (
                df_filtered.groupby(["year", "month", "indicateur", "sous_indicateur"])
                .agg({"value": "sum", "value_crise": "sum"})
                .reset_index()
            )
            st_module.dataframe(
                pivot_data.sort_values(["year", "month"]), use_container_width=True
            )

            csv = pivot_data.to_csv(index=False).encode("utf-8")
            st_module.download_button(
                label="Telecharger CSV",
                data=csv,
                file_name="activite_service_data.csv",
                mime="text/csv",
            )

    tabs_main = st_module.tabs(["Activites", "Services", "Tout"])
    with tabs_main[0]:
        render_section(df[df["categorie"] == "Activites"], "Activites")
    with tabs_main[1]:
        render_section(df[df["categorie"] == "Services"], "Services")
    with tabs_main[2]:
        render_section(df, "Tout")
