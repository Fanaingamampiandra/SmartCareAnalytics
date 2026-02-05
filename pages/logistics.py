# pages/logistics.py — Logistique page content
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

from utils import load_data, generate_forecast_2017

PALETTE_ANNEES = [
    "#003A8F",  # Bleu AP-HP foncé (institutionnel)
    "#0B5ED7",  # Bleu hospitalier standard
    "#1F77B4",  # Bleu scientifique (charts)
    "#4A90E2",  # Bleu moyen
    "#6BAED6",  # Bleu clair
    "#9ECAE1",  # Bleu très clair
    "#DCEAF7",  # Bleu blanc cassé
    "#F2F6FA",  # Blanc bleuté (fond)
    "#B0B7C3",  # Gris hospitalier clair
    "#6C757D",  # Gris technique
]

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
    if not data_path:
        st_module.info("Aucune donnée configurée pour cette page.")
        return
    try:
        df = load_data(data_path)
    except Exception as e:
        st_module.error(f"Impossible de charger les données : {e}")
        return

    # Choix de la colonne de valeur selon le mode (Normal / Crise)
    value_col = "value"
    if mode_choice == "Crise" and "value_crise" in df.columns:
        value_col = "value_crise"

    # Filtre site / total (colonne `site_code`)
    hospital_choice = kwargs.get("hospital_choice", "TOTAL")
    if "site_code" in df.columns:
        if hospital_choice in ("PLF", "CFX"):
            df = df[df["site_code"] == hospital_choice]
        # TOTAL => on garde PLF + CFX

    # Filtre année (données mensuelles : year)
    dff = df.copy()
    year_col = "ANNEE" if "ANNEE" in dff.columns else "year"
    if year_choice != "Toutes" and year_col in dff.columns:
        dff[year_col] = dff[year_col].astype(int)
        dff = dff[dff[year_col] == int(year_choice)]

    # Colonnes attendues (CSV mensuel)
    required_cols = {"indicateur", "sous_indicateur", "unite", value_col}
    if not required_cols.issubset(dff.columns):
        st_module.error(
            f"Le CSV ne contient pas les colonnes attendues ({', '.join(sorted(required_cols))})."
        )
        return

    indicateurs = sorted(dff["indicateur"].dropna().unique().tolist())
    tabs = st_module.tabs(indicateurs)

    for tab, indic in zip(tabs, indicateurs):
        with tab:
            st_module.subheader(indic)
            df_indic = dff[dff["indicateur"] == indic]
            unites = sorted(df_indic["unite"].dropna().unique().tolist())

            multi_year = year_choice == "Toutes"

            for unite in unites:
                df_u = df_indic[df_indic["unite"] == unite]
                sous_labels = sorted(df_u["sous_indicateur"].dropna().unique().tolist())

                for sous in sous_labels:
                    df_s = df_u[df_u["sous_indicateur"] == sous]
                    subtitle = f"{sous} ({unite})"
                    mode_label = "Situation normale" if mode_choice == "Normal" else "Crise (simulation)"

                    # 1) Toutes les années -> courbes mensuelles par année
                    if multi_year:
                        if not {"year", "month"}.issubset(df_s.columns):
                            continue
                        agg = (
                            df_s.groupby(["year", "month"])[value_col]
                            .sum()
                            .reset_index()
                            .rename(columns={value_col: "value"})
                        )
                        agg["year"] = agg["year"].astype(int)
                        agg["month"] = agg["month"].astype(int)
                        
                        # Ajouter les prédictions 2017 si demandé
                        forecast_data = None
                        if show_forecast:
                            hospital_choice = kwargs.get("hospital_choice", "TOTAL")
                            df_forecast = generate_forecast_2017(
                                df, hospital_choice, indic, sous, value_col
                            )
                            if not df_forecast.empty:
                                forecast_agg = (
                                    df_forecast.groupby("month")["value"]
                                    .sum()
                                    .reset_index()
                                )
                                forecast_agg["year"] = 2017
                                forecast_agg["month"] = forecast_agg["month"].astype(int)
                                forecast_data = forecast_agg
                        
                        chart_obj = (
                            alt.Chart(agg)
                            .mark_line(point=True)
                            .encode(
                                x=alt.X("month:O", title="Mois"),
                                y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                color=alt.Color("year:O", title="Année", scale=alt.Scale(domain=years, range=PALETTE_ANNEES)),
                            )
                        )
                        
                        # Ajouter la ligne de prédiction si disponible
                        if forecast_data is not None:
                            forecast_chart = (
                                alt.Chart(forecast_data)
                                .mark_line(point=True, strokeDash=[5, 5])
                                .encode(
                                    x=alt.X("month:O", title="Mois"),
                                    y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                    color=alt.value("#FF6B6B"),  # Rouge pour la prédiction
                                )
                            )
                            chart_obj = chart_obj + forecast_chart
                        
                        sub = f"{subtitle} — profil mensuel multi-années ({mode_label})"

                    # 2) Une seule année -> profil mensuel de l'année
                    else:
                        if "month" not in df_s.columns:
                            continue
                        agg = (
                            df_s.groupby("month")[value_col]
                            .sum()
                            .reset_index()
                            .rename(columns={value_col: "value"})
                        )
                        agg["month"] = agg["month"].astype(int)
                        # Couleur fixe liée à l'année sélectionnée
                        year_idx = years.index(int(year_choice)) if int(year_choice) in years else 0
                        color_annee = PALETTE_ANNEES[year_idx % len(PALETTE_ANNEES)]
                        chart_obj = (
                            alt.Chart(agg)
                            .mark_line(point=True)
                            .encode(
                                x=alt.X("month:O", title="Mois"),
                                y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                color=alt.value(color_annee),
                            )
                        )
                        
                        # Ajouter les prédictions 2017 si demandé et si on affiche 2017
                        if show_forecast and int(year_choice) == 2017:
                            hospital_choice = kwargs.get("hospital_choice", "TOTAL")
                            df_forecast = generate_forecast_2017(
                                df, hospital_choice, indic, sous, value_col
                            )
                            if not df_forecast.empty:
                                forecast_agg = (
                                    df_forecast.groupby("month")["value"]
                                    .sum()
                                    .reset_index()
                                )
                                forecast_agg["month"] = forecast_agg["month"].astype(int)
                                forecast_chart = (
                                    alt.Chart(forecast_agg)
                                    .mark_line(point=True, strokeDash=[5, 5])
                                    .encode(
                                        x=alt.X("month:O", title="Mois"),
                                        y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                        color=alt.value("#FF6B6B"),
                                    )
                                )
                                chart_obj = chart_obj + forecast_chart
                        
                        sub = f"{subtitle} — profil mensuel {year_choice} ({mode_label})"


                    chart = chart_obj.properties(
                        title={"text": indic, "subtitle": sub},
                        height=350,
                    )

                    st_module.altair_chart(chart, use_container_width=True)
                    st_module.markdown("<div style='margin-bottom: 3.5rem;'></div>", unsafe_allow_html=True)

            # Tableau détaillé annuel par sous-indicateur / année
            with st_module.expander("Voir le détail (tableau)"):
                table = (
                    df_indic
                    .groupby(["year", "unite", "sous_indicateur"])[value_col]
                    .sum()
                    .reset_index()
                    .rename(columns={
                        "year": "ANNEE",
                        "unite": "UNITE",
                        "sous_indicateur": "SOUS-INDICATEUR",
                        value_col: "Valeur annuelle",
                    })
                )
                cols = ["ANNEE", "UNITE", "SOUS-INDICATEUR", "Valeur annuelle"]
                st_module.dataframe(
                    table[cols].sort_values(["UNITE", "SOUS-INDICATEUR", "ANNEE"]),
                    use_container_width=True,
                )
