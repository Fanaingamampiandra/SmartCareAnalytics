# pages/quality.py — Qualité page content

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
    month_choice,
    mode_choice,
    normal_col,  # gardé pour compatibilité mais non utilisé avec le CSV journalier
    crise_col,   # idem
    years,
    **kwargs,
):
    show_forecast=False,
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

    # Filtre site / total à partir de la colonne daily `site_code`
    hospital_choice = kwargs.get("hospital_choice", "TOTAL")
    if "site_code" in df.columns:
        if hospital_choice in ("PLF", "CFX"):
            df = df[df["site_code"] == hospital_choice]
        # TOTAL => on garde PLF + CFX

    # Filtre année (colonne `year` dans le CSV journalier)
    dff = df.copy()
    year_col = "ANNEE" if "ANNEE" in dff.columns else "year"
    if year_choice != "Toutes" and year_col in dff.columns:
        dff[year_col] = dff[year_col].astype(int)
        dff = dff[dff[year_col] == int(year_choice)]

    # Filtre mois si demandé (colonne `month` créée dans la saisonnalité)
    if month_choice != "Tous" and "month" in dff.columns:
        dff["month"] = dff["month"].astype(int)
        dff = dff[dff["month"] == int(month_choice)]

    # Colonnes attendues dans le CSV journalier
    required_cols = {"indicateur", "sous_indicateur", "unite", value_col}
    if not required_cols.issubset(dff.columns):
        st_module.error(
            f"Le CSV journalier ne contient pas les colonnes attendues ({', '.join(sorted(required_cols))})."
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
            monthly_mode = month_choice == "Tous"

            for unite in unites:
                df_u = df_indic[df_indic["unite"] == unite]
                sous_labels = sorted(df_u["sous_indicateur"].dropna().unique().tolist())

                for sous in sous_labels:
                    df_s = df_u[df_u["sous_indicateur"] == sous]
                    subtitle = f"{sous} ({unite})"
                    mode_label = "Situation normale" if mode_choice == "Normal" else "Crise (simulation)"

                    # 1) Toutes les années + tous les mois -> courbes mensuelles par année
                    if multi_year and monthly_mode:
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

                    # 2) Une seule année + tous les mois -> profil mensuel de l'année
                    elif (not multi_year) and monthly_mode:
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

                    # 3) Mois spécifique -> série journalière
                    else:
                        if not {"date", "year"}.issubset(df_s.columns):
                            continue
                        agg = (
                            df_s.groupby(["year", "date"])[value_col]
                            .sum()
                            .reset_index()
                            .rename(columns={value_col: "value"})
                        )
                        agg["year"] = agg["year"].astype(int)
                        # Afficher le jour du mois sur l’axe X, couleur par année
                        if multi_year:
                            chart_obj = (
                                alt.Chart(agg)
                                .mark_line(point=True)
                                .encode(
                                    x=alt.X(
                                        "date:T",
                                        title="Jour",
                                        axis=alt.Axis(format="%d"),
                                    ),
                                    y=alt.Y("value:Q", title=f"Volume journalier ({unite})", axis=alt.Axis(format=",.2f")),
                                    color=alt.Color(
                                        "year:O",
                                        title="Année",
                                        scale=alt.Scale(domain=years, range=PALETTE_ANNEES),
                                    ),
                                )
                            )
                        else:
                            year_idx = years.index(int(year_choice)) if int(year_choice) in years else 0
                            color_annee = PALETTE_ANNEES[year_idx % len(PALETTE_ANNEES)]
                            chart_obj = (
                                alt.Chart(agg)
                                .mark_line(point=True)
                                .encode(
                                    x=alt.X(
                                        "date:T",
                                        title="Jour",
                                        axis=alt.Axis(format="%d"),
                                    ),
                                    y=alt.Y("value:Q", title=f"Volume journalier ({unite})", axis=alt.Axis(format=",.2f")),
                                    color=alt.value(color_annee),
                                )
                            )
                            
                            # Ajouter les prédictions 2017 si demandé et si on affiche 2017
                            if show_forecast and int(year_choice) == 2017:
                                hospital_choice = kwargs.get("hospital_choice", "TOTAL")
                                df_forecast = generate_forecast_2017(
                                    df, hospital_choice, indic, sous, value_col
                                )
                                if not df_forecast.empty and "month" in df_forecast.columns:
                                    forecast_filtered = df_forecast[
                                        df_forecast["month"] == int(month_choice)
                                    ]
                                    if not forecast_filtered.empty:
                                        forecast_chart = (
                                            alt.Chart(forecast_filtered)
                                            .mark_line(point=True, strokeDash=[5, 5])
                                            .encode(
                                                x=alt.X("date:T", title="Jour", axis=alt.Axis(format="%d")),
                                                y=alt.Y("value:Q", title=f"Volume journalier ({unite})", axis=alt.Axis(format=",.2f")),
                                                color=alt.value("#FF6B6B"),
                                            )
                                        )
                                        chart_obj = chart_obj + forecast_chart
                        
                        # Ajouter les prédictions 2017 si demandé (cas multi_year)
                        if show_forecast and multi_year:
                            hospital_choice = kwargs.get("hospital_choice", "TOTAL")
                            df_forecast = generate_forecast_2017(
                                df, hospital_choice, indic, sous, value_col
                            )
                            if not df_forecast.empty and "month" in df_forecast.columns:
                                forecast_filtered = df_forecast[
                                    df_forecast["month"] == int(month_choice)
                                ]
                                if not forecast_filtered.empty:
                                    forecast_chart = (
                                        alt.Chart(forecast_filtered)
                                        .mark_line(point=True, strokeDash=[5, 5])
                                        .encode(
                                            x=alt.X("date:T", title="Jour", axis=alt.Axis(format="%d")),
                                            y=alt.Y("value:Q", title=f"Volume journalier ({unite})", axis=alt.Axis(format=",.2f")),
                                            color=alt.value("#FF6B6B"),
                                        )
                                    )
                                    chart_obj = chart_obj + forecast_chart
                        
                        if multi_year:
                            sub = f"{subtitle} — série journalière (mois {month_choice}, toutes années) ({mode_label})"
                        else:
                            sub = f"{subtitle} — série journalière {year_choice}-{int(month_choice):02d} ({mode_label})"

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
