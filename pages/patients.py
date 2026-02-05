# pages/patients.py — Patients page content
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

from utils import load_data

# Couleur dédiée pour l'année 2017 (prévision)
COULEUR_2017 = "#E67E22"  # Orange

PALETTE_ANNEES = [
    "#003A8F",  # Bleu AP-HP foncé (institutionnel)
    "#0B5ED7",  # Bleu hospitalier standard
    "#1F77B4",  # Bleu scientifique (charts)
    "#4A90E2",  # Bleu moyen
    "#6BAED6",  # Bleu clair
    "#9ECAE1",  # Bleu très clair
    "#E67E22",  # Gris technique
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

    # Filtre année (données mensuelles : year ; patients-all contient déjà 2017)
    dff = df.copy()
    year_col = "ANNEE" if "ANNEE" in dff.columns else "year"
    if year_choice != "Toutes" and year_col in dff.columns:
        dff[year_col] = dff[year_col].astype(int)
        dff = dff[dff[year_col] == int(year_choice)]
    # Par défaut les données 2017 sont masquées ; la case « Afficher prévision 2017 » les active
    if not show_forecast and year_col in dff.columns:
        dff = dff[dff[year_col].astype(int) != 2017]
    has_2017 = year_col in df.columns and (df[year_col].astype(int) == 2017).any()

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

                        # 2017 en orange (déjà inclus dans patients-all.csv)
                        color_range = [COULEUR_2017 if y == 2017 else PALETTE_ANNEES[i % len(PALETTE_ANNEES)] for i, y in enumerate(years)]
                        chart_obj = (
                            alt.Chart(agg)
                            .mark_line(point=True)
                            .encode(
                                x=alt.X("month:O", title="Mois"),
                                y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                color=alt.Color("year:O", title="Année", scale=alt.Scale(domain=years, range=color_range)),
                            )
                        )

                        sub = f"{subtitle} — profil mensuel multi-années ({mode_label})"
                        if has_2017 and multi_year:
                            sub += " — 2017 = prévision SARIMA"

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
                        # Couleur fixe : orange pour 2017 (déjà dans le CSV), sinon palette
                        color_annee = COULEUR_2017 if int(year_choice) == 2017 else PALETTE_ANNEES[years.index(int(year_choice)) % len(PALETTE_ANNEES)] if int(year_choice) in years else PALETTE_ANNEES[0]
                        chart_obj = (
                            alt.Chart(agg)
                            .mark_line(point=True)
                            .encode(
                                x=alt.X("month:O", title="Mois"),
                                y=alt.Y("value:Q", title=f"Volume mensuel ({unite})", axis=alt.Axis(format=",.2f")),
                                color=alt.value(color_annee),
                            )
                        )

                        sub = f"{subtitle} — profil mensuel {year_choice} ({mode_label})"
                        if has_2017 and int(year_choice) == 2017:
                            sub += " (prévision SARIMA)"

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
