# pages/logistics.py — Logistique page content
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

from utils import load_data

PALETTE_ANNEES = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]


def render(
    st_module,
    *,
    data_path,
    year_choice,
    mode_choice,
    normal_col,
    crise_col,
    years,
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

    indicateurs = sorted(cmp["INDICATEUR"].dropna().unique().tolist())
    tabs = st_module.tabs(indicateurs)

    for tab, indic in zip(tabs, indicateurs):
        with tab:
            st_module.subheader(indic)
            df_indic = cmp[cmp["INDICATEUR"] == indic]
            unites = sorted(df_indic["UNITE"].dropna().unique().tolist())

            for unite in unites:
                df_u = df_indic[df_indic["UNITE"] == unite]
                sous_labels = sorted(df_u["SOUS-INDICATEUR"].dropna().unique().tolist())
                subtitle = f"{sous_labels[0]} ({unite})" if len(sous_labels) == 1 else f"({unite})"
                by_year = year_choice == "Toutes"

                if by_year:
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
                            y=alt.Y("Valeur:Q", title=f"Volume ({unite})", axis=alt.Axis(format=",.2f")),
                            color=alt.Color("ANNEE:O", title="Année", scale=alt.Scale(domain=years, range=PALETTE_ANNEES)),
                        )
                    )
                    sub = f"{subtitle} — par année ({mode_label})"
                    chart = bars.properties(title={"text": indic, "subtitle": sub}, height=450)
                else:
                    col_to_show = normal_col if mode_choice == "Normal" else crise_col
                    mode_label = "Situation normale" if mode_choice == "Normal" else "Crise (simulation)"
                    year_idx = years.index(int(year_choice)) if int(year_choice) in years else 0
                    color_annee = PALETTE_ANNEES[year_idx % len(PALETTE_ANNEES)]
                    agg = df_u.groupby("SOUS-INDICATEUR")[[col_to_show]].sum().reset_index()
                    agg = agg.rename(columns={col_to_show: "Valeur"})
                    bars = (
                        alt.Chart(agg)
                        .mark_bar()
                        .encode(
                            x=alt.X("SOUS-INDICATEUR:N", title="Sous-indicateur"),
                            y=alt.Y("Valeur:Q", title=f"Volume ({unite})", axis=alt.Axis(format=",.2f")),
                            color=alt.value(color_annee),
                        )
                    )
                    chart = bars.properties(
                        title={"text": indic, "subtitle": f"{subtitle} — {mode_label}"},
                        height=450,
                    )
                st_module.altair_chart(chart, use_container_width=True)
                st_module.markdown("<div style='margin-bottom: 3.5rem;'></div>", unsafe_allow_html=True)

            with st_module.expander("Voir le détail (tableau)"):
                out = df_indic.rename(columns={
                    normal_col: "Valeur (normal)",
                    crise_col: "Valeur (crise)",
                    "CHANGEMENT": "Changement",
                    "EVOLUTION_%": "Évolution (%)",
                    "TENDANCE": "Tendance",
                })
                cols = ["ANNEE", "UNITE", "SOUS-INDICATEUR", "Valeur (normal)", "Valeur (crise)", "Changement", "Évolution (%)", "Tendance"]
                st_module.dataframe(
                    out[[c for c in cols if c in out.columns]].sort_values(["UNITE", "SOUS-INDICATEUR", "ANNEE"]),
                    use_container_width=True,
                )
