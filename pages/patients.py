# pages/patients.py — Patients page content
"""
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

from utils import load_data

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
"""sumary_line

"""
patients.py - Module Dashboard Patients PSL-CFX
Analyse de l'activité hospitalière avec simulation de crise sanitaire
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from typing import Optional, List, Dict


# ============================================================================
# CONFIGURATION ET CONSTANTES
# ============================================================================

PALETTE_APHP = {
    "bleu_fonce": "#003A8F",
    "bleu_hospital": "#0B5ED7",
    "bleu_clair": "#4A90E2",
    "bleu_tres_clair": "#9ECAE1",
    "rouge_alerte": "#DC3545",
    "vert_normal": "#28A745",
    "orange_warning": "#FFA500",
    "gris": "#6C757D",
    "blanc_bleute": "#F2F6FA"
}

PALETTE_ANNEES = [
    "#003A8F", "#0B5ED7", "#1F77B4", "#4A90E2",
    "#6BAED6", "#9ECAE1", "#DCEAF7", "#B0B7C3"
]


# ============================================================================
# FONCTIONS UTILITAIRES
# ============================================================================

def format_nombre(nombre: float) -> str:
    """Formate les nombres pour affichage (1234 -> 1 234)"""
    if pd.isna(nombre):
        return "N/A"
    return f"{nombre:,.0f}".replace(",", " ")


def calcul_variation(val_normal: float, val_crise: float) -> float:
    """Calcule la variation en pourcentage"""
    if val_normal == 0 or pd.isna(val_normal):
        return 0
    return ((val_crise - val_normal) / val_normal) * 100


def get_tendance_icon(variation: float) -> str:
    """Retourne l'icône de tendance selon la variation"""
    if variation > 10:
        return "⬆️ Forte hausse"
    elif variation > 0:
        return "↗️ Hausse"
    elif variation < -10:
        return "⬇️ Forte baisse"
    elif variation < 0:
        return "↘️ Baisse"
    return "➡️ Stable"


def load_data(data_path: str) -> Optional[pd.DataFrame]:
    """Charge les données avec gestion d'erreurs"""
    try:
        df = pd.read_csv(data_path, encoding='utf-8')
        return df
    except Exception as e:
        st.error(f"❌ Erreur de chargement : {e}")
        return None


# ============================================================================
# COMPOSANTS UI
# ============================================================================

def apply_custom_css():
    """Applique le CSS personnalisé pour le thème hospitalier"""
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
        
        * { font-family: 'Inter', sans-serif; }
        
        .main-title {
            background: linear-gradient(135deg, #003A8F 0%, #0B5ED7 100%);
            color: white;
            padding: 2rem;
            border-radius: 10px;
            margin-bottom: 2rem;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .main-title h1 {
            margin: 0;
            font-size: 2.5rem;
            font-weight: 700;
        }
        
        .main-title p {
            margin: 0.5rem 0 0 0;
            font-size: 1.1rem;
            opacity: 0.9;
        }
        
        .section-header {
            display: flex;
            align-items: center;
            gap: 1rem;
            margin: 2rem 0 1rem 0;
            padding-bottom: 0.5rem;
            border-bottom: 3px solid #003A8F;
        }
        
        .section-header h2 {
            margin: 0;
            color: #003A8F;
            font-weight: 700;
        }
        
        .alert-box {
            padding: 1rem;
            border-radius: 8px;
            margin: 1rem 0;
            border-left: 4px solid;
        }
        
        .alert-box.info {
            background: #D1ECF1;
            border-color: #0C5460;
            color: #0C5460;
        }
        
        .alert-box.danger {
            background: #F8D7DA;
            border-color: #721C24;
            color: #721C24;
        }
        
        .metric-highlight {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            border-left: 5px solid #003A8F;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            margin-bottom: 1rem;
        }
        
        .footer {
            margin-top: 3rem;
            padding: 2rem;
            background: #F2F6FA;
            border-radius: 10px;
            text-align: center;
            color: #6C757D;
        }
    </style>
    """, unsafe_allow_html=True)


def render_header():
    """Affiche l'en-tête du dashboard"""
    st.markdown("""
    <div class="main-title">
        <h1>🏥 Hôpitaux Universitaires Pitié Salpêtrière Charles Foix</h1>
        <p>Dashboard d'Analyse de l'Activité des Patients - Projet Data</p>
    </div>
    """, unsafe_allow_html=True)


def render_sidebar(df: pd.DataFrame) -> Dict:
    """Affiche la sidebar et retourne les paramètres sélectionnés"""
    with st.sidebar:
        st.image("https://upload.wikimedia.org/wikipedia/fr/thumb/e/ef/Logo_AP-HP_2020.svg/1200px-Logo_AP-HP_2020.svg.png", 
                 width=200)
        
        st.markdown("### ⚙️ Paramètres de Simulation")
        
        # Mode Normal / Crise
        mode_choice = st.radio(
            "**Mode d'analyse**",
            options=["Normal", "Crise Sanitaire"],
            help="Basculez entre situation normale et simulation de crise sanitaire (+80% activité)",
            horizontal=False
        )
        
        # Affichage du statut
        if mode_choice == "Crise Sanitaire":
            st.error("🚨 **MODE CRISE ACTIVÉ**")
            st.markdown("Impact simulé : **+80% d'activité**")
            col_valeur = "TOTAL_CRISE"
            couleur_mode = PALETTE_APHP["rouge_alerte"]
        else:
            st.success("✅ **MODE NORMAL**")
            st.markdown("Activité hospitalière standard")
            col_valeur = "TOTAL_NORMAL"
            couleur_mode = PALETTE_APHP["vert_normal"]
        
        st.markdown("---")
        
        # Sélection année
        years = sorted(df["ANNEE"].unique().tolist())
        year_choice = st.selectbox(
            "**Année d'analyse**",
            options=["Toutes"] + [str(y) for y in years],
            help="Sélectionnez une année spécifique ou toutes les années"
        )
        
        st.markdown("---")
        
        # Filtres avancés
        with st.expander("🔍 Filtres Avancés"):
            indicateurs_disponibles = sorted(df["INDICATEUR"].dropna().unique().tolist())
            filtre_indicateur = st.multiselect(
                "Filtrer par indicateur",
                options=indicateurs_disponibles,
                default=indicateurs_disponibles
            )
        
        st.markdown("---")
        
        # Informations projet
        st.markdown("### 📊 Informations Projet")
        st.info("""
        **Objectifs :**
        - Anticiper les pics d'activité
        - Optimiser les ressources
        - Simuler les crises sanitaires
        
        **Modèle :** ARIMA/SARIMA  
        **Période :** 2011-2016
        """)
    
    return {
        "mode": mode_choice,
        "col_valeur": col_valeur,
        "couleur_mode": couleur_mode,
        "year": year_choice,
        "years": years,
        "filtres": filtre_indicateur
    }


def render_kpi_section(df: pd.DataFrame, col_valeur: str, mode_choice: str):
    """Affiche la section des KPI principaux"""
    st.markdown("""
    <div class="section-header">
        <h2>📈 Indicateurs Clés de Performance (KPI)</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Calcul des KPI
    total_patients = df[df["UNITE"] == "patients"][col_valeur].sum()
    total_sejours = df[df["UNITE"] == "sejours"][col_valeur].sum()
    passages_urgences = df[df["INDICATEUR"] == "Urgences"][col_valeur].sum()
    
    # Calcul variations
    if mode_choice == "Crise Sanitaire":
        variation_patients = calcul_variation(
            df[df["UNITE"] == "patients"]["TOTAL_NORMAL"].sum(),
            total_patients
        )
        variation_sejours = calcul_variation(
            df[df["UNITE"] == "sejours"]["TOTAL_NORMAL"].sum(),
            total_sejours
        )
    else:
        variation_patients = None
        variation_sejours = None
    
    # Affichage métriques
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="👥 Total Patients",
            value=format_nombre(total_patients),
            delta=f"{variation_patients:.1f}%" if variation_patients else None,
            delta_color="inverse"
        )
    
    with col2:
        st.metric(
            label="🛏️ Séjours Hospitaliers",
            value=format_nombre(total_sejours),
            delta=f"{variation_sejours:.1f}%" if variation_sejours else None,
            delta_color="inverse"
        )
    
    with col3:
        st.metric(
            label="🚑 Passages Urgences",
            value=format_nombre(passages_urgences)
        )
    
    with col4:
        age_moyen = df[df["SOUS-INDICATEUR"].str.contains("Âge moyen", case=False, na=False)][col_valeur].mean()
        st.metric(
            label="⏱️ Âge Moyen Patients",
            value=f"{age_moyen:.0f} ans" if not pd.isna(age_moyen) else "N/A"
        )
    
    # Alerte contextuelle
    if mode_choice == "Crise Sanitaire":
        st.markdown("""
        <div class="alert-box danger">
            <strong>⚠️ SIMULATION CRISE SANITAIRE ACTIVÉE</strong><br>
            Impact estimé : Augmentation moyenne de <strong>80%</strong> de l'activité hospitalière.
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-box info">
            <strong>ℹ️ MODE NORMAL</strong><br>
            Visualisation de l'activité hospitalière en conditions standards.
        </div>
        """, unsafe_allow_html=True)


def render_graphique_evolution(df: pd.DataFrame, indic: str, unite: str, 
                               col_valeur: str, year_choice: str, years: List[int]):
    """Génère un graphique d'évolution pour un indicateur"""
    
    if year_choice == "Toutes":
        # Vue multi-années
        agg = df.groupby(["SOUS-INDICATEUR", "ANNEE"])[[col_valeur]].sum().reset_index()
        
        if agg.empty:
            return None
        
        fig = go.Figure()
        sous_indicateurs = agg["SOUS-INDICATEUR"].unique()
        
        for idx, sous_indic in enumerate(sous_indicateurs):
            df_temp = agg[agg["SOUS-INDICATEUR"] == sous_indic]
            
            fig.add_trace(go.Bar(
                name=sous_indic,
                x=df_temp["ANNEE"].astype(str),
                y=df_temp[col_valeur],
                marker_color=PALETTE_ANNEES[idx % len(PALETTE_ANNEES)],
                text=df_temp[col_valeur].apply(lambda x: format_nombre(x)),
                textposition='outside',
                hovertemplate='<b>%{fullData.name}</b><br>' +
                              'Année: %{x}<br>' +
                              f'Valeur: %{{y:,.0f}} {unite}<br>' +
                              '<extra></extra>'
            ))
        
        fig.update_layout(
            title={
                'text': f"{indic} ({unite}) - Évolution temporelle",
                'font': {'size': 18, 'color': PALETTE_APHP["bleu_fonce"], 'family': 'Inter'}
            },
            xaxis_title="Année",
            yaxis_title=f"Volume ({unite})",
            barmode='group',
            template='plotly_white',
            height=500,
            hovermode='x unified',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            font=dict(family='Inter', size=12)
        )
        
    else:
        # Vue année unique
        agg = df.groupby("SOUS-INDICATEUR")[[col_valeur]].sum().reset_index()
        
        if agg.empty:
            return None
        
        year_idx = years.index(int(year_choice)) if int(year_choice) in years else 0
        color_year = PALETTE_ANNEES[year_idx % len(PALETTE_ANNEES)]
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=agg["SOUS-INDICATEUR"],
            y=agg[col_valeur],
            marker_color=color_year,
            text=agg[col_valeur].apply(lambda x: format_nombre(x)),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>' +
                          f'Valeur: %{{y:,.0f}} {unite}<br>' +
                          '<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': f"{indic} ({unite}) - Année {year_choice}",
                'font': {'size': 18, 'color': PALETTE_APHP["bleu_fonce"], 'family': 'Inter'}
            },
            xaxis_title="Sous-indicateur",
            yaxis_title=f"Volume ({unite})",
            template='plotly_white',
            height=500,
            font=dict(family='Inter', size=12)
        )
    
    return fig


def render_tableau_comparaison(df: pd.DataFrame):
    """Affiche le tableau de comparaison Normal/Crise"""
    df_compare = df.copy()
    df_compare["CHANGEMENT"] = df_compare["TOTAL_CRISE"] - df_compare["TOTAL_NORMAL"]
    df_compare["VARIATION_%"] = df_compare.apply(
        lambda row: calcul_variation(row["TOTAL_NORMAL"], row["TOTAL_CRISE"]), 
        axis=1
    )
    df_compare["TENDANCE"] = df_compare["VARIATION_%"].apply(get_tendance_icon)
    
    colonnes = [
        "ANNEE", "SOUS-INDICATEUR", "UNITE",
        "TOTAL_NORMAL", "TOTAL_CRISE", 
        "CHANGEMENT", "VARIATION_%", "TENDANCE"
    ]
    
    df_display = df_compare[colonnes].rename(columns={
        "ANNEE": "Année",
        "SOUS-INDICATEUR": "Sous-indicateur",
        "UNITE": "Unité",
        "TOTAL_NORMAL": "Valeur Normale",
        "TOTAL_CRISE": "Valeur Crise",
        "CHANGEMENT": "Écart Absolu",
        "VARIATION_%": "Variation (%)",
        "TENDANCE": "Tendance"
    })
    
    df_display["Valeur Normale"] = df_display["Valeur Normale"].apply(format_nombre)
    df_display["Valeur Crise"] = df_display["Valeur Crise"].apply(format_nombre)
    df_display["Écart Absolu"] = df_display["Écart Absolu"].apply(format_nombre)
    df_display["Variation (%)"] = df_display["Variation (%)"].apply(lambda x: f"{x:.1f}%")
    
    st.dataframe(
        df_display.sort_values(["Unité", "Sous-indicateur", "Année"]),
        use_container_width=True,
        height=400
    )


def render_analyse_indicateurs(df: pd.DataFrame, col_valeur: str, year_choice: str, years: List[int]):
    """Affiche l'analyse détaillée par indicateur avec onglets"""
    st.markdown("""
    <div class="section-header">
        <h2>🔬 Analyse Détaillée par Indicateur</h2>
    </div>
    """, unsafe_allow_html=True)
    
    indicateurs = sorted(df["INDICATEUR"].dropna().unique().tolist())
    tabs = st.tabs(indicateurs)
    
    for tab, indic in zip(tabs, indicateurs):
        with tab:
            st.subheader(f"📊 {indic}")
            
            df_indic = df[df["INDICATEUR"] == indic]
            unites = sorted(df_indic["UNITE"].dropna().unique().tolist())
            
            for unite in unites:
                df_unite = df_indic[df_indic["UNITE"] == unite]
                
                # Générer graphique
                fig = render_graphique_evolution(df_unite, indic, unite, col_valeur, year_choice, years)
                
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
            
            # Tableau détaillé
            with st.expander("📋 Voir le tableau détaillé avec comparaisons"):
                render_tableau_comparaison(df_indic)


def render_comparaison_globale(df: pd.DataFrame):
    """Affiche la comparaison globale Normal vs Crise"""
    st.markdown("""
    <div class="section-header">
        <h2>⚖️ Comparaison Situation Normale vs Crise Sanitaire</h2>
    </div>
    """, unsafe_allow_html=True)
    
    df_comparison = df.groupby("INDICATEUR")[["TOTAL_NORMAL", "TOTAL_CRISE"]].sum().reset_index()
    df_comparison["ECART"] = df_comparison["TOTAL_CRISE"] - df_comparison["TOTAL_NORMAL"]
    df_comparison["VARIATION_%"] = df_comparison.apply(
        lambda row: calcul_variation(row["TOTAL_NORMAL"], row["TOTAL_CRISE"]), 
        axis=1
    )
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='Situation Normale',
        x=df_comparison["INDICATEUR"],
        y=df_comparison["TOTAL_NORMAL"],
        marker_color=PALETTE_APHP["vert_normal"],
        text=df_comparison["TOTAL_NORMAL"].apply(lambda x: format_nombre(x)),
        textposition='outside'
    ))
    
    fig.add_trace(go.Bar(
        name='Crise Sanitaire (+80%)',
        x=df_comparison["INDICATEUR"],
        y=df_comparison["TOTAL_CRISE"],
        marker_color=PALETTE_APHP["rouge_alerte"],
        text=df_comparison["TOTAL_CRISE"].apply(lambda x: format_nombre(x)),
        textposition='outside'
    ))
    
    fig.update_layout(
        title={
            'text': "Comparaison Globale : Activité Normale vs Crise Sanitaire",
            'font': {'size': 20, 'color': PALETTE_APHP["bleu_fonce"], 'family': 'Inter'}
        },
        xaxis_title="Indicateurs Hospitaliers",
        yaxis_title="Volume Total",
        barmode='group',
        template='plotly_white',
        height=550,
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family='Inter', size=13)
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Synthèse variations
    st.markdown("### 📊 Synthèse des Variations par Indicateur")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Top 3 Hausses**")
        top_hausses = df_comparison.nlargest(3, "ECART")[["INDICATEUR", "VARIATION_%"]]
        for idx, row in top_hausses.iterrows():
            st.markdown(f"- **{row['INDICATEUR']}** : +{row['VARIATION_%']:.1f}%")
    
    with col2:
        st.markdown("**Indicateurs les Plus Impactés**")
        top_impact = df_comparison.nlargest(3, "TOTAL_CRISE")[["INDICATEUR", "TOTAL_CRISE"]]
        for idx, row in top_impact.iterrows():
            st.markdown(f"- **{row['INDICATEUR']}** : {format_nombre(row['TOTAL_CRISE'])} unités")


def render_origine_geographique(df: pd.DataFrame, col_valeur: str, mode_choice: str):
    """Affiche l'analyse géographique des patients"""
    st.markdown("""
    <div class="section-header">
        <h2>🗺️ Origine Géographique des Patients</h2>
    </div>
    """, unsafe_allow_html=True)
    
    df_geo = df[df["INDICATEUR"] == "Origine géographique"]
    
    if not df_geo.empty:
        geo_agg = df_geo.groupby("SOUS-INDICATEUR")[[col_valeur]].sum().reset_index()
        geo_agg = geo_agg.sort_values(col_valeur, ascending=False).head(10)
        
        fig = go.Figure(data=[go.Pie(
            labels=geo_agg["SOUS-INDICATEUR"],
            values=geo_agg[col_valeur],
            hole=0.4,
            marker=dict(colors=PALETTE_ANNEES),
            textinfo='label+percent',
            textposition='outside',
            hovertemplate='<b>%{label}</b><br>Valeur: %{value:.1f}%<br><extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': f"Répartition Géographique des Patients - {mode_choice}",
                'font': {'size': 18, 'color': PALETTE_APHP["bleu_fonce"], 'family': 'Inter'}
            },
            height=500,
            font=dict(family='Inter', size=12),
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucune donnée géographique disponible pour cette sélection.")


def render_profil_patients(df: pd.DataFrame, col_valeur: str):
    """Affiche le profil démographique des patients"""
    st.markdown("""
    <div class="section-header">
        <h2>👤 Profil Démographique des Patients</h2>
    </div>
    """, unsafe_allow_html=True)
    
    df_profil = df[df["INDICATEUR"] == "Profil patients"]
    
    if not df_profil.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            # Répartition Hommes/Femmes
            sexe_data = df_profil[df_profil["SOUS-INDICATEUR"].str.contains("Hommes|Femmes", case=False, na=False)]
            sexe_agg = sexe_data.groupby("SOUS-INDICATEUR")[[col_valeur]].sum().reset_index()
            
            if not sexe_agg.empty:
                fig_sexe = go.Figure(data=[go.Pie(
                    labels=sexe_agg["SOUS-INDICATEUR"],
                    values=sexe_agg[col_valeur],
                    marker=dict(colors=[PALETTE_APHP["bleu_hospital"], PALETTE_APHP["rouge_alerte"]]),
                    textinfo='label+percent',
                    hole=0.3
                )])
                
                fig_sexe.update_layout(
                    title="Répartition Hommes/Femmes",
                    height=400,
                    font=dict(family='Inter', size=12)
                )
                
                st.plotly_chart(fig_sexe, use_container_width=True)
        
        with col2:
            # Âge moyen
            age_data = df_profil[df_profil["SOUS-INDICATEUR"].str.contains("Âge moyen", case=False, na=False)]
            age_agg = age_data.groupby("SOUS-INDICATEUR")[[col_valeur]].mean().reset_index()
            
            if not age_agg.empty:
                fig_age = go.Figure(data=[go.Bar(
                    x=age_agg["SOUS-INDICATEUR"],
                    y=age_agg[col_valeur],
                    marker_color=PALETTE_APHP["bleu_clair"],
                    text=age_agg[col_valeur].apply(lambda x: f"{x:.0f} ans"),
                    textposition='outside'
                )])
                
                fig_age.update_layout(
                    title="Âge Moyen des Patients",
                    xaxis_title="",
                    yaxis_title="Âge (années)",
                    height=400,
                    font=dict(family='Inter', size=12),
                    template='plotly_white'
                )
                
                st.plotly_chart(fig_age, use_container_width=True)
    else:
        st.info("Aucune donnée de profil disponible pour cette sélection.")


def render_urgences(df: pd.DataFrame, col_valeur: str, mode_choice: str):
    """Affiche l'activité des urgences"""
    st.markdown("""
    <div class="section-header">
        <h2>🚑 Activité des Urgences</h2>
    </div>
    """, unsafe_allow_html=True)
    
    df_urgences = df[df["INDICATEUR"] == "Urgences"]
    
    if not df_urgences.empty:
        urgences_agg = df_urgences.groupby("SOUS-INDICATEUR")[[col_valeur]].sum().reset_index()
        
        fig = go.Figure(data=[go.Bar(
            x=urgences_agg["SOUS-INDICATEUR"],
            y=urgences_agg[col_valeur],
            marker_color=PALETTE_APHP["rouge_alerte"],
            text=urgences_agg[col_valeur].apply(lambda x: format_nombre(x)),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Volume: %{y:,.0f} patients<br><extra></extra>'
        )])
        
        fig.update_layout(
            title={
                'text': f"Activité des Urgences - {mode_choice}",
                'font': {'size': 18, 'color': PALETTE_APHP["bleu_fonce"], 'family': 'Inter'}
            },
            xaxis_title="Type d'Activité",
            yaxis_title="Nombre de Patients",
            height=450,
            template='plotly_white',
            font=dict(family='Inter', size=12)
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistiques urgences
        col1, col2, col3 = st.columns(3)
        
        passages = df_urgences[df_urgences["SOUS-INDICATEUR"] == "Passages"][col_valeur].sum()
        admis = df_urgences[df_urgences["SOUS-INDICATEUR"] == "Patients admis"][col_valeur].sum()
        uhcd = df_urgences[df_urgences["SOUS-INDICATEUR"] == "UHCD"][col_valeur].sum()
        
        with col1:
            st.metric("Passages Totaux", format_nombre(passages))
        
        with col2:
            st.metric("Patients Admis", format_nombre(admis))
            if passages > 0:
                taux_admission = (admis / passages) * 100
                st.markdown(f"**Taux d'admission :** {taux_admission:.1f}%")
        
        with col3:
            st.metric("UHCD", format_nombre(uhcd))
    else:
        st.info("Aucune donnée d'urgences disponible.")


def render_recommandations(mode_choice: str):
    """Affiche les recommandations et insights"""
    st.markdown("""
    <div class="section-header">
        <h2>💡 Recommandations et Insights</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🎯 Gestion Normale")
        st.markdown("""
        - Optimisation des plannings selon saisonnalité
        - Capacité tampon de 15-20% pour pics d'activité
        - Révision trimestrielle des stocks
        - Formation continue du personnel
        """)
    
    with col2:
        st.markdown("### 🚨 Préparation Crises")
        st.markdown("""
        - Plan de continuité avec +80% de capacité
        - Constitution d'une réserve sanitaire
        - Stocks stratégiques (3 mois minimum)
        - Simulations régulières
        """)
    
    # Alertes
    if mode_choice == "Crise Sanitaire":
        st.markdown("""
        <div class="alert-box danger">
            <strong>🔴 ALERTE CRITIQUE :</strong> Capacité d'accueil atteindrait 95%.
            Activation du plan blanc recommandée.
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-box info">
            <strong>🟢 SITUATION NORMALE :</strong> Capacités suffisantes.
            Surveillance continue recommandée.
        </div>
        """, unsafe_allow_html=True)


def render_export(df: pd.DataFrame, year_choice: str, mode_choice: str):
    """Affiche la section d'export des données"""
    st.markdown("""
    <div class="section-header">
        <h2>📥 Export des Données</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        csv = df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📊 Télécharger les données (CSV)",
            data=csv,
            file_name=f"psl_cfx_data_{year_choice}_{mode_choice.lower().replace(' ', '_')}.csv",
            mime="text/csv"
        )
    
    with col2:
        df_comparison = df.groupby("INDICATEUR")[["TOTAL_NORMAL", "TOTAL_CRISE"]].sum().reset_index()
        resume = df_comparison.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📈 Télécharger le résumé",
            data=resume,
            file_name=f"psl_cfx_resume_{year_choice}.csv",
            mime="text/csv"
        )
    
    with col3:
        st.info("📄 Rapports techniques disponibles séparément")


def render_footer():
    """Affiche le footer du dashboard"""
    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p><strong>Hôpitaux Universitaires Pitié Salpêtrière Charles Foix (PSL-CFX)</strong></p>
        <p>Dashboard développé dans le cadre du Projet Data - EPITECH Digital School</p>
        <p>Données : 2011-2016 | Modèle : ARIMA/SARIMA | Simulation : +80% (impact COVID-19)</p>
        <p style="font-size: 0.9rem; color: #999; margin-top: 1rem;">
            ⚕️ Outil d'aide à la décision pour l'optimisation de la gestion hospitalière
        </p>
    </div>
    """, unsafe_allow_html=True)


# ============================================================================
# FONCTION PRINCIPALE RENDER
# ============================================================================

def render(st_module, *, data_path: str, **kwargs):
    """
    Fonction principale de rendu du dashboard patients
    
    Args:
        st_module: Module streamlit
        data_path: Chemin vers le fichier CSV des données
        **kwargs: Arguments supplémentaires (compatibilité)
    """
    # Configuration de la page
    st_module.set_page_config(
        page_title="Dashboard PSL-CFX | Patients",
        page_icon="🏥",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Appliquer le CSS
    apply_custom_css()
    
    # Charger les données
    if not data_path:
        st_module.info("Aucune donnée configurée pour cette page.")
        return
    
    df = load_data(data_path)
    if df is None:
        return
    
    # Afficher l'en-tête
    render_header()
    
    # Afficher la sidebar et récupérer les paramètres
    params = render_sidebar(df)
    
    # Filtrer les données
    dff = df.copy()
    if params["year"] != "Toutes":
        dff = dff[dff["ANNEE"] == int(params["year"])]
    
    if params["filtres"]:
        dff = dff[dff["INDICATEUR"].isin(params["filtres"])]
    
    # Sections du dashboard
    render_kpi_section(dff, params["col_valeur"], params["mode"])
    render_analyse_indicateurs(dff, params["col_valeur"], params["year"], params["years"])
    render_comparaison_globale(dff)
    render_origine_geographique(dff, params["col_valeur"], params["mode"])
    render_profil_patients(dff, params["col_valeur"])
    render_urgences(dff, params["col_valeur"], params["mode"])
    render_recommandations(params["mode"])
    render_export(dff, params["year"], params["mode"])
    render_footer()


