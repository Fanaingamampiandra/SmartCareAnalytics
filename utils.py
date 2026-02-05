# utils.py — fonctions partagées (chargement de données, etc.)
import pandas as pd
import streamlit as st
import numpy as np
from statsmodels.tsa.statespace.sarimax import SARIMAX


@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    if "ANNEE" in df.columns:
        df["ANNEE"] = df["ANNEE"].astype(int)
    return df


@st.cache_data
def generate_forecast_2017(
    df: pd.DataFrame,
    site_code: str,
    indicateur: str,
    sous_indicateur: str,
    value_col: str = "value",
) -> pd.DataFrame:
    """
    Génère des prédictions pour l'année 2017 en utilisant SARIMAX.
    
    Args:
        df: DataFrame avec les données historiques
        site_code: Code du site (PLF, CFX, ou TOTAL)
        indicateur: Nom de l'indicateur
        sous_indicateur: Nom du sous-indicateur
        value_col: Nom de la colonne de valeur
        
    Returns:
        DataFrame avec les prédictions pour 2017 (date, year, month, value)
    """
    try:
        # Filtrer les données historiques (jusqu'à 2016)
        df_filtered = df[
            (df["indicateur"] == indicateur) &
            (df["sous_indicateur"] == sous_indicateur)
        ].copy()
        
        # Gérer le filtre site_code
        if site_code in ("PLF", "CFX"):
            df_filtered = df_filtered[df_filtered["site_code"] == site_code]
        elif site_code == "TOTAL":
            # Pour TOTAL, on agrège PLF + CFX par date
            if "date" in df_filtered.columns:
                df_filtered = (
                    df_filtered.groupby("date")[value_col]
                    .sum()
                    .reset_index()
                )
                df_filtered["site_code"] = "TOTAL"
            else:
                return pd.DataFrame()
        
        # S'assurer qu'on a la colonne date
        if "date" not in df_filtered.columns:
            return pd.DataFrame()
        
        df_filtered["date"] = pd.to_datetime(df_filtered["date"])
        
        # Filtrer jusqu'à 2016 inclus
        year_col = "ANNEE" if "ANNEE" in df_filtered.columns else "year"
        if year_col in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[year_col] <= 2016]
        else:
            df_filtered = df_filtered[df_filtered["date"].dt.year <= 2016]
        
        if len(df_filtered) < 100:  # Pas assez de données pour entraîner
            return pd.DataFrame()
        
        # Créer une série temporelle journalière
        df_filtered = df_filtered.sort_values("date")
        ts = df_filtered.set_index("date")[value_col].asfreq("D")
        
        # Remplir les valeurs manquantes (forward fill)
        ts = ts.fillna(method="ffill").fillna(method="bfill")
        
        if len(ts) < 100 or ts.isna().all():
            return pd.DataFrame()
        
        # Entraîner le modèle SARIMAX
        model = SARIMAX(
            ts,
            order=(1, 1, 1),  # tendance
            seasonal_order=(1, 1, 1, 7)  # saison hebdomadaire
        )
        
        results = model.fit(disp=False)
        
        # Générer les prédictions pour 2017 (365 jours)
        forecast_2017 = results.get_forecast(steps=365)
        pred_2017 = forecast_2017.predicted_mean
        
        # Créer un DataFrame avec les prédictions
        dates_2017 = pd.date_range("2017-01-01", "2017-12-31", freq="D")
        df_forecast = pd.DataFrame({
            "date": dates_2017,
            "year": 2017,
            "month": dates_2017.month,
            "value": pred_2017.values,
            "indicateur": indicateur,
            "sous_indicateur": sous_indicateur,
            "site_code": site_code if site_code != "TOTAL" else "PLF",  # Pour compatibilité
        })
        
        return df_forecast
        
    except Exception as e:
        # En cas d'erreur, retourner un DataFrame vide
        return pd.DataFrame()
