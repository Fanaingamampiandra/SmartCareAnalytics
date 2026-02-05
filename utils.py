# utils.py — fonctions partagées (chargement de données, etc.)
import pandas as pd
import streamlit as st


@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    if "ANNEE" in df.columns:
        df["ANNEE"] = df["ANNEE"].astype(int)
    return df
