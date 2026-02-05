"""
Script de generation du Rapport de Mise en Place -- PSL-CFX
Version synthetique avec donnees mensuelles
Sections : Logistique + Patients
"""

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


# --- Configuration ---
DATA_PATH_LOGISTICS = "data/logistics/logistics-donnees_mensuelles_reconstituees.csv"
DATA_PATH_PATIENTS = "data/patients/patients-donnees_mensuelles_reconstituees.csv"
OUTPUT_FILE = "Rapport_Mise_En_Place_PSL-CFX_v2.docx"
CHARTS_DIR = "charts_temp"

# Couleurs
BLEU = RGBColor(31, 119, 180)
ROUGE = RGBColor(214, 39, 40)
VERT = RGBColor(44, 160, 44)
GRIS = RGBColor(100, 100, 100)
BLEU_FONCE = RGBColor(0, 51, 102)
BLANC = RGBColor(255, 255, 255)

MOIS_LABELS = ["Jan", "Fev", "Mar", "Avr", "Mai", "Jun", "Jul", "Aou", "Sep", "Oct", "Nov", "Dec"]


# --- Fonctions utilitaires ---

def setup_charts_dir():
    os.makedirs(CHARTS_DIR, exist_ok=True)


def save_chart(fig, name):
    path = os.path.join(CHARTS_DIR, f"{name}.png")
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def set_cell_shading(cell, color_hex):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def add_styled_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = BLANC
        set_cell_shading(cell, "003366")

    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
            if r_idx % 2 == 1:
                set_cell_shading(cell, "EBF5FB")

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    return table


def add_heading_styled(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = BLEU_FONCE
    return h


def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.size = Pt(10)
        p.add_run(text)
    else:
        run = p.add_run(text)
        run.font.size = Pt(10)
    return p


def add_body(doc, text):
    p = doc.add_paragraph(text)
    for run in p.runs:
        run.font.size = Pt(10)
    return p


def add_constat_box(doc, constat_text):
    p = doc.add_paragraph()
    run = p.add_run("Constat : ")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = ROUGE
    run2 = p.add_run(constat_text)
    run2.font.size = Pt(10)
    return p


# --- Chargement des donnees ---

def load_monthly_data(path):
    """Charge un CSV mensuel."""
    df = pd.read_csv(path)
    df["year"] = df["year"].astype(int)
    df["month"] = df["month"].astype(int)
    return df


def agg_annual(df):
    """Agregation annuelle tous sites confondus."""
    agg = (
        df.groupby(["year", "indicateur", "sous_indicateur", "unite"])[["value", "value_crise"]]
        .sum()
        .reset_index()
    )
    agg["ecart"] = agg["value_crise"] - agg["value"]
    return agg


def agg_annual_by_site(df):
    """Agregation annuelle par site."""
    return (
        df.groupby(["year", "site_code", "indicateur", "sous_indicateur", "unite"])[["value", "value_crise"]]
        .sum()
        .reset_index()
    )


def agg_monthly_avg(df):
    """Moyenne mensuelle (profil saisonnier)."""
    return (
        df.groupby(["month", "indicateur"])[["value", "value_crise"]]
        .mean()
        .reset_index()
    )


# ============================================================
# GRAPHIQUES LOGISTIQUE
# ============================================================

def gen_chart_synthese_logistique(annual):
    """Barres horizontales : 6 indicateurs logistiques (2015)."""
    df_2015 = annual[annual["year"] == 2015]
    synthese = df_2015.groupby("indicateur").agg({"value": "sum", "value_crise": "sum"}).reset_index()
    synthese = synthese.sort_values("value", ascending=True)

    fig, ax = plt.subplots(figsize=(10, 4))
    y = np.arange(len(synthese))
    h = 0.35
    ax.barh(y - h/2, synthese["value"], h, label="Normal", color="#1f77b4", edgecolor="white")
    ax.barh(y + h/2, synthese["value_crise"], h, label="Crise", color="#d62728", edgecolor="white")
    ax.set_yticks(y)
    ax.set_yticklabels(synthese["indicateur"], fontsize=10)
    ax.set_title("Logistique -- Volumes par indicateur (2015)", fontsize=12, fontweight="bold")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend(fontsize=9)
    ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "logistique_synthese")


def gen_chart_saisonnalite_logistique(monthly_avg):
    """Profil mensuel moyen (saisonnalite)."""
    # Moyenne globale tous indicateurs
    agg = monthly_avg.groupby("month")[["value"]].mean().reset_index()

    fig, ax = plt.subplots(figsize=(8, 3.5))
    ax.bar(agg["month"], agg["value"], color="#1f77b4", edgecolor="white")
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels(MOIS_LABELS, fontsize=9)
    ax.set_ylabel("Volume moyen")
    ax.set_title("Profil saisonnier -- Logistique (moyenne mensuelle)", fontsize=12, fontweight="bold")
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "logistique_saisonnalite")


# ============================================================
# GRAPHIQUES PATIENTS
# ============================================================

def gen_chart_urgences(annual):
    """Barres : Passages et Admissions aux urgences (2015)."""
    df_urg = annual[(annual["indicateur"] == "Urgences") & (annual["year"] == 2015)]

    passages = df_urg[df_urg["sous_indicateur"] == "Passages"]
    admis = df_urg[df_urg["sous_indicateur"] == "Patients admis"]

    fig, ax = plt.subplots(figsize=(8, 4))
    x = np.arange(2)
    w = 0.35

    vals_n = [passages["value"].sum() if not passages.empty else 0,
              admis["value"].sum() if not admis.empty else 0]
    vals_c = [passages["value_crise"].sum() if not passages.empty else 0,
              admis["value_crise"].sum() if not admis.empty else 0]

    ax.bar(x - w/2, vals_n, w, label="Normal", color="#1f77b4", edgecolor="white")
    ax.bar(x + w/2, vals_c, w, label="Crise", color="#d62728", edgecolor="white")
    ax.set_xticks(x)
    ax.set_xticklabels(["Passages", "Patients admis"], fontsize=10)
    ax.set_title("Urgences -- Normal vs Crise (2015)", fontsize=12, fontweight="bold")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend(fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "patients_urgences")


def gen_chart_profil_patients(annual):
    """Repartition Hommes/Femmes et age moyen (2015)."""
    df_prof = annual[(annual["indicateur"] == "Profil patients") & (annual["year"] == 2015)]

    hommes = df_prof[df_prof["sous_indicateur"] == "Hommes - MCO"]["value"].sum()
    femmes = df_prof[df_prof["sous_indicateur"] == "Femmes - MCO"]["value"].sum()
    age_h = df_prof[df_prof["sous_indicateur"] == "Âge moyen - Hommes - MCO"]["value"].mean()
    age_f = df_prof[df_prof["sous_indicateur"] == "Âge moyen - Femmes - MCO"]["value"].mean()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))

    # Camembert H/F
    if hommes > 0 or femmes > 0:
        ax1.pie([hommes, femmes], labels=["Hommes", "Femmes"], autopct="%1.1f%%",
                colors=["#1f77b4", "#ff7f0e"], startangle=90)
        ax1.set_title("Repartition H/F (MCO)", fontsize=11, fontweight="bold")

    # Age moyen
    if not np.isnan(age_h) and not np.isnan(age_f):
        ax2.bar(["Hommes", "Femmes"], [age_h, age_f], color=["#1f77b4", "#ff7f0e"], edgecolor="white")
        ax2.set_ylabel("Age moyen (ans)")
        ax2.set_title("Age moyen par sexe (MCO)", fontsize=11, fontweight="bold")
        ax2.grid(axis="y", alpha=0.3)

    plt.tight_layout()
    return save_chart(fig, "patients_profil")


def gen_chart_origine_geo(annual):
    """Top 5 origines geographiques (2015)."""
    df_geo = annual[(annual["indicateur"] == "Origine géographique") & (annual["year"] == 2015)]

    # Filtrer les sous-indicateurs MCO principaux (pas SSR)
    df_mco = df_geo[df_geo["sous_indicateur"].str.contains("MCO", na=False)]
    top5 = df_mco.nlargest(5, "value")

    if top5.empty:
        return None

    fig, ax = plt.subplots(figsize=(9, 4))
    labels = [s.replace(" - MCO", "").replace("Île-de-France - ", "IDF ") for s in top5["sous_indicateur"]]
    y = np.arange(len(labels))
    ax.barh(y, top5["value"], color="#1f77b4", edgecolor="white")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel("Pourcentage (%)")
    ax.set_title("Top 5 origines geographiques (MCO, 2015)", fontsize=12, fontweight="bold")
    ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "patients_origine")


def gen_chart_pathologies(annual):
    """Evolution pathologies cancereuses."""
    df_patho = annual[annual["indicateur"] == "Causes d'hopitalisations"]
    df_total = df_patho[df_patho["sous_indicateur"] == "Pathologies cancéreuses - Total"]

    if df_total.empty:
        return None

    df_total = df_total.sort_values("year")

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(df_total["year"], df_total["value"], "o-", color="#1f77b4", linewidth=2, label="Normal")
    ax.plot(df_total["year"], df_total["value_crise"], "s--", color="#d62728", linewidth=2, label="Crise")
    ax.fill_between(df_total["year"], df_total["value"], df_total["value_crise"], alpha=0.15, color="red")
    ax.set_xlabel("Annee")
    ax.set_ylabel("Nombre de patients")
    ax.set_title("Pathologies cancereuses -- Evolution", fontsize=12, fontweight="bold")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend(fontsize=9)
    ax.grid(alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "patients_pathologies")


# ============================================================
# CONSTRUCTION DU DOCUMENT
# ============================================================

def build_document():
    setup_charts_dir()

    # Chargement
    df_log = load_monthly_data(DATA_PATH_LOGISTICS)
    df_pat = load_monthly_data(DATA_PATH_PATIENTS)

    annual_log = agg_annual(df_log)
    monthly_avg_log = agg_monthly_avg(df_log)
    annual_pat = agg_annual(df_pat)

    doc = Document()

    # Style
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(4)

    # ============================
    # PAGE DE TITRE
    # ============================
    for _ in range(6):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Rapport de Mise en Place")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = BLEU_FONCE

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Analyse des donnees hospitalieres\nPSL-CFX (Pitie-Salpetriere / Charles Foix)")
    run.font.size = Pt(16)
    run.font.color.rgb = GRIS

    doc.add_paragraph()
    obj = doc.add_paragraph()
    obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = obj.add_run("Propositions pour gerer les afflux de patients\net se preparer aux crises sanitaires")
    run.font.size = Pt(12)
    run.font.italic = True
    run.font.color.rgb = GRIS

    doc.add_paragraph()
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run("Fevrier 2026")
    run.font.size = Pt(12)
    run.font.color.rgb = GRIS

    doc.add_page_break()

    # ============================
    # SOMMAIRE
    # ============================
    add_heading_styled(doc, "Sommaire", level=1)
    doc.add_paragraph()

    sommaire_items = [
        ("1.", "Logistique hospitaliere", "Synthese des flux materiels"),
        ("2.", "Patients", "Urgences, profil et parcours de soins"),
        ("3.", "Activite et services", "(a venir)"),
        ("4.", "Capacite d'accueil", "(a venir)"),
        ("5.", "Finance", "(a venir)"),
        ("6.", "Qualite", "(a venir)"),
        ("7.", "Ressources humaines", "(a venir)"),
        ("8.", "Synthese et recommandations", ""),
    ]

    for num, titre, desc in sommaire_items:
        p = doc.add_paragraph()
        run_num = p.add_run(f"{num} ")
        run_num.bold = True
        run_num.font.size = Pt(12)
        run_num.font.color.rgb = BLEU_FONCE
        run_titre = p.add_run(titre)
        run_titre.bold = True
        run_titre.font.size = Pt(12)
        if desc:
            run_desc = p.add_run(f" -- {desc}")
            run_desc.font.size = Pt(10)
            run_desc.font.color.rgb = GRIS

    doc.add_page_break()

    # ============================
    # 1. LOGISTIQUE (SYNTHESE)
    # ============================
    add_heading_styled(doc, "1. Logistique hospitaliere", level=1)

    add_body(doc,
        "La logistique hospitaliere couvre les flux materiels essentiels : restauration, "
        "gestion du linge, approvisionnements, traitement des dechets, entretien et courrier. "
        "En crise sanitaire, ces postes subissent une augmentation brutale de la demande."
    )

    # Graphique synthese
    chart_path = gen_chart_synthese_logistique(annual_log)
    doc.add_picture(chart_path, width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tableau recapitulatif
    df_2015 = annual_log[annual_log["year"] == 2015]
    synthese = df_2015.groupby("indicateur").agg({"value": "sum", "value_crise": "sum"}).reset_index()
    synthese["ecart"] = synthese["value_crise"] - synthese["value"]

    add_body(doc, "Tableau recapitulatif (donnees 2015) :")

    headers = ["Indicateur", "Normal", "Crise", "Ecart"]
    rows = []
    for _, r in synthese.iterrows():
        rows.append([
            r["indicateur"],
            f"{r['value']:,.0f}",
            f"{r['value_crise']:,.0f}",
            f"{r['ecart']:+,.0f}"
        ])
    add_styled_table(doc, headers, rows)

    doc.add_paragraph()

    # Profil saisonnier
    add_body(doc, "Profil saisonnier moyen :")
    chart_path = gen_chart_saisonnalite_logistique(monthly_avg_log)
    doc.add_picture(chart_path, width=Inches(4.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Plan d'action
    add_heading_styled(doc, "Plan d'action prioritaire", level=2)

    headers_pa = ["Priorite", "Action", "Responsable"]
    rows_pa = [
        ["HAUTE", "Stock strategique EPI et consommables (30 jours)", "Dir. Achats"],
        ["HAUTE", "Conventions prestataires de secours (restauration, linge)", "Dir. Logistique"],
        ["HAUTE", "Protocole bionettoyage niveau crise", "Service Hygiene"],
        ["MOYENNE", "Systeme d'alerte reapprovisionnement", "Magasin"],
        ["BASSE", "Dematerialisation courrier administratif", "DSI"],
    ]
    add_styled_table(doc, headers_pa, rows_pa, col_widths=[2, 9, 4])

    doc.add_page_break()

    # ============================
    # 2. PATIENTS
    # ============================
    add_heading_styled(doc, "2. Patients", level=1)

    add_body(doc,
        "Cette section analyse le profil des patients, les urgences et les causes d'hospitalisation. "
        "En crise sanitaire, l'afflux aux urgences et les hospitalisations augmentent significativement."
    )

    # Urgences
    add_heading_styled(doc, "2.1 Urgences", level=2)
    chart_path = gen_chart_urgences(annual_pat)
    doc.add_picture(chart_path, width=Inches(4.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Chiffres urgences
    df_urg_2015 = annual_pat[(annual_pat["indicateur"] == "Urgences") & (annual_pat["year"] == 2015)]
    passages = df_urg_2015[df_urg_2015["sous_indicateur"] == "Passages"]
    if not passages.empty:
        p_n = passages["value"].sum()
        p_c = passages["value_crise"].sum()
        add_constat_box(doc,
            f"En 2015, {p_n:,.0f} passages aux urgences en situation normale, "
            f"contre {p_c:,.0f} en crise (+{(p_c-p_n)/p_n*100:.0f}%)."
        )

    # Profil patients
    add_heading_styled(doc, "2.2 Profil des patients", level=2)
    chart_path = gen_chart_profil_patients(annual_pat)
    doc.add_picture(chart_path, width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Origine geographique
    add_heading_styled(doc, "2.3 Origine geographique", level=2)
    chart_path = gen_chart_origine_geo(annual_pat)
    if chart_path:
        doc.add_picture(chart_path, width=Inches(5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Pathologies
    add_heading_styled(doc, "2.4 Causes d'hospitalisation", level=2)
    chart_path = gen_chart_pathologies(annual_pat)
    if chart_path:
        doc.add_picture(chart_path, width=Inches(4.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Propositions patients
    add_heading_styled(doc, "Propositions pour gerer l'afflux de patients", level=2)

    add_bullet(doc, " Renforcer les effectifs aux urgences (medecins, infirmiers) avec un plan de rappel.",
               bold_prefix="Capacite d'accueil --")
    add_bullet(doc, " Activer les lits de reserve et les unites de debordement.",
               bold_prefix="Lits supplementaires --")
    add_bullet(doc, " Orienter les patients stables vers la medecine de ville ou les centres de crise.",
               bold_prefix="Filiere de soins --")
    add_bullet(doc, " Mettre en place un guichet unique d'orientation des patients.",
               bold_prefix="Coordination --")

    doc.add_page_break()

    # ============================
    # PLACEHOLDERS
    # ============================
    placeholder_sections = [
        ("3. Activite et services", "Volumes d'activite et organisation des soins"),
        ("4. Capacite d'accueil", "Lits, equipements et plateaux techniques"),
        ("5. Finance", "Equilibre budgetaire et investissements"),
        ("6. Qualite", "Satisfaction et indicateurs qualite"),
        ("7. Ressources humaines", "Effectifs et organisation du personnel"),
    ]

    for titre, desc in placeholder_sections:
        add_heading_styled(doc, titre, level=1)
        p = doc.add_paragraph()
        run = p.add_run(f"{desc} -- Section a completer.")
        run.font.size = Pt(11)
        run.font.italic = True
        run.font.color.rgb = GRIS
        doc.add_paragraph()

    # ============================
    # SYNTHESE GENERALE
    # ============================
    doc.add_page_break()
    add_heading_styled(doc, "8. Synthese et recommandations", level=1)
    p = doc.add_paragraph()
    run = p.add_run("Cette section sera completee une fois l'ensemble des categories analysees.")
    run.font.italic = True
    run.font.color.rgb = GRIS

    # --- Sauvegarde ---
    doc.save(OUTPUT_FILE)
    print(f"Rapport genere avec succes : {OUTPUT_FILE}")

    # Nettoyage
    import shutil
    shutil.rmtree(CHARTS_DIR, ignore_errors=True)


if __name__ == "__main__":
    build_document()
