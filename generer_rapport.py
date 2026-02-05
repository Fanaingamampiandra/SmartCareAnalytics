"""
Script de generation du Rapport de Mise en Place -- PSL-CFX
Genere un document Word (.docx) avec la section Logistique remplie
et des placeholders pour les autres categories.
"""

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os
import io

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT


# --- Configuration ---
DATA_PATH = "data/logistics/logistics-data-with-crise.csv"
OUTPUT_FILE = "Rapport_Mise_En_Place_PSL-CFX.docx"
CHARTS_DIR = "charts_temp"

# Couleurs
BLEU = RGBColor(31, 119, 180)
ROUGE = RGBColor(214, 39, 40)
VERT = RGBColor(44, 160, 44)
GRIS = RGBColor(100, 100, 100)
BLEU_FONCE = RGBColor(0, 51, 102)
BLANC = RGBColor(255, 255, 255)


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

    # En-tete
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

    # Lignes
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
        for run in p.runs:
            if not run.bold:
                run.font.size = Pt(10)
    else:
        run = p.add_run(text)
        run.font.size = Pt(10)
    return p


def add_sub_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet 2")
    run = p.add_run(text)
    run.font.size = Pt(9)
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


# ============================================================
# GENERATION DES GRAPHIQUES
# ============================================================

def gen_chart_restauration(df):
    df_repas = df[
        (df["INDICATEUR"] == "Restauration") &
        (df["SOUS-INDICATEUR"] == "Nombre de Repas")
    ].copy()

    fig, ax = plt.subplots(figsize=(10, 4.5))
    x = np.arange(len(df_repas))
    w = 0.35
    ax.bar(x - w / 2, df_repas["TOTAL_NORMAL"], w, label="Situation normale", color="#2ca02c", edgecolor="white")
    ax.bar(x + w / 2, df_repas["TOTAL_CRISE"], w, label="Crise sanitaire (+70%)", color="#d62728", edgecolor="white")
    ax.set_xlabel("Annee", fontsize=11)
    ax.set_ylabel("Nombre de repas", fontsize=11)
    ax.set_title("Nombre total de repas par an - Normal vs Crise sanitaire", fontsize=13, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(df_repas["ANNEE"].astype(int))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend(fontsize=10)
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "restauration")


def gen_chart_dechets(df):
    df_d = df[(df["INDICATEUR"] == "Dechets") | (df["INDICATEUR"] == "D\u00e9chets")]
    df_d = df_d[df_d["ANNEE"] == 2015].copy()

    labels_courts = []
    for s in df_d["SOUS-INDICATEUR"].values:
        if "Infectieux" in s:
            labels_courts.append("DASRI")
        elif "M\u00e9nagers Assimil" in s:
            labels_courts.append("D\u00e9chets m\u00e9nagers (DAE)")
        elif "assimil\u00e9s aux ordures" in s:
            labels_courts.append("Ordures m\u00e9nag\u00e8res")
        elif "lectriques" in s:
            labels_courts.append("DEEE")
        else:
            labels_courts.append(s)

    fig, ax = plt.subplots(figsize=(11, 5))
    y = np.arange(len(labels_courts))
    h = 0.35
    ax.barh(y - h / 2, df_d["TOTAL_NORMAL"].values, h, label="Normal", color="#1f77b4", edgecolor="white")
    ax.barh(y + h / 2, df_d["TOTAL_CRISE"].values, h, label="Crise (+70%)", color="#ff7f0e", edgecolor="white")
    ax.set_yticks(y)
    ax.set_yticklabels(labels_courts, fontsize=9)
    ax.set_xlabel("Tonnes", fontsize=11)
    ax.set_title("Volume de d\u00e9chets par type (2015) - Normal vs Crise", fontsize=13, fontweight="bold")
    ax.legend(fontsize=10)
    ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "dechets")


def gen_chart_hygiene(df):
    df_h = df[df["INDICATEUR"] == "Hygi\u00e8ne"].copy()
    locaux = df_h[df_h["SOUS-INDICATEUR"] == "Locaux"]
    espaces = df_h[df_h["SOUS-INDICATEUR"] == "Espaces verts"]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4.5))

    ax1.plot(locaux["ANNEE"], locaux["TOTAL_NORMAL"], "o-", color="#1f77b4", linewidth=2, label="Normal")
    ax1.plot(locaux["ANNEE"], locaux["TOTAL_CRISE"], "s--", color="#d62728", linewidth=2, label="Crise (+70%)")
    ax1.fill_between(locaux["ANNEE"], locaux["TOTAL_NORMAL"], locaux["TOTAL_CRISE"], alpha=0.15, color="red")
    ax1.set_title("Locaux (m\u00b2)", fontsize=12, fontweight="bold")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax1.legend(fontsize=9)
    ax1.grid(alpha=0.3)

    ax2.plot(espaces["ANNEE"], espaces["TOTAL_NORMAL"], "o-", color="#1f77b4", linewidth=2, label="Normal")
    ax2.plot(espaces["ANNEE"], espaces["TOTAL_CRISE"], "s--", color="#d62728", linewidth=2, label="Crise (+70%)")
    ax2.fill_between(espaces["ANNEE"], espaces["TOTAL_NORMAL"], espaces["TOTAL_CRISE"], alpha=0.15, color="red")
    ax2.set_title("Espaces verts (m\u00b2)", fontsize=12, fontweight="bold")
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax2.legend(fontsize=9)
    ax2.grid(alpha=0.3)

    fig.suptitle("Hygiene : surfaces a entretenir - Normal vs Crise", fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    return save_chart(fig, "hygiene")


def gen_chart_lingerie(df):
    df_l = df[
        (df["INDICATEUR"] == "Lingerie") &
        (df["SOUS-INDICATEUR"] == "Quantit\u00e9 de linge")
    ].copy()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4.5))

    ax1.fill_between(df_l["ANNEE"], df_l["TOTAL_NORMAL"], alpha=0.3, color="#1f77b4")
    ax1.fill_between(df_l["ANNEE"], df_l["TOTAL_NORMAL"], df_l["TOTAL_CRISE"], alpha=0.3, color="#d62728")
    ax1.plot(df_l["ANNEE"], df_l["TOTAL_NORMAL"], "o-", color="#1f77b4", linewidth=2, label="Normal")
    ax1.plot(df_l["ANNEE"], df_l["TOTAL_CRISE"], "s--", color="#d62728", linewidth=2, label="Crise (+70%)")
    ax1.set_title("Quantit\u00e9 de linge trait\u00e9 (kg/an)", fontsize=12, fontweight="bold")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax1.legend(fontsize=9)
    ax1.grid(alpha=0.3)

    # Repartition par site 2015
    l2015 = df_l[df_l["ANNEE"] == 2015].iloc[0]
    sites = ["Pitie-Salpetriere", "Charles Foix"]
    x = np.arange(2)
    ax2.bar(x - 0.2, [l2015["PLF_NORMAL"], l2015["CFX_NORMAL"]], 0.35, label="Normal", color="#1f77b4")
    ax2.bar(x + 0.2, [l2015["PLF_CRISE"], l2015["CFX_CRISE"]], 0.35, label="Crise", color="#d62728")
    ax2.set_xticks(x)
    ax2.set_xticklabels(sites)
    ax2.set_title("R\u00e9partition par site (2015, kg)", fontsize=12, fontweight="bold")
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax2.legend(fontsize=9)
    ax2.grid(axis="y", alpha=0.3)

    plt.tight_layout()
    return save_chart(fig, "lingerie")


def gen_chart_magasin(df):
    df_m = df[(df["INDICATEUR"] == "Magasin") & (df["ANNEE"] == 2015)].copy()

    fig, ax = plt.subplots(figsize=(8, 4.5))
    cats = ["Hygiene &\nEnvironnement", "Hotelier", "Imprimes"]
    x = np.arange(len(cats))
    w = 0.35
    ax.bar(x - w / 2, df_m["TOTAL_NORMAL"].values, w, label="Normal", color="#2ca02c", edgecolor="white")
    ax.bar(x + w / 2, df_m["TOTAL_CRISE"].values, w, label="Crise (+70%)", color="#d62728", edgecolor="white")
    ax.set_xticks(x)
    ax.set_xticklabels(cats, fontsize=10)
    ax.set_ylabel("Nombre de references")
    ax.set_title("References gerees par le magasin (2015)", fontsize=13, fontweight="bold")
    ax.legend(fontsize=10)
    ax.grid(axis="y", alpha=0.3)

    for i, (n, c) in enumerate(zip(df_m["TOTAL_NORMAL"].values, df_m["TOTAL_CRISE"].values)):
        ax.text(i - w / 2, n + 5, f"{n:.0f}", ha="center", va="bottom", fontsize=9)
        ax.text(i + w / 2, c + 5, f"{c:.0f}", ha="center", va="bottom", fontsize=9, color="#d62728")

    plt.tight_layout()
    return save_chart(fig, "magasin")


def gen_chart_vaguemestre(df):
    df_v = df[(df["INDICATEUR"] == "Vaguemestre") & (df["ANNEE"] == 2015)].copy()
    plis = df_v[df_v["SOUS-INDICATEUR"] == "Plis affranchis"].iloc[0]
    colis = df_v[df_v["SOUS-INDICATEUR"] == "Colis"].iloc[0]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))
    ax1.bar(["Normal", "Crise"], [plis["TOTAL_NORMAL"], plis["TOTAL_CRISE"]],
            color=["#1f77b4", "#d62728"], edgecolor="white")
    ax1.set_title("Plis affranchis/an (2015)", fontsize=11, fontweight="bold")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax1.grid(axis="y", alpha=0.3)

    ax2.bar(["Normal", "Crise"], [colis["TOTAL_NORMAL"], colis["TOTAL_CRISE"]],
            color=["#1f77b4", "#d62728"], edgecolor="white")
    ax2.set_title("Colis/an (2015)", fontsize=11, fontweight="bold")
    ax2.grid(axis="y", alpha=0.3)

    fig.suptitle("Service courrier - Normal vs Crise", fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    return save_chart(fig, "vaguemestre")


def gen_chart_synthese(df):
    df_2015 = df[df["ANNEE"] == 2015].copy()
    synthese = df_2015.groupby("INDICATEUR").agg({
        "TOTAL_NORMAL": "sum",
        "TOTAL_CRISE": "sum",
    }).reset_index()

    fig, ax = plt.subplots(figsize=(10, 5))
    y = np.arange(len(synthese))
    h = 0.35
    ax.barh(y - h / 2, synthese["TOTAL_NORMAL"], h, label="Normal", color="#1f77b4", edgecolor="white")
    ax.barh(y + h / 2, synthese["TOTAL_CRISE"], h, label="Crise (+70%)", color="#d62728", edgecolor="white")
    ax.set_yticks(y)
    ax.set_yticklabels(synthese["INDICATEUR"], fontsize=10)
    ax.set_title("Synthese Logistique - Tous indicateurs (2015)", fontsize=13, fontweight="bold")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend(fontsize=10)
    ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    return save_chart(fig, "synthese")


# ============================================================
# CONSTRUCTION DU DOCUMENT
# ============================================================

def build_document():
    setup_charts_dir()
    df = pd.read_csv(DATA_PATH)
    doc = Document()

    # -- Style par defaut --
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
        ("1.", "Logistique hospitaliere", "Gestion des flux materiels et des approvisionnements"),
        ("2.", "Activite et services", "Volumes d'activite et organisation des soins (a venir)"),
        ("3.", "Capacite d'accueil", "Lits, equipements et plateaux techniques (a venir)"),
        ("4.", "Finance", "Equilibre budgetaire et investissements (a venir)"),
        ("5.", "Patients", "Profil des patients et parcours de soins (a venir)"),
        ("6.", "Qualite", "Satisfaction et indicateurs qualite (a venir)"),
        ("7.", "Ressources humaines", "Effectifs et organisation du personnel (a venir)"),
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
    # 1. LOGISTIQUE
    # ============================
    add_heading_styled(doc, "1. Logistique hospitaliere", level=1)

    add_body(doc,
        "La logistique hospitaliere est un pilier essentiel du bon fonctionnement de l'hopital. "
        "Elle couvre l'ensemble des flux materiels necessaires aux soins et a la vie quotidienne "
        "de l'etablissement : restauration des patients et du personnel, gestion du linge, "
        "approvisionnement en fournitures, traitement des dechets, entretien des locaux et "
        "gestion du courrier."
    )

    add_body(doc,
        "Enjeu principal : En situation de crise sanitaire (epidemie, pandemie, afflux massif "
        "de patients), chacun de ces postes logistiques peut etre soumis a une augmentation "
        "brutale de la demande. L'hopital doit etre en mesure d'anticiper ces pics pour maintenir "
        "la qualite des soins et la securite des patients et du personnel."
    )

    add_body(doc,
        "Les donnees analysees couvrent les deux sites du groupe hospitalier :"
    )
    add_bullet(doc, " Hopital de reference, forte activite MCO (Medecine, Chirurgie, Obstetrique)", bold_prefix="Pitie-Salpetriere (PSL) --")
    add_bullet(doc, " Specialise en geriatrie et soins de suite", bold_prefix="Charles Foix (CFX) --")

    add_body(doc,
        "La simulation de crise sanitaire utilisee dans cette analyse repose sur une augmentation "
        "de +70% de l'activite sur l'ensemble des postes logistiques, conformement aux scenarios "
        "observes lors des pandemies recentes (COVID-19)."
    )

    # --- 1.1 RESTAURATION ---
    add_heading_styled(doc, "1.1 Restauration hospitaliere", level=2)

    add_body(doc,
        "La restauration est un service critique : chaque patient hospitalise doit recevoir "
        "ses repas quotidiens. En cas d'afflux de patients, le nombre de repas a preparer "
        "augmente proportionnellement. La cuisine centrale doit etre dimensionnee pour absorber "
        "ces variations."
    )

    chart_path = gen_chart_restauration(df)
    doc.add_picture(chart_path, width=Inches(5.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Chiffres cles
    r2015 = df[(df["INDICATEUR"] == "Restauration") & (df["SOUS-INDICATEUR"] == "Nombre de Repas") & (df["ANNEE"] == 2015)].iloc[0]
    rq = df[(df["INDICATEUR"] == "Restauration") & (df["SOUS-INDICATEUR"] == "Nombre de Repas Quotidien") & (df["ANNEE"] == 2015)].iloc[0]

    add_constat_box(doc,
        f"En situation de crise, le nombre de repas a fournir augmenterait de +70%, soit environ "
        f"{r2015['ECART_TOTAL']:,.0f} repas supplementaires par an (donnees 2015). Au quotidien, "
        f"cela represente le passage de {rq['TOTAL_NORMAL']:,.0f} repas/jour a environ "
        f"{rq['TOTAL_CRISE']:,.0f} repas/jour."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions pour gerer l'afflux :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Prevoir des conventions avec des prestataires de restauration externe "
               "(traiteurs agrees par l'ARS) activables sous 48h. Constituer un stock tampon "
               "de repas longue conservation pour couvrir les 72 premieres heures de crise.",
               bold_prefix="Protocole de montee en charge de la cuisine centrale --")

    add_bullet(doc, " Passer a un menu unique simplifie (au lieu de 2-3 choix) pour accelerer "
               "la production. Privilegier les plats en barquettes individuelles pour limiter "
               "les contaminations croisees.",
               bold_prefix="Adaptation des menus en situation de crise --")

    add_bullet(doc, " Identifier dans le plan de crise les agents mobilisables "
               "(personnel en back-office, stagiaires, prestataires). Prevoir un planning de "
               "rotation specifique en mode crise.",
               bold_prefix="Renforcement des effectifs --")

    add_bullet(doc, " PSL assure environ 62% des repas, CFX 38%. En crise, envisager un "
               "reequilibrage si l'un des sites est plus impacte.",
               bold_prefix="Repartition entre sites --")

    # --- 1.2 DECHETS ---
    add_heading_styled(doc, "1.2 Gestion des dechets hospitaliers", level=2)

    add_body(doc,
        "Les dechets hospitaliers sont un enjeu majeur de securite sanitaire. Les Dechets "
        "d'Activites de Soins a Risques Infectieux (DASRI) necessitent un traitement specifique "
        "et reglemente. En situation de crise, l'augmentation de l'activite de soins entraine "
        "mecaniquement une hausse des dechets, notamment infectieux."
    )

    chart_path = gen_chart_dechets(df)
    doc.add_picture(chart_path, width=Inches(5.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    dasri = df[(df["INDICATEUR"] == "D\u00e9chets") & (df["SOUS-INDICATEUR"].str.contains("Infectieux")) & (df["ANNEE"] == 2015)].iloc[0]
    menagers = df[(df["INDICATEUR"] == "D\u00e9chets") & (df["SOUS-INDICATEUR"].str.contains("M\u00e9nagers Assimil")) & (df["ANNEE"] == 2015)].iloc[0]

    add_constat_box(doc,
        f"Les DASRI (dechets infectieux) passeraient de {dasri['TOTAL_NORMAL']:,.0f} tonnes a "
        f"{dasri['TOTAL_CRISE']:,.0f} tonnes en situation de crise (+{dasri['ECART_TOTAL']:,.0f} t). "
        f"Les dechets menagers augmenteraient de {menagers['TOTAL_NORMAL']:,.0f} tonnes a "
        f"{menagers['TOTAL_CRISE']:,.0f} tonnes (+{menagers['ECART_TOTAL']:,.0f} t)."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Installer des conteneurs supplementaires de stockage DASRI sur chaque site "
               "(zone dediee, ventilee, securisee). Prevoir des capacites de stockage froid pour "
               "les dechets anatomiques si les filieres d'elimination sont saturees.",
               bold_prefix="Augmentation des capacites de stockage temporaire --")

    add_bullet(doc, " Negocier des clauses de montee en charge avec le prestataire de collecte DASRI "
               "(passage quotidien au lieu de 3 fois/semaine). Prevoir un prestataire secondaire "
               "en cas de defaillance.",
               bold_prefix="Renforcement des contrats d'enlevement --")

    add_bullet(doc, " Constituer un stock d'EPI (Equipements de Protection Individuelle) specifiques "
               "pour les agents de collecte. Renforcer les formations aux gestes de securite.",
               bold_prefix="Protection du personnel de collecte --")

    add_bullet(doc, " Renforcer les consignes de tri pour eviter les contaminations croisees. "
               "Assurer une tracabilite complete des DASRI (bordereau de suivi des dechets).",
               bold_prefix="Tri renforce et tracabilite --")

    # --- 1.3 HYGIENE ---
    add_heading_styled(doc, "1.3 Hygiene et entretien des locaux", level=2)

    add_body(doc,
        "Le nettoyage et la desinfection des locaux sont des elements determinants dans la lutte "
        "contre les infections nosocomiales. En situation de crise sanitaire, les protocoles "
        "d'hygiene doivent etre intensifies, tant au niveau de la frequence que des produits utilises."
    )

    chart_path = gen_chart_hygiene(df)
    doc.add_picture(chart_path, width=Inches(5.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    locaux = df[(df["INDICATEUR"] == "Hygi\u00e8ne") & (df["SOUS-INDICATEUR"] == "Locaux") & (df["ANNEE"] == 2015)].iloc[0]

    add_constat_box(doc,
        f"En 2015, les surfaces de locaux a entretenir s'elevent a {locaux['TOTAL_NORMAL']:,.0f} m\u00b2 "
        f"en situation normale et augmenteraient a {locaux['TOTAL_CRISE']:,.0f} m\u00b2 en crise. "
        f"Les espaces verts (147 014 m\u00b2), les vitres (92 701 m\u00b2) et la voirie sont egalement concernes."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Definir un protocole de desinfection de niveau crise avec des produits virucides "
               "et bactericides adaptes. Augmenter la frequence de nettoyage : passage de 1 a 3 fois/jour "
               "dans les zones de soins. Attention particuliere aux surfaces frequemment touchees "
               "(poignees, rampes, boutons d'ascenseur).",
               bold_prefix="Plan de bionettoyage renforce --")

    add_bullet(doc, " Maintenir un stock strategique de produits desinfectants pour couvrir 30 jours "
               "de consommation en mode crise. Diversifier les fournisseurs.",
               bold_prefix="Stocks de produits d'entretien --")

    add_bullet(doc, " Convention avec des entreprises de nettoyage pour un deploiement rapide d'equipes "
               "supplementaires (delai < 24h). Formation prealable aux protocoles hospitaliers.",
               bold_prefix="Mobilisation de personnel supplementaire --")

    add_bullet(doc, " Concentrer les efforts sur les zones de soins et les unites de reanimation. "
               "Reduire temporairement la frequence d'entretien des zones administratives.",
               bold_prefix="Priorisation des zones --")

    # --- 1.4 LINGERIE ---
    add_heading_styled(doc, "1.4 Lingerie hospitaliere", level=2)

    add_body(doc,
        "Le linge hospitalier (draps, blouses, serviettes, tenues du personnel) est un element "
        "essentiel du confort des patients et de l'hygiene de l'etablissement. Un afflux de patients "
        "entraine une hausse directe des besoins en linge propre."
    )

    chart_path = gen_chart_lingerie(df)
    doc.add_picture(chart_path, width=Inches(5.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    l2015 = df[(df["INDICATEUR"] == "Lingerie") & (df["SOUS-INDICATEUR"] == "Quantit\u00e9 de linge") & (df["ANNEE"] == 2015)].iloc[0]

    add_constat_box(doc,
        f"La quantite de linge traite en 2015 est de {l2015['TOTAL_NORMAL']:,.0f} kg en situation "
        f"normale. En crise, elle passerait a {l2015['TOTAL_CRISE']:,.0f} kg, soit pres de "
        f"{l2015['ECART_TOTAL']:,.0f} kg supplementaires par an. "
        f"Pitie-Salpetriere represente environ {l2015['PLF_NORMAL']/l2015['TOTAL_NORMAL']*100:.0f}% du volume."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Identifier une blanchisserie industrielle agree pouvant absorber le surplus. "
               "Tester cette convention une fois par an (exercice grandeur nature).",
               bold_prefix="Convention avec une blanchisserie de secours --")

    add_bullet(doc, " Constituer une reserve de draps et blouses a usage unique (non-tisses) pour "
               "les 15 premiers jours de crise. Ces articles limitent aussi le risque de transmission "
               "d'agents infectieux.",
               bold_prefix="Stock de linge a usage unique --")

    add_bullet(doc, " Augmenter la frequence de ramassage du linge souille (de 2 a 4 passages/jour). "
               "Prevoir des sacs hydrosolubles specifiques pour le linge contamine.",
               bold_prefix="Circuits de collecte renforces --")

    add_bullet(doc, " En crise, fournir des tenues a usage unique au personnel soignant. "
               "Prevoir des vestiaires temporaires aux entrees des unites d'isolement.",
               bold_prefix="Linge du personnel soignant --")

    # --- 1.5 MAGASIN ---
    add_heading_styled(doc, "1.5 Magasin et approvisionnements", level=2)

    add_body(doc,
        "Le magasin central gere les fournitures hotelieres, les produits d'hygiene et les "
        "imprimes. En crise, la demande en fournitures de protection et d'hygiene augmente "
        "considerablement."
    )

    chart_path = gen_chart_magasin(df)
    doc.add_picture(chart_path, width=Inches(5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    df_mag = df[(df["INDICATEUR"] == "Magasin") & (df["ANNEE"] == 2015)]
    total_n = df_mag["TOTAL_NORMAL"].sum()
    total_c = df_mag["TOTAL_CRISE"].sum()

    add_constat_box(doc,
        f"Le magasin gere environ {total_n:.0f} references en situation normale. "
        f"En crise, le volume de commandes augmente de 70% (soit {total_c:.0f} references a gerer), "
        f"meme si le nombre de references catalogue reste stable."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Maintenir en permanence un stock tampon de 30 jours pour les consommables "
               "critiques : gants, masques, surblouses, solutions hydroalcooliques, desinfectants. "
               "Systeme de rotation premier entre, premier sorti pour eviter la peremption.",
               bold_prefix="Stock strategique de 30 jours --")

    add_bullet(doc, " Pour chaque categorie de produit critique, avoir au minimum 2 fournisseurs "
               "references et une source alternative. Lecon du COVID-19 : eviter la dependance "
               "a un fournisseur unique.",
               bold_prefix="Diversification des fournisseurs --")

    add_bullet(doc, " Definir des seuils d'alerte pour chaque reference critique. Declenchement "
               "automatique d'une commande lorsqu'un stock passe sous le seuil.",
               bold_prefix="Systeme d'alerte de reapprovisionnement --")

    add_bullet(doc, " En cas de rupture, activer les circuits de depannage au sein du GHU AP-HP "
               "pour mutualiser les stocks disponibles entre etablissements.",
               bold_prefix="Coordination inter-etablissements --")

    # --- 1.6 VAGUEMESTRE ---
    add_heading_styled(doc, "1.6 Service courrier (Vaguemestre)", level=2)

    add_body(doc,
        "Le service courrier assure la distribution du courrier aux patients, l'envoi des "
        "documents administratifs et la reception des colis. Bien que secondaire face aux enjeux "
        "sanitaires, la continuite du service courrier contribue au bien-etre des patients "
        "hospitalises et au bon fonctionnement administratif."
    )

    chart_path = gen_chart_vaguemestre(df)
    doc.add_picture(chart_path, width=Inches(5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    plis = df[(df["INDICATEUR"] == "Vaguemestre") & (df["SOUS-INDICATEUR"] == "Plis affranchis") & (df["ANNEE"] == 2015)].iloc[0]
    colis = df[(df["INDICATEUR"] == "Vaguemestre") & (df["SOUS-INDICATEUR"] == "Colis") & (df["ANNEE"] == 2015)].iloc[0]

    add_constat_box(doc,
        f"Plus de {plis['TOTAL_NORMAL']:,.0f} plis affranchis par an et environ "
        f"{colis['TOTAL_NORMAL']:,.0f} colis (2015). En crise, les volumes augmenteraient "
        f"significativement, notamment les envois administratifs."
    )

    p = doc.add_paragraph()
    run = p.add_run("Propositions :")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BLEU_FONCE

    add_bullet(doc, " Privilegier l'envoi par e-mail ou via le portail patient pour les "
               "comptes-rendus et resultats. Reduire la dependance au courrier physique.",
               bold_prefix="Dematerialisation des courriers administratifs --")

    add_bullet(doc, " Prevoir un agent supplementaire mobilisable en cas de crise. "
               "Amenager les horaires de distribution pour les adapter aux contraintes de confinement.",
               bold_prefix="Renforcement ponctuel du service --")

    # --- SYNTHESE LOGISTIQUE ---
    doc.add_page_break()
    add_heading_styled(doc, "Synthese -- Logistique hospitaliere", level=2)

    chart_path = gen_chart_synthese(df)
    doc.add_picture(chart_path, width=Inches(5.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tableau de synthese
    add_body(doc, "Tableau recapitulatif de l'impact d'une crise sanitaire sur la logistique (donnees 2015) :")

    df_2015 = df[df["ANNEE"] == 2015]
    synthese = df_2015.groupby("INDICATEUR").agg({
        "TOTAL_NORMAL": "sum",
        "TOTAL_CRISE": "sum",
        "ECART_TOTAL": "sum"
    }).reset_index()

    headers = ["Domaine", "Volume normal", "Volume en crise", "Ecart", "Hausse"]
    rows = []
    for _, r in synthese.iterrows():
        rows.append([
            r["INDICATEUR"],
            f"{r['TOTAL_NORMAL']:,.0f}",
            f"{r['TOTAL_CRISE']:,.0f}",
            f"+{r['ECART_TOTAL']:,.0f}",
            "+70%"
        ])
    add_styled_table(doc, headers, rows)

    doc.add_paragraph()

    # Plan d'action prioritaire
    add_heading_styled(doc, "Plan d'action prioritaire -- Logistique", level=3)

    headers_pa = ["Priorite", "Action", "Responsable"]
    rows_pa = [
        ["HAUTE", "Constituer un stock strategique d'EPI et de consommables (30 jours)", "Direction des Achats"],
        ["HAUTE", "Negocier des conventions avec prestataires de secours (restauration, linge, dechets)", "Direction Logistique"],
        ["HAUTE", "Definir un protocole de bionettoyage de niveau crise", "Service Hygiene"],
        ["HAUTE", "Augmenter les capacites de stockage DASRI", "Services Techniques"],
        ["MOYENNE", "Mettre en place un systeme d'alerte de reapprovisionnement automatique", "Pharmacie / Magasin"],
        ["MOYENNE", "Former le personnel logistique aux procedures de crise", "DRH / Formation Continue"],
        ["MOYENNE", "Prevoir des menus simplifies activables en 48h", "Service Restauration"],
        ["BASSE", "Dematerialiser le courrier administratif", "Direction des Systemes d'Information"],
        ["BASSE", "Reduire la frequence d'entretien des zones non-critiques en crise", "Direction Logistique"],
    ]
    add_styled_table(doc, headers_pa, rows_pa, col_widths=[2.5, 10, 4.5])

    # ============================
    # SECTIONS A VENIR (placeholders)
    # ============================
    doc.add_page_break()

    placeholder_sections = [
        ("2. Activite et services", "Volumes d'activite et organisation des soins"),
        ("3. Capacite d'accueil", "Lits, equipements et plateaux techniques"),
        ("4. Finance", "Equilibre budgetaire et investissements"),
        ("5. Patients", "Profil des patients et parcours de soins"),
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

    # Nettoyage des charts temporaires
    import shutil
    shutil.rmtree(CHARTS_DIR, ignore_errors=True)


if __name__ == "__main__":
    build_document()
