"""
config/settings.py
------------------
Paramètres centraux du projet : chemins, couleurs, URL des données.
"""

import os


# ─── CHEMINS ──────────────────────────────────────────────────────────────────
DATA_URL = "https://minio.lab.sspcloud.fr/nerojeni10/DATA_PROJET_SMA/Teen_Mental_Health_Dataset.csv"
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(_ROOT, "template")
PATH_FILE = os.path.join(TEMPLATE_DIR, "Projet_ODD_SIVARAJAH.xlsx")

# ─── PALETTE DE COULEURS ──────────────────────────────────────────────────────
COLORS = {
    "title_bg": "FF003D6B",  # Bleu foncé  → titres principaux / en-têtes
    "header_bg": "FF008080",  # Teal        → en-têtes de colonnes
    "row_header_bg": "FFB3E5E0",  # Teal clair  → en-têtes de lignes
    "white": "FFFFFF",
    "black": "000000",
    "male": "040459",  # Bleu très foncé → série Male
    "female": "7D093F",  # Bordeaux        → série Female
    "scatter_teal": "20B2AA",  # Teal vif        → scatter (Low / No Dep)
    "scatter_rose": "FF6B7B",  # Rose vif        → scatter (High / Dep)
    "scatter_blue": "4169E1",  # Bleu roi        → scatter Male
    "scatter_pink": "FF1493",  # Rose chaud      → scatter Female
}

# ─── FILTRES EXCEL (Chaînes réutilisées dans les formules) ────────────────────
FILTER_TDB1 = (
    ', DATA!$B$2:$B$1201, IF(TDB1!$G$3="Tous", "<>", TDB1!$G$3)'
    ', DATA!$A$2:$A$1201, IF(TDB1!$D$3="Tous", "<>", TDB1!$D$3)'
    ', DATA!$D$2:$D$1201, IF(TDB1!$J$3="Tous", "<>", TDB1!$J$3)'
    ', DATA!$I$2:$I$1201, IF(TDB1!$M$3="Tous", "<>", TDB1!$M$3)'
)

FILTER_ARR_TDB1 = (
    ' * IF(TDB1!$G$3="Tous", 1, DATA!$B$2:$B$1201=TDB1!$G$3)'
    ' * IF(TDB1!$D$3="Tous", 1, DATA!$A$2:$A$1201=TDB1!$D$3)'
    ' * IF(TDB1!$J$3="Tous", 1, DATA!$D$2:$D$1201=TDB1!$J$3)'
    ' * IF(TDB1!$M$3="Tous", 1, DATA!$I$2:$I$1201=TDB1!$M$3)'
)

FILTER_TDB2 = (
    ', DATA!$B$2:$B$1201, IF(TDB2!$G$3="Tous", "<>", TDB2!$G$3)'
    ', DATA!$A$2:$A$1201, IF(TDB2!$D$3="Tous", "<>", TDB2!$D$3)'
    ', DATA!$D$2:$D$1201, IF(TDB2!$J$3="Tous", "<>", TDB2!$J$3)'
    ', DATA!$I$2:$I$1201, IF(TDB2!$M$3="Tous", "<>", TDB2!$M$3)'
)

FILTER_ARR_TDB2 = (
    ' * IF(TDB2!$G$3="Tous", 1, DATA!$B$2:$B$1201=TDB2!$G$3)'
    ' * IF(TDB2!$D$3="Tous", 1, DATA!$A$2:$A$1201=TDB2!$D$3)'
    ' * IF(TDB2!$J$3="Tous", 1, DATA!$D$2:$D$1201=TDB2!$J$3)'
    ' * IF(TDB2!$M$3="Tous", 1, DATA!$I$2:$I$1201=TDB2!$M$3)'
)

# ─── COLONNES POUR len_dict ───────────────────────────────────────────────────
COLS_FOR_LEN = ["age", "gender", "platform_usage", "social_interaction_level"]

# ─── COLONNES NUMÉRIQUES POUR LA MATRICE DE CORRÉLATION ──────────────────────
COLS_CORR = [
    "age",
    "sleep_hours",
    "daily_social_media_hours",
    "academic_performance",
    "physical_activity",
    "social_num",
    "stress_level",
    "anxiety_level",
    "addiction_level",
    "depression_label",
]
