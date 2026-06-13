import os

# Ce fichier stocke toutes les constantes, pour ne les définir qu'une seule fois.

# Chemins
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH_FILE = os.path.join(BASE_DIR, "data", "processed", "Projet_ODD_SIVARAJAH.xlsx")
CSV_URL = "https://minio.lab.sspcloud.fr/nerojeni10/DATA_PROJET_SMA/Teen_Mental_Health_Dataset.csv"

# Couleurs standard
COLOR_BLUE_DARK = "FF003D6B"
COLOR_TEAL = "FF008080"
COLOR_LIGHT_TEAL = "FFB3E5E0"
COLOR_MALE = "4169E1"
COLOR_FEMALE = "FF1493"
COLORS_INTERACTION = ["20B2AA", "FF6B7B", "FFD700"]
