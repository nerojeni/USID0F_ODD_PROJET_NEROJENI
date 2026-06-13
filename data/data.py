import pandas as pd


def fetch_and_clean_data(url: str) -> pd.DataFrame:
    """
    Télécharge le dataset depuis une URL, vérifie les données manquantes,
    supprime les lignes incomplètes si nécessaire, et affiche un rapport.

    Parameters
    ----------
    url : str
        Le lien (URL) vers le fichier CSV.

    Returns
    -------
    pd.DataFrame
        Le DataFrame nettoyé, prêt pour l'analyse.
    """
    print("Téléchargement des données en cours...")

    try:
        # Chargement des données
        df = pd.read_csv(url)

        # Comptage du nombre de lignes avant nettoyage
        initial_rows = len(df)

        # Suppression des lignes avec des valeurs manquantes (NaN)
        df_cleaned = df.dropna()

        # Comptage du nombre de lignes après nettoyage
        final_rows = len(df_cleaned)

        # Calcul du nombre de lignes supprimées
        missing_rows = initial_rows - final_rows

        # Affichage des résultats dans la console
        print("-" * 50)
        if missing_rows > 0:
            print(
                f"Attention : {missing_rows} ligne(s) avec des données manquantes ont été supprimées."
            )
        else:
            print("Zéro donnée manquante : Le dataset est parfaitement complet !")

        print(
            f"Taille finale du dataset : {df_cleaned.shape[0]} lignes | {df_cleaned.shape[1]} colonnes."
        )
        print("-" * 50 + "\n")

        return df_cleaned

    except Exception as e:
        print(f"Erreur lors du téléchargement ou de la lecture des données : {e}")
        return None
