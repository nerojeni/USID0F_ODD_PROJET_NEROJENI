# Projet Santé Mentale des Adolescents
## Dashboard Excel automatisé — openpyxl

---

## Arborescence du projet

```
projet_sante_mentale/
│
├── main.py                    ← Point d'entrée unique
│
├── config/
│   ├── __init__.py
│   └── settings.py            ← Chemins, couleurs, filtres Excel, constantes
│
├── data/
│   ├── __init__.py
│   └── loader.py              ← Chargement CSV, len_dict, feuilles DATA/CALC/Indicateurs/Correlations
│
├── sheets/
│   ├── __init__.py
│   ├── dashboards.py          ← TDB1, TDB2, TDB3 (titre + filtres)
│   ├── tcd.py                 ← TCD  : tableaux croisés + KPI pour TDB1
│   ├── tcd2.py                ← TCD2 : perf scolaire + dépressifs par âge
│   └── tcd3.py                ← TCD3 : 4 tableaux pour les scatter plots
│
├── charts/
│   ├── __init__.py
│   └── charts.py              ← Tous les graphiques (BarChart, PieChart, ScatterChart, Heatmap)
│
└── utils/
    ├── __init__.py
    ├── styles.py              ← Objets de style openpyxl centralisés
    └── excel_helpers.py       ← Fonctions génériques (write_table, add_filter, setup_dashboard_page…)
```

---

## Feuilles Excel générées

| Feuille       | Rôle                                              |
|---------------|---------------------------------------------------|
| `DATA`        | Données brutes (1 200 lignes)                     |
| `CALC`        | Valeurs distinctes (UNIQUE ArrayFormula)           |
| `Indicateurs` | 4 indicateurs simples (COUNTIFS / AVERAGEIF)       |
| `Correlations`| Matrice de corrélation (pandas → Excel)            |
| `TCD`         | Tableaux croisés + KPI + helpers pour TDB1        |
| `TCD2`        | Tableaux croisés pour TDB2                        |
| `TCD3`        | Tableaux pour les scatter plots de TDB3           |
| `TDB1`        | Dashboard page 1 (filtres + 3 graphiques + KPI)  |
| `TDB2`        | Dashboard page 2 (filtres + 2 graphiques + KPI)  |
| `TDB3`        | Dashboard page 3 (4 scatter plots)                |
| `TDB4`        | Heatmap des corrélations                          |

---

## Installation

```bash
pip install openpyxl==3.1.3 pandas==3.0.2 s3fs==2026.3.0
```

---

## Utilisation

```bash
# Pipeline complet
python main.py

# Étapes séparées
python main.py --step data      # Chargement + CALC + Indicateurs + Correlations
python main.py --step sheets    # TDB1/2/3 + TCD/TCD2/TCD3
python main.py --step charts    # Tous les graphiques
```

---

## Fonctions réutilisables

### `utils/excel_helpers.py`

| Fonction                  | Description                                            |
|---------------------------|--------------------------------------------------------|
| `write_table()`           | Écrit un DataFrame dans une feuille avec auto-largeur  |
| `write_section_title()`   | Bandeau titre bleu foncé avec merge                    |
| `write_col_headers()`     | Ligne d'en-têtes style teal                            |
| `write_row_header()`      | En-tête de ligne style teal clair                      |
| `add_filter()`            | Filtre interactif (titre + dropdown + colonne cachée)  |
| `setup_dashboard_page()`  | Init une page TDB (grille cachée + titre principal)    |
| `setup_filter_row()`      | Ajoute les 4 filtres horizontaux (Âge/Genre/Plateforme/Interaction) |

### `utils/styles.py`

Tous les objets `PatternFill`, `Font`, `Border`, `Alignment` en constantes pré-construites.

### `charts/charts.py`

| Fonction                  | Description                              |
|---------------------------|------------------------------------------|
| `build_all_charts()`      | Lance TDB1 + TDB2 + TDB3 en un appel    |
| `build_charts_tdb1()`     | BarChart addiction + PieChart + Boxplot  |
| `build_charts_tdb2()`     | KPI + BarChart perf + BarChart dépressifs|
| `build_charts_tdb3()`     | 4 scatter plots                          |
| `build_charts_tdb4()`     | Heatmap corrélations                     |
