import pandas as pd
from pathlib import Path

# Le chemin vers le dossier avec les CSV (au cas où ils bougent)
csv_originaux = Path("Originaux")

# Le fichier Excel de sortie pour les data propre
fichier_excel = "Rush1_nettoyé.xlsx"

# On récupère tous les CSV
csv_files = list(csv_originaux.glob("*.csv"))

with pd.ExcelWriter(fichier_excel, engine="openpyxl") as writer:
    for csv_file in csv_files:

        # Lecture du CSV avec le séparateur en virgule, si ça change un jour changer le ","
        df = pd.read_csv(csv_file, sep=",")

        # Suppression des lignes identiques
        df = df.drop_duplicates()

        # Nom de la feuille (maximum 31 caractères)
        nom_feuille = csv_file.stem[:31]

        # Export vers Excel
        df.to_excel(writer, sheet_name=nom_feuille, index=False)

print(f"{len(csv_files)} fichiers CSV copiés dans {fichier_excel}, sans doublons.")