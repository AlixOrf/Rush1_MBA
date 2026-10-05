import pandas as pd

# Nos fichiers d'entrée et de sortie
excel_nettoye = "Rush1_nettoyé.xlsx"
fichier_excel = "Rush1_regroupé.xlsx"


# On lit les feuilles de l'excel
feuilles = pd.read_excel(
    excel_nettoye,
    sheet_name=None
)

print("Feuilles trouvées :")
for nom, df in feuilles.items():
    print(f"  - {nom}: {len(df)} lignes, {len(df.columns)} colonnes")


# On fusionne les feuilles

def fusionner_feuilles(feuilles, cle):

    resultat = None

    for nom_feuille, df in feuilles.items():

        # On ignore les feuilles qui ne possèdent pas la clé
        if cle not in df.columns:
            continue

        print(f"  → {nom_feuille} ajoutée pour {cle}")

        df = df.copy()

        # Supprimer les lignes sans clé
        df = df.dropna(subset=[cle])

        # Supprimer les doublons dans cette feuille
        df = df.drop_duplicates()

        # Une seule ligne par clé dans chaque feuille
        df = df.drop_duplicates(subset=[cle])

        # Première feuille
        if resultat is None:
            resultat = df
            continue

        # Colonnes déjà présentes
        colonnes_communes = [
            col for col in df.columns
            if col != cle and col in resultat.columns
        ]

        # Colonnes uniquement présentes dans la nouvelle feuille
        nouvelles_colonnes = [
            col for col in df.columns
            if col != cle and col not in resultat.columns
        ]

  
        # Ajouter les nouvelles colonnes

        # On fait une jointure
        resultat = resultat.merge(
            df[[cle] + nouvelles_colonnes + colonnes_communes],
            on=cle,
            how="outer",
            suffixes=("", "_nouveau")
        )

        for col in colonnes_communes:

            nouvelle_colonne = col + "_nouveau"

            if nouvelle_colonne in resultat.columns:

                resultat[col] = (
                    resultat[col]
                    .combine_first(resultat[nouvelle_colonne])
                )

                resultat.drop(
                    columns=[nouvelle_colonne],
                    inplace=True
                )

    if resultat is None:
        return pd.DataFrame()


    resultat = resultat.drop_duplicates(
        subset=[cle]
    )

    return resultat


# ============================================================
# FEUILLE VIDÉOS
# ============================================================

print("VIDÉOS")

videos = fusionner_feuilles(
    feuilles,
    "video_id"
)

print(f"Nombre de vidéos uniques : {len(videos)}")


# ============================================================
# FEUILLE AUTEURS
# ============================================================

print("\n========== AUTEURS ==========")

auteurs = fusionner_feuilles(
    feuilles,
    "author_uniqueId"
)

print(f"Nombre d'auteurs uniques : {len(auteurs)}")


# ============================================================
# EXPORT
# ============================================================

with pd.ExcelWriter(
    fichier_excel,
    engine="openpyxl"
) as writer:

    videos.to_excel(
        writer,
        sheet_name="Vidéos",
        index=False
    )

    auteurs.to_excel(
        writer,
        sheet_name="Auteurs",
        index=False
    )


print("\n===================================")
print("FUSION TERMINÉE")
print("===================================")
print(f"Fichier créé : {fichier_excel}")
print(f"Vidéos  : {len(videos)}")
print(f"Auteurs : {len(auteurs)}")