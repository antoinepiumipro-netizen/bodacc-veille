import requests
import time
import pandas as pd
import openpyxl

INPUT_FILE = "entreprises.xlsx"
OUTPUT_FILE = "entreprises_enrichies.xlsx"
SIRET_COL = "Siret"
BASE_URL = "https://recherche-entreprises.api.gouv.fr/search"

def fetch_info(siret):
    siret = str(siret).strip().replace(" ", "")
    for tentative in range(3):
        try:
            r = requests.get(BASE_URL, params={"q": siret, "per_page": 1}, timeout=15)
            if r.status_code == 429:
                print("  Rate limit, attente 30s...")
                time.sleep(30)
                continue
            if r.status_code == 200:
                results = r.json().get("results", [])
                if results:
                    siege = results[0].get("siege", {})
                    return {
                        "Ville": siege.get("libelle_commune", ""),
                        "Departement": siege.get("departement", ""),
                        "Code Postal": siege.get("code_postal", ""),
                    }
                return {"Ville": "Non trouve", "Departement": "", "Code Postal": ""}
            return {"Ville": "Erreur " + str(r.status_code), "Departement": "", "Code Postal": ""}
        except Exception as e:
            if tentative == 2:
                return {"Ville": "Erreur: " + str(e)[:40], "Departement": "", "Code Postal": ""}
            time.sleep(5)

wb = openpyxl.load_workbook(INPUT_FILE, read_only=True)
print("Onglets disponibles :", wb.sheetnames)
wb.close()

df = pd.read_excel(INPUT_FILE, dtype={SIRET_COL: str}, sheet_name="Export CFNews")
print("Colonnes disponibles :", df.columns.tolist())
villes, depts, cps = [], [], []

for i, row in df.iterrows():
    siret = str(row[SIRET_COL]).strip()
    nom = str(row["Societe Cible ou Acteur"]) if "Societe Cible ou Acteur" in df.columns else str(row.iloc[0])

    if not siret or siret in ("nan", "None", ""):
        print(str(i+1) + "/" + str(len(df)) + " " + nom + " -> SIRET vide, ignore")
        villes.append("")
        depts.append("")
        cps.append("")
        continue

    print(str(i+1) + "/" + str(len(df)) + " " + nom + " (" + siret + ")...")
    info = fetch_info(siret)
    villes.append(info["Ville"])
    depts.append(info["Departement"])
    cps.append(info["Code Postal"])
    print("  -> " + info["Ville"] + " (" + info["Departement"] + ")")
    time.sleep(0.5)

df["Ville"] = villes
df["Departement"] = depts
df["Code Postal"] = cps
df.to_excel(OUTPUT_FILE, index=False)
print("Fichier genere : " + OUTPUT_FILE)
