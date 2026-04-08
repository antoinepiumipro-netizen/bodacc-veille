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
                print(f"  Rate limit, attente 30s...")
                time.sleep(30)
                continue
            if r.status_code == 200:
                results = r.json().get("results", [])
                if results:
                    siege = results[0].get("siege", {})
                    return {
                        "Ville": siege.get("libelle_commune", ""),
                        "Département": siege.get("departement", ""),
                        "Code Postal": siege.get("code_postal", ""),
                    }
                return {"Ville": "Non trouvé", "Département": "", "Code Postal": ""}
            return {"Ville": f"Erreur {r.status_code}", "Département": "", "Code Postal": ""}
        except Exception as e:
            if tentative == 2:
                return {"Ville": f"Erreur: {str(e)[:40]}", "Département": "", "Code Postal": ""}
            time.sleep(5)

wb = openpyxl.lo
