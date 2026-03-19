import requests
import os
import time
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]
BASE_URL = "https://recherche-entreprises.api.gouv.fr/search"

# TEST : 10 premières entreprises — remplacer par la liste complète pour le vrai run
ENTREPRISES = [
    "3DEUS DYNAMICS", "A-NSE", "AAS INDUSTRIES", "ABC", "AC DISMANTLING",
    "ACB", "ADB", "ADDEV MATERIALS AEROSPACE SAS", "AddUp", "ADHETEC",
    "ADSS", "AEGIS PLATING", "AEQUS AEROSPACE FRANCE", "AERIADES",
    "AERO NEGOCE INTERNATIONAL", "AEROCAMPUS AQUITAINE", "AEROCAST",
    "AEROCENTRE", "AEROMETAL SAS", "AEROMETALS & ALLOYS",
]

# ── Délais anti rate-limit ───────────────────────────────────────────
DELAI_SIRENE   = 2   # secondes entre chaque appel Sirene
DELAI_CLAUDE   = 3   # secondes entre chaque appel Claude
ATTENTE_429    = 30  # secondes d'attente si rate limit 429
# ────────────────────────────────────────────────────────────────────

def chercher_siren(nom):
    for tentative in range(3):
        try:
            r = requests.get(BASE_URL, params={"q": nom, "per_page": 3}, timeout=15)
            if r.status_code == 429:
                print(f"  Sirene 429, attente {ATTENTE_429}s...")
                time.sleep(ATTENTE_429)
                continue
            if r.status_code == 200:
                data = r.json()
                resultats = data.get("results", [])
                if resultats:
                    e = resultats[0]
                    return {
                        "siren": e.get("siren", ""),
                        "nom_trouve": e.get("nom_complet", "") or e.get("nom_raison_sociale", ""),
                        "ville": e.get("siege", {}).get("commune", "") if e.get("siege") else "",
                        "statut": "Trouve"
                    }
                return {"siren": "", "nom_trouve": "", "ville": "", "statut": "Non trouve"}
            return {"siren": "", "nom_trouve": "", "ville": "", "statut": f"Erreur {r.status_code}"}
        except Exception as e:
            return {"siren": "", "nom_trouve": "", "ville": "", "statut": f"Erreur: {str(e)[:40]}"}
    return {"siren": "", "nom_trouve": "", "ville": "", "statut": "Echec apres 3 tentatives"}

def appel_claude(payload, timeout=60):
    """Appel API Claude avec retry automatique si 429."""
    for tentative in range(3):
        try:
            r = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "x-api-key": ANTHROPIC_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json=payload,
                timeout=timeout,
            )
            if r.status_code == 429:
                print(f"  Claude 429, attente {ATTENTE_429}s... (tentative {tentative+1})")
                time.sleep(ATTENTE_429)
                continue
            r.raise_for_status()
            return r.json()
        except Exception as e:
            if tentative == 2:
                raise
            time.sleep(5)
    return None

def verifier_faux_positif(nom_recherche, nom_trouve):
    """Étape 2 : Claude sans web search — même entreprise ?"""
    prompt = f"""Est-ce que ces deux noms désignent la même entreprise ?
Nom recherché : "{nom_recherche}"
Nom trouvé dans Sirene : "{nom_trouve}"

Réponds UNIQUEMENT en JSON :
{{"faux_positif": true ou false, "raison": "explication en 1 phrase"}}"""

    try:
        res = appel_claude({
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 200,
            "messages": [{"role": "user", "content": prompt}],
        })
        texte = res["content"][0]["text"].strip().strip("```json").strip("```").strip()
        data = json.loads(texte)
        return data.get("faux_positif", True), data.get("raison", "")
    except Exception as e:
        return True, f"Erreur: {str(e)[:40]}"

def chercher_description(nom, siren):
    """Étape 3 : Claude avec web search — description 1 phrase + lien aéro."""
    prompt = f"""Recherche cette entreprise française.
Nom : "{nom}"
SIREN : {siren}

Réponds UNIQUEMENT en JSON :
{{
  "description": "1 phrase courte sur l'activité principale",
  "lien_aero": true si activité liée à l'aéronautique/spatial/défense/aviation, false sinon
}}"""

    try:
        res = appel_claude({
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 300,
            "tools": [{"type": "web_search_20250305", "name": "web_search"}],
            "messages": [{"role": "user", "content": prompt}],
        })
        texte = ""
        for bloc in res["content"]:
            if bloc.get("type") == "text":
                texte += bloc.get("text", "")
        texte = texte.strip().strip("```json").strip("```").strip()
        data = json.loads(texte)
        return data.get("description", ""), data.get("lien_aero", True)
    except Exception as e:
        return "", True  # En cas d'erreur on garde par défaut

def creer_excel(vrais_positifs, faux_positifs):
    """Génère un Excel avec 2 onglets."""
    wb = Workbook()

    HEADER_BG = "1F3864"
    HEADER_FP  = "8B0000"
    thin = Side(style="thin", color="C0C0C0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def ecrire_onglet(ws, lignes, couleur_header, headers, largeurs):
        for col, (h, w) in enumerate(zip(headers, largeurs), 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            cell.fill = PatternFill("solid", start_color=couleur_header)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 30

        for i, row_data in enumerate(lignes, 2):
            bg = "EEF2F7" if i % 2 == 0 else "FFFFFF"
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill("solid", start_color=bg)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
            ws.row_dimensions[i].height = 50

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Onglet 1 : Vrais positifs
    ws1 = wb.active
    ws1.title = "Vrais positifs"
    ecrire_onglet(
        ws1, vrais_positifs, HEADER_BG,
        ["Nom GIFAS", "SIREN", "Nom officiel Sirene", "Ville", "Description"],
        [30, 12, 35, 20, 60]
    )

    # Onglet 2 : Faux positifs
    ws2 = wb.create_sheet("A vérifier manuellement")
    ecrire_onglet(
        ws2, faux_positifs, HEADER_FP,
        ["Nom GIFAS", "SIREN trouvé", "Nom officiel Sirene", "Raison exclusion"],
        [30, 12, 35, 60]
    )

    wb.save("siren_enrichi_test.xlsx")
    print("Fichier genere : siren_enrichi_test.xlsx")
    print(f"  Onglet 1 - Vrais positifs : {len(vrais_positifs)} entreprises")
    print(f"  Onglet 2 - A verifier : {len(faux_positifs)} entreprises")

if __name__ == "__main__":
    vrais_positifs = []
    faux_positifs  = []
    total = len(ENTREPRISES)

    for i, nom in enumerate(ENTREPRISES, 1):
        print(f"\n[{i}/{total}] {nom}")

        # ── Étape 1 : Recherche SIREN ──────────────────────────────
        siren_res = chercher_siren(nom)
        print(f"  SIREN : {siren_res['statut']} → {siren_res.get('siren', '')} {siren_res.get('nom_trouve', '')}")
        time.sleep(DELAI_SIRENE)

        if siren_res["statut"] != "Trouve":
            faux_positifs.append([nom, "", "", f"SIREN non trouvé : {siren_res['statut']}"])
            continue

        # ── Étape 2 : Vérification faux positif (sans web search) ──
        time.sleep(DELAI_CLAUDE)
        faux_positif, raison = verifier_faux_positif(nom, siren_res["nom_trouve"])
        print(f"  Faux positif : {faux_positif} — {raison}")

        if faux_positif:
            faux_positifs.append([nom, siren_res["siren"], siren_res["nom_trouve"], f"Nom différent : {raison}"])
            continue

        # ── Étape 3 : Description + vérification lien aéro (web search) ──
        time.sleep(DELAI_CLAUDE)
        description, lien_aero = chercher_description(nom, siren_res["siren"])
        print(f"  Lien aéro : {lien_aero} — {description[:80]}")

        if not lien_aero:
            faux_positifs.append([nom, siren_res["siren"], siren_res["nom_trouve"], f"Hors secteur aéro : {description}"])
            continue

        # ── Vrai positif confirmé ──────────────────────────────────
        vrais_positifs.append([nom, siren_res["siren"], siren_res["nom_trouve"], siren_res["ville"], description])

    creer_excel(vrais_positifs, faux_positifs)
    print(f"\nTerminé : {len(vrais_positifs)} vrais positifs, {len(faux_positifs)} à vérifier.")
