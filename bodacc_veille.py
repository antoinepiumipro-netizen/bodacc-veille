import requests
import re
import smtplib
import os
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date

# ── CONFIG ──────────────────────────────────────────────
SEUIL = 100_000
EMAIL_FROM  = "u4356824811@gmail.com"
EMAIL_TO    = "u4356824811@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 587
SMTP_PASS   = os.environ["SMTP_PASS"]
# ────────────────────────────────────────────────────────

BASE_URL = (
    "https://bodacc-datadila.opendatasoft.com/api/records/1.0/search/"
    "?dataset=annonces-commerciales"
    "&rows=100"
    "&sort=dateparution"
    f"&refine.dateparution={date.today().isoformat()}"
)

def fetch_records(famille):
    url = BASE_URL + f"&refine.familleavis={famille}"
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        data = response.json()
        print(f"  → {data.get('nhits', 0)} annonces trouvées pour '{famille}' aujourd'hui")
        return data.get("records", [])
    except Exception as e:
        print(f"Erreur API pour {famille} : {e}")
        return []

def extraire_montant_texte(texte):
    """Extrait un montant en € depuis un texte libre."""
    patterns = [
        r"prix\s+de\s+([\d\s\u202f\.]+)\s*[€euros]",
        r"moyennant\s+(?:le\s+prix\s+de\s+)?([\d\s\u202f\.]+)\s*[€euros]",
        r"capital\s+(?:social\s+)?(?:de\s+|fix[eé]\s+[àa]\s+|port[eé]\s+[àa]\s+)?([\d\s\u202f\.]+)\s*[€euros]",
        r"apport[s]?\s+(?:de\s+)?([\d\s\u202f\.]+)\s*[€euros]",
        r"montant\s+(?:de\s+)?([\d\s\u202f\.]+)\s*[€euros]",
    ]
    for pattern in patterns:
        match = re.search(pattern, texte, re.IGNORECASE)
        if match:
            try:
                montant_str = re.sub(r"[\s\u202f\.]", "", match.group(1))
                val = int(montant_str)
                if val > 1000:  # ignore les montants absurdes < 1000€
                    return val
            except ValueError:
                continue
    return None

def extraire_montant_acte(acte_str):
    """Cherche un montant dans le JSON du champ 'acte'."""
    try:
        acte = json.loads(acte_str)
        # Convertit tout le JSON en texte et cherche les montants
        texte = json.dumps(acte, ensure_ascii=False)
        return extraire_montant_texte(texte)
    except Exception:
        return None

def get_texte_complet(fields):
    """Assemble tout le texte disponible d'une annonce."""
    parties = []
    for champ in ["acte", "listepersonnes", "listeetablissements", "commercant"]:
        val = fields.get(champ, "")
        if val:
            parties.append(str(val))
    return " ".join(parties)

def traiter_records(records, type_label):
    resultats = []
    for rec in records:
        fields = rec.get("fields", {})
        texte = get_texte_complet(fields)
        montant = extraire_montant_texte(texte)

        # Pour les cessions : cherche aussi dans le champ acte séparément
        if not montant and fields.get("acte"):
            montant = extraire_montant_acte(fields["acte"])

        if montant and montant >= SEUIL:
            resultats.append({
                "type": type_label,
                "date": fields.get("dateparution", ""),
                "numero": fields.get("numeroannonce", ""),
                "ville": fields.get("ville", ""),
                "dept": fields.get("numerodepartement", ""),
                "commercant": fields.get("commercant", ""),
                "tribunal": fields.get("tribunal", ""),
                "url": fields.get("url_complete", "https://www.bodacc.fr/pages/annonces-commerciales/"),
                "texte": texte,
                "montant": montant,
            })
    return resultats

def envoyer_email(toutes_annonces):
    if not toutes_annonces:
        print(f"Aucune annonce > {SEUIL:,} € aujourd'hui.")
        return

    toutes_annonces.sort(key=lambda x: x["montant"], reverse=True)

    corps = f"<h2>📋 {len(toutes_annonces)} annonce(s) &gt; {SEUIL:,} € — {date.today()}</h2>\n"
    for c in toutes_annonces:
        corps += f"""
        <hr>
        <b>{c['type']}</b><br>
        <b>🏢 Entreprise :</b> {c['commercant']}<br>
        <b>📅 Date :</b> {c['date']}<br>
        <b>📍 Ville :</b> {c['ville']} ({c['dept']})<br>
        <b>⚖️ Tribunal :</b> {c['tribunal']}<br>
        <b>💶 Montant détecté :</b> {c['montant']:,} €<br>
        <a href="{c['url']}">👉 Voir l'annonce complète sur BODACC</a>
        """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[BODACC] {len(toutes_annonces)} annonce(s) > {SEUIL:,} € — {date.today()}"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.attach(MIMEText(corps, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, SMTP_PASS)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"✅ Email envoyé : {len(toutes_annonces)} annonce(s) trouvée(s).")

if __name__ == "__main__":
    print("Récupération des cessions (vente)...")
    cessions  = traiter_records(fetch_records("vente"), "🏪 Cession de fonds de commerce")

    print("Récupération des créations...")
    creations = traiter_records(fetch_records("creation"), "🏗️ Création d'entreprise")

    print("Récupération des modifications de capital...")
    modifs    = traiter_records(fetch_records("modification"), "📈 Augmentation de capital")

    toutes = cessions + creations + modifs
    envoyer_email(toutes)
