import requests
import re
import smtplib
import os
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

def build_url(famille):
    return (
        "https://bodacc-datadila.opendatasoft.com/api/records/1.0/search/"
        "?dataset=annonces-commerciales"
        f"&refine.familleavis={famille}"
        "&rows=100"
        "&sort=dateparution"
        
    )

def extraire_montant_cession(texte):
    patterns = [
        r"prix\s+de\s+([\d\s\u202f]+)\s*[€euros]",
        r"moyennant\s+(?:le\s+prix\s+de\s+)?([\d\s\u202f]+)\s*[€euros]",
    ]
    for pattern in patterns:
        match = re.search(pattern, texte, re.IGNORECASE)
        if match:
            try:
                return int(re.sub(r"[\s\u202f]", "", match.group(1)))
            except ValueError:
                continue
    return None

def extraire_montant_capital(texte):
    patterns = [
        r"capital\s+(?:social\s+)?(?:de\s+|fix[eé]\s+[àa]\s+|port[eé]\s+[àa]\s+|est\s+de\s+)?([\d\s\u202f]+)\s*[€euros]",
        r"capital\s+(?:social\s+)?(?:initial\s+)?(?:de\s+)?([\d\s\u202f]+)\s*[€euros]",
        r"augment[eé]\s+(?:le\s+capital\s+)?(?:[àa]\s+|de\s+)([\d\s\u202f]+)\s*[€euros]",
        r"apport[s]?\s+(?:en\s+num[eé]raire\s+)?(?:de\s+)?([\d\s\u202f]+)\s*[€euros]",
        r"souscrit\s+(?:et\s+lib[eé]r[eé]\s+)?(?:de\s+)?([\d\s\u202f]+)\s*[€euros]",
    ]
    for pattern in patterns:
        match = re.search(pattern, texte, re.IGNORECASE)
        if match:
            try:
                return int(re.sub(r"[\s\u202f]", "", match.group(1)))
            except ValueError:
                continue
    return None

def fetch_records(famille):
    try:
        response = requests.get(build_url(famille), timeout=30)
        response.raise_for_status()
        return response.json().get("records", [])
    except Exception as e:
        print(f"Erreur API pour {famille} : {e}")
        return []

def traiter_cessions(records):
    resultats = []
    for rec in records:
        fields = rec.get("fields", {})
        texte = fields.get("contenu", "") or ""
        montant = extraire_montant_cession(texte)
        if montant and montant >= SEUIL:
            resultats.append({
                "type": "🏪 Cession de fonds de commerce",
                "date": fields.get("dateparution", ""),
                "numero": fields.get("numeroannonce", ""),
                "ville": fields.get("ville", ""),
                "dept": fields.get("numerodepartement", ""),
                "texte": texte,
                "montant": montant,
            })
    return resultats

def traiter_creations(records):
    resultats = []
    for rec in records:
        fields = rec.get("fields", {})
        texte = fields.get("contenu", "") or ""
        montant = extraire_montant_capital(texte)
        if montant and montant >= SEUIL:
            resultats.append({
                "type": "🏗️ Création d'entreprise",
                "date": fields.get("dateparution", ""),
                "numero": fields.get("numeroannonce", ""),
                "ville": fields.get("ville", ""),
                "dept": fields.get("numerodepartement", ""),
                "texte": texte,
                "montant": montant,
            })
    return resultats

def traiter_modifications(records):
    resultats = []
    for rec in records:
        fields = rec.get("fields", {})
        texte = fields.get("contenu", "") or ""
        if not re.search(r"capital|apport", texte, re.IGNORECASE):
            continue
        montant = extraire_montant_capital(texte)
        if montant and montant >= SEUIL:
            resultats.append({
                "type": "📈 Augmentation de capital",
                "date": fields.get("dateparution", ""),
                "numero": fields.get("numeroannonce", ""),
                "ville": fields.get("ville", ""),
                "dept": fields.get("numerodepartement", ""),
                "texte": texte,
                "montant": montant,
            })
    return resultats

def envoyer_email(toutes_annonces):
    if not toutes_annonces:
        print("Aucune annonce > 3M€ aujourd'hui.")
        return

    toutes_annonces.sort(key=lambda x: x["montant"], reverse=True)

    corps = f"<h2>📋 {len(toutes_annonces)} annonce(s) &gt; 3 000 000 € — {date.today()}</h2>\n"
    for c in toutes_annonces:
        corps += f"""
        <hr>
        <b>{c['type']}</b><br>
        <b>📅 Date :</b> {c['date']}<br>
        <b>📍 Ville :</b> {c['ville']} ({c['dept']})<br>
        <b>💶 Montant :</b> {c['montant']:,} €<br>
        <b>📄 Annonce n°{c['numero']} :</b><br>
        <blockquote style="font-size:12px;color:#444">{c['texte'][:800]}...</blockquote>
        <a href="https://www.bodacc.fr/pages/annonces-commerciales/">Voir sur BODACC</a>
        """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[BODACC] {len(toutes_annonces)} annonce(s) > 3M€ — {date.today()}"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.attach(MIMEText(corps, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, SMTP_PASS)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"✅ Email envoyé : {len(toutes_annonces)} annonce(s) trouvée(s).")

if __name__ == "__main__":
    print("Récupération des cessions...")
    cessions  = traiter_cessions(fetch_records("Vente"))

    print("Récupération des créations...")
    creations = traiter_creations(fetch_records("Immatriculation"))

    print("Récupération des modifications de capital...")
    modifs    = traiter_modifications(fetch_records("Modification"))

    toutes = cessions + creations + modifs
    envoyer_email(toutes)
