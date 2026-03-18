import requests
import re
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date

# ── CONFIG ──────────────────────────────────────────────
SEUIL = 3_000_000
EMAIL_FROM  = "u4356824811@gmail.com"
EMAIL_TO    = "u4356824811@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 587
SMTP_PASS   = os.environ["SMTP_PASS"]
# ────────────────────────────────────────────────────────

API_URL = (
    "https://bodacc-datadila.opendatasoft.com/api/records/1.0/search/"
    "?dataset=annonces-commerciales"
    "&refine.familleavis=Vente"
    "&rows=100"
    "&sort=dateparution"
    f"&refine.dateparution={date.today().isoformat()}"
)

def extraire_montant(texte):
    patterns = [
        r"prix\s+de\s+([\d\s\u202f]+)\s*[€euros]",
        r"moyennant\s+(?:le\s+prix\s+de\s+)?([\d\s\u202f]+)\s*[€euros]",
    ]
    for pattern in patterns:
        match = re.search(pattern, texte, re.IGNORECASE)
        if match:
            montant_str = re.sub(r"[\s\u202f]", "", match.group(1))
            try:
                return int(montant_str)
            except ValueError:
                continue
    return None

def fetch_cessions():
    response = requests.get(API_URL, timeout=30)
    response.raise_for_status()
    data = response.json()
    return data.get("records", [])

def filtrer_grosses_cessions(records):
    resultats = []
    for rec in records:
        fields = rec.get("fields", {})
        texte = fields.get("contenu", "") or ""
        montant = extraire_montant(texte)
        if montant and montant >= SEUIL:
            resultats.append({
                "date": fields.get("dateparution", ""),
                "numero": fields.get("numeroannonce", ""),
                "ville": fields.get("ville", ""),
                "dept": fields.get("numerodepartement", ""),
                "texte": texte,
                "montant": montant,
            })
    return resultats

def envoyer_email(cessions):
    if not cessions:
        print("Aucune cession > 3M€ aujourd'hui.")
        return

    corps = f"<h2>🏪 {len(cessions)} cession(s) &gt; 3 000 000 € — {date.today()}</h2>\n"
    for c in cessions:
        corps += f"""
        <hr>
        <b>📅 Date :</b> {c['date']}<br>
        <b>📍 Ville :</b> {c['ville']} ({c['dept']})<br>
        <b>💶 Montant :</b> {c['montant']:,} €<br>
        <b>📄 Annonce n°{c['numero']} :</b><br>
        <blockquote style="font-size:12px">{c['texte'][:800]}...</blockquote>
        <a href="https://www.bodacc.fr/pages/annonces-commerciales/">Voir sur BODACC</a>
        """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[BODACC] {len(cessions)} cession(s) > 3M€ — {date.today()}"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.attach(MIMEText(corps, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, SMTP_PASS)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"✅ Email envoyé : {len(cessions)} cession(s) trouvée(s).")

if __name__ == "__main__":
    records = fetch_cessions()
    cessions = filtrer_grosses_cessions(records)
    envoyer_email(cessions)
