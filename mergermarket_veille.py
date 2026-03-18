import os
import json
import base64
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ── CONFIG ──────────────────────────────────────────────
EMAIL_FROM    = "u4356824811@gmail.com"
EMAIL_TO      = "antoine.piumi@lazard.com"
SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_PASS     = os.environ["SMTP_PASS"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]
SEUIL_M       = 5  # millions €
# ────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

def get_gmail_service():
    """Authentification Gmail via token stocké en variable d'env."""
    token_data = json.loads(os.environ["GMAIL_TOKEN"])
    creds = Credentials(
        token=token_data["token"],
        refresh_token=token_data["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=SCOPES,
    )
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)

def get_mergermarket_emails(service):
    """Récupère les emails MergerMarket non lus."""
    results = service.users().messages().list(
        userId="me",
        q='from:mergermarket.com is:unread',
        maxResults=20
    ).execute()
    return results.get("messages", [])

def get_email_body(service, msg_id):
    """Extrait le texte d'un email."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    payload = msg["payload"]

    def extract_text(payload):
        if payload.get("mimeType") == "text/plain":
            data = payload.get("body", {}).get("data", "")
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
        for part in payload.get("parts", []):
            text = extract_text(part)
            if text:
                return text
        return ""

    return extract_text(payload), msg.get("snippet", "")

def mark_as_read(service, msg_id):
    """Marque l'email comme lu pour ne pas le retraiter."""
    service.users().messages().modify(
        userId="me",
        id=msg_id,
        body={"removeLabelIds": ["UNREAD"]}
    ).execute()

def analyser_avec_claude(texte_email):
    """Envoie l'email à Claude pour analyse."""
    prompt = f"""Tu es un assistant pour un banquier privé français. Analyse cette alerte MergerMarket et réponds UNIQUEMENT en JSON avec ce format exact :

{{
  "pertinent": true ou false,
  "raison_pertinence": "explication courte si pertinent, sinon null",
  "acheteur": "nom de l'acheteur",
  "vendeur": "nom du vendeur ou de la cible",
  "pays_acheteur": "pays",
  "pays_vendeur": "pays",
  "montant_estime_millions_eur": nombre ou null,
  "secteur": "secteur d'activité",
  "resume": "2-3 phrases de résumé du deal",
  "action_recommandee": "action concrète à faire si pertinent, sinon null"
}}

Un deal est PERTINENT si :
- L'acheteur OU le vendeur est français (entreprise française ou basée en France)
- ET le montant est supérieur à {SEUIL_M}M€ (ou inconnu mais deal significatif)

Email MergerMarket à analyser :
---
{texte_email[:4000]}
---"""

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1000,
            "messages": [{"role": "user", "content": prompt}],
        },
        timeout=30,
    )
    response.raise_for_status()
    content = response.json()["content"][0]["text"]

    # Nettoie le JSON si Claude ajoute des backticks
    content = content.strip().strip("```json").strip("```").strip()
    return json.loads(content)

def envoyer_alerte(analyse, texte_original):
    """Envoie l'email de synthèse."""
    montant_str = f"{analyse['montant_estime_millions_eur']}M€" if analyse.get("montant_estime_millions_eur") else "montant non précisé"

    corps = f"""
    <h2>🔔 Alerte MergerMarket — Deal pertinent</h2>
    <hr>
    <b>🏢 Acheteur :</b> {analyse.get('acheteur', 'N/A')} ({analyse.get('pays_acheteur', 'N/A')})<br>
    <b>🎯 Cible / Vendeur :</b> {analyse.get('vendeur', 'N/A')} ({analyse.get('pays_vendeur', 'N/A')})<br>
    <b>💶 Montant :</b> {montant_str}<br>
    <b>🏭 Secteur :</b> {analyse.get('secteur', 'N/A')}<br>
    <br>
    <b>📋 Résumé :</b><br>
    {analyse.get('resume', 'N/A')}
    <br><br>
    <b>✅ Action recommandée :</b><br>
    <div style="background:#f0f7ff;padding:10px;border-left:4px solid #0066cc">
    {analyse.get('action_recommandee', 'N/A')}
    </div>
    <br>
    <b>💡 Pourquoi pertinent :</b> {analyse.get('raison_pertinence', 'N/A')}
    <hr>
    <details>
    <summary style="cursor:pointer;color:#666">Voir l'email original</summary>
    <pre style="font-size:11px;color:#444">{texte_original[:3000]}</pre>
    </details>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[MergerMarket] {analyse.get('acheteur', '?')} acquiert {analyse.get('vendeur', '?')} — {montant_str}"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.attach(MIMEText(corps, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, SMTP_PASS)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"✅ Alerte envoyée : {analyse.get('acheteur')} / {analyse.get('vendeur')}")

if __name__ == "__main__":
    print("Connexion Gmail...")
    service = get_gmail_service()

    emails = get_mergermarket_emails(service)
    print(f"{len(emails)} email(s) MergerMarket non lu(s) trouvé(s)")

    for msg in emails:
        msg_id = msg["id"]
        texte, snippet = get_email_body(service, msg_id)

        if not texte:
            texte = snippet

        print(f"Analyse : {snippet[:80]}...")

        try:
            analyse = analyser_avec_claude(texte)
            if analyse.get("pertinent"):
                envoyer_alerte(analyse, texte)
            else:
                print(f"  → Non pertinent, ignoré")
        except Exception as e:
            print(f"  → Erreur analyse : {e}")

        mark_as_read(service, msg_id)

    print("Terminé.")
