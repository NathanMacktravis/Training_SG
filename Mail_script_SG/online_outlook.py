"""
------ CONFIGURATIONS ------
Va sur https://portal.azure.com :

Azure Active Directory → "Enregistrements d’applications"

→ Nouvelle inscription

Nom : ScriptMailDownloader

Rediriger URI : http://localhost:8000

Copie :

Application (client) ID

Directory (tenant) ID

Dans "Certificats et secrets" → Créer un secret client

Exécuter cette ligne de commande pour l'installation de la librairie : pip install msal requests 

"""

import os
import requests
import msal

# === CONFIG ===
CLIENT_ID = "VOTRE_CLIENT_ID"
TENANT_ID = "VOTRE_TENANT_ID"
CLIENT_SECRET = "VOTRE_SECRET"
EMAIL_COMMUN = "adresse_mail_commune@votre_domaine.com"
EXPEDITEUR_AUTORISE = "expediteur@domaine.com"
DOSSIER_DEST = r"C:\chemin\vers\dossier"

# Auth
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_URL = "https://graph.microsoft.com/v1.0"

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

token = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in token:
    raise Exception("Erreur d'authentification MSAL")

headers = {
    "Authorization": f"Bearer {token['access_token']}",
    "Accept": "application/json"
}

# 🔍 Obtenir les 20 derniers mails de la boîte partagée
url = f"{GRAPH_URL}/users/{EMAIL_COMMUN}/mailFolders/inbox/messages?$top=20&$orderby=receivedDateTime desc"

resp = requests.get(url, headers=headers).json()
messages = resp.get("value", [])

for mail in messages:
    sender = mail.get("from", {}).get("emailAddress", {}).get("address", "").lower()
    if sender != EXPEDITEUR_AUTORISE.lower():
        continue

    mail_id = mail["id"]
    att_url = f"{GRAPH_URL}/users/{EMAIL_COMMUN}/messages/{mail_id}/attachments"
    att_resp = requests.get(att_url, headers=headers).json()

    for att in att_resp.get("value", []):
        if att["@odata.type"] != "#microsoft.graph.fileAttachment":
            continue

        nom = att["name"]
        contenu = att["contentBytes"]

        chemin = os.path.join(DOSSIER_DEST, nom)
        with open(chemin, "wb") as f:
            f.write(bytes(contenu, encoding="utf-8"))
        print(f"✅ Sauvegardé : {chemin}")
