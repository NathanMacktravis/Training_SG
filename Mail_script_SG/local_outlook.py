import os
import win32com.client
from datetime import datetime, timedelta

# === CONFIGURATION ===
EMAIL_COMMUN = "adresse_mail_commune@votre_organisation.com"
EXPEDITEUR_AUTORISE = "expediteur@domaine.com"
DOSSIER_DESTINATION = r"C:\Chemin\Vers\Le\Répertoire"

# === PÉRIODE À CONSIDÉRER ===
# Pour aujourd’hui uniquement :
date_debut = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
date_fin = datetime.now()

# === OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Trouve la boîte partagée
boite_mail = None
for account in outlook.Folders:
    if EMAIL_COMMUN in account.Name:
        boite_mail = account
        break

if not boite_mail:
    raise Exception("Boîte mail commune non trouvée.")

inbox = boite_mail.Folders["Boîte de réception"]

# Tri des mails par date décroissante
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

# Filtrage des mails par période
for message in messages:
    try:
        if message.Class != 43:
            continue

        date_reception = message.ReceivedTime
        if not (date_debut <= date_reception <= date_fin):
            continue

        if message.SenderEmailAddress.lower() != EXPEDITEUR_AUTORISE.lower():
            continue

        # Récupération des pièces jointes
        if message.Attachments.Count > 0:
            for i in range(1, message.Attachments.Count + 1):
                pièce = message.Attachments.Item(i)
                nom_fichier = pièce.FileName

                chemin_complet = os.path.join(DOSSIER_DESTINATION, nom_fichier)

                if not os.path.exists(chemin_complet):
                    pièce.SaveAsFile(chemin_complet)
                    print(f"Fichier sauvegardé : {chemin_complet}")
    except Exception as e:
        print(f"Erreur : {e}")
