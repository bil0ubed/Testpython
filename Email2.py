import os
import email
from email.policy import default
import pandas as pd

def analyser_emails(dossier_eml, fichier_excel):
    # Liste pour stocker les données des e-mails
    donnees_emails = []

    # Parcourir tous les fichiers .eml dans le dossier
    for fichier in os.listdir(dossier_eml):
        if fichier.endswith('.eml'):
            chemin_fichier = os.path.join(dossier_eml, fichier)

            # Ouvrir et lire le fichier .eml
            with open(chemin_fichier, 'r', encoding='utf-8') as f:
                contenu = f.read()

            # Analyse du contenu de l'e-mail
            msg = email.message_from_string(contenu, policy=default)

            # Récupérer les données de l'e-mail
            destinataires = msg.get('To', '')
            cc = msg.get('Cc', '')
            sujet = msg.get('Subject', '')

            # Récupérer le corps de l'e-mail au format HTML (s'il existe)
            corps_html = ''
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/html':
                        corps_html = part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
            else:
                if msg.get_content_type() == 'text/html':
                    corps_html = msg.get_payload(decode=True).decode(msg.get_content_charset(), errors='ignore')

            # Ajouter les données dans la liste
            donnees_emails.append({
                'Destinataires': destinataires,
                'CC': cc,
                'Sujet': sujet,
                'Corps HTML': corps_html
            })

    # Créer un DataFrame pandas
    df = pd.DataFrame(donnees_emails)

    # Enregistrer les données dans un fichier Excel
    df.to_excel(fichier_excel, index=False, engine='openpyxl')
    print(f'Données enregistrées dans le fichier : {fichier_excel}')

# Exemple d'utilisation
dossier_eml = r'C:\chemin\vers\dossier_eml'  # Remplacez par le chemin du dossier contenant les fichiers .eml
fichier_excel = r'C:\chemin\vers\fichier.xlsx'  # Remplacez par le chemin de sortie du fichier Excel

analyser_emails(dossier_eml, fichier_excel)
