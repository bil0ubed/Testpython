import os
import smtplib
from email import policy
from email.parser import BytesParser

def envoyer_emails(dossier, serveur_smtp, port, email_utilisateur, mot_de_passe):
    # Liste tous les fichiers dans le dossier
    fichiers = [f for f in os.listdir(dossier) if f.endswith('.eml')]

    with smtplib.SMTP(serveur_smtp, port) as serveur:
        serveur.starttls()  # Chiffrement TLS
        serveur.login(email_utilisateur, mot_de_passe)  # Connexion au serveur SMTP

        for fichier in fichiers:
            chemin_fichier = os.path.join(dossier, fichier)
            
            # Lecture du fichier EML
            with open(chemin_fichier, 'rb') as f:
                email_message = BytesParser(policy=policy.default).parse(f)
            
            # Ajout de l'en-tête de confidentialité
            email_message['Sensitivity'] = 'company-confidential'  # Définit la confidentialité en tant qu'interne

            # Envoi de l'email
            try:
                serveur.send_message(email_message)
                print(f"Email envoyé depuis le fichier : {fichier}")
            except Exception as e:
                print(f"Erreur lors de l'envoi depuis le fichier {fichier} : {e}")

if __name__ == "__main__":
    dossier_emails = input("Entrez le chemin du dossier contenant les fichiers .eml : ")
    serveur_smtp = input("Entrez l'adresse du serveur SMTP (ex : smtp.gmail.com) : ")
    port = int(input("Entrez le port SMTP (ex : 587) : "))
    email_utilisateur = input("Entrez votre adresse email : ")
    mot_de_passe = input("Entrez votre mot de passe : ")

    envoyer_emails(dossier_emails, serveur_smtp, port, email_utilisateur, mot_de_passe)
