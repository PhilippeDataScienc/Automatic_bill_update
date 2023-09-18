import yagmail
from datetime import datetime
import glob
import os
from pathlib import Path
from generateur_facture import get_facture_name, parse_txt


mois = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']

with open('./path.txt', 'r',encoding="utf8") as f:
    chemin_dossier_dropbox=Path(f.readline())

dossier_mois_en_cours = (f"Factures {(str(datetime.now().month).zfill(2))} {datetime.now().year}")

def send_mail(your_email,to_somebody, password): 
    yag = yagmail.SMTP(your_email, password) 
    destinataire = liste_mail[to_somebody] 
    titre = f"SCI LA ROTONDE ALISA - Facture du mois {mois[datetime.now().month-1]} {datetime.now().year}"
    message = f"""Bonjour {to_somebody},
    Voici la facture du mois {mois[datetime.now().month-1]} {datetime.now().year}
    Bonne journée
    """
    pj=liste_pdf_path[to_somebody]
    yag.send(to = destinataire,
             subject =  titre,
             contents = message,
             attachments= pj)

pdf_folder = os.path.join(chemin_dossier_dropbox,dossier_mois_en_cours, "*.pdf")
pdf_path = glob.glob(pdf_folder)
noms, montants, email = parse_txt("./liste_des_locataires.txt")
liste_mail = dict(zip(noms, email))
name_within_pdf = [get_facture_name(el) for el in pdf_path]
liste_pdf_path = dict(zip(name_within_pdf,pdf_path))