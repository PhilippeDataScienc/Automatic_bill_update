import docx
from datetime import datetime
from pathlib import Path
import glob
import os
from unidecode import unidecode
import os
import win32com.client
import pythoncom


wdFormatPDF = 17


mois = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']

with open('./path.txt', 'r',encoding="utf8") as f:
    chemin_dossier_dropbox=Path(f.readline())

def get_facture_name(filepath):
    name = Path(filepath).stem.split('-')[-1][4:]
    return name


def change_and_rename(filepath, index, montant = 0):
    pythoncom.CoInitialize()
    try:
        
        # Recupère le nom du fichier et le nom associé à la facture (dans le nom de facture)
        nom_fichier = Path(filepath).stem
        nom_facture = ''.join([i for i in nom_fichier[1:] if not i.isdigit() and i != "-"])

        # Transforme le docx en objet parsable dans python
        doc = docx.Document(filepath)
        paragraphs = doc.paragraphs

        try:
            indice_lieu = [i for i in range(len(paragraphs)) if paragraphs[i].text.startswith("Loubeyrat")][0]
            indice_facture = [i for i in range(len(paragraphs)) if paragraphs[i].text.startswith("Facture")][0]
            indice_mois = [i for i in range(len(paragraphs)) if paragraphs[i].text.startswith("Mois")][0]
            indice_montant = [i for i in range(len(paragraphs)) if paragraphs[i].text.startswith("Montant")][0]
            indice_recu = [i for i in range(len(paragraphs)) if paragraphs[i].text.startswith("Reçu")][0]
        except IndexError:
            return 0
        
        if montant != 0:
            loyer = montant
        else :
            try:
                loyer = (int(paragraphs[indice_recu].text.split(" ")[5]))
            except ValueError:
                loyer = (int(paragraphs[indice_recu].text.split(" ")[5][:-1]))

        try:
            paragraphs[indice_lieu].text = f"Loubeyrat, le {datetime.now().day} {mois[datetime.now().month-1]} {datetime.now().year}"
            paragraphs[indice_facture].text = f"Facture : F{index}-{mois[datetime.now().month-1]} {datetime.now().year}"
            paragraphs[indice_mois].text = f"Mois de {mois[datetime.now().month-1]} {datetime.now().year}"
            paragraphs[indice_montant].text = f"Montant forfaitaire mensuel de {loyer:.2f}€"
            paragraphs[indice_recu].text = f"Reçu la somme totale de {loyer}€ par virement"
        except:
            return 0

        facture_name = f"F{index}-{(str(datetime.now().month).zfill(2))}-{datetime.now().year}{nom_facture}"

        nom_dossier = (f"Factures {(str(datetime.now().month).zfill(2))} {datetime.now().year}")
        chemin_dossier = os.path.join(chemin_dossier_dropbox,f"{nom_dossier}")
        if not os.path.exists(chemin_dossier):
            os.mkdir(chemin_dossier)

        doc.save(os.path.join(chemin_dossier,f"{facture_name}.docx"))
        
        in_file = os.path.join(chemin_dossier,f"{facture_name}.docx")
        out_file = os.path.join(chemin_dossier,f"{facture_name}.pdf")

        word = win32com.client.Dispatch('Word.Application')
        doc2 = word.Documents.Open(in_file)
        doc2.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc2.Close()
        word.Quit()
        return 1
    finally:
        pythoncom.CoUninitialize()


def indice_facture_max(list_of_files):
    facture_nr_list=[]
    for file in list_of_files:
        F_nr = int((Path(file).name.split('-')[0][1:]))
        facture_nr_list.append(F_nr)
    return max(facture_nr_list)

def folder_of_max_facture(list_of_files,max_number):
    for file in list_of_files:
        if max_number == int((Path(file).name.split('-')[0][1:])):
            return Path(file).parent


def parse_txt(file):
    with open(file, 'r',encoding="utf8") as f:
        lines=f.readlines()
    
    noms = [unidecode(lines[i].strip('\n').split(",")[0]) for i in range(1,len(lines))]
    montants = [lines[i].strip('\n').strip().split(",")[1] for i in range(1,len(lines))]
    mails = [lines[i].strip('\n').split(",")[2] for i in range(1,len(lines))]
    return noms, montants, mails

def execute_facture():
    docx_path = os.path.join(chemin_dossier_dropbox,"**", "*.docx")
    doc_path = os.path.join(chemin_dossier_dropbox, "**", "*.doc")

    all_files = glob.glob(docx_path, recursive=True) + glob.glob(doc_path, recursive=True)

    max_idx = indice_facture_max(all_files)
    folder_to_modify = folder_of_max_facture(all_files, max_idx)
    print(folder_to_modify)
    files_to_modify = glob.glob(os.path.join(folder_to_modify,"*.docx"))

    noms, montants, email = parse_txt("./liste_des_locataires.txt")

    dossier_mois_en_cours = (f"Factures {(str(datetime.now().month).zfill(2))} {datetime.now().year}")

    counter=1
    if not os.path.exists(dossier_mois_en_cours):
        message = f"Dossier {dossier_mois_en_cours} crée !\n"
        print(message)
        for i,file in enumerate(files_to_modify):
            if unidecode(get_facture_name(file)) in noms:
                try:
                    loyer = int(montants[noms.index(unidecode(get_facture_name(file)))])
                except ValueError:
                    loyer = 0
                change_and_rename(file, max_idx+counter, loyer)
                montant_facture = loyer
                if loyer !=0:
                    print(f"\nFacture de {montant_facture:.2f}€ pour {unidecode(get_facture_name(file))} générée")
                    message = message + f"\nFacture de {montant_facture:.2f}€ pour {unidecode(get_facture_name(file))} générée"
                else:
                    print(f"\nmontant pour {unidecode(get_facture_name(file))} non renseigné dans le fichier txt")
                    message = message + f"\nmontant pour {unidecode(get_facture_name(file))} non renseigné dans le fichier txt"
                counter+=1
    else:
        return "\nDossier déjà existant. Veuillez l'effacer si vous voulez le générer à nouveau"
    return message + "\n\ntoutes les facures ont été générées"

