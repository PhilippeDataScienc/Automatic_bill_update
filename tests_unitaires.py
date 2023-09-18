import unittest
from generateur_facture import get_facture_name, change_and_rename, indice_facture_max, folder_of_max_facture, parse_txt
from unidecode import unidecode
from pathlib import Path
from unittest.mock import patch
from envoi_mail_gmail import send_mail, mois, datetime

class Test_generateur_facture(unittest.TestCase):

    def test_get_facture_name(self):
        filepath = "test/F1-01-2023NomFacture.docx"
        self.assertEqual(get_facture_name(filepath), "NomFacture")

    def test_indice_facture_max(self):
        list_of_files = ["test/F1-01-2023Nom1.docx", "test/F2-01-2023Nom2.docx", "test/F3-01-2023Nom3.docx"]
        self.assertEqual(indice_facture_max(list_of_files), 3)

    def test_folder_of_max_facture(self):
        list_of_files = ["test/F1-01-2023Nom1.docx", "test/F2-01-2023Nom2.docx", "test/F3-01-2023Nom3.docx"]
        max_number = 3
        self.assertEqual(folder_of_max_facture(list_of_files, max_number), Path("test"))

    def test_parse_txt(self):
        file = "liste_des_locataires.txt"
        expected_noms = ["Vitrey", "Fouchère", "Bettinelli"]
        expected_noms = [unidecode(nom) for nom in expected_noms]
        expected_montants = ["1552", "324", "22"]
        expected_emails = ["email1@example.com", "email2@example.com", "email3@example.com"]
        noms, montants, emails = parse_txt(file)
        self.assertEqual(noms[:3], expected_noms)
        self.assertEqual(montants[:3], expected_montants)
        self.assertEqual(emails[:3], expected_emails)


class Test_envoi_mail_gmail(unittest.TestCase):

    @patch("yagmail.SMTP")
    def test_send_mail(self, mock_smtp):
        # Données de test
        your_email = "votre_email@gmail.com"
        to_somebody = "Gagliardi"
        password = "votre_mot_de_passe"
        
        # Appel à la fonction
        send_mail(your_email, to_somebody, password)
        
        # Vérification que le constructeur SMTP a été appelé avec les bons arguments
        mock_smtp.assert_called_with(your_email, password)
        
        # Vérification que la méthode send a été appelée avec les bons arguments
        mock_smtp.return_value.send.assert_called_with(
            to="philippeacquier@yahoo.fr",
            subject=f"SCI LA ROTONDE ALISA - Facture du mois {mois[datetime.now().month-1]} {datetime.now().year}",
            contents=f"""Bonjour {to_somebody},
    Voici la facture du mois {mois[datetime.now().month-1]} {datetime.now().year}
    Bonne journée
    """,
            attachments="C:\\Users\\pacquier\\Desktop\\Auto-implement\\Factures 09 2023\\F688-09-2023Gagliardi.pdf"
        )


if __name__ == '__main__':
    unittest.main()
