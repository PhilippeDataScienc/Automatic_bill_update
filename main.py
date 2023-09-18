# coding: utf-8
 
import tkinter as tk
from tkinter import ttk, font
import threading
from envoi_mail_gmail import *
import smtplib
import webbrowser
from generateur_facture import *

# pyinstaller -F main.py

fenetre = tk.Tk()

fenetre.title("Générateur de facture")  # Définir le titre de la fenêtre

# Personnaliser le style des widgets ttk
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12), background="#3498db", foreground="#37d67a", borderwidth=0, relief=tk.SUNKEN)
style.map("TButton", background=[("active", "#37d67a")])
style.configure("TEntry", font=("Helvetica", 12))

# Définir le fond en couleur
fenetre.configure(bg="#8ed1fc")
label = tk.Label(fenetre, text="Générateur de facture", font=("Helvetica", 16, "bold"), bg="#ecf0f1")
label.pack(pady=20)

def run_facture():
    progress_bar.start()
    output_text.set("Génération des factures en cours")
    fenetre.update_idletasks()
    result_message = execute_facture()  # Capture the result message from execute_facture
    output_text.set(result_message)  # Set the output_text to the captured result message
    progress_bar.stop()  # Stop the progress bar

def run_mail():
    password = password_entry.get()
    your_email = email_var.get()
    output_text.set("Envoi des e-mails en cours")

    def send_emails():
        try:
            for el in pdf_path:
                destinataire = get_facture_name(el)
                if destinataire in noms:
                    if email[noms.index(destinataire)] != "":
                        print(f"sending {el} to {email[noms.index(destinataire)]}")
                        send_mail(your_email, destinataire, password)
                        output_text.set(output_text.get() + f"\nE-mail envoyé à {destinataire}.")
                        fenetre.update_idletasks()  # Forcer la mise à jour de l'interface
        except smtplib.SMTPAuthenticationError:
            print("Mauvais couple nom d'utilisateur et mot de passe")
            output_text.set("Mauvais couple nom d'utilisateur et mot de passe")
            fenetre.update_idletasks()  # Forcer la mise à jour de l'interface
        except:
            print("adresse mail au format non standard détecté")
            output_text.set("adresse mail au format non standard détecté")
            fenetre.update_idletasks()  # Forcer la mise à jour de l'interface

        print("Tous les e-mails ont été envoyés.")
        output_text.set(output_text.get() + "\nTous les e-mails ont été envoyés.")
        fenetre.update_idletasks()  # Forcer la mise à jour de l'interface

    threading.Thread(target=send_emails).start()


# Créer un cadre pour regrouper la barre de progression et le bouton
progress_frame = tk.Frame(fenetre,bd=1, relief=tk.SOLID, pady=5)
progress_frame.pack(pady=10, padx=10)  # Ajoute un espacement vertical autour du cadre


progress_bar = tk.ttk.Progressbar(progress_frame, mode="indeterminate")
progress_bar.pack(side=tk.LEFT, padx=10)  # Ajoute un espacement horizontal autour de la barre de progression

button_run_facture = tk.Button(progress_frame, text="Générer les factures du mois en cours", command=run_facture)
button_run_facture.pack(side=tk.RIGHT, padx=10)  # Ajoute un espacement horizontal autour du bouton


# Champ d'entrée pour le mot de passe Gmail
email_label = tk.Label(fenetre, text="vérifiez votre adresse mail d'envoi", bg="#8ed1fc")
email_label.pack(pady=3)

# Champ d'entrée pour l'adresse e-mail avec texte persistant
email_var = tk.StringVar()
email_var.set("rotonde.alisa@gmail.com")
email_entry = tk.Entry(fenetre, textvariable=email_var, width=40)
email_entry.pack(pady=10)

def open_link():
    webbrowser.open("https://support.google.com/accounts/answer/185833?hl=fr")

style = ttk.Style()
style.configure("Link.TLabel", foreground="blue", cursor="hand2")

# Champ d'entrée pour le mot de passe Gmail
password_label = tk.Label(fenetre, text="Mot de passe d'application du compte google:", bg="#8ed1fc")
password_label.pack(pady=3)

link_label = ttk.Label(fenetre, text="Comment générer un mot de passe d'application google ?", style="Link.TLabel")
link_label.pack(pady=5)
link_label.bind("<Button-1>", lambda event: open_link())

password_entry = tk.Entry(fenetre, show="*", width=40)  # Pour cacher les caractères du mot de passe
password_entry.pack(pady=10)



# Créer un bouton
button_run_mail = tk.Button(fenetre, text="Envoi des factures du mois en cours", command=run_mail)
button_run_mail.pack(pady=10)

output_text = tk.StringVar()
output_label = tk.Label(fenetre, textvariable=output_text, background="#8ed1fc", foreground="red")
output_label.pack()


# bouton de sortie
bouton=tk.Button(fenetre, text="Fermer", command=fenetre.quit)
bouton.pack(pady=10)


# Obtenir la largeur et la hauteur de l'écran
screen_width = fenetre.winfo_screenwidth()
screen_height = fenetre.winfo_screenheight()

# Calculer les coordonnées pour centrer la fenêtre
x = (screen_width - fenetre.winfo_reqwidth()) // 2
y = (screen_height - fenetre.winfo_reqheight()) // 2

# Définir la position de la fenêtre pour la centrer
fenetre.geometry("+{}+{}".format(x, y))

fenetre.mainloop()