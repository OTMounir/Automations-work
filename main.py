import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import Fiche_Produit_PGC as fp  # Importez votre fichier de fonctions ici


# Fonctions pour les boutons de l'interface
def generate_all_fiches():
    fp.Extract_sheets(fp.wb_source, fp.Static_Sheets_names, fp.output_file_annexes)
    print("Toutes les annexes ont été générées.")
    categories = [
        ("ALIMENTATION_ANIMALE", fp.output_file_fp_AA),
        ("ALIMENTATION_HUMAINE", fp.output_file_fp_AH),
        ("DROGUERIE_PARFUMERIE_&HYGIENE", fp.output_file_fp_DPH),
        ("BOISSONS_ALCOOLISEES", fp.output_file_fp_BA),
        ("TABAC", fp.output_file_fp_tabac),
        ("Vins & Spi B2C", fp.output_file_fp_VS),
        ("Cosmétiques", fp.output_file_fp_COS),
        ("PGC", fp.output_file_fp_PGC)
    ]
    for name, file_path in categories:
        fp.creation_Fiche(name, file_path, fp.Dynamic_Sheets_name, fp.df)
        print(f"Fiche produit {name} générée.")
    messagebox.showinfo("Succès", "Toutes les fiche-produits ont été générées.")

def generate_specific_fiche():
    selected_fiche = fiche_var.get()
    output_files = {
        "ALIMENTATION_ANIMALE": fp.output_file_fp_AA,
        "ALIMENTATION_HUMAINE": fp.output_file_fp_AH,
        "DROGUERIE_PARFUMERIE_HYGIENE": fp.output_file_fp_DPH,
        "BOISSONS_ALCOOLISEES": fp.output_file_fp_BA,
        "TABAC": fp.output_file_fp_tabac,
        "Vins & Spi B2C": fp.output_file_fp_VS,
        "Cosmétiques": fp.output_file_fp_COS,
        "PGC": fp.output_file_fp_PGC
    }
    output_file = output_files.get(selected_fiche, "")
    if output_file:
        fp.creation_Fiche(selected_fiche, output_file, fp.Dynamic_Sheets_name, fp.df)
        messagebox.showinfo("Succès", f"Fiche produit {selected_fiche} générée avec succès.")
    else:
        messagebox.showerror("Erreur", "Erreur dans la génération de la fiche sélectionnée.")

def view_parameters():
    chemin = fp.chemin
    version_number = fp.version_number
    messagebox.showinfo("Paramètres", f"Chemin : {chemin}\nVersion : {version_number}")
def maj_fp():
    messagebox.showinfo("Succès", "Toutes les fiches produit ont été mises à jour!.")

    

def quit_app():
    root.destroy()

# Création de la fenêtre principale
root = tk.Tk()
root.title("Outil d'exportation de fiche-produits")
root.geometry("700x500")

# Load the logo and resize it
logo_image = Image.open("Assets/logo.png")  # Remplacez par le chemin de votre image
logo_image = logo_image.resize((150, 80), Image.ANTIALIAS)  # Redimensionnez le logo
logo_photo = ImageTk.PhotoImage(logo_image)

# Place the logo in a Label widget in the top-left corner
logo_label = tk.Label(root, image=logo_photo, bg="#f2f2f2")
logo_label.place(x=10, y=10)  # Position en haut à gauche

# Titre de l'application
label_title = tk.Label(root, text="Gestion des Fiche-Produits", font=("Helvetica", 18, "bold"), fg="#333", bg="#f2f2f2")
label_title.pack(pady=15)

# Bouton pour générer toutes les fiches produit
btn_generate_all = tk.Button(root, text="Générer toutes les fiche-produits", font=("Helvetica", 12), bg="#4CAF50", fg="white",
                             activebackground="#45a049", padx=10, pady=5, command=generate_all_fiches)
btn_generate_all.pack(pady=10)

# Bouton pour générer une fiche spécifique
fiche_var = tk.StringVar(value="ALIMENTATION_ANIMALE")
label_specific = tk.Label(root, text="Choisir une fiche spécifique :", font=("Helvetica", 12), bg="#f2f2f2")
label_specific.pack(pady=5)
options = ["ALIMENTATION_ANIMALE", "ALIMENTATION_HUMAINE", "DROGUERIE_PARFUMERIE_HYGIENE", "BOISSONS_ALCOOLISEES", "TABAC", "Vins & Spi B2C", "Cosmétiques","PGC"]
dropdown = tk.OptionMenu(root, fiche_var, *options)
dropdown.config(font=("Helvetica", 10), bg="#f9f9f9", fg="#333", width=25)
dropdown.pack(pady=5)
btn_generate_specific = tk.Button(root, text="Générer la fiche sélectionnée", font=("Helvetica", 12), bg="#2196F3", fg="white",
                                  activebackground="#1976D2", padx=10, pady=5, command=generate_specific_fiche)
btn_generate_specific.pack(pady=10)

# Bouton pour voir les paramètres
btn_view_parameters = tk.Button(root, text="Voir les paramètres", font=("Helvetica", 12), bg="#FF9800", fg="white",
                                activebackground="#FB8C00", padx=10, pady=5, command=view_parameters)
btn_view_parameters.pack(pady=10)

# Bouton pour voir les paramètres
btn_MAJ_FP = tk.Button(root, text="Mettre à jour toutes les fiches-produits (hors pgc)", font=("Helvetica", 12), bg="black", fg="white",
                                activebackground="#FB8C00", padx=10, pady=5, command=maj_fp)
btn_MAJ_FP.pack(pady=10)


# Bouton pour quitter l'application
btn_quit = tk.Button(root, text="Quitter", font=("Helvetica", 12), bg="#f44336", fg="white",
                     activebackground="#e53935", padx=10, pady=5, command=quit_app)
btn_quit.pack(pady=20)

# Lancer l'interface graphique
root.mainloop()
