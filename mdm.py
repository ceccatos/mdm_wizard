import tkinter as tk
from tkinter import ttk
import sv_ttk
from tkinter import messagebox
import openpyxl
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("MASTER DATA WIZARD")
        self.create_home_screen()

    def create_home_screen(self):

        # Clear the screen
        for widget in self.root.winfo_children():
            widget.destroy()

        label = ttk.Label(self.root, text="Fornitore")
        label.grid(row=0, column=0, padx=10, pady=15, sticky="e")

        self.entrySupplier = ttk.Entry(self.root, width=25)
        self.entrySupplier.grid(row=0, column=1, padx=(10, 5), pady=15)  # Reduced right padding

        # Add an info icon
        supplier_info_icon = ttk.Button(self.root, text="?", width=2, command=self.supplier_infobox)
        supplier_info_icon.grid(row=0, column=2, padx=(5, 10), pady=15, sticky="w")  # Reduced left padding

        # Create three clickable cards with images
        cards = [
            {"text": "Template generico", "image": "img\\card_gen.png", "command": lambda: self.check_supplier_and_proceed("GEN")},
            {"text": "Template AL-CH-VI", "image": "img\\card_al.png", "command": lambda: self.check_supplier_and_proceed("AL")},
            {"text": "Template MO", "image": "img\\card_mo.png", "command": lambda: self.check_supplier_and_proceed("MO")},
        ]

        for i, card in enumerate(cards):
            frame = ttk.Frame(self.root, borderwidth=2, relief="groove", padding=10)
            frame.grid(row=1, column=i, padx=10, pady=10)

            # Display the text
            text_label = ttk.Label(frame, text=card["text"], font=("Tahoma", 16))
            text_label.pack()

            # Load and display the image
            try:
                image = tk.PhotoImage(file=card["image"])
                image_label = ttk.Label(frame, image=image)
                image_label.image = image  # Keep a reference to avoid garbage collection
                image_label.pack(pady=(10, 10))
            except Exception as e:
                label = ttk.Label(frame, text="[ERR]", font=("Tahoma", 12))
                label.pack()

            # Button to open the form
            button = ttk.Button(frame, text="Compila", command=card["command"], style="Accent.TButton")
            button.pack()

    def supplier_infobox(self):
        messagebox.showinfo("INFO", "Ragione sociale del Fornitore.\nIl campo non può essere vuoto")

    def check_supplier_and_proceed(self, card_name):
        # Check if the entrySupplier field is empty
        supplier = self.entrySupplier.get().strip()
        if not supplier:
            messagebox.showerror("ERRORE", "Il Fornitore non può essere vuoto.")
            return
        
        # Proceed to the form screen if validation passes
        self.create_form_screen(card_name)

    def create_form_screen(self, card_name):

        # Clear the screen
        for widget in self.root.winfo_children():
            widget.destroy()

        # Define a mapping of card_name to title
        card_name_to_title = {
            "GEN": "generico",
            "AL": "AL_CH_VI",
            "MO": "MO"
        }

        title = card_name_to_title.get(card_name, "[ERR]")

        # Form title
        title_label = ttk.Label(self.root, text=f"Template {title}", font=("Tahoma", 16))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Labels and Entry fields
        self.entries = {}
        labels = ["Nome", "Cognome", "Età"]
        for i, label_text in enumerate(labels):
            label = ttk.Label(self.root, text=label_text)
            label.grid(row=i+1, column=0, padx=10, pady=5, sticky="e")

            entry = ttk.Entry(self.root, width=25)
            entry.grid(row=i+1, column=1, padx=10, pady=5)
            self.entries[label_text] = entry

        # Save button
        save_button = ttk.Button(self.root, text="Salva", command=self.save_data)
        save_button.grid(row=len(labels)+1, column=0, columnspan=2, pady=20)

        # Back button
        back_button = ttk.Button(self.root, text="Home", command=self.create_home_screen)
        back_button.grid(row=len(labels)+2, column=0, columnspan=2, pady=5)

    def save_data(self):
        # Validate fields
        data = {}
        for label, entry in self.entries.items():
            value = entry.get().strip()
            if not value:
                messagebox.showerror("ERRORE", f"Il campo '{label}' non può essere vuoto.")
                return
            data[label] = value

        # Save to Excel
        file_name = "dati.xlsx"
        if not os.path.exists(file_name):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["Nome", "Cognome", "Età"])
            workbook.save(file_name)

        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active
        sheet.append([data["Nome"], data["Cognome"], data["Età"]])
        workbook.save(file_name)

        messagebox.showinfo("Successo", "Dati salvati con successo!")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)

    sv_ttk.set_theme("light")
    root.mainloop()
