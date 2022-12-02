from tkinter import *
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import customtkinter
import pandas as pd

customtkinter.set_appearance_mode("dark")  # Modes:  system, dark light
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()

root.title("Abrechnungsanalyse")
icon = tk.PhotoImage(file="image/abrechnung.png")
root.wm_iconphoto(False, icon)
root.geometry("500x700")
root.resizable(False, False)


def combobox_callback(choice):
    results = choice
    print("combobox dropdown clicked:", choice)
    return results


def open_excel():
    # open a file
    file = filedialog.askopenfilename(
        title="Open File", filetypes=(
            ("Excel Datei", ".xlsx .xls"),
            ("Alle Dateien", "*.*")
            )
        )

    # grab the file
    try:
        df = pd.read_excel(file)
    except Exception as e:
        messagebox.showerror("Woah!", f"Da ist ein Problem aufgetreten {e}")


# Dropdown
combobox = customtkinter.CTkComboBox(master=root,
                                     width=200,
                                     fg_color="#395E9C",
                                     values=[
                                        "Kunden wählen ",
                                        "SKA",
                                        "ZVO",
                                        "STD",
                                        "SWQ",
                                        "BBS",
                                     ],
                                     command=combobox_callback)
combobox.place(x=20, y=20, width=260)


# Button
load_button = customtkinter.CTkButton(
    root, text="Export laden...", command=open_excel)
load_button.place(x=300, y=20, width=180)


# Textview
textbox = customtkinter.CTkTextbox(root, width=460, height=610)
textbox.place(x=20, y=70)


results = """Here are the results
second line
third line"""

textbox.insert("0.0", results)  # insert at line 0 character 0






root.mainloop()
