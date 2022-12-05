from tkinter import *
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import customtkinter
import pandas as pd
import sys
import os

# pyinstaller
# C:\Users\skornber\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\LocalCache\local-packages\Python310\Scripts\pyinstaller.exe
# C:\Users\skornber\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\LocalCache\local-packages\Python310\Scripts\auto-py-to-exe


customers = ["--- Kunden wählen ---", "SKA", "ZVO", "STD", "SWQ", "BBS", "SWNO",
             "TOR"]

workflows = ["Transfer AG",
             "Transfer an AG",
             "abgeschlossen"]

abort_codes = ["2000 - techn. Mangel",
               "2419 - falsche Verbraucherdaten",
               "2412 - Technischer Mangel",
               "2436 - Fehlanfahrt",
               "2407 - Wechsel nicht möglich",
               "2401 - Zähler bereits gewechselt",
               "2011 - Versorgerventil",
               "2004 - TW nicht mögl",
               "2003 - Zutritt verweigert",
               "2002 - falsche Verbraucherdaten",
               "2001 - Leerstand",
               "147 - Covid 19 Abbruch",
               "138 - Zähler bereits gewechselt",
               "124 - Langzeiturlaub",
               "123 - Leerstand",
               "116 - Zähler n. auffindbar",
               "112 - Haus abgerissen"
               ]
results = {}

save = True


# ------------------------------------ GUI -----------------------------------
class GUI(customtkinter.CTk):
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue")

    def __init__(self):
        super().__init__()
        self.geometry("500x690")
        self.resizable(0, 0)
        self.title("Abrechnungsanalyse")
        self.icon = tk.PhotoImage(file="image/abrechnung.png")
        self.wm_iconphoto(False, self.icon)

    def widgets(self):
        # Combobox
        self.combobox = customtkinter.CTkComboBox(
            self, width=200, fg_color="#395E9C", values=sorted(customers))
        self.combobox.place(x=20, y=20, width=260)

        # Button
        self.button = customtkinter.CTkButton(
            self, text="Export laden...", command=open_dialog)
        self.button.place(x=300, y=20, width=180)

        # Frame
        self.frame = Frame(self, borderwidth=0, bg="#395E9C", relief="ridge")
        self.frame.place(x=20, y=70, width=460, height=600)

        # Treeview
        self.treeview = ttk.Treeview(
            self, columns=("key", "value"), show="tree")
        self.treeview.place(x=30, y=80, width=440, height=580)
        # Treeview Columns
        self.treeview.column('#0', minwidth=10, width=10, stretch=NO)
        self.treeview.column("key", minwidth=230, width=230, stretch=NO)
        self.treeview.column("value", minwidth=200, width=200, stretch=NO)

# ----------------------------------- PANDAS ----------------------------------


def open_dialog():
    # Clear results
    results.clear()

    # Clear treeview
    window.treeview.delete(*window.treeview.get_children())

    # load adms export
    try:
        file = filedialog.askopenfilename(
            title="Export laden...", filetypes=(("Excel Datei", ".xlsx .xls"),
                                                ("Alle Dateien", "*.*")))
        global df
        df = pd.read_excel(file)

        get_file_path(file)
        get_file_name(file)

        template()
    except Exception as e:
        messagebox.showerror("Woah!", f"Da ist ein Problem aufgetreten {e}")


def template():
    # Common query's
    empty_row()
    periode()
    count_data()

    # Read customer from Combobox
    selection = str(window.combobox.get())
    # Normal Customer
    if selection.startswith(("SKA", "ZVO")):
        finished()
    # Foreign System
    elif selection.startswith(("BBS", "SWQ")):
        foreign()
    # Other
    elif selection.startswith(("STD", "SWNO")):
        other()

    # Display results
    show_results()


def get_file_path(file):
    global path
    path = os.path.dirname(file)


def get_file_name(file):
    global filename
    filename = os.path.basename(file).split(".")[0]


def empty_row():
    # Make first Row empty
    results[""] = ""


def periode():
    # Get minimum date
    date_start = df["Einbaudatum"].min().strftime("%d.%m.")
    # Get maximum date
    date_end = df["Einbaudatum"].max().strftime("%d.%m.%Y")
    # write into results
    results["Abrechnungszeitraum"] = (
            f"{date_start} - {date_end}")


def finished():
    # hide lines with empty values
    finished = df[
        (df["ZählerNr"].notnull()) &
        (df["Workflowschritt"].isin(workflows))]

    results["Abgeschlossen"] = len(finished)

    extension = ("_abgeschlossene_Aufträge")
    write_results(finished, extension)


def foreign():
    # results["Foreign"] = "this is foreign"
    pass


def other():
    # results["Other"] = "this is other"
    pass


def count_data():
    # read dataquantity and write into results
    results["Anzahl Datensätze"] = (len(df))


def show_results():
    # Read the values from the results and show them into treeview
    for index, (key, value) in enumerate(results.items()):
        window.treeview.insert(
            "", tk.END, iid=index, text="", values=(key, value))


def write_results(data, extension):
    # Write excel file without index
    if (save):
        data.to_excel(path + "/" + filename + extension + ".xlsx", index=False)


if __name__ == "__main__":
    window = GUI()
    window.widgets()
    window.mainloop()
