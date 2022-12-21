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

# empty list
customers = []
# first dropdown selection
cus_0 = ["--- Kunden wählen ---"]
# ! normal invoice customer
cus_1 = ["SKA", "ZVO", "STD", "TOR", "SWQ", "BBS", "KAKI", "EWS", "GEW", "OHS",
         "SWNO", "SNF", "SSN", "SNY"]
# ! foreign system customer
cus_2 = ["EMD", "SWN"]
# !other customer
cus_3 = ["SHH"]

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

        # MenuBar
        self.menubar = Menu(self)
        self.config(menu=self.menubar)
        self.file_menu = Menu(self.menubar, tearoff=False,
                              activebackground="#395E9C",
                              activeforeground="#ffffff")

        # MenuBar Menus
        self.menubar.add_cascade(label="Datei", menu=self.file_menu)

        # MenuBar Items
        self.save = IntVar()  # // declare Variable as Integer
        self.save.set(1)      # // set to default on
        self.file_menu.add_checkbutton(label="Auswertung speichern",
                                       variable=self.save,
                                       onvalue=1, offvalue=0)
        # MenuBar SeparatorLine
        self.file_menu.add_separator()

        self.file_menu.add_command(label="Beenden", command=self.quit)


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
        pass


def template():
    # Common query's
    empty_row(1)
    periode()
    count_data()

    # Read customer from Combobox
    selection = str(window.combobox.get())
    # Normal Customer
    if selection in cus_1:
        empty_row(2)
        finished()
        failed_tour()
        empty_row(3)
        same_place()
        empty_row(4)
        meter_size()
    # Foreign System
    elif selection in cus_2:
        foreign()
    # Other
    elif selection in cus_3:
        other()

    # Display results
    print(results)
    show_results()


def get_file_path(file):
    global path
    path = os.path.dirname(file)


def get_file_name(file):
    global filename
    filename = os.path.basename(file).split(".")[0]


def empty_row(i):
    # Make empty Rows
    results[(" " * i)] = ""


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


def failed_tour():
    # show only lines with blank value
    failed = df[
        (df["ZählerNr"].isnull()) &
        (df["Workflowschritt"].isin(workflows))
    ]

    found = (failed[
        (failed["Kommentar"].str.contains(
            '|'.join(abort_codes), na=False))])

    extension = ("_Fehlfahrten_eindeutig")
    write_results(found, extension)

    not_found = (failed[
        ~(failed["Kommentar"].str.contains(
            '|'.join(abort_codes), na=False))])

    extension = ("_Fehlfahrten_prüfen")
    write_results(not_found, extension)

    results["Fehlfahrten"] = len(found)
    results["Fehlfahrten prüfen"] = len(not_found)


def same_place():
    # avoid panda warning
    pd.options.mode.chained_assignment = None

    finished = df[(df["Workflowschritt"].isin(workflows))]

    finished["tmp_same"] = finished["Strasse"].str.cat(finished["HausNr"])
    even = finished[finished.duplicated(subset=["tmp_same"], keep=False)]
    even = even.loc[:, even.columns != "tmp_same"]

    extension = ("_gleiche_Abnahmestelle")
    write_results(even, extension)

    results["Einzelne Abnahmestelle"] = (len(finished) - len(even))
    results["Gleiche Abnahmestelle"] = len(even)


def meter_size():
    m_size = pd.DataFrame(
        df.loc[(df["Workflowschritt"].isin(workflows)) &
               (df["ZählerNr"].notnull())], columns=["Zählergröße"])
    results.update(
        m_size["Zählergröße"].value_counts().sort_index().to_dict())


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
    writer = pd.ExcelWriter(path + "/" + filename + "_Auswerrtung" + ".xlsx",
                            engine='xlsxwriter')
    data.to_excel(writer, sheet_name=extension, index=False)
    writer.save()


if __name__ == "__main__":
    # Merge customer lists
    customers = cus_0 + cus_1 + cus_2 + cus_3
    window = GUI()
    window.widgets()
    window.mainloop()
