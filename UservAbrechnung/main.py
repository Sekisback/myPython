from tkinter import *
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import customtkinter
import pandas as pd

customer = ["--- Kunden wählen ---", "SKA", "ZVO", "STD", "SWQ", "BBS", "SWNO",
            "TOR"]

results = {}


class GUI(customtkinter.CTk):
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue")

    def __init__(self):
        super().__init__()
        self.geometry("500x700")
        self.resizable(0, 0)
        self.title("Abrechnungsanalyse")
        self.icon = tk.PhotoImage(file="image/abrechnung.png")
        self.wm_iconphoto(False, self.icon)

    def widgets(self):
        # Combobox
        self.combobox = customtkinter.CTkComboBox(
            self, width=200, fg_color="#395E9C", values=sorted(customer))
        self.combobox.place(x=20, y=20, width=260)

        # Button
        self.button = customtkinter.CTkButton(
            self, text="Export laden...", command=open_dialog)
        self.button.place(x=300, y=20, width=180)

        # Frame
        self.frame = Frame(self, borderwidth=0, bg="white", relief="ridge")
        self.frame.place(x=20, y=70, width=460, height=610)

        # Treeview
        self.treeview = ttk.Treeview(
            self, columns=("key", "value"), show="tree")
        self.treeview.place(x=30, y=90, width=440, height=570)
        # Treeview Columns
        self.treeview.column('#0', minwidth=10, width=10, stretch=NO)
        self.treeview.column("key", minwidth=230, width=230, stretch=NO)
        self.treeview.column("value", minwidth=200, width=200, stretch=NO)


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
        template()
    except Exception as e:
        messagebox.showerror("Woah!", f"Da ist ein Problem aufgetreten {e}")


def template():
    count_data()
    # TODO You are here
    selection = str(window.combobox.get())
    if selection.startswith(("SKA", "ZVO")):
        print("ZVO or SKA")
    else:
        print("Other AG")

    show_results()


def count_data():
    # Make first Row empty
    results[""] = ""
    results["Anzahl Datensätze"] = (len(df))


def show_results():
    for index, (key, value) in enumerate(results.items()):
        window.treeview.insert(
            "", tk.END, iid=index, text="", values=(key, value))


if __name__ == "__main__":
    window = GUI()
    window.widgets()
    window.mainloop()
