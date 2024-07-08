# Při převodu na exe je potřeba v nástroji auto-py-to-exe v Advanced vyplnit do --hidden-import  hodnotu xlrd

from tkinter import *
from tkinter import filedialog
import pandas as pd
import glob
import os

def select_folder():
    try:
        load_entry.delete(0,END)
        checkbutton.deselect()
        info_duplicity["text"] = ""
        source_path = filedialog.askdirectory(title='Výběr adresáře s .xlsx soubory')
        load_entry.insert(0, source_path)
        folder = load_entry.get()
        filenames = glob.glob(folder + "\*.xlsx")
        cnt = len(filenames)
        count_label["text"] = f"Počet nalezených .xlsx souborů: {cnt}\n\n"
        if cnt > 0:
            all_dfs = pd.DataFrame()
            for file in filenames:
                df = pd.read_excel(file, engine="openpyxl")
                all_dfs = pd.concat([all_dfs, df], ignore_index=True, sort=False)
            duplicate_rows = all_dfs[all_dfs.duplicated()]
            duplicity_count = len(duplicate_rows)
            count_label["text"] = f"Počet nalezených .xlsx souborů: {cnt}\n\nPočet nalezených duplicitních řádků: {duplicity_count}"
            if duplicity_count > 0:
                create_checkbutton()
            else:
                hide_checkbutton()
    except:
        count_label["text"] = 0

def create_checkbutton():
    checkbutton.pack()

def hide_checkbutton():
    checkbutton.pack_forget()

def count_files():
    try:
        count_label["text"] = ""
        folder = load_entry.get()
        filenames = glob.glob(folder + "\*.xlsx")
        cnt = len(filenames)
        count_label["text"] = f"Počet nalezených .xlsx souborů: {cnt}\n\n"
    except:
        count_label["text"] = 0

def merge_files():
    folder = load_entry.get()
    file_name = name_entry.get()
    filenames = glob.glob(folder + "\*.xlsx")
    cnt = len(filenames)
    if file_name:
        if cnt > 0:
            print(check.get())
            if check.get() == 0:
                all_dfs = pd.DataFrame()
                for file in filenames:
                    df = pd.read_excel(file, engine="openpyxl")
                    all_dfs = pd.concat([all_dfs, df], ignore_index=True, sort=False)
                all_dfs.to_excel(folder + f"\{file_name}.xlsx", index=None, engine="openpyxl")
                count_label["text"] = ""
                count_label["text"] = f"Soubory úspěšně sloučeny.\n\nDuplicity nebyly odebrány!\n\nNázev souboru: {file_name}.xlsx"
            else:
                all_dfs = pd.DataFrame()
                for file in filenames:
                    df = pd.read_excel(file, engine="openpyxl")
                    all_dfs = pd.concat([all_dfs, df], ignore_index=True, sort=False)
                all_dfs.drop_duplicates(keep='last', inplace=True)
                all_dfs.to_excel(folder + f"\{file_name}.xlsx", index=None, engine="openpyxl")
                count_label["text"] = ""
                count_label[
                    "text"] = f"Soubory úspěšně sloučeny.\n\nDuplicity odebrány!\n\nNázev souboru: {file_name}.xlsx"

        else:
            count_label["text"] = "Soubory nenalezeny\n\n"
    else:
        count_label["text"] = "Není zadán název výstupního souboru\n\n"

def open_folder():
    folder = load_entry.get()
    try:
        if folder:
            os.startfile(folder)
        else:
            count_label["text"] = "Žádná složka se soubory není vybraná.\n\n"
    except:
        count_label["text"] = "Žádná složka se soubory není vybraná.\n\n"

win = Tk()
win.title("XLSX spojovač")
win.minsize(600,160)
win.resizable(True, False)
main_font = ("Sans Serif",11)

# define frames
load_frame = Frame(win)
load_frame.pack()
file_name_frame = Frame(win)
file_name_frame.pack()
output_frame = Frame(win)
output_frame.pack()
duplicity_frame = Frame(win)
duplicity_frame.pack()
button_frame = Frame(win)
button_frame.pack()

# load
load_title = Label(load_frame, text="Cesta k souborům .xlsx", font=main_font)
load_title.grid(row=0, column=0, padx=5, pady=5)
load_entry = Entry(load_frame, width=50, font=main_font)
load_entry.grid(row=0, column=1, padx=5, pady=5)
browse_button = Button(load_frame, text="Browse", font=main_font, command=select_folder)
browse_button.grid(row=0, column=2, padx=5, pady=5)

#output file name
name_title = Label(file_name_frame, text="Název výstupního souboru", font=main_font)
name_title.grid(row=0, column=0, padx=5, pady=5)
name_entry = Entry(file_name_frame, width=30, font=main_font)
name_entry.grid(row=0, column=1, padx=5, pady=5)
extension_label = Label(file_name_frame, text=".xlsx", font=main_font)
extension_label.grid(row=0, column=2, padx=5, pady=5)

# output
count_label = Label(output_frame, text="\n\n", font=main_font)
count_label.grid(row=1, column=1, padx=5, pady=5)
info_duplicity = Label(output_frame, text = "", font=main_font)
info_duplicity.grid(row=2, column=1, padx=5, pady=5)

#duplicity
check = IntVar()
checkbutton = Checkbutton(duplicity_frame, text="Odebrat duplicity", onvalue=1, offvalue=0, variable=check, font=main_font)
checkbutton.pack_forget()

# buttons
submit_button = Button(button_frame, text = "Proveď sloučení", font=main_font, command=merge_files, bg="lightgray")
submit_button.grid(row=2, column=0, padx=5, pady=5, ipady=2, ipadx=3)
folder_button = Button(button_frame, text = "Otevřít adresář", font=main_font, command=open_folder, bg="lightgray")
folder_button.grid(row=2, column=1, padx=5, pady=5, ipady=2, ipadx=3)

win.mainloop()