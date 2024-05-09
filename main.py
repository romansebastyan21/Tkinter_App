import tkinter as tk
from tkinter import ttk
import openpyxl


def load_data():
    path = "D:\\Python_Course\\Tkinter Excel App\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)

    for i, name_col in enumerate(list_values[0]):
        treeview.heading(i, text=name_col)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def insert_row():
    name_1 = name_tata.get()
    name_2 = name_mama.get()
    name_3 = name_child.get()
    sex_1 = status_combobox.get()
    dob_1 = dob.get()
    birth_note_1 = birth_note.get()

    print(name_1, name_2, name_3, sex_1, dob_1, birth_note_1)

    path = "D:\\Python_Course\\Tkinter Excel App\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name_1, name_2, name_3, sex_1, dob_1, birth_note_1]
    sheet.append(row_values)
    workbook.save(path)

    treeview.insert('', tk.END, values=row_values)

    name_tata.delete(0, "end")
    name_tata.insert(0, "NUME PRENUME TATA")
    name_mama.delete(0, "end")
    name_mama.insert(0, "NUME PRENUME MAMA")
    name_child.delete(0, "end")
    name_child.insert(0, "NUME PRENUME COPIL")
    status_combobox.delete(0, "end")
    status_combobox.insert(0, "SEX")
    dob.delete(0, "end")
    dob.insert(0, "DATA NASTERII")
    birth_note.delete(0, "end")
    birth_note.insert(0, "NOTA LA NASTERE")


def toggle_mode():
    if mode_swich.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


root = tk.Tk()
treeview = ttk.Treeview(root)

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list = ["FATA", "BAIAT", "ALTCEVA"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="DATE NOU-NASCUTI")
widgets_frame.grid(row=0, column=0, padx=100, pady=10)

name_tata = ttk.Entry(widgets_frame)
name_tata.insert(0, "NUME PRENUME TATA")
name_tata.bind("<FocusIn>", lambda e: name_tata.delete('0', 'end'))
name_tata.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

name_mama = ttk.Entry(widgets_frame)
name_mama.insert(0, "NUME PRENUME MAMA")
name_mama.bind("<FocusIn>", lambda e: name_mama.delete('0', 'end'))
name_mama.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

name_child = ttk.Entry(widgets_frame)
name_child.insert(0, "NUME PRENUME COPIL")
name_child.bind("<FocusIn>", lambda e: name_child.delete('0', 'end'))
name_child.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.insert(0, "SEX")
status_combobox.grid(row=3, column=0, padx=5, pady=(0, 5), sticky="ew")

dob = ttk.Entry(widgets_frame)
dob.insert(0, "DATA NASTERII")
dob.bind("<FocusIn>", lambda e: dob.delete('0', 'end'))
dob.grid(row=4, column=0, padx=5, pady=(0, 5), sticky="ew")

birth_note = ttk.Spinbox(widgets_frame, from_=1, to=10)
birth_note.insert(0, "NOTA LA NASTERE")
birth_note.grid(row=5, column=0, padx=5, pady=(0, 5), sticky="ew")

button = ttk.Button(widgets_frame, text="ADAUGARE", command=insert_row)
button.grid(row=6, column=0, padx=5, pady=(0, 5), sticky="nsew")

mode_swich = ttk.Checkbutton(
    widgets_frame, text="MOD", style="Switch", command=toggle_mode)
mode_swich.grid(row=7, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
tree_Scrol = ttk.Scrollbar(treeFrame)
tree_Scrol.pack(side="right", fill="y")

cols = ("TATA", "MAMA", "COPIL", "SEX", "DATA NASTERII", "NOTA",)
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=tree_Scrol.set, columns=cols, height=30)
treeview.column("TATA", width=100)
treeview.column("MAMA", width=100)
treeview.column("COPIL", width=100)
treeview.column("SEX", width=50)
treeview.column("DATA NASTERII", width=100)
treeview.column("NOTA", width=100)
treeview.pack()
tree_Scrol.config(command=treeview.yview)
load_data()

root.mainloop()
