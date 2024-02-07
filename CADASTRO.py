import tkinter as tk
from tkinter import ttk

try:
    import openpyxl
except ImportError:

    import subprocess
    subprocess.call(['pip', 'install', 'openpyxl'])
    import openpyxl

def load_data():
    path = 'eleicao_cipa.xlsx'
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():
    nome = nome_entrada.get()
    idade = int(idade_spinbox.get())
    cipa = status_combobox.get()
    colaborador_status="Ativa" if a.get() else "Inativo"

    path = "eleicao_cipa.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [nome, idade, cipa, colaborador_status]
    sheet.append(row_values)
    workbook.save(path)

    treeview.insert("", tk.END, values=row_values)

    nome_entrada.delete(0, "end")
    nome_entrada.insert(0, "Nome")
    idade_spinbox.delete(0, "end")
    idade_spinbox.insert(0, "idade")
    status_combobox.set(combo_list[0])
    a.set(False)

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("clam")
    else:
        style.theme_use("default")

root = tk.Tk()

style = ttk.Style(root)
root.tk_setPalette(background="#ececec")
style.configure("Treeview", background="#ececec", fieldbackground="#ececec")
style.theme_use("default")

combo_list = ["Inscrito", "Não inscrito"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Inserir linha")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

nome_entrada = ttk.Entry(widgets_frame)
nome_entrada.insert(0, "Nome")
nome_entrada.bind("<FocusIn>", lambda e: nome_entrada.delete("0", "end"))
nome_entrada.grid(row=0, column=0, padx=5, pady=(0,5), sticky="ew")

idade_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
idade_spinbox.insert(0, "Idade")
idade_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=3, column=0, padx=0, pady=0)

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Ativo", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="Inserir", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20,10), pady=5, sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, padx=10)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("NOME", "IDADE", "CIPA", "COLABORADOR")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("NOME", width=100)
treeview.column("IDADE", width=50)
treeview.column("CIPA", width=100)
treeview.column("COLABORADOR", width=100)
treeview.pack()

treeScroll.config(command=treeview.yview)

load_data()

root.mainloop()

