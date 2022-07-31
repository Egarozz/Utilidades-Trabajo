import tkinter as tk
from tkinter import StringVar, ttk
from tkinter import filedialog
import Inventarios
import pandas as pd
import os

app = tk.Tk()
ruta_mesan = StringVar()
ruta_mesac = StringVar()

def browse_files(var: StringVar):
    extension = (("Excel Files","*.xls*"),)
    filename = filedialog.askopenfilename(title = "Selecciona un excel", filetypes = extension)
    print(os.path.abspath(os.getcwd()))
    var.set(filename)
def copiar_clipboard(app, str):
    app.clipboard_clear()
    app.clipboard_append(str)
def process_inv_mensual(ruta, nombre):
    df = Inventarios.procesar_inv_prime(ruta)
    path = os.path.abspath(os.getcwd()) + "\\" + nombre + ".xlsx"
    df.to_excel(path)  
def process_ec(anterior, actual, var: StringVar):
    ec = Inventarios.procesar_esfuerzo_comercial(anterior, actual)
    var.set(ec)

def open_inv(root):
    window = tk.Toplevel(root)
    window.title("Inventario Mensual")
    nombre_inv = StringVar()
    ruta_inv = StringVar()
    

    ttk.Label(window, text="Procesar Inventario Mensual", font=("Arial",10)).grid(row=0, column=0, columnspan=3)
    ttk.Label(window, text="Instrucciones", font=("Arial",10, "bold")).grid(row=1, column=0, columnspan=3)
    ttk.Label(window, text="Cargar el archivo .xls o xlsx salido del spring mensual y en dólares", font=("Arial",10), wraplength=150, justify="left").grid(row=2, column=0, columnspan=3,pady=5)
    ttk.Label(window, text="Archivo:", font=("Arial",10)).grid(row=3, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv)
    entry.config(state=tk.DISABLED)
    entry.grid(row=4, column=1,padx=5)
    ttk.Label(window, text=" ", font=("Arial",10)).grid(row=5, column=0, columnspan=3)
    ttk.Label(window, text="Nombre del archivo:", font=("Arial",10)).grid(row=6, column=0, columnspan=3, pady=5)
    ttk.Entry(window, width=23, textvariable=nombre_inv).grid(row=7, column=0, columnspan=3)
    ttk.Button(window, text="Procesar", width=10, command=lambda: process_inv_mensual(ruta_inv.get(),nombre_inv.get())).grid(row=8, column=0, columnspan=3,pady=10)
    ttk.Button(window, text="...", width=4, command=lambda: browse_files(ruta_inv)).grid(row=4, column=2)
    window.geometry("200x300")  
    window.grab_set()

def open_ec(root):
    window = tk.Toplevel(root)
    window.title("Esfuerzo Comercial")

    ruta_inv1 = StringVar()
    ruta_inv2 = StringVar()
    ec_var = StringVar()

    ttk.Label(window, text="Procesar Esfuerzo Comercial", font=("Arial",10)).grid(row=0, column=0, columnspan=3)
    ttk.Label(window, text="Instrucciones", font=("Arial",10, "bold")).grid(row=1, column=0, columnspan=3)
    ttk.Label(window, text="Cargar el archivo .xls o xlsx salido del spring semestral y en dólares del mes actual y anterior", font=("Arial",10), wraplength=150, justify="left").grid(row=2, column=0, columnspan=3,pady=5)
    ttk.Label(window, text="Archivo mes anterior:", font=("Arial",10)).grid(row=3, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv1)
    entry.config(state=tk.DISABLED)
    entry.grid(row=4, column=1,padx=5)
    ttk.Button(window, text="...", width=4,command=lambda: browse_files(ruta_inv1)).grid(row=4, column=2)

    ttk.Label(window, text="Archivo mes actual:", font=("Arial",10)).grid(row=5, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv2)
    entry.config(state=tk.DISABLED)
    entry.grid(row=6, column=1,padx=5)
    ttk.Button(window, text="...", width=4, command=lambda: browse_files(ruta_inv2)).grid(row=6, column=2)
    ttk.Button(window, text="Procesar", command=lambda: process_ec(ruta_inv1.get(), ruta_inv2.get(), ec_var)).grid(row=7, column=0, columnspan=3)
    
    ttk.Label(window, text=" ", font=("Arial",10)).grid(row=8, column=0, columnspan=3)
    ttk.Label(window, text="Esfuerzo Comercial:  ", font=("Arial",10)).grid(row=9, column=0, columnspan=3)
    ttk.Label(window, text="", font=("Arial",10), textvariable=ec_var).grid(row=10, column=0, columnspan=3)
    ttk.Button(window, text="Copiar", command=lambda:copiar_clipboard(root, ec_var.get())).grid(row=11, column=0, columnspan=3)
    
    window.geometry("200x350") 
    window.grab_set() 

def open_ctg(root):
    window = tk.Toplevel(root)
    window.title("Cambio de tramo grupo")

    ruta_inv = StringVar()
    nombre_inv = StringVar()
    ttk.Label(window, text="Procesar Cambio de tramo grupal", font=("Arial",10)).grid(row=0, column=0, columnspan=3)
    ttk.Label(window, text="Instrucciones", font=("Arial",10, "bold")).grid(row=1, column=0, columnspan=3)
    ttk.Label(window, text="Cargar el archivo .xlsx con dos pestañas \"inicio\" y \"fin\" ambos sin encabezado y con el mismo # de filas", font=("Arial",10), wraplength=150, justify="left").grid(row=2, column=0, columnspan=3,pady=5)
    ttk.Label(window, text="Archivo:", font=("Arial",10)).grid(row=3, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv)
    entry.config(state=tk.DISABLED)
    entry.grid(row=4, column=1,padx=5)
    ttk.Label(window, text=" ", font=("Arial",10)).grid(row=5, column=0, columnspan=3)
    ttk.Label(window, text="Nombre del archivo:", font=("Arial",10)).grid(row=6, column=0, columnspan=3, pady=5)
    nombre = ttk.Entry(window, width=23, textvariable=nombre_inv).grid(row=7, column=0, columnspan=3)
    ttk.Button(window, text="Procesar", width=10).grid(row=8, column=0, columnspan=3,pady=10)
    ttk.Button(window, text="...", width=4).grid(row=4, column=2)
    window.geometry("200x300") 

def open_ctp(root):
    window = tk.Toplevel(root)
    window.geometry("500x600") 

titulo = ttk.Label(app, text="Utilidades Trabajo", font=("Arial", 20)).pack(pady=2)
creditos = ttk.Label(app, text="Creado por: Emilio Garofolin   ver 1.0.0", font=("Arial", 9, "italic")).pack(pady=5)
btn_invp = ttk.Button(app, text="Inv. mensual", width=20, command=lambda: open_inv(app)).pack(pady=5)
btn_ec = ttk.Button(app, text="Esfuerzo comercial", width=20, command=lambda:open_ec(app)).pack(pady=5)
btn_tr = ttk.Button(app, text="Cambio tramo grupo", width=20, command=lambda:open_ctg(app)).pack(pady=5)
btn_tp = ttk.Button(app, text="Cambio tramo prime", width=20, command=lambda: open_ctp(app)).pack(pady=5)

app.geometry("250x350")
app.resizable(False, False)
app.title("Utilidades Trabajo")
app.mainloop()
