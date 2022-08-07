import tkinter as tk
from tkinter import CENTER, NO, StringVar, ttk
from tkinter import filedialog

from numpy import pad, save
import Inventarios
import Tramos
import pandas as pd
import CambioTramoPrime
import os

app = tk.Tk()
ruta_mesan = StringVar()
ruta_mesac = StringVar()
df_tramos = pd.DataFrame()
df_errores = pd.DataFrame()
df_cambio = pd.DataFrame()

def browse_files(var: StringVar):
    extension = (("Excel Files","*.xls*"),)
    filename = filedialog.askopenfilename(title = "Selecciona un excel", filetypes = extension)
    print(os.path.abspath(os.getcwd()))
    var.set(filename)
    
def save_excel():
    extension = (("Excel Files","*.xlsx"),)
    filename = filedialog.asksaveasfilename(title="Guardar excel", filetypes=extension, defaultextension="xlsx")
    return filename
def save_df(df: pd.DataFrame):
    extension = (("Excel Files","*.xlsx"),)
    filename = filedialog.asksaveasfilename(title="Guardar excel", filetypes=extension, defaultextension="xlsx")
    df.to_excel(filename)
def copiar_clipboard(app, str):
    app.clipboard_clear()
    app.clipboard_append(str)

def process_inv_mensual(ruta, tipo):
    path = save_excel()
    if tipo == "Prime":
        df = Inventarios.procesar_inv_prime(ruta)
    if tipo == "Repuestos":
        df = Inventarios.procesar_inv_repuestos(ruta)
    df.to_excel(path)
def process_ec(anterior, actual, var: StringVar):
    ec = Inventarios.procesar_esfuerzo_comercial(anterior, actual)
    var.set(ec)
def process_ctg(ruta):
    path = save_excel()
    df = Tramos.processExcel(ruta)
    df.to_excel(path)
def process_ctp(mesanterior, mesactual, entradas:ttk.Treeview, salidas:ttk.Treeview, errores:ttk.Treeview, cambiocodigo:ttk.Treeview, diferencias:ttk.Treeview, cambiotramo: ttk.Treeview, output:ttk.Treeview):
    extension = (("Excel Files","*.xls*"),)
    ruta = filedialog.askopenfilename(title = "Selecciona un excel", filetypes = extension)

    procesado = CambioTramoPrime.process_excel(ruta, mesanterior, mesactual)
    ent = procesado[1]
    ent = ent[["Item", "Rango", "Total"]]
    for i in range(len(ent["Item"].values)):
        row = (ent.iloc[i][0], ent.iloc[i][1], str(round(ent.iloc[i][2],2)))
        entradas.insert("", tk.END, values=row)

    sal = procesado[2]
    sal = sal[["Item", "Rango", "Total"]]
    for i in range(len(sal["Item"].values)):
        row = (sal.iloc[i][0], sal.iloc[i][1], str(round(sal.iloc[i][2],2)))
        salidas.insert("", tk.END, values=row)

    dif = procesado[3]
    dif = dif[["Item", "Rango Actual", "Total Actual", "dif"]]
    for i in range(len(dif["Item"].values)):
        row = (dif.iloc[i][0], dif.iloc[i][1], str(round(dif.iloc[i][2],2)), str(round(dif.iloc[i][3],2)))
        diferencias.insert("", tk.END, values=row)

    cc = procesado[4]
    cc = cc[["Item", "Nuevo Codigo", "Rango", "Total"]]
    for i in range(len(cc["Item"].values)):
        row = (cc.iloc[i][0], cc.iloc[i][1], cc.iloc[i][2], str(round(cc.iloc[i][3],2)))
        cambiocodigo.insert("", tk.END, values=row)
    
    global df_cambio
    df_cambio = procesado[6]
    ct = procesado[6]
    ct = ct[["Item", "Total Actual", "Rango Anterior", "Rango Actual"]]
    for i in range(len(ct["Item"].values)):
        row = (ct.iloc[i][0], str(round(ct.iloc[i][1],2)), ct.iloc[i][2], ct.iloc[i][3])
        cambiotramo.insert("", tk.END, values=row)
    
    err = procesado[5]
    global df_errores
    df_errores = procesado[7]
    err = err[["Item", "Total Actual"]]
    for i in range(len(err["Item"].values)):
        row = (err.iloc[i][0], str(round(err.iloc[i][1],2)))
        errores.insert("", tk.END, values=row)
    
    tr = procesado[0]
    global df_tramos
    df_tramos = tr
    tr = tr[["Ingresos", "Egresos", "Camb. Tram. Ing.", "Camb. Tram. Egr.", "TC"]]
    for i in range(len(tr["Ingresos"].values)):
        row = (str(round(tr.iloc[i][0],2)), str(round(tr.iloc[i][1],2)), str(round(tr.iloc[i][2],2)), str(round(tr.iloc[i][3],2)), str(round(tr.iloc[i][4],2)))
        output.insert("", tk.END, values=row)
    
def open_inv(root):
    window = tk.Toplevel(root)
    window.title("Inventario Mensual")
    ruta_inv = StringVar()
    opciones = StringVar()
    op = ["Prime", "Repuestos"]
    opciones.set(op[0])

    ttk.Label(window, text="Procesar Inventario Mensual", font=("Arial",10)).grid(row=0, column=0, columnspan=3)
    ttk.Label(window, text="Instrucciones", font=("Arial",10, "bold")).grid(row=1, column=0, columnspan=3)
    ttk.Label(window, text="Cargar el archivo .xls o xlsx salido del spring mensual y en dólares", font=("Arial",10), wraplength=150, justify="left").grid(row=2, column=0, columnspan=3,pady=5)
    ttk.Label(window, text="Archivo:", font=("Arial",10)).grid(row=3, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv)
    entry.config(state=tk.DISABLED)
    entry.grid(row=4, column=1,padx=5)
    
    ttk.Label(window, text="Tipo:", font=("Arial",10)).grid(row=5, column=0, columnspan=3, pady=5)
    tk.OptionMenu(window, opciones, *op).grid(row=6, column=0, columnspan=3)

    ttk.Label(window, text=" ", font=("Arial",10)).grid(row=7, column=0, columnspan=3)
    ttk.Button(window, text="Procesar", width=10, command=lambda: process_inv_mensual(ruta_inv.get(),opciones.get())).grid(row=8, column=0, columnspan=3,pady=10)
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

    ttk.Label(window, text="Procesar Cambio de tramo grupal", font=("Arial",10)).grid(row=0, column=0, columnspan=3)
    ttk.Label(window, text="Instrucciones", font=("Arial",10, "bold")).grid(row=1, column=0, columnspan=3)
    ttk.Label(window, text="Cargar el archivo .xlsx con dos pestañas \"Inicio\" y \"Fin\" ambos sin encabezado y con el mismo # de filas", font=("Arial",10), wraplength=150, justify="left").grid(row=2, column=0, columnspan=3,pady=5)
    ttk.Label(window, text="Archivo:", font=("Arial",10)).grid(row=3, column=0, columnspan=3)
    entry = ttk.Entry(window, width=23, textvariable=ruta_inv)
    entry.config(state=tk.DISABLED)
    entry.grid(row=4, column=1,padx=5)
    ttk.Label(window, text=" ", font=("Arial",10)).grid(row=5, column=0, columnspan=3)
    ttk.Button(window, text="Procesar", width=10, command=lambda: process_ctg(ruta_inv.get())).grid(row=6, column=0, columnspan=3,pady=10)
    ttk.Button(window, text="...", width=4, command=lambda: browse_files(ruta_inv)).grid(row=4, column=2)
    window.geometry("200x300") 
    window.grab_set()
def open_ctp(root):
    window = tk.Toplevel(root)
    window.columnconfigure(0, weight=2)
    window.columnconfigure(1, weight=2)
    window.columnconfigure(2, weight=2)
    window.columnconfigure(3, weight=2)
    window.columnconfigure(4, weight=2)
    window.columnconfigure(5, weight=2)
    window.columnconfigure(6, weight=2)
    window.columnconfigure(7, weight=2)
    window.columnconfigure(8, weight=2)
    window.columnconfigure(9, weight=2)
    window.columnconfigure(10, weight=2)
    window.columnconfigure(11, weight=2)
    window.columnconfigure(12, weight=2)
    window.columnconfigure(13, weight=2)



    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    op_ant = StringVar()
    op_ant.set(meses[0])
    op_act = StringVar()
    op_act.set(meses[0])

    ruta_tci = StringVar()

    ttk.Label(window, text="Procesar Cambio de tramo prime", font=("Arial",10)).grid(row=0, columnspan=13)
    
    ttk.Label(window, text="Instrucciones: Cargar el archivo .xlsx TCI del drive", font=("Arial",10, "bold")).grid(row=1, column=1, columnspan=13)

    ttk.Label(window, text="Mes Anterior:", font=("Arial",10, "bold")).grid(row=2, column=4,columnspan=1)
    ttk.OptionMenu(window, op_ant, *meses).grid(row=2, column=5)

    ttk.Label(window, text="Mes Actual:", font=("Arial",10, "bold")).grid(row=2, column=6,columnspan=1)
    ttk.OptionMenu(window, op_act, *meses).grid(row=2, column=7)
    
    ttk.Label(window, text="Archivo:", font=("Arial",10, "bold")).grid(row=3, column=4)
    ttk.Entry(window, width=23, textvariable=ruta_tci, state=tk.DISABLED).grid(row=3, column=5,columnspan=2)
    ttk.Button(window, text="...", width=4, command=lambda:process_ctp(op_ant.get(), op_act.get(), entradas, salidas, errores, cc, diferencias, ct, tramos)).grid(row=3, column=7)
    
    ttk.Label(window, text="Entradas:", font=("Arial",10, "bold")).grid(row=4, column=1,columnspan=3)
    ttk.Label(window, text="Salidas:", font=("Arial",10, "bold")).grid(row=4, column=4,columnspan=3)
    ttk.Label(window, text="Diferencias:", font=("Arial",10, "bold")).grid(row=4, column=7,columnspan=7)


    entradas = ttk.Treeview(window, height=8)
    entradas["columns"] = ("Item", "Rango", "Monto")
    entradas.column("#0", width=0, stretch=NO)
    entradas.column("Item", width=80, stretch=NO)
    entradas.column("Rango", width=80, stretch=NO)
    entradas.column("Monto", width=80, stretch=NO)

    entradas.heading("Item", text="Item", anchor=CENTER)
    entradas.heading("Rango", text="Rango", anchor=CENTER)
    entradas.heading("Monto", text="Monto", anchor=CENTER)
    
    entradas.grid(row=5, column=1, columnspan=3)

    salidas = ttk.Treeview(window, height=8)
    salidas["columns"] = ("Item", "Rango", "Monto")
    salidas.column("#0", width=0, stretch=NO)
    salidas.column("Item", width=80, stretch=NO)
    salidas.column("Rango", width=80, stretch=NO)
    salidas.column("Monto", width=80, stretch=NO)

    salidas.heading("Item", text="Item", anchor=CENTER)
    salidas.heading("Rango", text="Rango", anchor=CENTER)
    salidas.heading("Monto", text="Monto", anchor=CENTER)
    salidas.grid(row=5, column=4, columnspan=3)

    diferencias = ttk.Treeview(window, height=8)
    diferencias["columns"] = ("Item", "Rango", "Monto", "Diferencia")
    diferencias.column("#0", width=0, stretch=NO)
    diferencias.column("Item", width=80, stretch=NO)
    diferencias.column("Rango", width=80, stretch=NO)
    diferencias.column("Monto", width=80, stretch=NO)
    diferencias.column("Diferencia", width=80, stretch=NO)

    diferencias.heading("Item", text="Item", anchor=CENTER)
    diferencias.heading("Rango", text="Rango Act.", anchor=CENTER)
    diferencias.heading("Monto", text="Monto Act.", anchor=CENTER)
    diferencias.heading("Diferencia", text="Diferencia", anchor=CENTER)

    diferencias.grid(row=5, column=7, columnspan=7)  

    
    ttk.Label(window, text="Cambios de Código:", font=("Arial",10, "bold")).grid(row=7, column=1,columnspan=3)
    cc = ttk.Treeview(window, height=4)
    cc["columns"] = ("CAntes", "CAhora", "Rango", "Monto")
    cc.column("#0", width=0, stretch=NO)
    cc.column("Rango", width=80, stretch=NO)
    cc.column("Monto", width=80, stretch=NO)
    cc.column("CAntes", width=80, stretch=NO)
    cc.column("CAhora", width=80, stretch=NO)

    cc.heading("Rango", text="Rango", anchor=CENTER)
    cc.heading("Monto", text="Monto", anchor=CENTER)
    cc.heading("CAntes", text="Código Ant.", anchor=CENTER)
    cc.heading("CAhora", text="Código Act.", anchor=CENTER)
    
    cc.grid(row=8, column=1, columnspan=3,pady=10)

    ttk.Label(window, text="Cambios de tramo:", font=("Arial",10, "bold")).grid(row=9, column=1,columnspan=2)
    ttk.Button(window, text="Excel", command=lambda:save_df(df_cambio)).grid(row=9, column=2,columnspan=2,padx=10)
    ct = ttk.Treeview(window, height=4)
    ct["columns"] = ("Item", "Monto", "RangoAnt", "RangoAct")
    ct.column("#0", width=0, stretch=NO)
    ct.column("Item", width=80, stretch=NO)
    ct.column("Monto", width=80, stretch=NO)
    ct.column("RangoAnt", width=80, stretch=NO)
    ct.column("RangoAct", width=80, stretch=NO)

    ct.heading("Item", text="Item", anchor=CENTER)
    ct.heading("Monto", text="Monto", anchor=CENTER)
    ct.heading("RangoAnt", text="Rango Ant.", anchor=CENTER)
    ct.heading("RangoAct", text="Rango Act.", anchor=CENTER)
    ct.grid(row=10, column=1,columnspan=3)

    ttk.Label(window, text="Errores:", font=("Arial",10, "bold")).grid(row=7, column=4,pady=10,sticky="ne")
    ttk.Button(window, text="Excel", command=lambda:save_df(df_errores)).grid(row=7, column=5,columnspan=2,sticky="w",padx=10)
    errores = ttk.Treeview(window, height=8)
    errores["columns"] = ("Item", "Monto")
    errores.column("#0", width=0, stretch=NO)
    errores.column("Item", width=80, stretch=NO)
    errores.column("Monto", width=80, stretch=NO)

    errores.heading("Item", text="Item", anchor=CENTER)
    errores.heading("Monto", text="Monto", anchor=CENTER)
    errores.grid(row=8, column=4,columnspan=3,rowspan=3,padx=10,sticky="n")

    ttk.Label(window, text="Tramos:", font=("Arial",10, "bold")).grid(row=7, column=8)
    ttk.Button(window, text="Excel", command=lambda:save_df(df_tramos)).grid(row=7, column=9,columnspan=2,sticky="w",padx=10)
    tramos = ttk.Treeview(window, height=8)
    tramos["columns"] = ("Ing", "Egr", "CTI", "CTE", "Dif")
    tramos.column("#0", width=0, stretch=NO)
    tramos.column("Ing", width=80, stretch=NO)
    tramos.column("Egr", width=80, stretch=NO)
    tramos.column("CTI", width=80, stretch=NO)
    tramos.column("CTE", width=80, stretch=NO)
    tramos.column("Dif", width=80, stretch=NO)

    tramos.heading("Ing", text="Ingr.", anchor=CENTER)
    tramos.heading("Egr", text="Egre.", anchor=CENTER)
    tramos.heading("CTI", text="CT Ingr.", anchor=CENTER)
    tramos.heading("CTE", text="CT Egre.", anchor=CENTER)
    tramos.heading("Dif", text="Diff", anchor=CENTER)

    tramos.grid(row=8, column=7, rowspan=2, columnspan=5)  

    window.geometry("1000x700")
    window.grab_set()

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
