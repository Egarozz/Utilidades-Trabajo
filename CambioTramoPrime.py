import pandas as pd
import csv
pd.options.mode.chained_assignment = None  # default='warn'

#Colocar el rango en una columna nueva y si se tiene 2 rangos en un mismo item 
#Saldra error por lo que se puede utilizar el cambio de tramo como grupo para cubrir el error
def col_tramos(t1, t2, t3, t4, t5):
  valores = 0
  rango = ""
  if t1 != 0:
    rango = "0 a 6"
    valores += 1
  if t2 != 0:
    rango = "7 a 12"
    valores += 1
  if t3 != 0:
    rango = "13 a 18"
    valores += 1
  if t4 != 0:
    rango = "19 a 24"
    valores += 1
  if t5 != 0:
    rango = "mayor a 24"
    valores += 1
  if valores > 1:
    rango = "Error"
  return rango
#Limpiar la tabla y cambiar nombres una vez unidos los codigos de ambos meses para compararlos
def arreglar_merged(df: pd.DataFrame):
  df = df[["Linea_x","Item", "Nombre Item_x", "Total_x", "Rango_x", "Total_y", "Rango_y"]]
  df.columns = ["Linea", "Item", "Nombre Item", "Total Actual", "Rango Actual", "Total Anterior", "Rango Anterior"]
  return df
#Al haber items que cambian de codigo se utiliza un archivo csv con items y su respectivo cambio
#Lo que hace esta funcion es unificar los codigos en uno solo osea cambiar todo al nombre del item mas actual
#Para el caso de equipos en flota se cambia todos los equipos a este codigo
def cambiar_codigos(df: pd.DataFrame, cc: pd.DataFrame):
  col = []
  for c in df["Item"]:
    if c in cc["Item"].values:
      a = cc[cc["Item"]==c].values
      col.append(a[0][1]) #a[0][1] -> equivale al codigo a cambiar | a[0][0] -> equivale al codigo que se esta buscando osea al "Item"
    else:
      col.append(c)
  df["Item"] = col
  return df 

def get_cc_table(df: pd.DataFrame, cc: pd.DataFrame):
  col = []
  cambios = df
  for c in df["Item"]:
    if c in cc["Item"].values:
      a = cc[cc["Item"]==c].values
      col.append(a[0][1]) #a[0][1] -> equivale al codigo a cambiar | a[0][0] -> equivale al codigo que se esta buscando osea al "Item"
    else:
      col.append(" ")
  cambios["Nuevo Codigo"] = col
  cambios = cambios[["Item", "Nuevo Codigo", "Rango", "Total"]]
  cambios = cambios[cambios["Nuevo Codigo"]!= " "]
  cambios = cambios.drop_duplicates(subset=['Nuevo Codigo'])
  return cambios 

def process_excel(ruta, mesanterior, mesactual):
    df = pd.read_excel(ruta, sheet_name="Trex Prime Real")
    cc = pd.read_csv("CambioCodigo.csv",delimiter=";")

    codigos = df
    codigos["Rango"] = codigos.apply(lambda x: col_tramos(x["de 0 a 6 m"], x["de 7 a 12"], x["de 13 a 18"], x["de 19 a 24"], x["mayor a 24"]), axis=1)
    cambio_codigos = get_cc_table(codigos,cc)

    df = cambiar_codigos(df, cc)

    #Se rellenan los campos vacios con 0 y aquellos que tengan un espacio tambien
    #Tener cuidado ya que puede haber mas espacios en una celda por lo que terminara saliendo en los errores
    df["de 0 a 6 m"] = df["de 0 a 6 m"].fillna(0)
    df.loc[df["de 0 a 6 m"]==" ", "de 0 a 6 m"] = 0
    df["de 7 a 12"] = df["de 7 a 12"].fillna(0)
    df.loc[df["de 7 a 12"]==" ", "de 7 a 12"] = 0
    df["de 13 a 18"] = df["de 13 a 18"].fillna(0)
    df.loc[df["de 13 a 18"]==" ", "de 13 a 18"] = 0
    df["de 19 a 24"] = df["de 19 a 24"].fillna(0)
    df.loc[df["de 19 a 24"]==" ", "de 19 a 24"] = 0
    df["mayor a 24"] = df["mayor a 24"].fillna(0)
    df.loc[df["mayor a 24"]==" ", "mayor a 24"] = 0

    #Se cambia el item como string, pero para evitar problemas es mejor cambiar toda la columna del drive como string
    #Ya que aunque se cambie el tipo a string igual surgen problemas
    df["Item"] = df["Item"].astype(str)

    #Al haber mas columnas a la derecha del drive como apoyo, se corta el dataframe solo a los valores que se necesitan
    df = df.iloc[:,:12]

    #Se filtra la tabla a los meses que se necesitan, en el caso nuestro para la prueba Mayo y Junio, ademas, se quita la linea
    #Otros y transito ya que transito no se utiliza para el calculo y otros se tiene que resolver con el script para resolver tramos en grupos
    junio = df[df["Mes"]==mesactual]
    junio = junio[junio["Linea"]!="Otros"]
    junio = junio[junio["Linea"] != "Transito"]
    mayo = df[df["Mes"]==mesanterior]
    mayo = mayo[mayo["Linea"]!="Otros"]
    mayo = mayo[mayo["Linea"]!="Transito"]

    #Se colocan los rangos en una nueva columna rango
    junio["Rango"] = junio.apply(lambda x: col_tramos(x["de 0 a 6 m"], x["de 7 a 12"], x["de 13 a 18"], x["de 19 a 24"], x["mayor a 24"]), axis=1)
    mayo["Rango"] = mayo.apply(lambda x: col_tramos(x["de 0 a 6 m"], x["de 7 a 12"], x["de 13 a 18"], x["de 19 a 24"], x["mayor a 24"]), axis=1)

    #Se filtran las entradas y salidas de equipos
    #Entradas: todo item que este en el mes actual pero no este en el mes anterior
    #Salidas: todo item que este en el mes anterior pero no este en el mes actual
    entradas = junio[~junio["Item"].isin(mayo["Item"].values)]
    salidas = mayo[~mayo["Item"].isin(junio["Item"].values)]

    merged = junio.merge(mayo, how="left", on="Item")

    #Merged not null para eliminar los que no se encuentran en el mes anterior debido a que es una entrada
    cambios = merged[(merged["Rango_x"] != merged["Rango_y"]) & ~merged["Rango_y"].isnull()]
    cambios = cambios[cambios["Rango_x"] != "Error"]
    cambios = cambios[cambios["Rango_y"] != "Error"]
    cambios = arreglar_merged(cambios)
    #Se le añade una columna diferencia para hallar el tipo de cambio restandole el total de mes actual - el total de mes anterior
    #No se cuentan los errores ya que si suceden se estaria restando mas de un solo rango
    cambios["dif"] = cambios["Total Actual"] - cambios["Total Anterior"]

    iguales = merged[(merged["Rango_x"] == merged["Rango_y"]) & ~merged["Rango_y"].isnull()]
    iguales = iguales[iguales["Rango_x"] != "Error"]
    iguales = iguales[iguales["Rango_y"] != "Error"]
    iguales = arreglar_merged(iguales)
    iguales["dif"] = iguales["Total Actual"] - iguales["Total Anterior"]
    diferencias = iguales[iguales["dif"] != 0]

    errores = merged[(merged["Rango_x"] == "Error") | (merged["Rango_y"] == "Error") ]
    errores = errores[~errores["Rango_y"].isnull()]
    errores = arreglar_merged(errores)
    #Una vez recopilada la informacion se crea el diccionario para una vez asi crear el DataFrame final
    dic = {}

    #Tipo de cambio
    #Para esto se utiliza la diferencia de la tabla donde no hay cambio de tramo sumado a 
    #La diferencia de la tabla donde hay cambio de tramo, se puede quitar la diferencia del cambio de tramo, pero se tendria que colocar
    #En cambio de tramo ingreso el monto con el cambio en el monto
    tc1 = diferencias[diferencias["Rango Actual"] == "0 a 6"]["dif"].sum() + cambios[cambios["Rango Anterior"] == "0 a 6"]["dif"].sum()
    tc2 = diferencias[diferencias["Rango Actual"] == "7 a 12"]["dif"].sum() + cambios[cambios["Rango Anterior"] == "7 a 12"]["dif"].sum()
    tc3 = diferencias[diferencias["Rango Actual"] == "13 a 18"]["dif"].sum() + cambios[cambios["Rango Anterior"] == "13 a 18"]["dif"].sum()
    tc4 = diferencias[diferencias["Rango Actual"] == "19 a 24"]["dif"].sum() + cambios[cambios["Rango Anterior"] == "19 a 24"]["dif"].sum()
    tc5 = diferencias[diferencias["Rango Actual"] == "mayor a 24"]["dif"].sum() + cambios[cambios["Rango Anterior"] == "mayor a 24"]["dif"].sum()

    #Ingresos
    ing1 = entradas[entradas["Rango"] == "0 a 6"]["Total"].sum()
    ing2 = entradas[entradas["Rango"] == "7 a 12"]["Total"].sum()
    ing3 = entradas[entradas["Rango"] == "13 a 18"]["Total"].sum()
    ing4 = entradas[entradas["Rango"] == "19 a 24"]["Total"].sum()
    ing5 = entradas[entradas["Rango"] == "mayor a 24"]["Total"].sum()

    #Egresos
    eg1 = salidas[salidas["Rango"] == "0 a 6"]["Total"].sum()
    eg2 = salidas[salidas["Rango"] == "7 a 12"]["Total"].sum()
    eg3 = salidas[salidas["Rango"] == "13 a 18"]["Total"].sum()
    eg4 = salidas[salidas["Rango"] == "19 a 24"]["Total"].sum()
    eg5 = salidas[salidas["Rango"] == "mayor a 24"]["Total"].sum()

    #Cambio de tramo Ingresos -> No puede ingresar de 0 a 6 ya que no hay un rango menor
    #Se utiliza para ambos cambios de tramo (Ingreso y egreso) el total actual ya que la diferencias (Tipo de cambio)
    #Se esta sumando en la columna de tipo de cambio
    ct2i = cambios[cambios["Rango Actual"] == "7 a 12"]["Total Actual"].sum()
    ct3i = cambios[cambios["Rango Actual"] == "13 a 18"]["Total Actual"].sum()
    ct4i = cambios[cambios["Rango Actual"] == "19 a 24"]["Total Actual"].sum()
    ct5i = cambios[cambios["Rango Actual"] == "mayor a 24"]["Total Actual"].sum()

    #Cambio de tramo Egresos -> No puede salir de mayor a 24 ya que no hay un rango mayor
    ct1e = cambios[cambios["Rango Anterior"] == "0 a 6"]["Total Actual"].sum()
    ct2e = cambios[cambios["Rango Anterior"] == "7 a 12"]["Total Actual"].sum()
    ct3e = cambios[cambios["Rango Anterior"] == "13 a 18"]["Total Actual"].sum()
    ct4e = cambios[cambios["Rango Anterior"] == "19 a 24"]["Total Actual"].sum()

    dic["Ingresos"] = [round(ing1,3), round(ing2,3), round(ing3,3), round(ing4,4), round(ing5,5)]
    dic["Egresos"] = [round(eg1,3), round(eg2,3), round(eg3,3), round(eg4,3), round(eg5, 3)]
    dic["Camb. Tram. Ing."] = [0, round(ct2i,3), round(ct3i,3), round(ct4i,3), round(ct5i,3)]
    dic["Camb. Tram. Egr."] = [round(ct1e,3), round(ct2e,3), round(ct3e,3), round(ct4e,3), 0]
    dic["TC"] = [round(tc1,3), round(tc2,3), round(tc3,3), round(tc4,3), round(tc5,3)]
    output = pd.DataFrame(dic)
    return (output, entradas, salidas, diferencias, cambio_codigos, errores, cambios)