import pandas as pd
import csv
import os
# ESFUERZO COMERCIAL, TRAMOS PRIME Y REPUESTOS

# Script para poder realizar el procesado de inventarios de la data que se extrae del sistema spring de repuestos y prime, ademas 
# De poder calcular el esfuerzo comercial de repuestsos

# Se coge la data en crudo salida del sistema y se procesa para obtener una tabla con los rangos identificados y 
# Eliminando aquellos que no tienen monto listo para colocar al drive en el cierre del mes
# El objetivo es agilizar el proceso como tambien eliminar manualidades y perdida de tiempo buscando en otros exceles la formula

# Este script viene con la funcion para procesar los tramos de inventario prime e inventario de repuestos ya que ambos se
# Manejan de distina forma

# Install xlrd para poder trabajar con archivos xls, seria bueno añadir la opcion de trabajar tambien con archivos xlsx 
# Para evitar errores al leer los archivos xls
def procesar_inv_prime(ruta):
    split = os.path.splitext(ruta)
    if(split[1] == ".xls"):
        df = pd.read_excel(ruta, engine='xlrd')
        return procesar_tramos(df)
    if(split[1] == ".xlsx"):
        df = pd.read_excel(ruta)
        return procesar_tramos(df)
def procesar_inv_repuestos(ruta):
    split = os.path.splitext(ruta)
    if(split[1] == ".xls"):
        df = pd.read_excel(ruta, engine='xlrd')
        df = procesar_tramos(df)
        return filtrar_repuestos(df)
    if(split[1] == ".xlsx"):
        df = pd.read_excel(ruta)
        df = procesar_tramos(df)
        return filtrar_repuestos(df)
def procesar_esfuerzo_comercial(anterior, actual):
    split1 = os.path.splitext(anterior)
    split2 = os.path.splitext(actual)
    if(split1[1] == ".xls"):
        df1 = pd.read_excel(anterior, engine='xlrd')
    if(split1[1] == ".xlsx"):
        df1 = pd.read_excel(anterior)

    if(split2[1] == ".xls"):
        df2 = pd.read_excel(actual, engine='xlrd')
    if(split2[1] == ".xlsx"):
        df2 = pd.read_excel(actual)
    
    return get_esfuerzo_comercial(df1, df2)

# Lo que hace esta funcion es eliminar la columna de stock actual, convertir a numeros y colocar los rangos sumando las 
# columnas correspondientes, para luego hallar el total sumando los rangos y filtrando con aquellos cuyo total sea > 0
def procesar_tramos(df):
    df = df.drop(["stockactual"], axis=1)
    df.iloc[:,7:].apply(pd.to_numeric)
    df["mayor a 25"] = df.iloc[:,6:15].sum(axis=1)
    df["de 19 a 24"] = df.iloc[:,15:21].sum(axis=1)
    df["de 13 a 18"] = df.iloc[:,21:27].sum(axis=1)
    df["de 7 a 12"] = df.iloc[:,27:33].sum(axis=1)
    df["de 0 a 6 m"] = df.iloc[:,33:39].sum(axis=1)
    df["Total"] = df["de 0 a 6 m"] + df["de 7 a 12"] + df["de 13 a 18"] + df["de 19 a 24"] + df["mayor a 25"]
    df = df[["familia_name", "item", "item_name", "de 0 a 6 m", "de 7 a 12", "de 13 a 18", "de 19 a 24", "mayor a 25", "Total"]]
    df = df[df["Total"]>0]
    return df

# Para trabajar con repuestos es diferente ya que se tienen que agrupar, para ello se cuenta con un archivo CSV como base de datos
# Para poder agrupar. Estos datos se colocan en un diccionario. Se itera por cada fila de la tabla de repuestos y dependiendo de si
# Esta en el diccionario se coloca en un array ese valor, para luego colocar una nueva columna en la tabla con esos valores 
# Como ultimo paso se agrupan por esa nueva columna
def filtrar_repuestos(df):
    diccionario = {}
    with open('Filtros.csv', mode='r') as infile:
        reader = csv.reader(infile)
        diccionario = dict((rows[0],rows[1]) for rows in reader)
    valores = []
    for row in df["familia_name"].values:
        valores.append(diccionario[row])
    df["Linea2"] = valores
    df = df.groupby(["Linea2"]).sum().reset_index()
    df = df.rename(columns = {'Linea2':'Linea'})
    return df

# Esta funcion procesa los tramos de un archivo excel con los datos semestrales los cuales se van a agrupar en años
# Esto con el objetivo de poder hallar el esfuerzo comercial
def procesar_tramos_semestral(df: pd.DataFrame):
    df = df.drop(["stockactual"], axis=1)
        
    df["2016"] = df.iloc[:, 26:28].sum(axis=1)
    df["2017"] = df.iloc[:, 28:30].sum(axis=1)
    df["2018"] = df.iloc[:, 30:32].sum(axis=1)
    df["2019"] = df.iloc[:, 32:34].sum(axis=1)
    df["2020"] = df.iloc[:, 34:36].sum(axis=1)
    df["2021"] = df.iloc[:, 36:38].sum(axis=1)
    df["2022"] = df.iloc[:,38]

    df["Total"] = df["2016"] +df["2017"] +df["2018"] +df["2019"] +df["2020"] +df["2021"] +df["2022"]
    df = df[df["Total"]>0]
    return df

# df_mant -> excel mes anterior
# df_mact -> excel mes actual
# Para hallar el esfuerzo comercial se procesa por los años y luego se halla el total que se obtiene al sumar los años
# 2016 + 2017 + 2018 + 2019 + 2020 + 2021 -> para el esfuerzo comercial no se concidera el año 2022
# se hace igual para cada mes y se restan ambos totales para hallar el esfuerzo comerical 
def get_esfuerzo_comercial(df_mant: pd.DataFrame, df_mact: pd.DataFrame):
    df_mant = procesar_tramos_semestral(df_mant)
    df_mact = procesar_tramos_semestral(df_mact)

    df_mant["Total"] = df_mant["2016"] +df_mant["2017"] +df_mant["2018"] +df_mant["2019"] +df_mant["2020"] +df_mant["2021"]
    total_mant = df_mant["Total"].sum()

    df_mact["Total"] = df_mact["2016"] +df_mact["2017"] +df_mact["2018"] +df_mact["2019"] +df_mact["2020"] +df_mact["2021"]
    total_mact = df_mact["Total"].sum()

    esfuerzo_comercial = round(total_mant - total_mact,3)
    return esfuerzo_comercial
