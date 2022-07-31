# Script para automatizar reconocimiento de que montos entraron, salieron y cuales cambiaron de tramo
# Los tramos son de (0 a 6 - 7 a 12 - 13 a 18 - 19 a 24 - mayor a 25)

# Este programa esta hecho con la finalidad de lograr velocidad y automatizacion a la hora de reconocer 
# cambio de tramos de elementos AGRUPADOS

import pandas
import tkinter

# Tabla que contiene la columna ingresos, egresos, ctingresos, ctegresos ----> ct: Cambio de tramo
# Cada columna cuenta con un array que debe tener 5 espacios (Rangos)
class Tabla:
    def __init__(self):
        self.invInicial = [0,0,0,0,0]
        self.ingresos = [0,0,0,0,0]
        self.egresos = [0,0,0,0,0]
        self.ctingresos = [0,0,0,0,0]
        self.ctegresos = [0,0,0,0,0]
        self.invFinal = [0,0,0,0,0]
    def getPandasTable(self):
        dictio = {"Rangos": ["0 a 6", "7 a 12", "13 a 18", "19 a 24", "mayor a 24"],
                "Inv. Inicial": self.invInicial,
                "Ingresos": self.ingresos,
                "Egresos": self.egresos,
                "CTIngresos": self.ctingresos,
                "CTEgresos": self.ctegresos,
                "Inv. Final": self.invFinal

                }
        df = pandas.DataFrame(dictio)
        return df
def sumarTablas(tabla1, tabla2):
    suma = Tabla()
    for i in range(5):
        suma.invInicial[i] = tabla1.invInicial[i] + tabla2.invInicial[i]
        suma.ingresos[i] = tabla1.ingresos[i] + tabla2.ingresos[i]
        suma.egresos[i] = tabla1.egresos[i] + tabla2.egresos[i]
        suma.ctingresos[i] = tabla1.ctingresos[i] + tabla2.ctingresos[i]
        suma.ctegresos[i] = tabla1.ctegresos[i] + tabla2.ctegresos[i]
        suma.invFinal[i] = tabla1.invFinal[i] + tabla2.invFinal[i]
    return suma
# Para tener la tabla resuelta se necesita la diferencia de los montos de ambos meses para asi poder identificar
# Cuanto se esta moviendo, ingresando o saliendo
# dif1: diferencia del rango 0 a 6
# dif2: diferencia del rango 7 a 12
# dif3: diferencia del rango 13 a 18
# dif4: diferencia del rango 19 a 24
# dif5: diferencia del rango mayor a 25

def getTabla(dif1, dif2, dif3, dif4, dif5):
    temp1 = dif1
    temp2 = dif2
    temp3 = dif3
    temp4 = dif4

    ingresos = [0,0,0,0,0]
    egresos = [0,0,0,0,0]
    ctingresos = [0,0,0,0,0]
    ctegresos = [0,0,0,0,0]

    if dif5 > 0:
        temp4 = dif4 + dif5
    if temp4 > 0:
        temp3 = temp4 + dif3
    if temp3 > 0:
        temp2 = temp3 + dif2
    if temp2 > 0:
        temp1 = temp2 + dif1
    
    if temp1 < 0:
        egresos[0] = round(temp1*-1,4)
    else:
        ingresos[0] = round(temp1,4)

    if temp2 < 0:
        egresos[1] = round(temp2*-1,4)
    else:
        ctegresos[0] = round(temp2,4)
        ctingresos[1] = round(temp2,4)
    
    if temp3 < 0:
        egresos[2] = round(temp3*-1,4)
    else:
        ctegresos[1] = round(temp3,4)
        ctingresos[2] = round(temp3,4)
    
    if temp4 < 0:
        egresos[3] = round(temp4*-1,4)
    else:
        ctegresos[2] = round(temp4,4)
        ctingresos[3] = round(temp4,4)
    
    if dif5 < 0:
        egresos[4] = round(dif5*-1,4)
    else:
        ctegresos[3] = round(dif5,4)
        ctingresos[4] = round(dif5,4)
    tabla = Tabla()
    tabla.ingresos = ingresos
    tabla.egresos = egresos
    tabla.ctingresos = ctingresos
    tabla.ctegresos = ctegresos
    return tabla

# Hallar la diferencia entre los montos de ambos meses
# Si es negativo es porque hubo ventas
# Si es positivo es porque hubo compras
# Si es positivo en un tramo distinto a 0 a 6 entonces hubo cambio de tramo
def getTramos(row1, row2):
    dif = [0,0,0,0,0]
    for i in range(5):
        dif[i] = row2[i] - row1[i]
    table = getTabla(dif[0], dif[1], dif[2], dif[3], dif[4])
    table.invInicial = row1
    table.invFinal = row2
    return table
# Procesar el excel
# El excel a leer no debe tener encabezados pero debe estar de la siguiente forma

# Item | 0 a 6 | 7 a 12 | 13 a 18 | 19 a 24 | mayor a 25
# ambas hojas deben tener las mismas filas y con nombre de item identicos para asi poder hacer la resta

# La hoja "Inicio" contiene las filas con los datos agrupados del mes anterior
# La hoja "Fin" contiene las filas con los datos agrupados del mes actual

def processExcel(excel):
    inicio = pandas.read_excel("Prueba.xlsx", sheet_name="Inicio", header=None)
    fin = pandas.read_excel("Prueba.xlsx", sheet_name="Fin", header=None)
    rows = inicio.shape[0]
    tabla = Tabla()
    
    # Iterar todas las filas
    for i in range(rows):
        # Nombre del item de la tabla inicio
        name = inicio.iloc[i][0]
        # Iterar otra vez todas las filas pero de la tabla fin
        for j in range(rows):
            # Si se haya la coincidencia se coge toda la fila de ambos meses y se procesa
            if name == fin.iloc[j][0]:
                row1 = [inicio.iloc[i][1], inicio.iloc[i][2], inicio.iloc[i][3], inicio.iloc[i][4], inicio.iloc[i][5]]
                row2 = [fin.iloc[j][1], fin.iloc[j][2], fin.iloc[j][3], fin.iloc[j][4], fin.iloc[j][5]]
                table = getTramos(row1,row2)
                tabla = sumarTablas(tabla, table)
    df = tabla.getPandasTable()
    df.to_excel("Output.xlsx", sheet_name="Output")
    
processExcel("Prueba.xlsx")