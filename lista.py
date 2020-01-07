import openpyxl
from funciones.funciones import rango,delete_format,normalize

doc = openpyxl.load_workbook('archivos/DOSIS_SEPTIEMBRE_2019.xlsx')
wsa = doc.active

#a = wsa['S4'].value
#a = a.splitlines()
#print(a)

a = rango(wsa,'S4','S11')
resul = delete_format(a)
for algo in resul:
    print(algo[0].value)
"""
[1] Convertir todo a mayusculas o minusculas
[1] Aliminar espacios y numeros
[1] Eliminar los ascentos
[0] Si es casilla naranja en un archivo ignorar
[0] Si una casilla amarilla ignorar en el otro archivo
[0] Saber separar todo bien por turnos.                                                                                                                                                                                                                                                                                                                                                                                                                 
"""