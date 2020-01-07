import openpyxl

## comparacion de color en la casilla
"""
def compara_naranja(casilla):

def compara_amarilla(casilla):
"""
def normalize(s):
    replacements = (
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
    )
    for a, b in replacements:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s

## Rango de los nombres de los archivos
def rango(ws,r_initial,r_final):
    # Param: ws = hoja con la que estamos trabajando
    # Value: Variable activa donde asignemos el valor dado por la funcion
        #    openpyxl.load_workbook
    # Param: r_initial = celda donde comienzan los nombres
    # Value: String con el formato cordenada de excel, ejem A2
    # Param: r_final = celda donde terminan los nombres
    # Value: Strign con el formato cordenada de excel, ejemplo C4
    
    cell_range = ws[r_initial:r_final]
    return cell_range

def delete_format(cell_range):
    for cell in cell_range:
        cell[0].value = cell[0].value.upper()
        cell[0].value = normalize(cell[0].value)
        cell[0].value = cell[0].value.replace(' ','')
        cell[0].value = cell[0].value.replace('1','')
        cell[0].value = cell[0].value.replace('2','')
        cell[0].value = cell[0].value.replace('3','')
        cell[0].value = cell[0].value.replace('4','')
        cell[0].value = cell[0].value.replace('5','')
        cell[0].value = cell[0].value.replace('6','')
        cell[0].value = cell[0].value.replace('7','')
        cell[0].value = cell[0].value.replace('8','')
        cell[0].value = cell[0].value.replace('9','')
        cell[0].value = cell[0].value.replace('0','')
        cell[0].value = cell[0].value.replace('.','')
    return cell_range
