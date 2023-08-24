"""
Codigo por Ethan Grané
25/08/2023
"""

import xlsxwriter
import requests
from bs4 import BeautifulSoup

# Abre el archivo en modo "a" (añadir) si ya existe, o en modo "w" (crear) si no existe
workbook = xlsxwriter.Workbook("file.xlsx")  # Crea un archivo de Excel llamado "file.xlsx"
sheetData = workbook.add_worksheet("Data sheet")  # Agrega una hoja llamada "Data sheet"
sheetGraph = workbook.add_worksheet("Graph sheet")  # Agrega una hoja llamada "Graph sheet"
sheetGraphData = workbook.add_worksheet("Graph sheet data (readonly)")  # Agrega una hoja llamada "Graph sheet data (readonly)"

# Variables de Formato
cellFormat_Title = workbook.add_format()
cellFormat_Title.set_bold()
cellFormat_Title.font_size = 14

### Funciones
def eliminar_espacios(texto):
    """
    Elimina espacios en blanco, tabulaciones y saltos de línea del texto.
    """
    texto_sin_espacios = ""
    for caracter in texto:
        if caracter != "\t" and caracter != "\n":
            texto_sin_espacios += caracter
    return texto_sin_espacios

def WriteExcel(Header, Column, List):
    """
    Escribe los datos en el archivo de Excel y agrega una fórmula para contar los elementos en la hoja de gráficos.
    """
    sheetData.write(0, Column, Header + " Info", cellFormat_Title)  # Escribe el encabezado en la primera fila de la columna
    asciiColumn = chr(65 + Column)  # Convierte el índice de columna a letra ASCII (A, B, C, ...)

    rowCount = 2  # Comenzar desde la segunda fila para escribir datos
    sheetData.set_column(Column, Column, 25)  # Establece el ancho de columna para los datos
    for item in List:
        text = eliminar_espacios(item.text)

        sheetData.write(asciiColumn + str(rowCount), text)  # Escribe el texto en la celda correspondiente
        sheetGraphData.write(Column, 0, f"=COUNTA('Data sheet'!{asciiColumn}2:{asciiColumn}100)")  # Fórmula para contar los elementos en la hoja de gráficos

        rowCount += 1

url = "https://youtube.fandom.com/es/wiki/Mangelrogel"
response = requests.get(url)
soup = BeautifulSoup(response.text, features="html.parser")

headerCount = 5
while headerCount > 0:
    headerList = soup.find_all("h" + str(headerCount))
    WriteExcel("H" + str(headerCount), headerCount - 1, headerList)  # Llama a la función para escribir los datos en Excel
    headerCount -= 1

workbook.close()  # Cierra el archivo de Excel al finalizar
