"""
Codigo por Ethan Grané
26/08/2023
"""

import xlsxwriter
import requests
from bs4 import BeautifulSoup
import nltk
from rake_nltk import Rake
from nltk.corpus import stopwords

nltk.download('stopwords')  # Descargar el conjunto de stop words

# Abre el archivo en modo "a" (añadir) si ya existe, o en modo "w" (crear) si no existe
workbook = xlsxwriter.Workbook("null.xlsx")  # Crea un archivo de Excel llamado "null.xlsx", mas tared se cambiara el nombre
sheetData = workbook.add_worksheet("Data sheet")  # Agrega una hoja llamada "Data sheet"
sheetGraph = workbook.add_worksheet("Graph sheet")  # Agrega una hoja llamada "Graph sheet"
sheetGraphData = workbook.add_worksheet("Graph sheet data (readonly)")  # Agrega una hoja llamada "Graph sheet data (readonly)"

# Variables de Formato
cellFormat_Title = workbook.add_format()
cellFormat_Title.set_bold()
cellFormat_Title.font_size = 14

#Hace el request a la pagina web
url = "https://youtube.fandom.com/es/wiki/Mangelrogel"
response = requests.get(url)
soup = BeautifulSoup(response.text, features="html.parser")

_test = " "

### Funciones
"""
Recoje el dominio del Url
"""
def GetDomain(url):
    domain = ""
    splitDoubleSlash = url.split("//")
    splitSlash = splitDoubleSlash[1].split("/")
    domain = splitSlash[0]
    return domain
"""
Elimina los espacios extra de un string
"""
def eliminar_espacios(texto):
    """
    Elimina espacios en blanco, tabulaciones y saltos de línea del texto.
    """
    texto_sin_espacios = ""
    for caracter in texto:
        if caracter != "\t" and caracter != "\n":
            texto_sin_espacios += caracter
    return texto_sin_espacios
"""
Escribe los datos en el archivo de Excel y agrega una fórmula para contar los elementos en la hoja de gráficos.
"""
def WriteExcel(Header, Column, List):
    #Header: El titulo de la columna
    #Column: La columna donde se escribira
    #List: Lista de textos que se escribiran (str array)

    sheetData.write(0, Column, Header, cellFormat_Title)  # Escribe el encabezado en la primera fila de la columna
    asciiColumn = chr(65 + Column)  # Convierte el índice de columna a letra ASCII (A, B, C, ...)

    rowCount = 2  # Comenzar desde la segunda fila para escribir datos
    sheetData.set_column(Column, Column, 35)  # Establece el ancho de columna para los datos
    for item in List:
        text = eliminar_espacios(item)

        sheetData.write(asciiColumn + str(rowCount), text)  # Escribe el texto en la celda correspondiente
        sheetGraphData.write(Column, 0, f"=COUNTA('Data sheet'!{asciiColumn}2:{asciiColumn}100)")  # Fórmula para contar los elementos en la hoja de gráficos

        rowCount += 1
"""
Comprueba si contiene un Substring dentro de un string
"""
def ContainsSubstring(link, substring):
    contains = False

    splittedText = str(link).split("/")
    for item in splittedText:
        if(item == substring):
            contains = True

    return contains
"""
"""
def SetupHeaders():
    #Recoje los Headers
    headerColumnCount = 6 + 1
    while headerColumnCount > 1:
        headerList = soup.find_all("h" + str(headerColumnCount))
        headerListString = []
        for item in headerList:
            headerListString.append(item.text)
        WriteExcel("H" + str(headerColumnCount) + f" ({len(headerList)})", headerColumnCount - 1, headerListString)  # Llama a la función para escribir los datos en Excel
        headerColumnCount -= 1

def SetupLinks():
    #Recoje todos los links
    linksList = soup.find_all("a")
    internalLinks = []              #Guardara los links internos
    externalLinks = []              #Guardara los links externos
    #Comprueba los links internos y externos
    for link in linksList:
        currentLink = link.get("href")
        if(currentLink != None):
            if(ContainsSubstring(currentLink,GetDomain(url))):
                internalLinks.append(str(currentLink))
            else:
                externalLinks.append(str(currentLink))

    internalLinks.sort()
    externalLinks.sort()

    #Write in excel internal links
    WriteExcel("Interal Links" + f" ({len(internalLinks)})",7,internalLinks)
    WriteExcel("External Links" + f" ({len(externalLinks)})",8,externalLinks)

def SetupImages():
    imagesObject = soup.find_all("img")
    imagesSrc = []
    imagesAlt = []
    for item in imagesObject:
        src = item.get("src")

        if(src == None):
            continue

        alt = item.get("alt")
        if(alt == ""):
            alt = "Null"

        imagesSrc.append(src)
        imagesAlt.append(alt)
    
    WriteExcel(f"Images Src ({len(imagesSrc)})",9,imagesSrc)
    WriteExcel(f"Images Alt ({len(imagesAlt)})",10,imagesAlt)
    
def KeywordAnalisis():
    stopWords = set(stopwords.words('spanish'))

    body = soup.find("body").text
    r = Rake()

    r.extract_keywords_from_text(body)
    density = []
    word_freq_distribution = r.get_word_frequency_distribution()
    values = list(word_freq_distribution.values())
    keys = list(word_freq_distribution.keys())

    #Si el key es un stopword se elimina el key X y el value X
    i = 0
    for key in keys:
        if(str(key) not in stopWords and int(values[i]) > 3 and len(str(key)) != 1):
            item = [values[i],keys[i]]
            density.append(item)
        i += 1
    density.sort(reverse=True)
    densityString = [(str(densityStr[1]).upper() + f" ({densityStr[0]}t)") for densityStr in density ]

    WriteExcel("Keywords Density",11,densityString)

def main():
    dominio = GetDomain(url)

    SetupHeaders()
    SetupLinks()
    SetupImages()
    KeywordAnalisis()

    #Recoje el title
    workbook.filename = f"{dominio}.xlsx"
    workbook.close()  # Cierra el archivo de Excel al finalizar

main()