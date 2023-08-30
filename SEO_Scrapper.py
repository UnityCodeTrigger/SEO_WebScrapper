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

import graphs as graph

nltk.download('stopwords')  # Descargar el conjunto de stop words

# Abre el archivo en modo "a" (añadir) si ya existe, o en modo "w" (crear) si no existe
workbook = xlsxwriter.Workbook("null.xlsx")  # Crea un archivo de Excel llamado "null.xlsx", mas tared se cambiara el nombre
sheetData = workbook.add_worksheet("Data sheet")  # Agrega una hoja llamada "Data sheet"

# Variables de Formato
cellFormat_Title = workbook.add_format()
cellFormat_Title.set_bold()
cellFormat_Title.font_size = 14

#Hace el request a la pagina web
url = "https://youtube.fandom.com/es/wiki/Mangelrogel"
url = url.lower()
response = requests.get(url)
soup = BeautifulSoup(response.text, features="html.parser")

#DataStorgae - Not a list
bodyDataList = []
headerDataList = []
headerDataString = ""

#Guarada la cantidad (int) de headers que se encuentran en la pagina
headerCount = []

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
Elimina los espacios extra de un string (no una lista)
"""
def eliminar_espacios(texto):
    """
    Elimina espacios en blanco, tabulaciones y saltos de línea del texto.
    """
    texto_sin_espacios = ""
    lastIsSpace = False

    for caracter in texto:
        if caracter != "\t" and caracter != "\n":
            texto_sin_espacios += caracter
            lastIsSpace = False
        elif lastIsSpace == False:
            texto_sin_espacios += " "
            lastIsSpace = True

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

        rowCount += 1
"""
Comprueba si contiene un Substring dentro de un string
"""
def GetContainsSubstring(link, substring):
    contains = False

    splittedText = str(link).split("/")
    for item in splittedText:
        if(item == substring):
            contains = True

    return contains

def GetCalculateDensity(listItem):
    r = Rake()
    stopWords = set(stopwords.words('spanish'))

    r.extract_keywords_from_text(listItem)
    density = []
    word_freq_distribution = r.get_word_frequency_distribution()
    values = list(word_freq_distribution.values())
    keys = list(word_freq_distribution.keys())

    #Si el key es un stopword se elimina el key X y el value X
    i = 0
    for key in keys:
        if(str(key) not in stopWords and len(str(key)) != 1):
            item = [values[i],keys[i]]
            density.append(item)
        i += 1
    density.sort(reverse=True)
    print(density[:10])
    densityString = [(str(densityStr[1]).upper() + f" ({densityStr[0]}t)") for densityStr in density ]

    return densityString

def GetCalculateDensityOrganicKeywords(header, body):
    stopWords = set(stopwords.words('spanish'))

    organicKeywords = [(0,"fasd")]
    importanciaHeader = 5
    importanciaURL = 10

    #Body density
    body_r = Rake()
    body_r.extract_keywords_from_text(body)
    body_WordFreqDistribution = body_r.get_word_frequency_distribution()
    body_values = list(body_WordFreqDistribution.values())
    body_keys = list(body_WordFreqDistribution.keys())
    body_Keywords = []

    #header density
    header_r = Rake()
    header_r.extract_keywords_from_text(header)
    header_WordFreqDistribution = header_r.get_word_frequency_distribution()
    header_values = list(header_WordFreqDistribution.values())
    header_keys = list(header_WordFreqDistribution.keys())

    #sumar las palabras que hay por orden de densidad
    allKeywords = []
    for i in range(len(header_WordFreqDistribution)):
        if(header_values[i] >= 3):
            if(header_keys[i] not in stopWords):
                allKeywords += [(header_values[i],header_keys[i])]
    
    #Comprueba si la keywords esta presente en los headers
    for i in range(len(allKeywords)):
        if(allKeywords[i][1] in body_keys):
            allKeywords[i] = (allKeywords[i][0] + importanciaHeader, allKeywords[i][1])
            
    for i in range(len(body_WordFreqDistribution)):
        if(body_values[i] >= 3):
            if(body_keys[i] not in stopWords and len(body_keys[i]) > 1):
                allKeywords += [(body_values[i], body_keys[i])]
    
    #Esta alguna de estas palabras dentro del url? dependiendo cual se le suma importancia
    for i in range(len(allKeywords)):
        if(url.__contains__(allKeywords[i][1])):
            allKeywords[i] = (allKeywords[i][0] + importanciaURL, allKeywords[i][1])

    #Ordena por importancia las keywords
    sorted_keywords = sorted(allKeywords, key=lambda x: x[0])
    sorted_keywords.reverse()
    sorted_keywords = sorted_keywords[:20]

    #Se queda con los ultimos 5 para mostrar
    organicKeywords = sorted_keywords[:10]

    return organicKeywords

def GetHeaderKeywordsString():
    headerList = []
    headerDataString = ""
    for sub_lista in headerDataList:
        headerList.extend(sub_lista)

    headerDataString =  " ".join(headerList)
    return headerDataString
        
def GetBodyKeywordString():
    return eliminar_espacios(soup.find("body").text)
     
"""
"""
def SetupHeaders():
    #Recoje los Headers
    headerColumnCount = 1

    while headerColumnCount <= 7:
        headerList = soup.find_all("h" + str(headerColumnCount))
        headerListString = []
        headerCount.append(len(headerList))

        #Append the list with all header items
        for item in headerList:
            headerListString.append(eliminar_espacios(item.text))
        
        #Write the excel
        WriteExcel("H" + str(headerColumnCount) + f" ({len(headerList)})", headerColumnCount - 1, headerListString)  # Llama a la función para escribir los datos en Excel
        
        #Add header to info list
        headerDataList.append(headerListString)

        #Breaks the while loop
        headerColumnCount += 1   
    
def SetupLinks():
    #Recoje todos los links
    linksList = soup.find_all("a")
    internalLinks = []              #Guardara los links internos
    externalLinks = []              #Guardara los links externos
    #Comprueba los links internos y externos
    for link in linksList:
        currentLink = link.get("href")
        if(currentLink != None):
            if(GetContainsSubstring(currentLink,GetDomain(url))):
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
    body = soup.find("body").text
    densityString = GetCalculateDensity(body)
    WriteExcel("Keywords Density",11,densityString)

def HeaderAnalisis():
    headerList = []
    for sub_lista in headerDataList:
        headerList.extend(sub_lista)

    headerDataString =  " ".join(headerList)
    densityList = GetCalculateDensity(headerDataString)    

    WriteExcel("Header Density",12,densityList)

def MainKeywords():
    bodyKeywords = GetBodyKeywordString()
    headerKeywords = GetHeaderKeywordsString()

    mainKeywords = GetCalculateDensityOrganicKeywords(headerKeywords,bodyKeywords)
    mainKeywordsString = [f"{i+1} {word.capitalize()} ({count})points" for i, (count, word) in enumerate(sorted(mainKeywords, key=lambda x: x[0], reverse=True))]
    WriteExcel("Organic Keywords", 13, mainKeywordsString)

def main():
    dominio = GetDomain(url)

    SetupHeaders()
    SetupLinks()
    SetupImages()
    HeaderAnalisis()
    KeywordAnalisis()
    MainKeywords()

    headerNames=["H1","H2","H3","H4","H5","H6","H7"]
    graph.GeneratePieGraph(headerNames,headerCount,"Test")

    #Recoje el title
    workbook.filename = f"{dominio}.xlsx"
    workbook.close()  # Cierra el archivo de Excel al finalizar

main()