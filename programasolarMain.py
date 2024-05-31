# IMPORTAR AS BIBLIOTECAS
from tkinter import * # biblioteca gráfica TKinter pip install tk
from tkinter import ttk
from docx import Document  #biblioteca que importa documentos  pip install python-docx
from openpyxl import load_workbook #Biblioteca para abrir o excel  pip install openpyxl

#Criação de dicionário para a substituição das variaveis
def editword():
    #aqui pode-se fazer calculo para as variaveis
    document = Document("MEMORIAL.docx")
    dictionary = {
        #Aqui as palavras serão substituidas por variaveis da parte da gráfica.
        "MEMORIAL":"CACETA"
    }

    for p in document.paragraphs:
        for key, word in dictionary.items():
            if p.text.find(key) >= 0:
                p.text = p.text.replace(key, word)

    document.save("CACETA.docx")
    print('Deu tudo certo')

editword()