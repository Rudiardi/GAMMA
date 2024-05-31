# IMPORTAR AS BIBLIOTECAS
from tkinter import * # biblioteca gráfica TKinter pip install tk
from tkinter import ttk
from docx import Document  #biblioteca que importa documentos  pip install python-docx
from docx.shared import Inches
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

def editimg():
    # Carregue o arquivo Word existente
    document = Document("MEMORIAL.docx")

    # Itere pelos parágrafos e substitua a imagem com a tag "replace_me"
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if 'Fig1' in run.text:
                            # Substitua pelo caminho da nova imagem
                            for existing_run in paragraph.runs:
                                paragraph._element.remove(existing_run._element)
                            # Adicione a nova imagem (substitua pelo caminho correto)
                            paragraph.add_run().add_picture('fig1.png',width=Inches(5))
                        if 'Fig2' in run.text:
                            # Substitua pelo caminho da nova imagem
                            for existing_run in paragraph.runs:
                                paragraph._element.remove(existing_run._element)
                            # Adicione a nova imagem (substitua pelo caminho correto)
                            paragraph.add_run().add_picture('fig2.png',width=Inches(4))
                        if 'Fig3' in run.text:
                            # Substitua pelo caminho da nova imagem
                            for existing_run in paragraph.runs:
                                paragraph._element.remove(existing_run._element)
                            # Adicione a nova imagem (substitua pelo caminho correto)
                            paragraph.add_run().add_picture('fig3.png',width=Inches(4))
                        if 'Fig4' in run.text:
                            # Substitua pelo caminho da nova imagem
                            for existing_run in paragraph.runs:
                                paragraph._element.remove(existing_run._element)
                            # Adicione a nova imagem (substitua pelo caminho correto)
                            paragraph.add_run().add_picture('fig4.png',width=Inches(3))    

    # Salve o documento modificado
    document.save("CACETA.docx")
    print("imagem subistituida")

editimg()