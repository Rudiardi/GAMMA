# IMPORTAR AS BIBLIOTECAS
from tkinter import * # biblioteca gráfica TKinter pip install tk
from tkinter import ttk
from docx import Document  #biblioteca que importa documentos  pip install python-docx
from docx.shared import Inches
from openpyxl import load_workbook #Biblioteca para abrir o excel  pip install openpyxl

from editimg import editimg
from editword import editword


editword()

editimg()