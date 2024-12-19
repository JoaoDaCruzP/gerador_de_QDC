import pandas as pd
from openpyxl import load_workbook
from funcoes import nome_item, quant_item, msg_colorida

doc = 'prototipo.xlsx'
book = load_workbook(doc)

sheet = book.active

def selecionar_celula(lin):

    lin = lin
    col1= 'A'; col2 = 'B'; col3 = 'C'; col4 = 'E'; col5 = 'G'

    new_cell1, new_cell3, new_cell4, new_cell5 = nome_item()
    new_cell2 = quant_item()

    name = new_cell1
    quant = new_cell2
    potencia = new_cell3
    fp = new_cell4
    fd = new_cell5

    modifica = name
    modifica = quant
    modifica = potencia
    modifica = fp
    modifica = fd


    modifica = sheet[col1 + str(lin + 1)]
    modifica.value = name
    modifica = sheet[col2 + str(lin + 1)]
    modifica.value = quant
    modifica = sheet[col3 + str(lin + 1)]
    modifica.value = potencia
    modifica = sheet[col4 + str(lin + 1)]
    modifica.value = fp
    modifica = sheet[col5 + str(lin + 1)]
    modifica.value = fd

def salva_arquivo():
    novo_arquivo = str(input(f'\nDIGITE O NOME DO ARQUIVO QUE DESEJA SALVAR: '))
    novo_arquivo = novo_arquivo + '.xlsx'
    msg_colorida(cor="green", text=" ARQUIVO SALVO COM SUCESSO. ")
    book.save(novo_arquivo)