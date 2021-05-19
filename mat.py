#!/usr/local/bin/python
# coding: utf-8

import xlrd
from docx import Document
from docx.shared import Inches
from docx import *

powerPoint = 'Pasta12.xlsx'
cAP = 'CONTRATO_AP.docx'
cBM = 'CONTRATO_BM.docx'

workBook = xlrd.open_workbook(powerPoint)
planilha = workBook.sheet_by_name('PRONTOS')


def coleta():
#Coletar nomes
    coleta.listaNomes = []

    for cell in planilha.col(1):
        if isinstance(cell.value, str):
            coleta.listaNomes.append(cell.value)

    del coleta.listaNomes[0:1]

######################################################################
    coleta.listaCPF = []

    for cell in planilha.col(16):
        if isinstance(cell.value, str):
            coleta.listaCPF.append(cell.value)

    del coleta.listaCPF[0:1]

######################################################################
    coleta.listaEndereço = []

    for cell in planilha.col(21):
        if isinstance(cell.value, str):
            coleta.listaEndereço.append(cell.value)

    del coleta.listaEndereço[0:1]

######################################################################
    coleta.listaBairro = []

    for cell in planilha.col(22):
        if isinstance(cell.value, str):
            coleta.listaBairro.append(cell.value)

    del coleta.listaBairro[0:1]

######################################################################
    coleta.listaNumeroProcesso = []
    coleta.novaListaNumeroProcesso = []

    for cell in planilha.col(17):
        if isinstance(cell.value, str):
            coleta.listaNumeroProcesso.append(cell.value)

    del coleta.listaNumeroProcesso[0:1]


    coleta.novaListaNumeroProcesso
    timer2 = 0
    timer3 = 0

    while timer2 != len(coleta.listaNumeroProcesso):
        if "\n" in coleta.listaNumeroProcesso[timer2]:
            a = coleta.listaNumeroProcesso[timer2]
            a = a.replace('\n', '-')
            coleta.novaListaNumeroProcesso.append(a)
            timer2 += 1

        else:
            coleta.novaListaNumeroProcesso.append(coleta.listaNumeroProcesso[timer2])
            timer2 += 1

######################################################################
#Estado Civil
    coleta.listaEstadoCivil = []

    for cell in planilha.col(6):
        if isinstance(cell.value, str):
            coleta.listaEstadoCivil.append(cell.value)

    del coleta.listaEstadoCivil[:1]

######################################################################
#Funcionou aqui pq Tem letra na celula (str)

    coleta.listaCI = []

    for cell in planilha.col(5):
        if isinstance(cell.value, str):
            coleta.listaCI.append(cell.value)

    del coleta.listaCI[:1]
######################################################################
    coleta.listaRenda = []

    for cell in planilha.col(8):
        if isinstance(cell.value, str):
            coleta.listaRenda.append(cell.value)

    del coleta.listaRenda[:1]
######################################################################

    coleta.listaRendaExt = []

    for cell in planilha.col(9):
        if isinstance(cell.value, str):
            coleta.listaRendaExt.append(cell.value)

    del coleta.listaRendaExt[:1]
######################################################################

    coleta.listaBT = []

    for cell in planilha.col(10):
        if isinstance(cell.value, str):
            coleta.listaBT.append(cell.value)

    del coleta.listaBT[:1]

######################################################################

    coleta.listaBTExt = []

    for cell in planilha.col(11):
        if isinstance(cell.value, str):
            coleta.listaBTExt.append(cell.value)

    del coleta.listaBTExt[:1]

######################################################################

    coleta.listaAno = []

    for cell in planilha.col(18):
        if isinstance(cell.value, int):
            coleta.listaAno.append(cell.value)

        else:
            coleta.listaAno.append(cell.value)

    del coleta.listaAno[:1]

######################################################################
    coleta.listaApBm = []

    for cell in planilha.col(3):
        if isinstance(cell.value, str):
            coleta.listaApBm.append(cell.value)

        else:
            coleta.listaApBm.append(cell.value)

    del coleta.listaApBm[:1]


'''
Strings Que serão Utilizadas
'''


#Write function
#Abrir documento
openDoc = docx.Document('CONTRATO_AP.docx')
print(len(openDoc))

'''
Início da função de alterações no arquivo
'''


'''
#Salvar Documento
saveDoc = document.save(f'{coleta.listaNomes[contador]}-{coleta.listaApBm[contador]}.docx')
'''













#Calling Function coleta
coleta()
