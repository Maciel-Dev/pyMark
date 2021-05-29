import xlrd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx.shared import Cm
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt

#Archives
powerPoint = 'Pasta12.xlsx'
cAP = 'CONTRATO_AP.docx'
cBM = 'CONTRATO_BM.docx'
#Start Run Archives Planilha
workBook = xlrd.open_workbook(powerPoint)
planilha = workBook.sheet_by_name('PRONTOS')

#Main Function

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

        elif isinstance(cell.value, float):
            a = cell.value
            b = int(a)
            c = str(b)
            coleta.listaNumeroProcesso.append(c)

    del coleta.listaNumeroProcesso[0:1]


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
#Funcionou aqui pq Tem letra na celula (str) - Porém, uma das células só possui número
    coleta.listaCI = []

    for cell in planilha.col(5):
        if isinstance(cell.value, str):
            coleta.listaCI.append(cell.value)

        else: #Acrescenta qualquer outro tipo de valor à minha lista
            coleta.listaCI.append(cell.value)

    del coleta.listaCI[:1]
######################################################################
    coleta.listaRenda = []

    for cell in planilha.col(8):
        if isinstance(cell.value, str):
            coleta.listaRenda.append(cell.value)

        else:
            coleta.listaRenda.append(cell.value)

    del coleta.listaRenda[:1]
######################################################################
    coleta.listaRendaExt = []

    for cell in planilha.col(9):
        if isinstance(cell.value, str):
            coleta.listaRendaExt.append(cell.value)

        else:
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
    coleta.novaListaAno = []

    for cell in planilha.col(18):
        if isinstance(cell.value, int):
            a = cell.value
            b = str(a)
            coleta.listaAno.append(cell.value)

        elif isinstance(cell.value, float):
            d = cell.value
            e = int(d)
            f = str(e)
            coleta.listaAno.append(f)

        else:
            coleta.listaAno.append(cell.value)

    del coleta.listaAno[:1]


    timer2 = 0
    timer3 = 0

    while timer2 != len(coleta.listaAno):
        if "\n" in coleta.listaAno[timer2]:
            recoloca = coleta.listaAno[timer2]
            recoloca = recoloca.replace('\n', '-')
            coleta.novaListaAno.append(recoloca)
            timer2 += 1

        else:
            coleta.novaListaAno.append(coleta.listaAno[timer2])
            timer2 += 1

######################################################################
    coleta.listaApBm = []

    for cell in planilha.col(3):
        if isinstance(cell.value, str):
            coleta.listaApBm.append(cell.value)

        else:
            coleta.listaApBm.append(cell.value)

    del coleta.listaApBm[:1]
###################################################################
#Start Code - Get Paragraphs

doc = Document(cAP)
paragraphs = doc.paragraphs
coleta()

lisTextos = []

for i in paragraphs:
    if "PROCESSO" in i.text:
        lisTextos.append(i.text)

    elif "NOME" in i.text:
        lisTextos.append(i.text)

    elif "RENDA" in i.text:
        lisTextos.append(i.text)

    elif "VALOR_BT" in i.text:
        lisTextos.append(i.text)


paragraphProcesso = lisTextos[0]
paragraphName = lisTextos[1]
paragraphRenda = lisTextos[2]
paragraphValorBt = lisTextos[3]

#Creates all My paragraghs

listParagraphName = []
listParagraphProcesso = []
listParagraphRenda = []
listaParagraphValorBt = []


for i in range (len(coleta.listaNomes)):
    a = paragraphName.replace('«NOME»', f'{coleta.listaNomes[i]}').replace("«ESTADO_CIVIL»", f'{coleta.listaEstadoCivil[i]}').replace("«CI»", f'{coleta.listaCI[i]}').replace("«CPF1»", f'{coleta.listaCPF[i]}').replace("«ENDEREÇO»", f'{coleta.listaEndereço[i]}').replace("«BAIRRO_DE_ORIGEM»", f'{coleta.listaBairro[i]}')
    listParagraphName.append(a) 

    b = paragraphProcesso.replace("«PROCESSO»", f'{coleta.novaListaNumeroProcesso[i]}').replace("«ANO»", f'{coleta.novaListaAno[i]}')
    listParagraphProcesso.append(b)

    c = paragraphRenda.replace("«RENDA»", f'{coleta.listaRenda[i]}').replace("(«EXT_RENDA»)", f'{coleta.listaRendaExt[i]}')
    listParagraphRenda.append(c)

    d = paragraphValorBt.replace("«VALOR_BT»", f'{coleta.listaBT[i]}').replace("(«EXT_BT»)", f'{coleta.listaBTExt[i]}')
    listaParagraphValorBt.append(d)

#Writing .txt file db
for j in range(len(listParagraphName)):
    print(listParagraphName[j])

for ano in range(len(listParagraphProcesso)):
    print(listParagraphProcesso[ano])

for renda in range(len(listParagraphRenda)):
    print(listParagraphRenda[renda])











