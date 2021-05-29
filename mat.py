from renamed import coleta
from renamed import Document
from renamed import apWord
from renamed import bmWord
from renamed import inputToWord

for contractsNum in range (len(coleta.listaNomes)):

    if "AP" in coleta.listaApBm[contractsNum]: #Processo de AP
        doc = Document(apWord)
        inputToWord(doc, contractsNum)

    elif "BM" in coleta.listaApBm[contractsNum]:
        doc = Document(bmWord)
        inputToWord(doc, contractsNum)

    else:
        continue













