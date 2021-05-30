from renamed import coleta
from renamed import Document
from renamed import apWord
from renamed import bmWord
from renamed import inputToWord

for contractsNum in range (len(coleta.listaNomes)):

    if "AP" in coleta.listaApBm[contractsNum]: #Processo de AP
        identity = "AP"
        doc = Document(apWord)
        inputToWord(doc, contractsNum, identity)

    elif "BM" in coleta.listaApBm[contractsNum]:
        identity = "BM"
        doc = Document(bmWord)
        inputToWord(doc, contractsNum, identity)

    elif ""

    else:
        continue













