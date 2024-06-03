from docx import Document  #biblioteca que importa documentos  pip install python-docx

#Criação de dicionário para a substituição das variaveis
def editword():
    #aqui pode-se fazer calculo para as variaveis
    
    vPWPK= int(vpplc.get())*int(vqplc.get())/1000
    vpplcAcima = int(vPWPK)*1.25
    vpplcAbaixo = int(vPWPK)*0.75
    if vtipoligacao=="TRIFÁSICO":
        vpdkva=(380*(int(vdisentr.get())*1.73)/1000)
        vpdkw=float(vpdkva)*0.92
    else:
        vpdkva=220*(int(vdisentr.get())/1000)
        vpdkw=float(vpdkva)*0.92

    document = Document("MEMORIAL.docx")    

    dictionary = {
        #Aqui as palavras serão substituidas por variaveis da parte da gráfica.
        "$POTINV":"vpotinv.get()",
        "$CLIENTE": "vcliente.get()",
        "$RG":"vrg.get()",
        "$EMISSORRG":"vemissorrg.get()",
        "$CIDADE":"vcidade.get()",
        "$ESTADO": "vestado.get()",
        "$MES":"vmes.get()",
        "$ANO":"vano.get()",
        "$QTMOD":"vqtmod.get()",#quantidade de módulos
        "$POTPLC":"vpotplc.get()",#potencia placa
        "$CC":"vcc.get()",#conta contrato
        "$CLASSE":"vclasse.get()",
        "ENDERECO":"vendereco.get()",
        "$RATEIO":"vrateio.get()",
        "$POSTE":"vposte.get()",
        "$UTME": "vutme.get()",
        "$UTMS": "vutms.get()",
        "$RAMAL": "vramal.get()",#tri ou mono
        "$CONDUTORES": "vcondutores.get()", #quantidade de condutores carregados
        "$DIAMETRO":"vdiametro.get()", #diametro dos cabos
        "$NUMPOLOS" :"vnumpolos.get()", #numero polos disjuntor
        "$CORDISJ": "vcordisj.get()", #CORRENTE disjuntor
        "$PDKVA" : str(vpdkva),
        "$PDKW": str(vpdkw)
        

        



        






    }

    for p in document.paragraphs:
        for key, word in dictionary.items():
            if p.text.find(key) >= 0:
                
                p.text = p.text.replace(key, word)
                print(f"Substituindo '{key}' por '{word}' em: {p.text}")

    document.save("CACETA.docx")
    print('Deu tudo certo')