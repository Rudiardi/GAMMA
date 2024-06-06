from docx import Document  #biblioteca que importa documentos  pip install python-docx

#Criação de dicionário para a substituição das variaveis
def editword():
    #aqui pode-se fazer calculo para as variaveis
    

    '''
    vPWPK= int(vpplc.get())*int(vqplc.get())/1000
    vpplcAcima = int(vPWPK)*1.25
    vpplcAbaixo = int(vPWPK)*0.75
    if vtipoligacao=="TRIFÁSICO":
        vpdkva=(380*(int(vdisentr.get())*1.73)/1000)
        vpdkw=float(vpdkva)*0.92
    else:
        vpdkva=220*(int(vdisentr.get())/1000)
        vpdkw=float(vpdkva)*0.92
    '''
    document = Document("MEMORIAL.docx")    

    dictionary = {
        #Aqui as palavras serão substituidas por variaveis da parte da gráfica.
        "$POTINV":"vpotinv.get()",#POTENCIA INVERSOR
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
        #"$PDKVA" : str(vpdkva),
        #"$PDKW": str(vpdkw),
        "$KWH":"vkwh.get()", #kwh
        "$POTSIS":"vpotsis.get()", #potencia do sistema
        
        "$MODPLACA": "vmodplaca.get()", # modelo módulo
        "$FABPLC":"vfabplc.get()",#fabricante das placas
        "$VOC":"vvoc.get()",#tensão de circuito aberto módulo
        "$ISC":"visc.get()", #corrente de curto circuito módulo
        "$VPMP":"vvpmp.get()",#tensão de máxima potência módulo
        "$IPMP":"vipmp.get()", #corrente de máxima potência módulo
        "$EFCMOD":"vfcmod.get()",#eficiência módulo
        "$COMMOD":"vcommod.get()", #comprimento módulo
        "$LARGMOD":"vlargmod.get()",#largura módulo
        "$AREAMOD":"vareamod.get()", #área do módulo
        "$PESOMOD":"vpesomod.get()",#peso do módulo         
        
        "$FABINV": "vfabinv",#fabricante inversor
        "$MODINV":"vmodinv.get()",#modelo do inversor
        "$PMAX":"$vpmax.get()",#Máxima potência na entrada CC – Pmax-cc [kW]
        "$VCCMAX":"$vvccmax.get()",#Máxima tensão CC – Vcc-máx [V]
        "$ICCMAX":"viccmax.get()",#Máxima corrente CC – Icc-máx [V]
        "$MPPTMAX":"vmpptmax.get()",#Máxima tensão MPPT – Vpmp-máx [V]
        "$MPPTMIM":"vmpptmin.get()",#Mínima tensão MPPT – Vpmp-min [V]
        "$VCCPART":"vvccpart.get()",#Tensão CC de partida – Vcc-part [V]
        "$STRINGS":"vstrings.get()",#Quantidade de Strings
        "$QTDMPPT":"vqtdmppt.get()",#quantidade de mmpts
        "$PSMAX" : "vpsmax.get()", #Máxima potência na saída CA INVERSOR
        "$IMAX" : "vimax.get()",#máxima corrente de saída
        "$VCAMAX":"vcamax.get()", #Máxima tensão CA – Vca-máx [V]
        "$VCAMIN":"vcamin.get()", #Mínima tensão CA – Vca-min [V]
        "$VSTR":"vvstr.get()",#tensão das strings
        "$QTDDPS": "vqtddps.get()",#num de DPS



    }

    for p in document.paragraphs:
        for key, word in dictionary.items():
            if p.text.find(key) >= 0:
                
                p.text = p.text.replace(key, word)
                print(f"Substituindo '{key}' por '{word}' em: {p.text}")

    document.save("CACETA.docx")
    print('Deu tudo certo')