import json
import pandas as pd
from datetime import date
from lxml import etree
import os.path

def encontrarexcel():
    if os.path.exists(".\\@input\\"+ 'Short.xlsx') == False:
        nomearquivo = ''
    else:
        nomearquivo = ".\\@input\\"+ 'Short.xlsx'
    return nomearquivo

def encontrarjson():
    notfound = True
    i = 2
    while (notfound == True and i <=20):
        if i <= 9:
            nomearquivo = ".\\@input\\"+'produtos_cp'+'0'+str(i)+'.json'
        else:
            nomearquivo = ".\\@input\\"+ 'produtos_cp'+str(i)+'.json'
        if os.path.exists(nomearquivo) == False:
            i = i + 1
            nomearquivo = ''
        else:
            notfound = False
    return nomearquivo
  
def encontrarxml():
    notfound = True
    i = 2
    while (notfound == True and i <=20):
        if i <= 9:
            nomearquivo = ".\\@input\\"+ 'produtos_cp'+'0'+str(i)+'.xml'
        else:
            nomearquivo = ".\\@input\\"+ 'produtos_cp'+str(i)+'.xml'
        if os.path.exists(nomearquivo) == False:
            i = i + 1
            nomearquivo = ''
        else:
            notfound = False
    return nomearquivo

def ler_json(nomearquivo):
    with open(nomearquivo,'r',encoding='utf8') as arquivojson:
        return json.load(arquivojson)
       
def escrever_json(dados,nomearquivo):
    with open(nomearquivo,'w',encoding='utf8') as arquivojson:
        json.dump(dados,arquivojson,ensure_ascii=False,sort_keys=False,indent=2,separators=(',',':'))

def remove_duplicados(lista):
    listareturn = []
    for i in lista:
        if i not in listareturn:
            listareturn.append(i)
    return listareturn

def ler_excel(nomearquivo):
    excelshort = pd.read_excel(nomearquivo,usecols = "A")
    excelshort["Line Number"] = excelshort["Line Number"].astype("object")
    excelshort = excelshort.values.tolist()
    excelshort = remove_duplicados(excelshort)
    excelshort = [str(item) for sublist in excelshort for item in sublist]
    return excelshort

def listdifference(list1,list2):
    newlist = []
    for i in list1:
        if(i not in list2):
            newlist.append(i)
    return newlist
                
def limpalinenumber(list1,list2):
    list3 = []
    for i in list1:
        listaux = i.get('lineNumber').split(",")
        listdif = listdifference(listaux,list2)
        listdif = ",".join(listdif)
        i['lineNumber'] = listdif
        if i['lineNumber'] != "":
            list3.append(i)
    return list3

def dataatual():
    dia = date.today().day
    mes = date.today().month
    ano = str(date.today().year)
    if dia > 9:
        dia = str(dia)
    else:
        dia = '0'+ str(dia)
    if mes > 9:
        mes = str(mes)
    else:
        mes = '0'+ str(mes)
    return dia+mes+ano

def lerxml(nomearquivo):
    with open(nomearquivo,'r',encoding='utf8') as origem:
        parser = etree.XMLParser(resolve_entities=False, strip_cdata=False)
        tree = etree.parse(origem,parser)
    return tree

def escreverxml(elementtree,nomearquivo):
    with open(nomearquivo, 'wb') as destination:
        elementtree.write(destination, encoding='utf-8', xml_declaration=True, pretty_print=True)

def shortxml(elementtree,listexcel):
    root = elementtree.getroot()
    for el in root.xpath("//produtos[@lineNumber]"):
        listaaux = []
        listaaux = el.attrib['lineNumber'].split(",")
        listaresult = listdifference(listaaux,listexcel)
        stringresult = ",".join(listaresult)
        el.attrib['lineNumber'] = stringresult
        if el.attrib['lineNumber'] == "":
            el.getparent().remove(el)
    stringroot = etree.tostring(root, pretty_print=True).decode('utf-8')
    stringroot = stringroot.replace('\t<!--','\n<!--')
    stringroot = stringroot.replace('produtos ','produtos \n\t\t')
    stringroot = stringroot.replace('" ','" \n\t\t')
    newroot = etree.fromstring(stringroot)
    return newroot.getroottree()
print('<=======================================================================>')    
print('Bem vindo ao programa que realiza Short dos arquivos Json e XML')
print('<=======================================================================>')
print('Para que esse programa rode com sucesso faz-se necessario que o Short.xlsx, produto_cpXX.json e produto_cpXX.xml estejam na pasta @input')
print('<=======================================================================>')
jsonfile = encontrarjson()
excelfile = encontrarexcel()
xmlfile = encontrarxml()


if jsonfile == '' or excelfile == '' or xmlfile == '':
    print('Erro nao foi possivel encontrar os 3 arquivos necessarios na pasta @input')
    input()
else :
    print('Os arquivos foram encontrados')
    print('<=======================================================================>')
    produtos_cpjson = ler_json(jsonfile)
    excel = ler_excel(excelfile)
    produtos_cpxml = lerxml(xmlfile)
    
    data_atual = dataatual()
    
    escrever_json(produtos_cpjson, '.\\@output\\'+data_atual+'-'+jsonfile.replace('.\\@input\\',''))
    escreverxml(produtos_cpxml,'.\\@output\\'+data_atual+'-'+xmlfile.replace('.\\@input\\',''))
    
    newprodutos_cpjson = limpalinenumber(produtos_cpjson,excel)
    escrever_json(newprodutos_cpjson,'.\\@output\\'+jsonfile.replace('.\\@input\\',''))
    print('Json shortado se encontra disponivel na pasta @output')
    print('<=======================================================================>')
    newprodutos_cpxml = shortxml(produtos_cpxml,excel)
    escreverxml(newprodutos_cpxml,'.\\@output\\'+xmlfile.replace('.\\@input\\',''))
    print('XML shortado se encontra disponivel na pasta @output')
    print('<=======================================================================>')
    input()

    
    
    