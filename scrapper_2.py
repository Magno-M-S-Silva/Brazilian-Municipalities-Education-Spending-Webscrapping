from urllib.request import urlopen #para abrir a pagina
from bs4 import BeautifulSoup #para repartir o código
import openpyxl
import datetime
import time

data=datetime.datetime.now()
cidades=openpyxl.load_workbook('C:/Users/T-Gamer/Documents/Translog_DP/tabelas sql/COD_MUN.xlsx')
time.sleep(4)
sheet_cidade=cidades.get_sheet_by_name("COD_MUN")
erro="não transmitiu por meio do Siope"
erro_2="indisponível no momento."


for k in list(range(0,5568)):
    cod_mun=int(sheet_cidade["A"+str(2+k)].value)#pega o código da cidade
    estado=int(sheet_cidade["B"+str(2+k)].value)#pega o código do estado
    if cod_mun==260545:
        continue
    print("Preenchendo para a cidade "+str(sheet_cidade["C"+str(2+k)].value)) #checa a cidade que está sendo estudada
    er_bol=1#variável que denota se há erro
    ano=int(data.year)#ano corrente
    while er_bol==1: #enquanto houver erro
        connect=0
        if ano==2004:
            break
        while connect==0:
            try:
                html=urlopen("https://www.fnde.gov.br/siope/demonstrativoFuncaoEducacao.do?acao=pesquisar&pag=result&anos="+str(ano)+"&periodos=1&cod_uf="+str(estado)+"&municipios="+str(cod_mun))#abre a página
                time.sleep(4)#espera dois segundo
                connect=1
 #               print(connect)
            except Exception as e:
                print(e)
                time.sleep(60)
                continue
            try:
                soup=BeautifulSoup(html.read()) #le a página
 #               print(soup)
            except Exception as ex:
                print("erro de leitura")
                print(ex)
                time.sleep(60)
            if erro_2 in soup.get_text():
                print ("Site fora do ar. Esperando uma hora para tentar de novo")
                time.sleep(3600)
            if ano==2004:
                print("Não conseguiu achar a página para nenhum ano até 2004")
                break
        g=soup.findAll("p")
        if erro in soup.get_text() or  g==[]:#se a mensagem de erro estiver no texto da página
            ano=ano-1#tentar um ano anterior
            if ano==2004:
                print("Não conseguiu achar a página para nenhum ano até 2004")
                break
            continue#começar de novo
        else:#se a mensagem de erro não estiver no texto da página
            er_bol=0#pare o loop que abre páginas
            print("Sucesso em " + str(ano))#informa o ano do sucesso
    if ano==2004:
        print("Não conseguiu achar a página para nenhum ano até 2004")
        break
    tags=soup.findAll("div", {"class": "number"}) #acha os valores em que estamos interessados
    tags_2=soup.findAll("td", {"align":"left", "class": "text"}) #acha os nomes das linhas que estamos interessados
    tag_tot=soup.findAll("strong") #acha as tags dos totais
    modelo=openpyxl.load_workbook('C:/Users/T-Gamer/Documents/Translog_DP/modelo relatorio previo versao 1.xlsx') #abrir o modelo de arquivo a ser preenchido
    time.sleep(4)
    tam=len(tag_tot)
    sheet=modelo.get_sheet_by_name("P1") #abrir planilha do arquivo de excel
    sheet["B158"]=float(tag_tot[tam-1].text.replace(".","").replace(",",".")) #pegar o valor total pago e coloca no excel
    sheet["B158"].number_format='0.00' #colocar em formato com dois dígítos depois do ponto
    tot=float(tag_tot[tam-1].text.replace(".","").replace(",","."))  #anota o valor total no python

    for i in list(range(1,12)): #para cada linha a ser preenchida da tabela 5.1
        index_tag=0 #indexador para a lista de tags, reset o valor quando mudar a linha a ser preenchida do arquivo de excel
        for tag in tags_2: #para cada linha na tabela da página da internet
            index_tag=index_tag+1 #atualiza o indexador para saber aonde a tag está na lista
            if sheet["A"+ str(148+i)].value in tag.text: #Se a linha do excel for igual a tag da internet
                sheet["B"+str(148+i)].value=float(tags[3*index_tag-1].text.replace(".","").replace(",",".")) #atualize a célula do excel com o valor da internet
                sheet["B"+str(148+i)].number_format='0.00' #colocar em formato com dois dígítos depois do ponto
    for j in list(range(0,8)): #para cada linha a ser preenchida na terceira coluna
        try: #tente preencher a coluna, pode dar erro se o valor não existe
            sheet["C"+str(150+j)].value=((sheet["B"+ str(150+j)].value/sheet["B158"].value)) #pega o valor da linha e divide pelo total
            sheet["C"+str(150+j)].number_format='0.00%' ##colocar em formato percentual com dois dígitos depois do ponto
        except: #se der erro
            print("A linha "+str(150+j)+" não pode ser preenchida") #avisa ao usuário aonde deu o erro
    sheet["A159"].value="Fonte: Sistema SIOPE ano"+str(ano)+"_FNDE"

    modelo.save("C:/Users/T-Gamer/Documents/Translog_DP/Municípios/"+str(cod_mun)+".xlsx") #salva o arquivo de excel
    time.sleep(4)