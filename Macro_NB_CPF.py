import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook
from pdfminer import high_level
from os import listdir
from os.path import isfile, join

#pdf = 'C:\\Users\\pablo.pamf\\Documents\\2021\\Macro_NB\\Report1.pdf'
caminho = 'C:\\Users\\pablo.pamf\\Documents\\2021\\Macro_NB\\'
ind = list(range(1,5001))

files = [f for f in listdir(caminho) if isfile(join(caminho, f))]
#print(ind)
dados = pd.DataFrame()
dados['ind'] = ind
nb = re.compile(r"[0-9]{3}\.[0-9]{3}\.[0-9]{3}\-[0-9]{2}\s")
for i in range(len(files)):
    #v_pdf = files[i][-3:]
    v_html = files[i][-4:]
    #print(v_pdf)
    if v_html == 'html':
        html = open(caminho+files[i],'r',encoding="utf8").read()
        nome_chat = files[i][:-5]      
        res_nb = re.findall(nb, html)
        #print(res_nb)
        
        p = []
        nb_sem_ponto = []
        nb_sem_dig = []
        for i in range(len(res_nb)):
            #p = re.sub(".", res_nb[i], "")
            p = res_nb[i].replace(".","" )
            #p1 = p[i].replace("-","")
            nb_sem_ponto.append(p)

        for i in range(len(res_nb)):
            #p = re.sub(".", res_nb[i], "")
            p = nb_sem_ponto[i].replace("-","" )
            #p1 = p[i].replace("-","")
            nb_sem_dig.append(p)

        #print(nb_sem_dig)
        res_nb_unic = list(set(nb_sem_dig))
        res_nb_u = list(set(res_nb))
        #print(res_nb_unic)
        print(len(res_nb_unic))
        #print(len(nb_sem_dig))
        #print(nb_sem_dig)
        if len(res_nb_unic) > 0:
        
            dados[nome_chat] = pd.Series(res_nb_u)
        #dados.insert(nome_chat,res_nb_unic,True)
print(dados)
book = Workbook()
sheet = book.active
book.save('resultadosCPF'+'.xlsx')
try:
    cria_excel_total = pd.ExcelWriter(caminho+'resultadosCPF'+'.xlsx')
    dados.to_excel(cria_excel_total, sheet_name= 'Total', index=False) 
    cria_excel_total.close()
except ValueError:
            print("Erro ao gerar excel")