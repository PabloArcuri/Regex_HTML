import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook
from pdfminer import high_level
from os import listdir
from os.path import isfile, join

pdf = 'C:\\Users\\pablo.pamf\\Documents\\2021\\Macro_NB\\Report1.pdf'
caminho = 'C:\\Users\\pablo.pamf\\Documents\\2021\\Macro_NB\\'

files = [f for f in listdir(caminho) if isfile(join(caminho, f))]
print(files)

for i in range(len(files)):
    v_pdf = files[i][-3:]
    #print(v_pdf)
    if v_pdf == 'pdf':
        n_pg = PyPDF2.PdfFileReader(caminho+files[i]).numPages
        num_paginas = n_pg+1
        pages = list(range(num_paginas)) #NÚMERO DE PÁGINA DO PDF
        print(num_paginas)

        extracted_text = high_level.extract_text(caminho+files[i], '', pages, num_paginas)
        #print(extracted_text)   # TESTE DE CAPTURA PDF
        nome_chat = extracted_text[0:50].replace("\n","")

        #print(extracted_text[0:50])
        nb = re.compile(r"\d{3}\.\d{3}\.\d{3}-\d{1}|\d{9}-\d{1}\s")
        res_nb = re.findall(nb, extracted_text)
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
        #print(res_nb_unic)
        print(len(res_nb_unic))
        print(len(nb_sem_dig))
        #print(nb_sem_dig)

        dados = pd.DataFrame(columns=['NB'])
        dados['NB'] = res_nb_unic
        book = Workbook()
        sheet = book.active
        book.save(nome_chat+'.xlsx')
        try:
            cria_excel_total = pd.ExcelWriter(caminho+nome_chat+'.xlsx')
            dados.to_excel(cria_excel_total, sheet_name= 'Total', index=False) 
            cria_excel_total.close()
        except ValueError:
                    print("Erro ao gerar excel")