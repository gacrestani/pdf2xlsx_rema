# -*- coding: utf-8 -*-
"""
Created on Jul 2019

@author: gacrestani

Gera vários arquivos .xlsx a partir de um laudo .pdf
"""

import tabula
import os
import datetime

print("Este programa cria arquivos .xlsx das tabelas de um arquivo .pdf!")
print("Início:", datetime.datetime.now())

titulo_relatorio = input("Insira o titulo do laudo: ")
diretorio = input(r"Insira o diretório onde está o laudo. Uma nova pasta será criada. Exemplo 'C:\Users\usuariorema\Documents\laudo1': ")
pag_inicial = int(input("Pagina Inicial a ser analisada: "))
pag_final = int(input("Pagina Final a ser analisada: "))

lista_de_paginas = list(range(pag_inicial, pag_final+1))

while True:
    retirar = input("Página para retirar da análise: ")
    try:
        retirar_int = int(retirar)
        print("Retirar a página %s" %(retirar))
        lista_de_paginas.remove(retirar_int)
        print(lista_de_paginas)
    except ValueError:
        print("Mais nenhuma página para se retirar da análise.")
        break

print ('As seguintes páginas serão analisadas: \n')
print(lista_de_paginas, '\n')

print("Criando o diretório...")   
novo_diretorio = diretorio + '\\' + titulo_relatorio
print("Diretorio: %s \n" %(novo_diretorio))
os.makedirs(novo_diretorio, exist_ok = True)

print("Analisando o PDF...")

dicionario_dataframes = {}
for pagina in lista_de_paginas:
    dicionario_dataframes[pagina] = tabula.read_pdf(diretorio + "\\" + titulo_relatorio + ".pdf", pages=pagina, encoding='cp1252', multiple_tables = True)
    print("Página %s analisada." % (pagina))
    

print("\nCriando as planilhas...")

for page, dataframe in dicionario_dataframes.items():
    if type(dataframe) is list:
        for i in range(len(dataframe)):
            export_excel = dataframe[i].to_excel (r'%s\%s\%s_%s_%s.xlsx' %(diretorio,titulo_relatorio,titulo_relatorio,page,i), index = None, header = True)
            print("Planilha %s_%s.xlsx criada." % (titulo_relatorio,page))
    else:
        export_excel = dataframe.to_excel (r'%s\%s\%s_%s.xlsx' %(diretorio,titulo_relatorio,titulo_relatorio,page), index = None, header = True)
        print("Planilha %s_%s.xlsx criada." % (titulo_relatorio,page))
    
print("\nFinal: ", datetime.datetime.now())