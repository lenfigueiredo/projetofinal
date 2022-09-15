

# Projeto Final do Curso de Introdução em Python - UFRGS TCERS
# Programa conforme os itens do projeto final
# Autor: Luiz Eduardo Nascimento Figueiredo (luiznf)
# Data 15/09/2022
# Versão: 0.0.5



import requests 
import pandas 
import openpyxl 

# 1- O Portal de Dados Abertos do TCE-RS contem o Balancete de Despesa
#Consolidados 2022 no link disponível aqui.

# 2 -Use o pacote Requests para armazenar em memória baixar o arquivo do
#balancete na variável denominada ”dados”.

endereco = 'http://dados.tce.rs.gov.br/dados/municipal/balancete-despesa/2022.csv'
dados = requests.get(endereco)


# 3 - Grave em disco o conteúdo da variável ”dados” em um arquivo denominado
#”balancete.csv”

arquivo = open('balancete.csv', 'wb')
for data in dados.iter_content():
    arquivo.write(data)

# 4 - Use o pacote Pandas, para ler o arquivo ”balancete.csv” para a variável
#”balancete”.

balancete = pandas.read_csv('balancete.csv')

print (balancete)

# 5 - Usando o pacote Pandas, grave em disco o conteúdo da variável ”balancete”
# em um arquivo denominado ”balancete.xlsx”, e

balancete.to_excel("balancete.xlsx", sheet_name="balancete", index = False)

print("Documento balancete.xlsx gerado")

# 6 - Usando o pacote OpenPyXL, leia o conteúdo do arquivo ”balancete.xlsx”
#para a variável ”novo balancete”, e

novo_balancete = openpyxl.load_workbook('balancete.xlsx')

print('Load finalizado')

# 7 - Finalizando a aplicaç˜ao, usando o OpenPyXL, grave o conteúdo da variável
#”novo_balancete” no arquivo ”novo_balancete.xlsx”
      
novo_balancete.save('novo_balancete.xlsx')

print('Projeto final concluído')

