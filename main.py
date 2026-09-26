from openpyxl import Workbook, load_workbook
import os

arquivo = "Organização_Financeira.xlsx"

if not os.path.exists(arquivo):
    wb = Workbook() #Aqui ele cria uma nova planilha
    ws = wb.active #Ele pega a planilha aberta

#wb = WordBook
#ws = WorkSheet
ws.append([" Salário "])
ws.append(["Mes","Valor Liquido","Valor total das contas"])
wb.save(arquivo)

mes = input("Digite o mês referente: ")
valorliquido = float(input("Digite o Valor Líquido recebido (Após descontos): " ))
valor_contas = float(input("Digite o Valor Total das contas a pagar: "))

wb=load_workbook(arquivo) #abre a planilha 
ws = wb.active

ws.append([mes, valorliquido, valor_contas])
wb.save(arquivo)

print("Dados salvos com sucesso, agora é necessário abrir o arquivo na pasta em questão")

