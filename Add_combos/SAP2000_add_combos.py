import os
import win32com.client
import time
import pandas as pd

def dados():
    df = pd.read_excel('Combinações.xlsx',header=5)
    df = df.drop(columns=["Unnamed: 0",'Unnamed: 1'])
    df = df.fillna(0)
    print(df)
    df = df.rename(columns={'Unnamed: 2':"Load Partenner"})

    return df

#create Sap2000 object
SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")

#start Sap2000 application
SapObject.ApplicationStart()

#create SapModel object
SapModel = SapObject.SapModel

excel = dados()

nomes = excel['Load Partenner']

combos = list(excel.columns)[2:]

quantidade = len(combos)

print(nomes)

'''
#adicionar caso
for nome in nomes:

    if nome[0] == 'w' or nome[0] == 'W':
        SapModel.LoadPatterns.Add(nome,6,0)
    elif nome[0] == 'c' or nome[0] == 'C':
        SapModel.LoadPatterns.Add(nome,1,0)
    elif nome[0] == 's' or nome[0] =='S':
        SapModel.LoadPatterns.Add(nome,3,0)
    elif nome[0] =='n' or nome[0]== 'N':
        SapModel.LoadPatterns.Add(nome,12,0)
    elif 'T' in nome or 't' in nome: #possivel bug associado
        SapModel.LoadPatterns.Add(nome,10,0)
'''

#add combo
for itens in combos:
    SapModel.RespCombo.Add(itens, 0)

    for i,j in zip(nomes,excel[itens]):
        if j == 0:
            pass
        else:
            SapModel.RespCombo.SetCaseList(itens,0,i,j) 
    



