import os
import win32com.client

#create Sap2000 object
SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")

#start Sap2000 application
SapObject.ApplicationStart()

#create SapModel object
SapModel = SapObject.SapModel



class Ferramentas:

    #ferramentas para se usar no sap2000

    def __init__(self):
        pass


    def add_load_patterner():
        pass






nomes = []
numero = 5
combinacao = "ELS-FREQUENTES-vao"
tipos =["CPtalha","SCtalha"]
valores =[1,0.8]

for i in range(numero+1):
    if i == 0:
        pass
    else:
        nomes.append(combinacao+str(i))


for numero, item in enumerate(nomes):

    SapModel.RespCombo.Add(item,0)
    tipos[0] = tipos[0]+str(numero+1)
    tipos[1] = tipos[1]+str(numero+1)
    SapModel.LoadPatterns.add(tipos[0],1,0) #Dead = 1

    SapModel.LoadPatterns.Add(tipos[1],3,0) #Live = 3

    [SapModel.RespCombo.SetCaseList(item,0,tipos[i],valores[i]) for i in range(len(valores))]
    for i in range(len(tipos)):
        tipos[i] = tipos[i].replace(str(numero+1),"")

        



