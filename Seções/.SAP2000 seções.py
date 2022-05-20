import os
import time
import win32com.client


def multiplos(propriedades,*args):
    SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")
    SapModel = SapObject.SapModel
    enderecos = list(*args)
    for locais in enderecos:
        SapModel.File.OpenFile(locais)
        time.sleep(4)
        SapModel.SetPresentUnits(7)
        SapModel.PropFrame.SetISection(nome, aco, propriedades[0],propriedades[1],propriedades[2]
        ,propriedades[3],propriedades[4],propriedades[5])

def atual (propriedades,*args):
    #create Sap2000 object
    SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")

    #start Sap2000 application
    SapObject.ApplicationStart()

    #create SapModel object
    SapModel = SapObject.SapModel
    #unidades

    #ret = SapModel.SetPresentUnits(7)
    SapModel.PropFrame.SetISection(nome, aco, propriedades[0],propriedades[1],propriedades[2]
    ,propriedades[3],propriedades[4],propriedades[5])











#secao I

nome = "PS_TALHA"
aco = "A572"

"""
t3 = altura do perfil
t2 = largura do flange superior
tf = espessura do falnge superior
tw = espessura da alma
t2b = largura do falnge inferior
tfb = espessura do falnge inferior
"""
#...............t3,  t2  tf   tw   t2b  tfb
propriedades = [31,15,2.5,0.635,15,2.5]


enderecos =["C:/Users/breno\Desktop/testes/eee.sdb"]

atual(propriedades,enderecos)


