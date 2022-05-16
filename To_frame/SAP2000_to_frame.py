import os
import win32com.client
import time
import pandas as pd



#create Sap2000 object
SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")

#start Sap2000 application
SapObject.ApplicationStart()

#create SapModel object
SapModel = SapObject.SapModel




#Pegar dados de da aréa e carregamento delas
# Mesclar esse dados em um unico DataFrame



# Fazer a relação de Barras pelas áreas ---> bem dificil
# Verificar menor vão ou se é quadrada
# Calcular um vetor das cargas
#  Aplicar a carga na barra (talvez tenha q procurar o nome, ai tem q importar a tabela de nome das seções)

