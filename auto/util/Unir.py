#DEV 18/NOV/2022

import pandas as pd
import os

def unificarDiario(origen, destino, name):

    auxOrigen = os.listdir(origen)
    listXLSX = []
    for xlsx in auxOrigen:

        dir = origen +"/"+ xlsx
        df = pd.read_excel(dir)
        listXLSX.append(df)
        
    join = pd.concat(listXLSX)   
    join.to_excel(destino +"/"+ name)

def unificarDominicales(origen, destino, name):

    auxOrigen = os.listdir(origen)

    df_RV = pd.read_excel(origen +"/"+ auxOrigen[0])
    df_RB = pd.read_excel(origen +"/"+ auxOrigen[1])
    df_AA = pd.read_excel(origen +"/"+ auxOrigen[2])
    
    df_RV = df_RV.rename(columns={"DOCUMENTO":"CEDULA", 
                                    "NOMPDV":"NOMPDV",
                                    "ZONA":"ZONA",
                                    "NOMCCTO":"NOMCOSTO",
                                    "FECHA_VISITA":"FECHA"})

    df_RB = df_RB.rename(columns={"ID de Usuario":"CEDULA", 
                                    "Punto del evento":"NOMPDV",
                                    "Dispositivo":"ZONA",
                                    "Estado":"NOMCOSTO",
                                    "Tiempo":"FECHA"})

    df_AA = df_AA.rename(columns={"PERSONA":"CEDULA", 
                                    "SUCURSAL":"NOMPDV",
                                    "ZONA":"ZONA",
                                    "CCOSTO":"NOMCOSTO",
                                    "FECHA":"FECHA"})

    auxRB = df_RB[["CEDULA", "NOMPDV", "ZONA", "NOMCOSTO", "FECHA"]]
    auxAA = df_AA[["CEDULA", "NOMPDV", "ZONA", "NOMCOSTO", "FECHA", "HORA ENTRADA", "HORA SALIDA"]]
    auxRV = df_RV[["CEDULA", "NOMPDV", "ZONA", "NOMCOSTO", "FECHA"]]

    dataframes = [auxRB, auxAA, auxRV] 

    join = pd.concat(dataframes)
    join.to_excel(destino +"/"+ name)
