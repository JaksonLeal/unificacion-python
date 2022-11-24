#DEV 18/NOV/2022

import os
import util.Convertir as conv
import util.DeleteRows as elim
import util.Unir as unir

nomReporDiario ="2. apertura_asesores_venta.xlsx"
nomReporteUnificado="reporte_unificado.xlsx"

carpetaOrigen = "C:/Users/RSOC1098799204/Desktop/dominicales/descarga"
contenidoOrigen = os.listdir(carpetaOrigen)

carpetaDestino = "C:/Users/RSOC1098799204/Desktop/dominicales/generado"
contenidoDestino = os.listdir(carpetaDestino)

for i in range(2):

    carpOrigen = carpetaOrigen +"/"+ contenidoOrigen[i]
    contenido = os.listdir(carpOrigen)
    cont = 0
    for xls  in contenido:

        dir = carpOrigen +"/"+ xls 
        if(xls.endswith("xls")):
            name = carpetaDestino +"/"+ contenidoDestino[i] +f"/{cont}. "+ xls + str("x")    
            conv.ConvToXLSX(dir, name)
            elim.eliminarFilas(name)
            cont += 1

        elif(xls.endswith("xlsx")):
            name = carpetaDestino +"/"+ contenidoDestino[i] +f"/{cont}. "+ xls
            conv.ConvToXLSX(dir, name)
            elim.eliminarFilas(name)
            cont += 1

conv.apagarJVM()  
  
auxOrigen = carpetaDestino + "/2. reporte_diario"
auxDestino = "../generado/1. reporte_mensual"
unir.unificarDiario(auxOrigen, auxDestino, nomReporDiario)

auxDestino = "../generado"
auxOrigen = carpetaDestino + "/1. reporte_mensual"
unir.unificarDominicales(auxOrigen, auxDestino, nomReporteUnificado)

