import defRPAselenium
import moldesTerrenos
import modelsUnidadesvendidas
import models
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
import time

browser = Selenium()
Dt=moldesTerrenos.masterlibros()
try:
 #moldesTerrenos.task_Modelos()
 #moldesTerrenos.Asignaconsultafecha()
 #moldesTerrenos.task_Modelos()
 #moldesTerrenos.Asignaconsultafecha()
 pass
except:
    pass

urlbase=defRPAselenium.Pyasset(asset="base")
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")

def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida") 
        rmtree('Excel')     
        print("Eliminamos carpetas")

    except:
        pass

def Creacionescarpetas():
    print("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir('Excel')
        os.mkdir("Formato Solicitud")  
        os.mkdir("Log Scraping")
        os.mkdir("Salida") 

    except:
        pass

def task():
    
        for dtable in Dt:
            if dtable[15] == "SI":

                Rut=str(dtable[3])
                Inmobiliaria=dtable[1]
                Asset=dtable[2]
                
                Carpeta=str(dtable[2]+" "+dtable[7]+" -"+dtable[8])
                Hoja=dtable[8]
                Activo=dtable[5]
                region=dtable[1]
                rolmatriz=dtable[7]
                rol1=dtable[7]                               
                rol2=dtable[8]
                Codigo=dtable[14]
                comuna=dtable[14]

                cantidad=0
                consulta=True
                while consulta==True:

                    defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                    try:
                        tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                        stado=False
                        consulta=False
                    
                    except:
                        defRPAselenium.cerraNavegador()
                    
                    #Tercer reintento para garantizar continuidad si encuentra Recatchat
                        try: 
                            print("segundo reintento ") 
                            time.sleep(3)
                            tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                            stado=False
                            consulta=False
                        except:
                            
                            print("tercer reintento ") 
                            defRPAselenium.cerraNavegador()
                            try:
                                time.sleep(3)
                                tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                                stado=False
                                consulta=False
                                
                            except:
                                defRPAselenium.cerraNavegador()
                                pass
                        
                        finally:
                            pass   
                            defRPAselenium.cerraNavegador() 
                if consulta==stado :
                 consulta=False 
                else:
                    defRPAselenium.cerraNavegador()


def tgc():
 task()                   
    
         
if __name__ == "__main__":
   
   eliminarcarpetas()
   Creacionescarpetas()
   defRPAselenium.bakup()
   
   tgc()

   modelsUnidadesvendidas.task()
   defRPAselenium.salida()
   print('Ejecucion finalizada')
   
 
 





