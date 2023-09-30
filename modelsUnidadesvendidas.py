from RPA.Tables import Tables
from RPA.Excel.Files import Files;
lib = Files()
library = Tables()

def masterlibros():
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    DtableFinal=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    return DtableFinal

def Consolidado(txt,nombre_consolidado,Rut_Cliente,Region,comuna,ROMATRIZ,Rol_Unidad):
    
    ubicacion="Salida/resumen unidades vendidas.xlsx"
    nombre_consolidado_HOJA=nombre_consolidado[0:31]
    f=open("Log Scraping/"+txt+".txt","r")
    f=f.read()
    w=open("CSV/"+txt+".csv","w")
    w.write(f)
    w.close()

    orders = library.read_table_from_csv(
        "CSV/"+txt+".csv",header=False,delimiters=" ")


    CapturaSCRAPIADO= ([{ }])
        
            

    for x in orders:
        if len(x[0]) != 0:
            CUOTA=x[0]
            VALOR_CUOTA=x[1]
            NRO_FOLIO=x[2]
            VENCIMIENTO=x[3]
            TOTAL_A_PAGAR=str(x[4]).replace(".","")
            EMAIL=x[5]
            CapturaSCRAPIADO.append({
                'CUOTA':CUOTA,
                'VALOR CUOTA':int(VALOR_CUOTA), 
                'NRO FOLIO':NRO_FOLIO,
                'VENCIMIENTO':VENCIMIENTO,
                'TOTAL A PAGAR':int(TOTAL_A_PAGAR),
                'EMAIL':EMAIL,
                'NOMBRE CONSOLIDADO':nombre_consolidado,
                'RUT CLIENTE':Rut_Cliente,
                'REGION':Region,
                'COMUNA':comuna,
                'ROL MATRIZ':ROMATRIZ,
                'ROL UNIDAD':Rol_Unidad            
                    })
            
    try:lib.open_workbook(ubicacion) 
    except TypeError: lib.create_workbook(ubicacion) 
    except FileNotFoundError: lib.create_workbook(ubicacion) 

    if lib.worksheet_exists(nombre_consolidado_HOJA)==True:
        print("existe")
    else:
        print("no existe crea uno nuevo")
        lib.create_worksheet(nombre_consolidado_HOJA)



    #emcabezado
    lib.set_cell_value(1,1,"CUOTA'")
    lib.set_cell_value(1,2,"VALOR CUOTA'")
    lib.set_cell_value(1,3,"NRO FOLIO'")
    lib.set_cell_value(1,4,"VENCIMIENTO'")
    lib.set_cell_value(1,5,"TOTAL A PAGAR'")
    lib.set_cell_value(1,6,"EMAIL'")
    lib.set_cell_value(1,7,"NOMBRE CONSOLIDADO'")
    lib.set_cell_value(1,8,"RUT CLIENTE'")
    lib.set_cell_value(1,9,"REGION'")
    lib.set_cell_value(1,10,"COMUNA'")
    lib.set_cell_value(1,11,"ROL MATRIZ'")
    lib.set_cell_value(1,12,"ROL UNIDAD'")

    #Introduccimos los datos con append  
    lib.set_active_worksheet(nombre_consolidado_HOJA)
    lib.append_rows_to_worksheet(content=CapturaSCRAPIADO,header=False,start=1)

    #Limpiamos la data
    lib.set_active_worksheet(nombre_consolidado_HOJA)
    lib.read_worksheet(nombre_consolidado_HOJA)       #nombre de la hoja
    lista=lib.read_worksheet_as_table(name=nombre_consolidado_HOJA,header=True, start=1).data
    registros=(lib.find_empty_row()*10)

    #eliminamos los espacios vacios en las celdas
    for celdas in range(registros) :
        Buscar_vacias =lib.get_cell_value(1+celdas,"A")
        Buscar_Totales =lib.get_cell_value(1+celdas,"D")
        if Buscar_vacias is None or Buscar_Totales=="total":
            lib.delete_rows(celdas+1)

    #hacemos un segundo barrido
    for celdas in range(registros):
        Buscar_vacias =lib.get_cell_value(1+celdas,"A")
        if Buscar_vacias is None or Buscar_Totales=="total":
            lib.delete_rows(celdas+1)
    
   

    #creacion de la funcion de totales
    def total():
        a=0
        subtotal=0
        while a<2000: 
            lib.set_cell_format(1+a,"E",fmt=0.00)
            Monto=lib.get_cell_value(1+a,"B")            
            if a == 10 or a == 22 or a == 34 or a == 46 or a == 58 or a == 70 or a == 82 or a == 94 or a == 106 or a == 118:
                lib.insert_rows_before(row=2+a)
                lib.set_cell_value(2+a,"E",subtotal)         
                lib.set_cell_value(2+a,"D","total")  
                
                subtotal=0
                Monto=0

                a=3+a
                lib.save_workbook()

            else:
                a=1+a 
                try:subtotal = Monto + subtotal
                except:pass  
            lib.save_workbook()
                #Cuando no encontramos valores
            if Monto is None:
                print(a)
                print(subtotal)
                lib.insert_rows_before(row=2+a)
                lib.set_cell_value(2+a,"E",subtotal)         
                lib.set_cell_value(2+a,"D","total")
                
                try:     
                    subtotal = Monto + Monto
                except:
                    pass
                lib.set_cell_value(2+a,"E",subtotal)                   
                subtotal=0
                a=3+a
                lib.save_workbook()
                break
    #ejecutamos la funcion de totales
    total()      





    lib.save_workbook(ubicacion)

def task():
    Dt=masterlibros()
    for dtable in Dt:
            if dtable[15] == "SI":
             try:
                Rut=str(dtable[5])
                nombre_consolidado=str(dtable[3])
                Inmobiliaria=dtable[1]
                Asset=dtable[2]
                region=dtable[1]
                comuna_=dtable[2]
                Carpeta=str(dtable[2]+" "+dtable[7]+" -"+dtable[8])
                Hoja=dtable[8]
                Activo=dtable[5]
                region=dtable[1]
                rolmatriz=dtable[7]
                rol1=dtable[7]                               
                rol2=dtable[8]
                Codigo=dtable[14]
                comuna=dtable[14]
             except TypeError:print("no encontro nada .")
                

            try:Consolidado(Carpeta,nombre_consolidado,Rut,region,comuna_,rol1,rol2)
            except FileNotFoundError: print("El archivo ,"+ Carpeta + ".txt → No fue contrado ")
            except TypeError:print("no encontro nada .")
                




            
     
