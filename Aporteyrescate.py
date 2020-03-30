
"""
Creado el 31-01-2020

@author: Luciano Aguilera
"""


import pandas as pd
import numpy as np
import datetime as dt
import xlwings as xw
import os
import random
import locale
from copy import copy
import unidecode
import timeit
import math
from libreria_fdo import *
from decimal import Decimal
from PyQt5 import uic
from PyQt5 import QtCore
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import sys
from datetime import datetime



def redondeo(numero,cifras):
    numero = str(numero).split(".")
    
    entero = numero[0]
    decimal = numero[1]
    decimal = decimal[:cifras]
    numero = float(entero) + float(decimal)*10**(-len(decimal))
    return numero


def aporteRescate(path):  
    
 





    dictionario = pd.DataFrame(columns = ['codigo' , 'fondo'])



    fondosc =  ['ESTRATEGIA','SPREADCORP','INTERNAC' ,'RF LATAM','GLOBALESI' ,'LAMERICAEQ','FI_ALLIANZ' ,'DEUDA CORP' ,'RENTA','MACRO CLP3' ,'MACRO 1.5' ,'DEUDA 360','ALTOREND' ,'LIQUIDEZ','INDICE' ,'ACCIONES','SMALLCAP' ,'M_MARKET','IMT E-PLUS' ,'PGRE SEC I','ACONCA-III' ,'ACONCA-II','DEBTII' ,'RENTA_RESI','RENTARESII' ,'FI PLAZA-E','PG PRIVATE' ,'DIRECT_III','FI CC SLP' ,'FUNDSICAV','PG SEC II' ,'PGDIRECTII','PRIVATE IC' ,'MONEDA_LATAM','LV PATIO' ,'FONDO LINK','DEUDA ARG' ,'AGG','E-PLUS']

    codigoss =   ['CFM8982','CFI7275','CFM8442','CFI7225','CFM8427','CFI9623','CFI9582','CFI9251','CFM8421','CFI9107','CFI9310','CFM9056','CFI9800','CFM8401','CFM8929','CFM8982','CFI7176','CFM8945','CFI9108','CFI9442','CFI7264','CFI9542','CFI9567','CFI9377','CFI9626','CFI9468','CFI7694','CFI9359','CFI9438','CFI9653','CFI9170','CFI7276','CFI9216','CFIIMDL','CFILVPA','CFILCPC','CFM9535','ADR AGG','CFM9310']


    dictionario = dictionario.assign(codigo = codigoss)
    dictionario = dictionario.assign(fondo = fondosc)



 







    l=["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
        "cam_divisa",	"canones_div",	"cartera",	"comision_div",	"corretaje_div"	,"cot_recompra",	
        "cot_rf_cup",	"cotizacion",	"cupon_corrido_div"	,"cupon_per",	"cupon_per_div"	,"dias",	"divisa",
        "efectivo_div",	"efectivo_vto_div"	,"ent_contrapartida",	"ent_depositaria",	"ent_liquidadora",
        "ent_mediadora"	,"fec_comunicacion_be",	"fec_liquidacion",	"fec_operacion"	,"fec_recompra",
        "fec_valor",	"fec_vto",	"gasto_extra",	"gastos_div",	"gestor_ordenante",	"ind_nominal_srn",
        "ind_periodificar",	"instrumento",	"liquido_div",	"mercado",	"minusvalia",	"minusvalia_div",
        "neto_div",	"nominal_div",	"operacion"	,"otros_gastos_div"	,"previa_compromiso_plazo",
        "plusvalia"	,"plusvalia_div",	"prima_div"	,"primario_secundario"	,"regularizar_sn",	"repo_dia",
        "tipo_interes",	"titulos",	"ventana_tratamiento",	"ind_proceso",	"enlace_sucursal",	"enlace_tipo_ap",
        "enlace_plan_subplan",	"retencion_div"	,"ind_cambio_asegurado"	,"ind_depositario",	"simple_compuesto",
        "tex_pago",	"tex_pagare_cod",	"tex_pagares",	"num_comunica_be",	"iva",	"iva_div",	"cuenta_ccc",
        "ind_cobertura",	"fecha_saldo",	"libre_n1",	"libre_n2",	"libre_x1",	"libre_x2",	"fecha_ejecucion",
        "codigo_uti",	"comision_ecc_div","tir_mercado"]


    excel = path
    diccionario = []
    contador = 0
    with open(excel,"r") as datos:
        for linea in datos:
            oracion = linea.strip()
            oracion = oracion.split(";")
            diccionario.append(oracion)


    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select * from [MesaInversiones].[dbo].[Perfil Clientes] WHERE not Rut IS NULL")
    df = df.drop(["Nombre","Tipo","Orientacion","Perfil_riesgo","Tracking Objetivo","RVL","RVG","RFL","RFG","Liq","LIQ_USD","Fwd","Cuenta Pershing","RutConVerificador","Administracion","FechaTermino","Codigo_Recomendacion"],axis=1)
    df_fondos = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select * from [MesaInversiones].[dbo].[Clasificacion_Fondos]")
    moneda_fondos = {}
    for fila in df_fondos.iterrows():
        moneda_fondos[fila[1]['Codigo_Fdo']] = fila[1]['Moneda']
            #rownum = self.main.wb.sheets["Ejecut"].range('A1').current_region.last_cell.row
    new_df = pd.DataFrame(columns=['Folio Mvto','Fecha','Folio VC','Fondo','Serie','Rut','secuencia','monto','Precio','Cantidad','1','Operacion','Tipo_Operac.','instrumento','Emisor','Reajuste','fecha_liq','Rescate','Boveda','liquidacion','2','3','4','5','Monto Total','Aport'], data=diccionario)

    new_df = new_df.drop(["instrumento","1","2","3","4","5","Emisor","Reajuste","fecha_liq","Rescate","Emisor","Reajuste","fecha_liq","Rescate","Boveda"],axis=1)
    fecha_operaciones = new_df.iloc[0]['Fecha']

    dict_valor_cuota = {}
    valor_cuota = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Fdo as Fondo,Codigo_Ser, Valor_Cuota as Valor from [MesaInversiones].[dbo].[zhis_series_rd] where fecha = '{}'".format(fecha_operaciones))
    for row in valor_cuota.iterrows():
        if row[1]["Fondo"].strip() in dict_valor_cuota.keys():
            dict_valor_cuota[row[1]["Fondo"].strip()][row[1]["Codigo_Ser"]] = float(row[1]["Valor"])
        else:
            dict_valor_cuota[row[1]["Fondo"].strip()] = {row[1]["Codigo_Ser"]: float(row[1]["Valor"])}

    
    
    query_paridad = "SELECT VALOR_PARIDAD FROM VISTA_PARIDADES_CDP WHERE CODIGO_MONEDA_ORIGEN = '{}' AND FECHA_PARIDAD = '{}' AND CODIGO_MONEDA ='{}'"
    cambio_diario = get_frame_sql_user("Puyehue", "replicasCredicorp", "usrConsultaComercial", "Comercial1w","select RTRIM(LTRIM(CODIGO_MONEDA_ORIGEN)) AS Moneda,VALOR_PARIDAD from VISTA_PARIDADES_CDP WHERE FECHA_PARIDAD = '{}' AND CODIGO_MONEDA = 'CLP' AND (CODiGO_MONEDA_ORIGEN = 'USD' OR CODiGO_MONEDA_ORIGEN = 'EUR')".format(fecha_operaciones))
    
    #print("select RTRIM(LTRIM(CODIGO_MONEDA_ORIGEN)) AS Moneda,VALOR_PARIDAD from VISTA_PARIDADES_CDP WHERE FECHA_PARIDAD = '{}' AND CODIGO_MONEDA = 'CLP' AND (CODiGO_MONEDA_ORIGEN = 'USD' OR CODiGO_MONEDA_ORIGEN = 'EUR')  AND GRUPO_COTIZACION = 1   ".format(fecha_operaciones))

    df6 = pd.DataFrame(columns = ["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
                                "cam_divisa",	"canones_div",	"cartera",	"comision_div",	"corretaje_div"	,"cot_recompra",
                                "cot_rf_cup",	"cotizacion",	"cupon_corrido_div",	"cupon_per"	,"cupon_per_div",	"dias",	"divisa",
                                "efectivo_div",	"efectivo_vto_div"	,"ent_contrapartida","ent_depositaria",	"ent_liquidadora",
                                "ent_mediadora"	,"fec_comunicacion_be",	"fec_liquidacion",	"fec_operacion"	,"fec_recompra",
                                "fec_valor",	"fec_vto",	"gasto_extra",	"gastos_div",	"gestor_ordenante",	"ind_nominal_srn",
                                "ind_periodificar",	"instrumentox",	"liquido_div",	"mercado",	"minusvalia",	"minusvalia_div",
                                "neto_div",	"nominal_div",	"operacion"	,"otros_gastos_div"	,"previa_compromiso_plazo",
                                "plusvalia"	,"plusvalia_div",	"prima_div"	,"primario_secundario"	,"regularizar_sn",	"repo_dia",
                                "tipo_interes",	"titulos",	"ventana_tratamiento",	"ind_proceso",	"enlace_sucursal",	"enlace_tipo_ap",
                                "enlace_plan_subplan",	"retencion_div"	,"ind_cambio_asegurado"	,"ind_depositario",	"simple_compuesto",
                                "tex_pago",	"tex_pagareS_cod",	"tex_pagares",	"num_comunica_be",	"iva",	"iva_div",	"cuenta_ccc",
                                "ind_cobertura",	"fecha_saldo",	"libre_n1",	"libre_n2",	"libre_x1",	"libre_x2",	"fecha_ejecucion",
                                "codigo_uti","tir_mercado",'folio', 'instru'])
    vago = new_df["Rut"]
    new_df["codigo_secuencia"] = new_df.index
    df6 = df6.assign(codigo_secuencia = new_df.index) 
    df6 = df6.assign(estado = 'T')

    df6 = df6.assign(fecha_sesion = new_df['Fecha'] )
    df6 = df6.assign(apertura_cierre = 'NULL')
    df6 = df6.assign(base_calculo = 3) # aca solo rv o rv local igual
    df6 = df6.assign(broker = 'CCCAPITAL')



    df6 = df6.assign(canones_div = None)

    df = df.astype({"Rut":int,"Secuencia":int})
    new_df[['Rut','no']] = new_df['Rut'].str.split('-',expand=True)
    new_df = new_df.astype({"Rut":int,"secuencia":int})
    df["Dig Ver"] = df["Dig Ver"].str.strip()
    df_join = dataframe_join(df,new_df,'Rut')
    df_join = df_join.loc[(df_join["Secuencia"] == df_join["secuencia"])&(df_join["no"] == df_join["Dig Ver"])]
    df6 = dataframe_join(df_join,df6,'codigo_secuencia')
    df6['Codigo_Fdo'] = df6['Codigo_Fdo'].str.strip()  
    df6['Serie'] = df6['Serie'].str.strip()  

    df6 = df6.drop(["cartera"],axis=1)    
    df6 = df6.drop(["cotizacion"],axis=1)
    df6 = df6.drop(["efectivo_div"],axis=1)
    df6 = df6.drop(["ent_liquidadora"],axis=1)    
    df6 = df6.drop(["nominal_div"],axis=1)    
    df6 = df6.drop(["operacion"],axis=1)
    df6 = df6.drop(["instrumentox"],axis=1)
    df6 = df6.assign(folio = new_df['Folio Mvto']) 

    df6 = df6.rename(columns={"Codigo_Emi":"ent_liquidadora"}) 
    #df6['ent_liquidadora'] = df6['ent_liquidadora'].str.strip()
    df6 = df6.rename(columns={"Precio":"cotizacion"}) 
    df6 = df6.rename(columns={"monto":"efectivo_div"})     
    df6 = df6.rename(columns={"Cantidad":"nominal_div"}) 
    df6 = df6.rename(columns={"Operacion":"operacion"})
    df6 = df6.rename(columns={"Codigo_Fdo":"cartera"})
    df6 = df6.rename(columns={"Fondo":"instrumento"})
    for row in df6.iterrows():
        #print(row[1]["instrumento"])
        cotizacion = float(row[1]["cotizacion"].split(",")[0])+ float(row[1]["cotizacion"].split(",")[1])*10**(-len(row[1]["cotizacion"].split(",")[1]))
        nominal = float(row[1]["nominal_div"].split(",")[0])+ float(row[1]["nominal_div"].split(",")[1])*10**(-len(row[1]["nominal_div"].split(",")[1]))
        #print(row)
        #print(cambio_diario)
        
        if cotizacion == 0 and nominal == 0:
            if '-' in row[1]["efectivo_div"]:
                row[1]["efectivo_div"] = row[1]["efectivo_div"].replace('-','0')
            else: 
                pass
            
            
            df6.loc[row[0],"cotizacion"] = float(dict_valor_cuota[row[1]['instrumento'].strip()][row[1]['Serie']])


            efectivo =  float(row[1]["efectivo_div"].split(",")[0])+ float(row[1]["efectivo_div"].split(",")[1])*10**(-len(row[1]["efectivo_div"].split(",")[1]))
            df6.loc[row[0],"nominal_div"] = efectivo/df6.loc[row[0],"cotizacion"]
            df6.loc[row[0],"nominal_div"] = str(redondeo(df6.loc[row[0],"nominal_div"],4)).replace(".",",")
            df6.loc[row[0],"cotizacion"] = str(redondeo(df6.loc[row[0],"cotizacion"],4)).replace(".",",")
   
        if moneda_fondos[row[1]["instrumento"].strip()] == 'CLP':
            df6.loc[row[0],"cam_divisa"] = 1
        elif moneda_fondos[row[1]["instrumento"].strip()] == 'EUR':

            numero = str(cambio_diario.loc[cambio_diario["Moneda"] == 'EUR',"VALOR_PARIDAD"].values[0]).replace(".",",")
            df6.loc[row[0],"cam_divisa"] = float(cambio_diario.loc[cambio_diario["Moneda"] == 'EUR',"VALOR_PARIDAD"])
        else:           
            numero = str(cambio_diario.loc[cambio_diario["Moneda"] == 'USD',"VALOR_PARIDAD"].values[0]).replace(".",",")
            df6.loc[row[0],"cam_divisa"] = numero

            

    df6 = df6.assign(comision_div = 0)
    df6 = df6.assign(corretaje_div = 0)
    df6 = df6.assign(cot_recompra = 0)


    df6 = df6.assign(cot_rf_cup = 0) # cotizacion renta fija con cupon



    df6 = df6.assign(cupon_corrido_div = None)
    df6 = df6.assign(instru = df6['instrumento'])

    df6 = df6.assign(cupon_per_div = 0)

    df6 = df6.assign(cupon_per = 0)



    # dias

    df6 = df6.assign(dias=0)


    df6 = df6.assign(divisa = df6['cam_divisa'])

    df6.loc[df6["divisa"] == 1,'divisa'] = '39' 
    df6.loc[(df6["divisa"] != '39') & (df6["divisa"] != 1),"divisa"] = "5" 
    

    
    df6 = df6.assign(efectivo_vto_div =0) # ojo aca 
    df6 = df6.assign(ent_liquidadora ="CHILE") # revisR ACA 
    df6 = df6.assign(ent_contrapartida = 'CCCAPITAL') 

    df6 = df6.assign(ent_depositaria = 'DCV')     # deposito central de valores 
                                                    # sicav sociedad de inversion de capital variable
                                                    # sicaf socidad de inversion de capital fijo 



    df6 = df6.assign(ent_mediadora = None)  # consultar

    df6 = df6.assign(fec_comunicacion_be = None)  # consultar 
    ## ponerle el assign con el loc aca segun el tipo de liquidacion
    df6.loc[df6["liquidacion"] == "PH","fec_liquidacion"] = fecha_operaciones 

    df6 = df6.assign(fec_operacion = fecha_operaciones)        # consultar 
    df6 = df6.assign(fec_recompra = None)         # consultar 
    df6 = df6.assign(fec_valor = fecha_operaciones)            # consultar 

    df6 = df6.assign(fec_vto = None)


    df6 = df6.assign(gasto_extra = 0)

    df6 = df6.assign(gastos_div ="")

    df6 = df6.assign(gestor_ordenante = 2960) # Consultar

    # consultar suma o resta segun que

    df6.loc[df6["operacion"] == "I",'ind_nominal_srn'] = "R" 
    df6.loc[df6["operacion"] == "R","ind_nominal_srn"] = "S" 



    df6.loc[df6["operacion"] == "I",'operacion'] = "C" 
    df6.loc[df6["operacion"] == "R","operacion"] = "V" 


    df6 = df6.assign(ind_periodificar = None)

    # si el emisor es central el nemo = BNPDBC + fec_vto si no revisamos si es UD$, $ o UF
    #  para agregar un pre = F*, FN o FU respectivamente con un guion mas la fecha de vencimiento

    df6 = df6.assign(liquido_div = df6['efectivo_div'])

    df6 = df6.assign(mercado = 'RVN') #

    df6 = df6.assign(minusvalia = 0)
    df6 = df6.assign(minusvalia_div = 0)

    df6 = df6.assign(neto_div = df6['efectivo_div'])




    df6 = df6.assign(otros_gastos_div = '')

    df6 = df6.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que

    df6 = df6.assign(plusvalia = 0 )
    df6 = df6.assign(plusvalia_div = 0)

    df6 = df6.assign(prima_div = None)

    df6 = df6.assign(primario_secundario = 'P')
    df6 = df6.assign(regularizar_sn = 'S')     ##### CONDICIONES
    df6 = df6.assign(repo_dia = 'N')

    df6 = df6.assign(tipo_interes = 0)
    df6 = df6.assign(titulos = df6['nominal_div']) # consultar

    df6 = df6.assign(ventana_tratamiento = 'RVN') # CONSULTAR

    df6 = df6.assign(ind_proceso = 'D')

    df6 = df6.assign(enlace_sucursal = '')
    df6 = df6.assign(enlace_tipo_ap = '')
    df6 = df6.assign(enlace_plan_subplan = '')


    df6 = df6.assign(retencion_div = 0)

    df6.loc[df6["liquidacion"] != "PH","fec_liquidacion"] = "S"
    df6.loc[df6["liquidacion"] == "PH","ind_cambio_asegurado"] = "N"
    df6.loc[df6["liquidacion"] == "PM","ind_cambio_asegurado"] = "S"

    df6 = df6.assign(ind_depositario = '')
    df6 = df6.assign(simple_compuesto = '')
    df6 = df6.assign(tex_pago = '')

    df6 = df6.assign(tex_pagareS_cod = None)
    df6 = df6.assign(tex_pagares = None)

    df6 = df6.assign(num_comunica_be = None)
    df6 = df6.assign(iva = None)
    df6 = df6.assign(iva_div =None)

    df6 = df6.assign(cuenta_ccc = '')

    df6 = df6.assign(ind_cobertura = None)
    df6 = df6.assign(fecha_saldo = '')
    df6 = df6.assign(libre_n1 = '')
    df6 = df6.assign(libre_n2 = '')
    df6 = df6.assign(libre_x1 = '')
    df6 = df6.assign(libre_x2 = '')
    df6 = df6.assign(fecha_ejecucion = '')
    df6 = df6.assign(codigo_uti = '')

    for i in  range(len(dictionario['codigo'])):

        df6.loc[df6['instru'].str.strip() == dictionario['fondo'].iloc[i].strip(),'instru'] = dictionario['codigo'].iloc[i] 


    
    df6['instrumento'] = df6['instru'] + df6['Serie']

    

    df6 = df6.assign(tir_mercado = 0.000000000) # consultar
    df6 = df6.drop(["Rut","Dig Ver","Secuencia","Status","secuencia","liquidacion","no"],axis=1)
    columnas = list(df6.columns)
    
    for columna in columnas:
        if not columna in l:
            df6 = df6.drop(["{}".format(columna)],axis= 1)
    
    for col in l:
        if not col in df6.columns:
           pass
           #print(col)
    
    df6 = df6.reindex(columns=l)
    df6["tir_mercado"] = None
    df6["comision_ecc_div"] = None
  

    #df6.to_excel('out.xlsx')
    print(df6['operacion'])
    return df6


#aporteRescate('GEO_ap/Ap/2020-01-08ap.txt')