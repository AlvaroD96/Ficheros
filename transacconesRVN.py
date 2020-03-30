
"""
Creado el 31-01-2020

@author: Alvaro Duque
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
from math import isnan
from sqlalchemy import create_engine, types 
import math
from libreria_fdo import connect_database_user, query_database , get_schema_sql , disconnect_database,get_frame_sql_user, get_ndays_from_date, convert_string_to_date, convert_date_to_string, dataframe_join,get_frame_xl, print_full
from decimal import Decimal
from PyQt5 import uic
from PyQt5 import QtCore
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import sys
from datetime import datetime





def get_nex2_weekday(adate):
    '''
    Retorna la fecha en string del ultimo weekday dado una fecha en string.
    '''
    adate = convert_string_to_date(date_string = adate)
    adate += dt.timedelta(days=2)
    while adate.weekday() > 4: # Mon-Fri are 0-4
        adate += dt.timedelta(days=1)
    adate = convert_date_to_string(adate) 
    return adate

def get_next_weekday(adate):
	'''
	Retorna la fecha en string del proximo weekday dado una fecha en string.
	'''
	adate = convert_string_to_date(date_string = adate)
	adate += dt.timedelta(days=1)
	while adate.weekday() > 4: # Mon-Fri are 0-4
		adate += dt.timedelta(days=1)
	adate = convert_date_to_string(adate)
	return adate

def poblate_fondos(self):
    '''
    Llena la lista con los fondos disponibles en la base de datos
    '''
    self.model1.clear()
    
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "Select * From Fondos")
    for i in df["codigo_fdo"]:
        self.model1.appendRow(QStandardItem(i))

def convert_date_to_string(date):
    '''
    Retorna la fecha en formato de string, dado el objeto date.
    '''
    date_string=str(date.strftime('%Y-%m-%d'))
    return date_string

def open_workbook(path, screen_updating, visible):
    '''
    Abre un workbook y retorna un objecto workbook de xlwings. Sin screen_updating es false es mas rapido.
    '''
    wb = xw.Book(path)
    if screen_updating == False:        
        wb.app.screen_updating = False
    else:
        wb.app.screen_updating = True
    if visible == False:
        wb.app.visible = False
    else:
        wb.app.visible = True
    return wb

def transaccionesRVN(path, fecha):

   

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


    date = fecha[6:10] +'-'+ fecha[3:5]+'-'+ fecha[0:2]
    #instrumentos =  get_frame_sql_user("Puyehue", "MesaInversiones","usrConsultaComercial" , "Comercial1w","Select * from Instrumentos where  tipo_instrumento  not like '%RF LATAM%' AND tipo_instrumento  not like '%accion%' AND tipo_instrumento  not like '%cuota%' AND tipo_instrumento  not like '%ETF%'")

    instrumentos =  get_frame_sql_user("Puyehue", "MesaInversiones","usrConsultaComercial" , "Comercial1w","Select * from Instrumentos where tipo_instrumento  not like '%Bono%' AND tipo_instrumento  not like '%letra%' AND tipo_instrumento  not like '%deposito%' and tipo_instrumento != ';D15;' and Tipo_Instrumento != ''  AND tipo_instrumento  not like '%pagare%'  AND tipo_instrumento  not like '%Factura%'")
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select * from [MesaInversiones].[dbo].[Perfil Clientes] WHERE not Rut IS NULL")
    df = df.drop(["Nombre","Tipo","Orientacion","Perfil_riesgo","Tracking Objetivo","RVL","RVG","RFL","RFG","Liq","LIQ_USD","Fwd","Cuenta Pershing","RutConVerificador","Administracion","FechaTermino","Codigo_Recomendacion"],axis=1)

    nombres_ins = instrumentos["Codigo_Ins"].values.tolist()


    new_df = pd.DataFrame(columns=['Folio_operacion','Fecha operacion','1','2','3','Rut','secuencia','monto','Precio','Cantidad','4','Operacion','5','instrumento','6','Moneda','fecha_liq','7','ent_depositaria','liquidacion','8','9','10','11','12','13'], data=diccionario)
    new_df = new_df.drop(["1","2","3","4","5","6","7","8","9","10","11","12","13"],axis=1)
    
    new_df2 = new_df.copy()
    eliminados = 0
    eliminados2 = 0
    largo = len(new_df.index)
    for indice in range(largo):
        if not new_df.iloc[indice-eliminados]["instrumento"] in nombres_ins:
            new_df = new_df.drop(new_df.index[[indice - eliminados]])
            eliminados += 1
        else:
            new_df2 = new_df2.drop(new_df2.index[[indice - eliminados2]])
            eliminados2 += 1
    #print(new_df)
    if new_df.empty:  
        df1 = pd.DataFrame(columns = l)
        return df1 
    
    else: 
        fecha_operaciones = date
        cambio_diario = get_frame_sql_user("Puyehue", "replicasCredicorp", "usrConsultaComercial", "Comercial1w","select RTRIM(LTRIM(CODIGO_MONEDA_ORIGEN)) AS Moneda,VALOR_PARIDAD from VISTA_PARIDADES_CDP WHERE FECHA_PARIDAD = '{}' AND CODIGO_MONEDA = 'CLP' AND (CODiGO_MONEDA_ORIGEN = 'USD' OR CODiGO_MONEDA_ORIGEN = 'EUR') AND GRUPO_COTIZACION = 1  ".format(fecha_operaciones))

        df1 = pd.DataFrame(columns = ["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
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
                                    "codigo_uti","tir_mercado"])

        df1 = df1.assign(fecha_sesion = new_df['Fecha operacion'])
        
        vago = new_df["Rut"]
        
        new_df["codigo_secuencia"] = new_df.index
        
        df1 = df1.assign(codigo_secuencia = new_df.index) 
        
        df1 = df1.assign(estado = 'T')
        
        #df1 = df1.assign(fecha_sesion = new_df['Fecha operacion'])
        
        df1 = df1.assign(apertura_cierre = 'NULL')
        
        df1 = df1.assign(base_calculo = 3) # modificar


        df1 = df1.assign(broker = 'CCCAPITAL')

        df1 = df1.assign(cam_divisa = new_df['Moneda'])

        #print(cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']== 'USD'] )

        df1.loc[df1['cam_divisa'] == 'US$', 'cam_divisa'] = cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']==  'USD']
        df1.loc[df1['cam_divisa'] == '$',   'cam_divisa'] =   '1'
        df1.loc[df1['cam_divisa'] == 'UF', ' cam_divisa'] =  cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']==   'UF']
        df1.loc[df1['cam_divisa'] == 'EUR', 'cam_divisa'] =  cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']== 'EUR']




        df1 = df1.assign(canones_div = None)
        df = df.astype({"Rut":int,"Secuencia":int})
        new_df[['Rut','no']] = new_df['Rut'].str.split('-',expand=True)
        new_df = new_df.astype({"Rut":int,"secuencia":int})
        
        df["Dig Ver"] = df["Dig Ver"].str.strip()
        df_join = dataframe_join(df,new_df,'Rut')

        df_join = df_join.loc[(df_join["Secuencia"] == df_join["secuencia"])&(df_join["no"].to_string(index=False).upper() == df_join["Dig Ver"].to_string(index=False).upper() )]
    
        df1 = dataframe_join(df_join,df1,'codigo_secuencia')


        df1['Codigo_Fdo'] = df1['Codigo_Fdo'].str.strip() 
        
        instrumentos = instrumentos.rename(columns={"Codigo_Ins":"instrumento"}) 
        instrumentos["instrumento"] = instrumentos["instrumento"].str.strip()    
        instrumentos["Codigo_Emi"] = instrumentos["Codigo_Emi"].str.strip()
        
        df1 = dataframe_join(df1,instrumentos[["instrumento","Codigo_Emi"]],"instrumento") 
        
        
        df1 = df1.drop(["cartera"],axis=1)    
        df1 = df1.drop(["cotizacion"],axis=1)
        df1 = df1.drop(["instrumentox"],axis=1)
        df1 = df1.drop(["efectivo_div"],axis=1)
        df1 = df1.drop(["ent_liquidadora"],axis=1)    
        df1 = df1.drop(["nominal_div"],axis=1)    
        df1 = df1.drop(["operacion"],axis=1)

        df1 = df1.rename(columns={"Codigo_Emi":"ent_liquidadora"}) 
        
        df1 = df1.rename(columns={"Precio":"cotizacion"}) 

        #df1 = df1.assign(cotizacion =0 )
        
        df1 = df1.rename(columns={"monto":"efectivo_div"})     
        df1 = df1.rename(columns={"Cantidad":"nominal_div"}) 
        df1 = df1.rename(columns={"Operacion":"operacion"})
        df1 = df1.rename(columns={"Codigo_Fdo":"cartera"})


        
        df1 = df1.assign(comision_div = 0)
        df1 = df1.assign(corretaje_div = 0)
        df1 = df1.assign(cot_recompra = 0)


        df1 = df1.assign(cot_rf_cup = 0) # cotizacion renta fija con cupon



        df1 = df1.assign(cupon_corrido_div = None)


        df1 = df1.assign(cupon_per_div = 0)
        
        df1 = df1.assign(cupon_per = 0)


        df1 = df1.assign(dias=0)


      
        df1 = df1.assign(divisa = new_df['Moneda'])

        #print(cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']== 'USD'] )

        df1.loc[df1['divisa'] == 'US$', 'divisa'] = '5'
        df1.loc[df1['divisa'] == '$',   'divisa'] =  '39'
        df1.loc[df1['divisa'] == 'UF', ' divisa'] =  '84'
        df1.loc[df1['divisa'] == 'EUR', 'divisa'] =  '4'
        
        df1 = df1.assign(efectivo_vto_div =0) #  

        df1 = df1.assign(ent_contrapartida = 'CCCAPITAL') 

        df1 = df1.assign(ent_depositaria = 'DCV')     # deposito central de valores 
                                                    # sicav sociedad de inversion de capital variable
                                                    # sicaf socidad de inversion de capital fijo 



        df1 = df1.assign(ent_mediadora = None)  # consultar

        df1 = df1.assign(fec_comunicacion_be = None)  # consultar 
        ## ponerle el assign con el loc aca segun el tipo de liquidacion
        df1.loc[df1["liquidacion"] == "PH","fec_liquidacion"] = date
        df1.loc[df1["liquidacion"] == "PM","fec_liquidacion"] = get_next_weekday(date)
        df1.loc[df1["liquidacion"] == "PM","fec_liquidacion"] = get_nex2_weekday(date)   
        df1 = df1.assign(fec_operacion = date)        # consultar 
        df1 = df1.assign(fec_recompra = None)         # consultar 
        df1 = df1.assign(fec_valor = date)            # consultar 

        df1 = df1.assign(fec_vto = None)


        df1 = df1.assign(gasto_extra = 0)

        df1 = df1.assign(gastos_div = "")

        df1 = df1.assign(gestor_ordenante = 2960) # Consultar

        # consultar suma o resta segun que

        df1.loc[df1["operacion"] == "C",'ind_nominal_srn'] = "S" 
        df1.loc[df1["operacion"] == "V","ind_nominal_srn"] = "R" 

        df1 = df1.assign(ind_periodificar = None)


        
        df1 = df1.assign(liquido_div = df1['efectivo_div'])


        # Renta variable Nacional
        df1 = df1.assign(mercado = 'RVN')



        df1 = df1.assign(minusvalia = 0)
        df1 = df1.assign(minusvalia_div = 0)

        df1 = df1.assign(neto_div = df1['efectivo_div'])




        df1 = df1.assign(otros_gastos_div = '')

        df1 = df1.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que

        df1 = df1.assign(plusvalia = 0 )
        df1 = df1.assign(plusvalia_div = 0)

        df1 = df1.assign(prima_div = None)

        df1 = df1.assign(primario_secundario = 'P')
        df1 = df1.assign(regularizar_sn = 'S')     # CONDICIONES
        df1 = df1.assign(repo_dia = 'N')

        df1 = df1.assign(tipo_interes = 0)
        df1 = df1.assign(titulos = df1['nominal_div']) # consultar

        df1 = df1.assign(ventana_tratamiento = 'RVN') 


        df1 = df1.assign(ind_proceso = 'D')

        df1 = df1.assign(enlace_sucursal = '')
        df1 = df1.assign(enlace_tipo_ap = '')
        df1 = df1.assign(enlace_plan_subplan = '')


        df1 = df1.assign(retencion_div = 0)

        df1.loc[df1["liquidacion"] != "PH","fec_liquidacion"] = "S"
        df1.loc[df1["liquidacion"] == "PH","ind_cambio_asegurado"] = "N"

        df1 = df1.assign(ind_depositario = '')
        df1 = df1.assign(simple_compuesto = '')
        df1 = df1.assign(tex_pago = '')

        df1 = df1.assign(tex_pagareS_cod = None)
        df1 = df1.assign(tex_pagares = None)

        df1 = df1.assign(num_comunica_be = None)
        df1 = df1.assign(iva = None)
        df1 = df1.assign(iva_div =None)

        df1 = df1.assign(cuenta_ccc = '')

        df1 = df1.assign(ind_cobertura = None)
        df1 = df1.assign(fecha_saldo = '')
        df1 = df1.assign(libre_n1 = '')
        df1 = df1.assign(libre_n2 = '')
        df1 = df1.assign(libre_x1 = '')
        df1 = df1.assign(libre_x2 = '')
        df1 = df1.assign(fecha_ejecucion = '')
        df1 = df1.assign(codigo_uti = '')



        df1 = df1.assign(tir_mercado = 0.000000000) # consultar
        df1 = df1.drop(["Rut","Dig Ver","Secuencia","Status","Folio_operacion","Fecha operacion","secuencia","Moneda","liquidacion","no"],axis=1)
        
        columnas = list(df1.columns)
        for columna in columnas:
            if not columna in l:
                
                df1 = df1.drop(["{}".format(columna)],axis= 1)
        for col in l:
            if not col in df1.columns:
                pass
                #print(col)

        
        duplicateRowsDF = df1[df1.duplicated(['codigo_secuencia', 'efectivo_div'])]
        df1.drop_duplicates(subset =['codigo_secuencia','efectivo_div' ],keep = False, inplace = True)
        
        
        df1 = df1.reindex(columns=l)
        duplicateRowsDF = duplicateRowsDF.reindex(columns=l)
        df1["tir_mercado"] = None
        df1["comision_ecc_div"] = None
        
        data = pd.concat([df1,duplicateRowsDF])

        
        

        data["tir_mercado"] = None
        data["comision_ecc_div"] = None

        duplicateRowsDF1 = data[data.duplicated(['codigo_secuencia', 'efectivo_div'])]
        data.drop_duplicates(subset =['codigo_secuencia','efectivo_div' ],keep = False, inplace = True)
        
        
        data = data.reindex(columns=l)
        duplicateRowsDF1 = duplicateRowsDF1.reindex(columns=l)
        data["tir_mercado"] = None
        data["comision_ecc_div"] = None
        
        data1 = pd.concat([data,duplicateRowsDF1])
        #print_full(data1)
        
        print(data1['operacion'])
        #data1.to_csv("hola.csv", index=False,sep = ';',encoding='latin-1')

        return data1      
        

#transaccionesRVN('GEO_ap/Geo/2020-01-03geo.txt', '03-01-2020')