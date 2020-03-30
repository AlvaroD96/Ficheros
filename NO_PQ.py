
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
import warnings


with warnings.catch_warnings():
    warnings.simplefilter(action='ignore', category=FutureWarning)

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

def resource_path(relative_path):
    '''
    Retorna el path absoluto del path entregado.
    '''
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def poblate_fondos(self):
    '''
    Llena la lista con los fondos disponibles en la base de datos
    '''
    self.model1.clear()
    
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "Select * From Fondos")
    for i in df["codigo_fdo"]:
        self.model1.appendRow(QStandardItem(i))

def float_to_string(numero):
    numero = str(numero).split(".")
    numero = ",".join(numero)
    return numero



def fichero(fechas):

    #excel = resource_path("Copia de Fichero Operaciones_v5.xlsx")
    #app = xw.App()
    #hoja = app.books.open(excel)
    #ordenada = get_frame_xl(hoja,"Hoja1",1,1,0)
    #ordenada = ordenada["Nombre campo"].tolist()

    #hoja.close()
    #app.quit()
    l = ["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
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
    




    date =fechas[6:10] +'-'+ fechas[3:5]+'-'+ fechas[0:2]

    cambio_diario = get_frame_sql_user("Puyehue", "replicasCredicorp", "usrConsultaComercial", "Comercial1w","select RTRIM(LTRIM(CODIGO_MONEDA_ORIGEN)) AS Moneda,VALOR_PARIDAD from VISTA_PARIDADES_CDP WHERE FECHA_PARIDAD = '{}' AND CODIGO_MONEDA = 'CLP' AND (CODiGO_MONEDA_ORIGEN != 'CLP') AND GRUPO_COTIZACION = 1 ".format(date))
    instrumentos_1 =  get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Emi AS emisor, Codigo_Ins AS Instrumento, Nombre_Instrumento AS name, Tipo_instrumento AS tipo from Instrumentos WHERE not Codigo_Ins in (SELECT Codigo_Ins FROM instrumentos_RD) AND Codigo_Emi != '' ")
    instrumentos_2 = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Emi AS emisor, Codigo_Ins AS Instrumento, Nombre_Instrumento AS name, Tipo_instrumento AS tipo from Instrumentos_RD  WHERE not Codigo_Ins in (SELECT Codigo_Ins FROM instrumentos) AND Codigo_Emi != '' ")
    instrumentos_3 = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Emi AS emisor, Codigo_Ins AS Instrumento, Nombre_Instrumento AS name, Tipo_instrumento AS tipo from Instrumentos  WHERE Codigo_Ins  in (SELECT Codigo_Ins FROM instrumentos_RD) AND Codigo_Emi != ''")
    instrumentos = pd.concat([instrumentos_1,instrumentos_2,instrumentos_3],ignore_index=True)
    for col in instrumentos.columns:
        instrumentos[col] = instrumentos[col].str.strip()
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select * from TransaccionesIRF where  Fecha= '{}'".format(date))
    #print("select * from TransaccionesIRF where  Fecha= '{}'".format(date))
    
    #funcion que elimina los espacios en blanco
    fondos = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Fdo AS operado from [Perfil Clientes]")
    #      D_PARIDAD   ID_MONEDA_ORIGEN  CODIGO_MONEDA_ORIGEN  ID_MONEDA   CODIGO_MONEDA   FECHA_PARIDAD      VALOR_PARIDAD  GRUPO_COTIZACION
    # 0       30137                84           UF                 39    CLP           2020-01-01       28310.86                 1
    # 1       30796                 5           USD                39    CLP           2020-01-01         749.83                 6

    df3 = pd.DataFrame(columns = ["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
                                "cam_divisa",	"canones_div",	"cartera",	"comision_div",	"corretaje_div"	,"cot_recompra",
                                "cot_rf_cup",	"cotizacion",	"cupon_corrido_div",	"cupon_per"	,"cupon_per_div",	"dias",	"divisa",
                                "efectivo_div",	"efectivo_vto_div"	,"ent_contrapartida","ent_depositaria",	"ent_liquidadora",
                                "ent_mediadora"	,"fec_comunicacion_be",	"fec_liquidacion",	"fec_operacion"	,"fec_recompra",
                                "fec_valor",	"fec_vto",	"gasto_extra",	"gastos_div",	"gestor_ordenante",	"ind_nominal_srn",
                                "ind_periodificar",	"instrumentox",	"liquido_div",	"mercado",	"minusvalia",	"minusvalia_div",
                                "neto_div",	"nominal_div",	"operacionx"	,"otros_gastos_div"	,"previa_compromiso_plazo",
                                "plusvalia"	,"plusvalia_div",	"prima_div"	,"primario_secundario"	,"regularizar_sn",	"repo_dia",
                                "tipo_interes",	"titulos",	"ventana_tratamiento",	"ind_proceso",	"enlace_sucursal",	"enlace_tipo_ap",
                                "enlace_plan_subplan",	"retencion_div"	,"ind_cambio_asegurado"	,"ind_depositario",	"simple_compuesto",
                                "tex_pago",	"tex_pagareS_cod",	"tex_pagares",	"num_comunica_be",	"iva",	"iva_div",	"cuenta_ccc",
                                "ind_cobertura",	"fecha_saldo",	"libre_n1",	"libre_n2",	"libre_x1",	"libre_x2",	"fecha_ejecucion",
                                "codigo_uti","tir_mercado"])
    df["operado"] = None
    df["operacion"] = None
    for row in df.iterrows():
        if "OD" in row[1]["Liq"]:
            row_2 = copy(row[1])
            row_2["operado"] = row[1]["Compra"]
            df.loc[row[0],["operado","operacion"]] = [row[1]["Vende"],"V"]
            row_2["operacion"] = "C"
            df = df.append(row_2)
        elif row[1]["Vende"]:
            df.loc[row[0],["operado","operacion"]] = [row[1]["Vende"],"V"]
        else:
            df.loc[row[0],["operado","operacion"]] = [row[1]["Compra"],"C"]
    df.reset_index(inplace=True, drop=True) 
    df["codigo_secuencia"] = df.index
    df3["codigo_secuencia"] = df.index
    #df3 = df3.assign(codigo_secuencia = new_df.index) 
    df3 = df3.assign(estado = 'T')
    df3 = df3.assign(fecha_sesion = date )
    df3 = df3.assign(apertura_cierre = 'NULL')
    df3 = df3.assign(base_calculo = 2) # aca solo rv o rv local igual
    df3 = df3.assign(broker = 'CCCAPITAL')

    #df3 = df3.assign(cam_divisa = 1)
    df3 = df3.assign(canones_div = None)
    df['operado'] = df['operado'].str.strip() 
    fondos['operado'] = fondos['operado'].str.strip() 

    df_join = dataframe_join(df,fondos,"operado")
    df3 = df3.drop(["instrumentox"],axis = 1)
    df_join = df_join.drop(["ID","V","OpV","C","OpC","Rte","Folio","D","Plazo","fecha_liq","Vende","Compra"],axis=1)
    df_join = dataframe_join(df_join,instrumentos,"Instrumento")
    df3 = dataframe_join(df_join,df3,"codigo_secuencia")

    for row in df3.iterrows():
        df3.loc[row[0],["dias","fec_vto"]] = [int(float(row[1]["Duration"])*365), get_ndays_from_today(int(float(row[1]["Duration"])*365))]
        
    df3['dias'].astype(int)
    df3 = df3.assign(base_calculo = df3['Moneda'])
    df3['base_calculo'] = df3['base_calculo'].str.strip() 
    df3.loc[df3['base_calculo'] == 'US$', 'base_calculo'] = '3'
    df3.loc[df3['base_calculo'] == 'CH$', 'base_calculo'] = '14'
    df3.loc[df3['base_calculo'] == 'UF', 'base_calculo'] = '3'


    df3 = df3.assign(cam_divisa = df3['Moneda'])
    df3['cam_divisa'] = df3['cam_divisa'].str.strip() 
    df3.loc[df3['cam_divisa'] == 'US$', 'cam_divisa'] = cambio_diario.loc[cambio_diario["Moneda"] == 'US$']["VALOR_PARIDAD"]
    df3.loc[df3['cam_divisa'] == 'CH$', 'cam_divisa'] = '1'
    df3.loc[df3['cam_divisa'] == 'UF', 'cam_divisa'] = cambio_diario.loc[cambio_diario["Moneda"] == 'UF']["VALOR_PARIDAD"]


    df3 = df3.assign(canones_div = 0)


    df3 = df3.assign(comision_div = 0)
    df3 = df3.assign(corretaje_div = 0)
    df3 = df3.assign(cot_recompra = 0)


    df3 = df3.assign(cot_rf_cup = 0) # cotizacion renta fija con cupon
    df3 = df3.drop(["cotizacion","cartera","operacionx"],axis = 1)
    df3 = df3.rename(columns={"operado":"cartera"}) 
    df3 = df3.assign(cotizacion = 0)
    df3 = df3.assign(cupon_corrido_div = 0)


    df3 = df3.assign(cupon_per_div = 0)



    # dias
    df3 = df3.assign(divisa = df3['Moneda'])
    df3['divisa'] = df3['divisa'].str.strip() 
    df3.loc[df3['divisa'] == 'US$', 'divisa'] = '5'
    df3.loc[df3['divisa'] == 'CH$', 'divisa'] = '39'
    df3.loc[df3['divisa'] == 'UF', 'divisa'] = '84'

    df3 = df3.assign(efectivo_div = df3['Monto'])# consultar

    df3 = df3.assign(efectivo_vto_div =0)

    df3 = df3.assign(ent_contrapartida = 'CCCAPITAL') 

    df3 = df3.assign(ent_depositaria = 'DCV')     # deposito central de valores 
                                                  # sicav sociedad de inversion de capital variable
                                                  # sicaf socidad de inversion de capital fijo 

    df3 = df3.assign(ent_liquidadora = df3['emisor'])   
    df3['ent_liquidadora'] = df3['ent_liquidadora'].str.strip() 

    df3 = df3.assign(ent_mediadora = 3000)  # consultar


    df3 = df3.assign(fec_comunicacion_be = None)  # consultar 
    df3 = df3.assign(fec_operacion = date)        # consultar 
    df3 = df3.assign(fec_recompra = date)         # consultar 
    df3 = df3.assign(fec_valor = date)            # consultar 
    df3.loc[df3["Liq"] == "PHOD","Liq"] = "PH"
    df3.loc[df3["Liq"] == "PMOD","Liq"] = "PM"
    df3.loc[df3["Liq"] == "CNOD","Liq"] = "CN"
    convert_string_to_date

    df3.loc[df3["Liq"] == "PH","fec_liquidacion"] = date
    df3.loc[df3["Liq"] == "PM","fec_liquidacion"] = convert_date_to_string(convert_string_to_date(get_next_weekday(date)))
    df3.loc[df3["Liq"] == "CN","fec_liquidacion"] = convert_date_to_string(convert_string_to_date(get_nex2_weekday(date)))


    df3 = df3.assign(gasto_extra = 'NULL')

    df3 = df3.assign(gastos_div = '')

    df3 = df3.assign(gestor_ordenante = 2960) # Consultar

    # consultar suma o resta segun que
    df3.loc[df3['operacion'] == "V"  , 'ind_nominal_srn'] = 'R'
    df3.loc[df3['operacion'] != "V" , 'ind_nominal_srn'] = 'S'




    df3 = df3.assign(ind_periodificar = 'NULL')





    # si el emisor es central el nemo = BNPDBC + fec_vto si no revisamos si es UD$, $ o UF
    #  para agregar un pre = F*, FN o FU respectivamente con un guion mas la fecha de vencimiento
    
    
    df3['Instrumento'] = df3['Instrumento'].str.strip() 

    df3 = df3.assign(liquido_div = df3['Monto'])

    df3 = df3.assign(mercado = 'RFN') # Renta Fija Nacional


    df3 = df3.assign(minusvalia = 0)
    df3 = df3.assign(minusvalia_div = 0)

    df3 = df3.assign(neto_div = df3['Monto'])

    df3 = df3.assign(nominal_div = df3['Cantidad']) # consultar

    # venta o compra segun la tasa de compra, si existe es compra si no es venta
    df3 = df3.assign(otros_gastos_div = '')

    df3 = df3.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que

    df3 = df3.assign(plusvalia = 0 )
    df3 = df3.assign(plusvalia_div = 0)

    df3 = df3.assign(prima_div = 'NULL')

    df3 = df3.assign(primario_secundario = 'P')
    df3 = df3.assign(regularizar_sn = 'S')     ##### CONDICIONES
    df3 = df3.assign(repo_dia = 'N')

    df3 = df3.assign(tipo_interes = 0)
    df3 = df3.assign(titulos =  df['Cantidad']) # consultar

    df3 = df3.assign(ventana_tratamiento = 'RFN') # CONSULTAR

    df3 = df3.assign(ind_proceso = 'D')

    df3 = df3.assign(enlace_sucursal = '')
    df3 = df3.assign(enlace_tipo_ap = '')
    df3 = df3.assign(enlace_plan_subplan = '')


    df3 = df3.assign(retencion_div = 0)

    df3 = df3.assign(ind_cambio_asegurado = '')
    df3 = df3.assign(ind_depositario = '')
    df3 = df3.assign(simple_compuesto = '')
    df3 = df3.assign(tex_pago = '')

    df3 = df3.assign(tex_pagare_cod = 'NULL')
    df3 = df3.assign(tex_pagares = 'NULL')

    df3 = df3.assign(num_comunica_be = 'NULL')
    df3 = df3.assign(iva = 'NULL')
    df3 = df3.assign(iva_div =' NULL')

    df3 = df3.assign(cuenta_ccc = '')

    df3 = df3.assign(ind_cobertura = 'NULL')
    df3 = df3.assign(fecha_saldo = '')
    df3 = df3.assign(libre_n1 = '')
    df3 = df3.assign(libre_n2 = '')
    df3 = df3.assign(libre_x1 = '')
    df3 = df3.assign(libre_x2 = '')
    df3 = df3.assign(fecha_ejecucion = '')
    df3 = df3.assign(codigo_uti = '')
    df3 = df3.assign(comision_ecc_div = '')

    df3 = df3.assign(tir_mercado = 0.000000000) # consult
    df3 = df3.rename(columns = {"Instrumento":"instrumento"})

    df3 = df3.assign(tir_mercado = 0.000000000) # consultar
    columnas = list(df3.columns)
    for columna in columnas:
        if not columna in l:
            df3 = df3.drop(["{}".format(columna)],axis= 1)
    for col in l:
        if not col in df3.columns:
            pass
    df3 = df3.reindex(columns=l)
    df3["tir_mercado"] = None
    df3["comision_ecc_div"] = None
    for row in df3.iterrows():
        df3.loc[row[0],["efectivo_div","liquido_div","neto_div","nominal_div","titulos"]] = [float_to_string(row[1]["efectivo_div"]),float_to_string(row[1]["liquido_div"]),float_to_string(row[1]["neto_div"]),float_to_string(row[1]["nominal_div"]),float_to_string(row[1]["titulos"])] 
    #print('df3 listo')
    
    #date = fechas

    cambio_diario = get_frame_sql_user("Puyehue", "replicasCredicorp", "usrConsultaComercial", "Comercial1w","select RTRIM(LTRIM(CODIGO_MONEDA_ORIGEN)) AS Moneda,VALOR_PARIDAD from VISTA_PARIDADES_CDP WHERE FECHA_PARIDAD = '{}' AND CODIGO_MONEDA = 'CLP' AND (CODiGO_MONEDA_ORIGEN != 'CLP') AND GRUPO_COTIZACION = 1 ".format(date))
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select * from TransaccionesIIF where  Fecha= '{}'".format(date))
    #funcion que elimina los espacios en blanco
    fondos = get_frame_sql_user("Puyehue", "MesaInversiones", "usrConsultaComercial", "Comercial1w","select Codigo_Fdo AS operado from [Perfil Clientes]")
    #      D_PARIDAD   ID_MONEDA_ORIGEN  CODIGO_MONEDA_ORIGEN  ID_MONEDA   CODIGO_MONEDA   FECHA_PARIDAD      VALOR_PARIDAD  GRUPO_COTIZACION
    # 0       30137                84           UF                 39    CLP           2020-01-01       28310.86                 1
    # 1       30796                 5           USD                39    CLP           2020-01-01         749.83                 6

   
    df4 = pd.DataFrame(columns = ["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	
                                "cam_divisa",	"canones_div",	"cartera",	"comision_div",	"corretaje_div"	,"cot_recompra",
                                "cot_rf_cup",	"cotizacion",	"cupon_corrido_div",	"cupon_per"	,"cupon_per_div",	"divisa",
                                "efectivo_div",	"efectivo_vto_div"	,"ent_contrapartida","ent_depositaria",	"ent_liquidadora",
                                "ent_mediadora"	,"fec_comunicacion_be",	"fec_liquidacion",	"fec_operacion"	,"fec_recompra",
                                "fec_valor",	"fec_vto",	"gasto_extra",	"gastos_div",	"gestor_ordenante",	"ind_nominal_srn",
                                "ind_periodificar",	"instrumentox",	"liquido_div",	"mercado",	"minusvalia",	"minusvalia_div",
                                "neto_div",	"nominal_div",	"operacionx"	,"otros_gastos_div"	,"previa_compromiso_plazo",
                                "plusvalia"	,"plusvalia_div",	"prima_div"	,"primario_secundario"	,"regularizar_sn",	"repo_dia",
                                "tipo_interes",	"titulos",	"ventana_tratamiento",	"ind_proceso",	"enlace_sucursal",	"enlace_tipo_ap",
                                "enlace_plan_subplan",	"retencion_div"	,"ind_cambio_asegurado"	,"ind_depositario",	"simple_compuesto",
                                "tex_pago",	"tex_pagareS_cod",	"tex_pagares",	"num_comunica_be",	"iva",	"iva_div",	"cuenta_ccc",
                                "ind_cobertura",	"fecha_saldo",	"libre_n1",	"libre_n2",	"libre_x1",	"libre_x2",	"fecha_ejecucion",
                                "codigo_uti","tir_mercado"] )
    df["operado"] = None
    df["operacion"] = None
    for row in df.iterrows():
        if "OD" in row[1]["Liq"]:
            row_2 = copy(row[1])
            row_2["operado"] = row[1]["Compra"]
            df.loc[row[0],["operado","operacion"]] = [row[1]["Vende"],"V"]
            row_2["operacion"] = "C"
            df = df.append(row_2)
        elif row[1]["Vende"]:
            df.loc[row[0],["operado","operacion"]] = [row[1]["Vende"],"V"]
        else:
            df.loc[row[0],["operado","operacion"]] = [row[1]["Compra"],"C"]
    df.reset_index(inplace=True, drop=True) 
    df["codigo_secuencia"] = df.index
    df4["codigo_secuencia"] = df.index
    #df4 = df4.assign(codigo_secuencia = new_df.index) 
    df4 = df4.assign(estado = 'T')
    df4 = df4.assign(fecha_sesion = date )
    df4 = df4.assign(apertura_cierre = 'NULL')
    df4 = df4.assign(base_calculo = 14) # aca solo rv o rv local igual
    df4 = df4.assign(broker = 'CCCAPITAL')

    #df4 = df4.assign(cam_divisa = 1)
    df4 = df4.assign(canones_div = None)
    df['operado'] = df['operado'].str.strip() 
    fondos['operado'] = fondos['operado'].str.strip() 

    df_join = dataframe_join(df,fondos,"operado")
    df4 = df4.drop(["instrumentox"],axis = 1)
    df_join = df_join.drop(["ID","V","OpV","C","OpC","Rte","Folio","D","fecha_liq","Vende","Compra"],axis=1)
    df4 = dataframe_join(df_join,df4,"codigo_secuencia")

    for row in df4.iterrows():
        df4.loc[row[0],["fec_vto"]] = get_ndays_from_today(int(float(row[1]["dias"])*365))

    df4['dias'].astype(int)
    df4 = df4.assign(base_calculo = df4['Moneda'])
    df4['base_calculo'] = df4['base_calculo'].str.strip() 
    df4.loc[df4['base_calculo'] == 'US$', 'base_calculo'] = '3'
    df4.loc[df4['base_calculo'] == '$', 'base_calculo'] = '14'
    df4.loc[df4['base_calculo'] == 'UF', 'base_calculo'] = '3'


    df4 = df4.assign(cam_divisa = df4['Moneda'])
    df4['cam_divisa'] = df4['cam_divisa'].str.strip() 
    df4.loc[df4['cam_divisa'] == 'US$', 'cam_divisa'] = cambio_diario.loc[cambio_diario["Moneda"] == 'US$']["VALOR_PARIDAD"]
    df4.loc[df4['cam_divisa'] == '$', 'cam_divisa'] = '1'
    df4.loc[df4['cam_divisa'] == 'UF', 'cam_divisa'] = cambio_diario.loc[cambio_diario["Moneda"] == 'UF']["VALOR_PARIDAD"]



    df4 = df4.assign(canones_div = 0)


    df4 = df4.assign(comision_div = 0)
    df4 = df4.assign(corretaje_div = 0)
    df4 = df4.assign(cot_recompra = 0)


    df4 = df4.assign(cot_rf_cup = 0) # cotizacion renta fija con cupon
    df4 = df4.drop(["cotizacion","cartera","operacionx"],axis = 1)
    df4 = df4.rename(columns={"operado":"cartera"}) 
    df4 = df4.assign(cotizacion = 0)
    df4 = df4.assign(cupon_corrido_div = None)


    df4 = df4.assign(cupon_per_div = 0)



    # dias
    df4 = df4.assign(divisa = df4['Moneda'])
    df4['divisa'] = df4['divisa'].str.strip() 
    df4.loc[df4['divisa'] == 'US$', 'divisa'] = 5
    df4.loc[df4['divisa'] == '$', 'divisa'] = 39
    df4.loc[df4['divisa'] == 'UF', 'divisa'] = 84

    df4 = df4.assign(efectivo_div = df4['rescate'])# consultar

    df4 = df4.assign(efectivo_vto_div =0)

    df4 = df4.assign(ent_contrapartida = 'CCCAPITAL') 

    df4 = df4.assign(ent_depositaria = 'DCV')     # deposito central de valores 
                                                  # sicav sociedad de inversion de capital variable
                                                  # sicaf socidad de inversion de capital fijo 

    df4 = df4.assign(ent_liquidadora = df4['emisor'])   
    df4['ent_liquidadora'] = df4['ent_liquidadora'].str.strip() 

    df4 = df4.assign(ent_mediadora = 3000)  # consultar


    df4 = df4.assign(fec_comunicacion_be = None)  # consultar 
    df4 = df4.assign(fec_operacion = date)        # consultar 
    df4 = df4.assign(fec_recompra = date)         # consultar 
    df4 = df4.assign(fec_valor = date)            # consultar 
    df4.loc[df4["Liq"] == "PHOD","Liq"] = "PH"
    df4.loc[df4["Liq"] == "PMOD","Liq"] = "PM"
    df4.loc[df4["Liq"] == "CNOD","Liq"] = "CN"


    df4.loc[df4["Liq"] == "PH","fec_liquidacion"] = date
    df4.loc[df4["Liq"] == "PM","fec_liquidacion"] = get_next_weekday(date)
    df4.loc[df4["Liq"] == "CN","fec_liquidacion"] = get_nex2_weekday(date)  


    df4 = df4.assign(gasto_extra = 'NULL')

    df4 = df4.assign(gastos_div = '')

    df4 = df4.assign(gestor_ordenante = 2960) # Consultar

    # consultar suma o resta segun que
    df4.loc[df4['operacion'] == "V"  , 'ind_nominal_srn'] = 'R'
    df4.loc[df4['operacion'] != "C" , 'ind_nominal_srn'] = 'S'




    df4 = df4.assign(ind_periodificar = 'NULL')





    # si el emisor es central el nemo = BNPDBC + fec_vto si no revisamos si es UD$, $ o UF
    #  para agregar un pre = F*, FN o FU respectivamente con un guion mas la fecha de vencimiento
    
    
    df4['Instrumento'] = df4['Instrumento'].str.strip() 

    df4 = df4.assign(liquido_div = df4['rescate'])

    df4 = df4.assign(mercado = 'RFN') # Renta Fija Nacional


    df4 = df4.assign(minusvalia = 0)
    df4 = df4.assign(minusvalia_div = 0)

    df4 = df4.assign(neto_div = df4['rescate'])

    
    df4 = df4.assign(nominal_div = 1) # consultar
    # venta o compra segun la tasa de compra, si existe es compra si no es venta
    df4 = df4.assign(otros_gastos_div = '')

    df4 = df4.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que

    df4 = df4.assign(plusvalia = 0 )
    df4 = df4.assign(plusvalia_div = 0)

    df4 = df4.assign(prima_div = 'NULL')

    df4 = df4.assign(primario_secundario = 'P')
    df4 = df4.assign(regularizar_sn = 'S')     ##### CONDICIONES
    df4 = df4.assign(repo_dia = 'N')

    df4 = df4.assign(tipo_interes = 0)
    df4 = df4.assign(titulos = 1) # consultar

    df4 = df4.assign(ventana_tratamiento = 'PAG') # CONSULTAR

    df4 = df4.assign(ind_proceso = 'D')

    df4 = df4.assign(enlace_sucursal = '')
    df4 = df4.assign(enlace_tipo_ap = '')
    df4 = df4.assign(enlace_plan_subplan = '')


    df4 = df4.assign(retencion_div = 0)

    df4 = df4.assign(ind_cambio_asegurado = '')
    df4 = df4.assign(ind_depositario = '')
    df4 = df4.assign(simple_compuesto = '')
    df4 = df4.assign(tex_pago = '')

    df4 = df4.assign(tex_pagare_cod = 'NULL')
    df4 = df4.assign(tex_pagares = 'NULL')

    df4 = df4.assign(num_comunica_be = 'NULL')
    df4 = df4.assign(iva = 'NULL')
    df4 = df4.assign(iva_div =' NULL')

    df4 = df4.assign(cuenta_ccc = '')

    df4 = df4.assign(ind_cobertura = 'NULL')
    df4 = df4.assign(fecha_saldo = '')
    df4 = df4.assign(libre_n1 = '')
    df4 = df4.assign(libre_n2 = '')
    df4 = df4.assign(libre_x1 = '')
    df4 = df4.assign(libre_x2 = '')
    df4 = df4.assign(fecha_ejecucion = '')
    df4 = df4.assign(codigo_uti = '')
    df4 = df4.assign(comision_ecc_div = '')

    df4 = df4.assign(tir_mercado = 0.000000000) # consult
    df4 = df4.rename(columns = {"Instrumento":"instrumento"})

    df4 = df4.assign(tir_mercado = 0.000000000) # consultar
    columnas = list(df4.columns)
    for columna in columnas:
        if not columna in l:
            df4 = df4.drop(["{}".format(columna)],axis= 1)
    for col in l:
        if not col in df4.columns:
            pass
    df4 = df4.reindex(columns=l)
    df4["tir_mercado"] = None
    df4["comision_ecc_div"] = None
    for row in df4.iterrows():
        df4.loc[row[0],["efectivo_div","liquido_div","neto_div","nominal_div","titulos"]] = [float_to_string(row[1]["efectivo_div"]),float_to_string(row[1]["liquido_div"]),float_to_string(row[1]["neto_div"]),float_to_string(row[1]["nominal_div"]),float_to_string(row[1]["titulos"])] 
   

    
    con = pd.concat([df3,df4])
    
    con = con.reindex(columns=l)
    

    print(con['operacion'])
    return con
    #con.to_csv("no_pq.csv", index=False,sep = ';',encoding='latin-1')
    


