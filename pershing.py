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


def Pershing(fecha):
  


    strcuentas = ['HMT082415','HMT090053','HMT090418','HMT090830','HMT090848','HMT090921','HMT090939','HMT090947','HMT090962','HMT090970','HMT100019','HMT100043','HMT100068','HMT100092','HMT270002','HMT390008','HMT090996','HMT100100','HMT091010','HMT091036','HMT086655','HMT086663','HMT086671','HMT086689','HMT087083','HMT091028','HMT091044','HMT087430']


    strSQL = "SET DATEFORMAT DMY SELECT mov.ID , mov.Numero_Cuenta, cue.Nombre_cuenta,mov.Instrumento,mov.Simbolo_Instrumento, ins.CUSIP, ins.Codigo_Ins, ins.Nombre, ins.Codigo_Emi ,mov.Precio, mov.Simbolo_Instrumento, mov.Codigo_compraventa,mov.Cantidad, mov.Valorizacion , mov.Moneda, mov.fecha_transaccion, mov.fecha_liquidacion FROM [CONCILIACION_INTERNACIONAL].[dbo].[TBL_MOVIMIENTO] mov INNER JOIN [CONCILIACION_INTERNACIONAL].[dbo].[TBL_CUENTAS_INTERNACIONALES] cue ON mov.Numero_Cuenta=cue.Numero_Cuenta LEFT OUTER JOIN [lagunillas].PFMIMT3.dbo.FM_INSTRUMENTOS ins ON LTRIM(RTRIM(mov.Instrumento)) = LTRIM(RTRIM(ins.CUSIP)) collate SQL_Latin1_General_CP1_CI_AS WHERE mov.Fecha_transaccion = '"+fecha+"' and ins.CUSIP is not NULL"

    df = get_frame_sql_user("Farellones", "CONCILIACION_INTERNACIONAL", "usuario1", "usuario1", strSQL) 
    



    #Extrae la tabla con el valor de las paridades segun la fecha ingresada
  
    eliminados = 0
    largo = df.shape[0]
    for i in range(largo):
        
        row = df.iloc[i-eliminados]
        if not row['Numero_Cuenta'].strip() in strcuentas:
        
            df = df.drop(df.index[i-eliminados])
            eliminados += 1

        else:
            pass
    

    for row in df['Numero_Cuenta']:
 
       
        row = row.strip()
        rut = get_frame_sql_user("Farellones", "CONCILIACION_INTERNACIONAL", "usuario1", "usuario1", "SELECT  Rut,Nro_Secuencia FROM [CONCILIACION_INTERNACIONAL].[dbo].TBL_CUENTAS_INTERNACIONALES  where Numero_Cuenta = '"+row+"'"  ) 

        rut['Rut'] = rut['Rut'].to_string(index=False).strip()
        
        rut['Nro_Secuencia'] = rut['Nro_Secuencia'].astype(int)
        
        sec = rut['Nro_Secuencia']
       
        sec = sec.to_string(index=False)
        rut[['Rut','no']] = rut['Rut'].str.split('-',expand=True)

        numero = rut['Rut'].astype(int)

        digitover = rut['no'].astype(int)


        codigo = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "select * from [Perfil Clientes] where Rut = '"+numero.to_string(index=False).strip()+"' and [Dig Ver] = '"+digitover.to_string(index=False).strip()+"' and Secuencia = '"+sec.strip()+"'"  ) 
      
   
        df.loc[df["Numero_Cuenta"] == row,"Numero_Cuenta"] = codigo['Codigo_Fdo'].to_string(index=False).strip()
        

   
    
    df['Estado'] = 'T'
    ## creamos una columna con la fecha de la session, es decir con la fecha actual

    date = fecha[6:10] +'-'+ fecha[3:5]+'-'+ fecha[0:2]
    df['fechasess'] = date
    paridades =  get_frame_sql_user("Puyehue", "replicasCredicorp","usrConsultaComercial" , "Comercial1w","Select * from VISTA_PARIDADES_CDP where CODIGO_MONEDA = 'CLP' AND FECHA_PARIDAD ='"+ date +"' ORDER BY CODIGO_MONEDA_ORIGEN ASC")

    paridades['CODIGO_MONEDA_ORIGEN'] = paridades['CODIGO_MONEDA_ORIGEN'].str.strip()

    dolar = paridades['VALOR_PARIDAD'].loc[(paridades['CODIGO_MONEDA_ORIGEN'] == 'USD') & (paridades['GRUPO_COTIZACION'] == 1)]
    dolar = dolar.to_string(index=False)
    
    ## creamos el dataframe con las columnas que utiliza RD
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

    df2 = pd.DataFrame(columns = l)

    

    df['Preciooo'] = df['Precio']
    
    ## agregamos columna estado
    df2 = df2.assign(estado =df['Estado'])
    
    ## agregamos columna de fecha de session 
    df2 = df2.assign(fecha_sesion = date)
    ## agregamos columna con el codigo de secuencia
    df2 = df2.assign(codigo_secuencia = df['ID']) 
    df2 = df2.assign(apertura_cierre = 'NULL')
   
    #
    #
    df2 = df2.assign(base_calculo =  df['Moneda'])  #Agregamos columna base calculo segun las condiciones                                                     #Para RV dejaremos fijo "3". Para RFN debe ser "2" (Act/365). 
                                                    #Para IIF sería "14" (Act/30) y si el papel es en moneda distinta de $ debe ser "3" (Act/360).


    df2['base_calculo'] = df2['base_calculo'].str.strip() 
    df2.loc[df2['base_calculo'] == 'USD' , 'base_calculo'] = '3' #
    df2.loc[df2['base_calculo'] == 'EUR' , 'base_calculo'] = '3' #

    
    ## agregamos columna broker fija para cccapital
    df2 = df2.assign(broker = 'CCCAPITAL')
    #
    ## agregamos columna cambio de divisa con el tipo de cambio actualizado segun los datos que se obtienen
    ## de la tabla paridades
    df2 = df2.assign(cam_divisa = df['Moneda'])

    df2['cam_divisa'] = df2['cam_divisa'].str.strip() 

    df2.loc[df2['cam_divisa'] == 'USD', 'cam_divisa'] = float(dolar)
    df2.loc[df2['cam_divisa'] == 'EUR', 'cam_divisa'] = '860.44'

    df2 = df2.assign(canones_div = 'NULL')
    #
    ## agregamos la columna cartera con el codigo de fondo de cada instrumento
    df2 = df2.assign(cartera = df['Numero_Cuenta'])
    df2['cartera'] = df2['cartera'].str.strip() 

    df2 = df2.assign(comision_div =  0)
    df2 = df2.assign(corretaje_div = 0)
    df2 = df2.assign(cot_recompra =  0)
    df2 = df2.assign(cot_rf_cup =    0) # cotizacion renta fija con cupon

    ## agregamos cotizacion segun las condiciones parqa renta variable es el precio de la operacion y para renta fija es 0
    df2 = df2.assign(cotizacion = df['Preciooo']) 
      
    df2 = df2.assign(cupon_corrido_div = 0) # FALTA AGREGAR
    df2 = df2.assign(cupon_per_div = 0)
    df2 = df2.assign(cupon_per = 0)
    #
    ## agregamos la columna dias
    #df['duration'].astype(float)
    df2 = df2.assign(dias = 0) 

    
    ## agregamos la columna divisa con el ID de tipo de moneda
    df2 = df2.assign(divisa = df['Moneda'])
    df2['divisa'] = df2['divisa'].str.strip() 
    df2.loc[df2['divisa'] == 'USD', 'divisa'] = '5'
    df2.loc[df2['divisa'] == 'EUR', 'divisa'] = '4'

    #
    df2 = df2.assign(efectivo_div = df['Valorizacion'])# monto, falta validar
  
    df2 = df2.assign(efectivo_vto_div =0)
    df2 = df2.assign(ent_contrapartida = 'CCCAPITAL') # valor por defecto
    df2 = df2.assign(ent_depositaria = 'DCV')     # deposito central de valores, va como defecto v # sicav sociedad de inversion de capital variable  # sicaf socidad de inversion de capital fijo                                                                                       
    #
    df2 = df2.assign(ent_liquidadora = 'PERSHING') # valor por defecto  
    #
    df2 = df2.assign(ent_mediadora = None)  # Valor por defecto
    #
    df2 = df2.assign(fec_comunicacion_be = None)  # valores por defecto 
    df2 = df2.assign(fec_liquidacion =     df['fecha_liquidacion'] ) # valores por defecto 
    df2 = df2.assign(fec_operacion =       df['fechasess'])  # valores por defecto 


    df2 = df2.assign(fec_recompra =        df['fechasess'])  # valores por defecto 
    df2 = df2.assign(fec_valor    =        df['fechasess'])  # valores por defecto 
    df2 = df2.assign(fec_vto      =        df['fecha_liquidacion'])

    df2 = df2.assign(gasto_extra = 0)
    df2 = df2.assign(gastos_div = '')

    df2 = df2.assign(gestor_ordenante = 2960) # valor por defecto

    #  suma o resta segun si es venta o compra
    df2 = df2.assign(ind_nominal_srn = df['Codigo_compraventa'])
    df2['ind_nominal_srn'] = df2['ind_nominal_srn'].str.strip()
    #df2['operacion'] = df2['operacion'].str.strip()

    df2.loc[df2['ind_nominal_srn'] == 'B'  , 'ind_nominal_srn'] = 'R'
    df2.loc[df2['ind_nominal_srn'] == 'S' , 'ind_nominal_srn'] = 'S'

    df2 = df2.assign(ind_periodificar = 'NULL')

 

    df2 = df2.assign(instrumento = df['Codigo_Ins'])

    df2['instrumento'] = df2['instrumento'].str.strip() 

    df2 = df2.assign(liquido_div = df['Valorizacion']) # falta validar



    df2 = df2.assign(mercado ='RVE')

    df2 = df2.assign(minusvalia = 0)
    df2 = df2.assign(minusvalia_div = 0)
    df2 = df2.assign(neto_div = df['Valorizacion'])
    df2 = df2.assign(nominal_div = df['Cantidad']) # consultar

  

    # venta o compra segun la tasa de compra, si existe es compra si no es venta
    df2 = df2.assign(operacion = df['Codigo_compraventa'])

    df2['operacion'] = df2['operacion'].str.strip()

    df2.loc[df2['operacion'] == 'B'  , 'operacion'] = 'C'
    df2.loc[df2['operacion'] == 'S'  , 'operacion'] = 'V'



    df2 = df2.assign(otros_gastos_div = '')
    df2 = df2.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que
    df2 = df2.assign(plusvalia = 0 )
    df2 = df2.assign(plusvalia_div = 0)
    df2 = df2.assign(prima_div = 'NULL')
    df2 = df2.assign(primario_secundario = 'S')# VALORES POR DEFECTO    
    df2 = df2.assign(regularizar_sn = 'S')     # VALORES POR DEFECTO    
    df2 = df2.assign(repo_dia = 'N')# VALORES POR DEFECTO   

    df2 = df2.assign(tipo_interes = 0)

    df2 = df2.assign(titulos =  df2['nominal_div']) 

    df2 = df2.assign(ventana_tratamiento = 'RVE')

    df2 = df2.assign(ind_proceso = 'D')
    df2 = df2.assign(enlace_sucursal = '')
    df2 = df2.assign(enlace_tipo_ap = '')
    df2 = df2.assign(enlace_plan_subplan = '')
    df2 = df2.assign(retencion_div = 0)

    # Ind_cambio_asegurado: si la fecha_operacion es distinta a fecha_liquidacion el campo debe ser “S”, si son iguales debe ser “N”
    df2 = df2.assign(ind_cambio_asegurado =  df2['nominal_div']) 
    df2 = df2.assign(ind_cambio_asegurado = 'N')




    df2 = df2.assign(ind_depositario = '')
    df2 = df2.assign(simple_compuesto = '')
    df2 = df2.assign(tex_pago = '')
    df2 = df2.assign(tex_pagare_cod = 'NULL')
    df2 = df2.assign(tex_pagares = 'NULL')
    df2 = df2.assign(num_comunica_be = 'NULL')
    df2 = df2.assign(iva = 'NULL')
    df2 = df2.assign(iva_div =' NULL')
    df2 = df2.assign(cuenta_ccc = '')
    df2 = df2.assign(ind_cobertura = 'NULL')
    df2 = df2.assign(fecha_saldo = '')
    df2 = df2.assign(libre_n1 = '')
    df2 = df2.assign(libre_n2 = '')
    df2 = df2.assign(libre_x1 = '')
    df2 = df2.assign(libre_x2 = '')
    df2 = df2.assign(fecha_ejecucion = '')
    df2 = df2.assign(codigo_uti = '')
    df2 = df2.assign(comision_ecc_div = '')
    df2 = df2.assign(tir_mercado = 0) # consultar


    df2['cam_divisa'] = df2['cam_divisa'].astype(str) 
    df2['cotizacion'] = df2['cotizacion'].astype(str) 
    df2['efectivo_div'] = df2['efectivo_div'].astype(str) 
    df2['liquido_div'] = df2['liquido_div'].astype(str) 
    df2['neto_div'] = df2['neto_div'].astype(str) 
    df2['nominal_div'] = df2['nominal_div'].astype(str) 
    df2['titulos'] = df2['titulos'].astype(str) 
    df2['mercado'] = df2['mercado'].astype(str) 
    df2['codigo_secuencia'] = df2['codigo_secuencia'].astype(str)


    df2['cam_divisa'] =   df2['cam_divisa'].str.replace('.',',')
    df2['cotizacion'] =   df2['cotizacion'].str.replace('.',',')
    df2['efectivo_div'] = df2['efectivo_div'].str.replace('.',',')
    df2['liquido_div'] =  df2['liquido_div'].str.replace('.',',')
    df2['neto_div'] =     df2['neto_div'].str.replace('.',',')
    df2['nominal_div'] =  df2['nominal_div'].str.replace('.',',')
    df2['titulos'] =      df2['titulos'].str.replace('.',',')
    df2['mercado'] =      df2['mercado'].str.replace('.',',')
    df2['codigo_secuencia'] = df2['codigo_secuencia'].str.replace('.',',')




    df2 = df2.reindex(columns=l)
    print(df2['operacion'])


    return df2


