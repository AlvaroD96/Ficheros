import pandas as pd
import numpy as np
import datetime as dt
import xlwings as xw
import os
import random
import locale
from copy import copy
import unidecode
import time
import timeit
from math import isnan
from sqlalchemy import create_engine, types 
import math
from decimal import Decimal
from PyQt5 import uic
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import sys
from datetime import datetime
import xlwings as xw
import pymssql

from win32com import client
# from PyPDF2 import PdfFileReader, PdfFileMerger
import win32com.client
import json
import re
import matplotlib.pyplot as plt
# import seaborn as sns
from shutil import copyfile

'''FUNCIONES VARIAS'''

def get_ndays_from_today(days):
	'''
	Retorna la fecha n dias desde hoy.
	'''
	date = dt.datetime.now() - dt.timedelta( days = days )
	date=str(date.strftime('%Y-%m-%d'))
	return date


def custom_round(x, base):
    '''
    Redondea un porcentaje a un multiplo, sirve
    para aplicar una funcion landa a un dataframe.
    '''
    return int(10000 * base * round(float(x)/base))/10000


def format_separators(n):
    '''
    Formatea un float n con comas y puntos decimales
    '''
    return  "{:,}".format(n)


def convert_string_to_date(date_string):
    '''
    Retorna la fecha en formato date.
    '''
    date_formated=dt.datetime.strptime(date_string, '%Y-%m-%d').date()
    return date_formated

def convert_date_to_string(date):
    '''
    Retorna la fecha en formato de string, dado el objeto date.
    '''
    date_string=str(date.strftime('%Y-%m-%d'))
    return date_string

def truncate(f, n):
    '''
    Trunca un numero a n digitos
    '''
    s = '{}'.format(f)
    if 'e' in s or 'E' in s:
        return '{0:.{1}f}'.format(f, n)
    i, p, d = s.partition('.')
    return '.'.join([i, (d+'0'*n)[:n]])


def get_self_path():
	'''
	Retorna la ruta del codigo en python que lo llama.
	'''
	self_path=os.path.dirname(os.path.realpath(sys.argv[0]))+"\\"
	return self_path


def get_current_time():
	'''
	Retorna la hora actual.
	'''
	return dt.datetime.now().time().strftime("%H:%M:%S")


def get_prev_weekday(adate):
	'''
	Retorna la fecha en string del ultimo weekday dado una fecha en string.
	'''
	adate = convert_string_to_date(date_string = adate)
	adate -= dt.timedelta(days=1)
	while adate.weekday() > 4: # Mon-Fri are 0-4
		adate -= dt.timedelta(days=1)
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


def get_dates_between(date_inic, date_fin):
	'''
	Retorna una lista con todas las fechas entre dos fechas dadas.
	'''
	day_list=pd.date_range(date_inic, date_fin).tolist()
	return day_list


def get_weekdays_dates_between(date_inic, date_fin):
	'''
	Retorna una lista con todas las fechas habiles entre dos fechas dadas.
	'''
	day_list =pd.date_range(date_inic, date_fin)
	day_list = day_list[day_list.dayofweek < 5].tolist()
	return day_list


def get_current_days_week(date):
	'''
	Retorna la cantidad de dias hasta el ultimo dia de la semana pasada.
	'''
	days= date.timetuple().tm_wday+2 #tm_wday devuelve un día menos de los que se lleva en la semana (si es jueves, devuelve 3), por eso se le suma 2 (el día q falta de la semana, y el último día de la semana anterior)
	return days



def get_current_weekdays_month(date):
	'''
	Retorna la cantidad de dias habiles hasta el ultimo dia del mes pasado.
	'''
	businessdays = 0
	for i in range(1, 32):
		try:
			thisdate = dt.date(date.year, date.month, i)
		except(ValueError):
			break
		if thisdate.weekday() < 5 and thisdate <= date: # Monday == 0, Sunday == 6 
			businessdays += 1
	return businessdays +1


def get_current_weekdays_year(date):
	'''
	Retorna la cantidad de dias habiles hasta el anio pasado.
	'''
	return np.busday_count( dt.date(date.year, 1, 1), date ) + 2



def get_nweekdays_from_date(days, date_string):
	'''
	Retorna la fecha n dias desde hoy.
	'''
	adate = convert_string_to_date(date_string = date_string)
	counter = -1
	while True:
		if adate.weekday() <= 4:
			counter += 1
		if counter == days:
			break 
		adate -= dt.timedelta(days = 1)
	return str(adate)


def convert_date_to_string(date):
	'''
	Retorna la fecha en formato de string, dado el objeto date.
	'''
	date_string=str(date.strftime('%Y-%m-%d'))
	return date_string


def convert_string_to_date(date_string):
	'''
	Retorna la fecha en formato date.
	'''
	date_formated=dt.datetime.strptime(date_string, '%Y-%m-%d').date()
	return date_formated

def convert_json_to_dict(json_string):
	'''
	Lee un string en formato JSON y lo retorna en un dictionary.
	'''
	json_dict=json.loads(json_string)
	return json_dict

	
'''FUNCIONES MATRICIALES'''
def array_to_numpy(arr):
	'''
	Transforma una matriz normal en una de numpy.
	'''
	return np.array(arr)


def get_vect_column(matrix, col_number):
	'''
	Dada una lista de listas retorna un vector con la columna.
	'''
	matrix_np=np.array(matrix)
	column=matrix_np[:,col_number]
	return column


def format_tuples(df):
    '''
    Transforma un dataframe en una lista de tuplas.
    '''
    serie_tuplas = [tuple(x) for x in df.values]
    return serie_tuplas


def print_full(x):
    '''
    Imprime un dataframe entero
    '''
    pd.set_option('display.max_rows', len(x))
    print(x)
    pd.reset_option('display.max_rows')


'''FUNCIONES SQL'''

def connect_database(server, database):
	'''
	Se conecta a una base de datos y retorna el objecto de la conexion, usando windows authentication.
	'''
	conn= pymssql.connect(host=server, database=database)
	return conn


def connect_database_user(server, database, username, password):
	'''
	Se conecta a una base de datos y retorna el objecto de la conexion, usando un usuario.
	'''
	conn= pymssql.connect(host=server, database=database, user=username, password=password)
	return conn


def query_database(conn, query):
	'''
	Consulta la la base de datos(conn) y devuelve el cursor asociado a la consulta.
	'''
	cursor=conn.cursor()
	cursor.execute(query)
	return cursor


def get_table_sql(cursor):
	'''
	Recibe un cursor asociado a una consulta en la BDD y la transforma en una matriz.
	'''
	table=[]
	row = cursor.fetchone()
	ncolumns=len(cursor.description)
	while row:
		col=0	
		vect=[]
		while col<ncolumns:
			vect.append(row[col])
			col=col+1
		row = cursor.fetchone()
		table.append(vect)
	return table


def get_list_sql(cursor):
	'''
	Recibe un cursor asociado a una consulta en la BDD y la transforma en un lista, solo usarlo cuando la consulta retorne una columna.
	'''
	lista=[]
	row = cursor.fetchone()
	while row:	
		escalar=row[0]
		row = cursor.fetchone()
		lista.append(escalar)
	return lista


def get_schema_sql(cursor):
	'''
	Recibe un cursor asociado a una consulta en la BDD y retorna el esquema de la relacion en una lista.
	'''
	schema = []
	for i in range(len(cursor.description)):
		prop = cursor.description[i][0]
		schema.append(prop)
	return schema


def disconnect_database(conn):
	'''
	Se deconecta a una base de la datos.
	'''
	conn.close()


def run_sql(conn, query):
    '''
    Ejecuta un statement en SQL, por ejemplo borrar.
    '''
    cursor=conn.cursor()
    cursor.execute(query)
    conn.commit()


def get_frame_sql_user(server, database, username, password, query):
    '''
    Retorna el resultado de la query en un panda's dataframe con usuario y clave. 
    '''
    conn = connect_database_user(server = server, database = database, username = username, password = password)
    cursor = query_database(conn, query)
    schema = get_schema_sql(cursor = cursor)
    table = get_table_sql(cursor)
    dataframe = pd.DataFrame(data = table, columns = schema)
    disconnect_database(conn)
    return dataframe


def get_frame_sql(server, database, query):
    '''
    Retorna el resultado de la query en un panda's dataframe. 
    '''
    conn = connect_database(server = server, database = database)
    cursor = query_database(conn, query)
    schema = get_schema_sql(cursor = cursor)
    table = get_table_sql(cursor)
    dataframe = pd.DataFrame(data = table, columns = schema)
    disconnect_database(conn)
    return dataframe


def get_val_sql_user(server, database, username, password, query):
    '''
    Retorna el resultado de la query en una variable. 
    '''
    conn = connect_database_user(server = server, database = database, username = username, password = password)
    cursor = query_database(conn, query)
    table = get_table_sql(cursor)
    val = table[0][0]
    disconnect_database(conn)
    return val


def get_val_sql(server, database, query):
    '''
    Retorna el resultado de la query en una variable. 
    '''
    conn = connect_database_user(server = server, database = database)
    cursor = query_database(conn, query)
    table = get_table_sql(cursor)
    val = table[0][0]
    disconnect_database(conn)
    return val


'''FUNCIONES EXCEL'''

def create_workbook():
	'''
	Crea un workbook.
	'''
	wb = xw.Book()
	return wb


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


def save_workbook(wb, path = ""):
	'''
	Guarda un workbook.
	'''
	if path == "":
		wb.save()
	else:
		wb.save(path)


def close_workbook(wb):
	'''
	Cierra un workbook.
	'''
	wb.close()


def close_excel(wb):
	'''
	Cierra un Excel.
	'''
	app = wb.app
	app.quit()


def kill_excel():
	'''
	Mata el proceso de Excel.
	'''
	try:
		os.system('taskkill /f /im Excel.exe')
	except:
		print("no hay excels abiertos")


def clear_table_xl(wb, sheet, row, column):
	'''
	Recibe el rango asociado y borra la tabla considerando que el rango esta en el top-left corner de la tabla.
	'''
	wb.sheets[sheet].range(row,column).expand('table').clear_contents()


def clear_column_xl(wb, sheet, row, column):
	'''
	Recibe el rango asociado y borra la tabla considerando que el rango esta en el top-left corner de la tabla.
	'''
	wb.sheets[sheet].range(row,column).expand('down').clear_contents()


def paste_val_xl(wb, sheet, row, column,values):
	'''
	Inserta un valor en una celta de excel dada la hoja. 
	'''

	for fila in range(len(values)):
		for columna in range(len(values[fila])):
			wb.sheets[sheet].cell(fila+2,columna
			+1).value = values[fila][columna]


def paste_col_xl(wb, sheet, row, column, serie):
	'''
	Inserta un valor en una celda de excel dada la hoja. 
	'''
	for val in serie:
		wb.sheets[sheet].range(row,column).value = val
		row += 1

def paste_data_frame(wb,sheet,position,value):
	wb.sheets[sheet].range(position).value = value

def paste_query_xl(wb, server, database, query, sheet, row, column, with_schema):
	'''
	Consulta a la base de datos y pega el resultado en una hoja de excel en una determinada fila y columna. Si with_schema es verdadero, la pega con
	los nombres de columnas, en otro caso solo pega los valores. 
	'''
	conn= connect_database(server,database)
	cursor=query_database(conn, query)
	table=get_table_sql(cursor)
	disconnect_database(conn)
	if with_schema:
		schema=get_schema_sql(cursor)
		paste_val_xl(wb, sheet, row, column, schema)
		paste_val_xl(wb, sheet, row+1, column, table)
	else:		
		paste_val_xl(wb, sheet, row, column, table)


def paste_query_xl_user(wb, server, database, query, sheet, row, column, with_schema, username, password):
	'''
	Consulta a la base de datos y pega el resultado en una hoja de excel en una determinada fila y columna. Si with_schema es verdadero, la pega con
	los nombres de columnas, en otro caso solo pega los valores. Esta funcion es igual a la anterior pero se conecta a la BDD con usuario y pass.
	'''
	conn= connect_database_user(server,database, username, password)
	cursor=query_database(conn, query)
	table=get_table_sql(cursor)
	disconnect_database(conn)
	if with_schema:
		schema=get_schema_sql(cursor)
		paste_val_xl(wb, sheet, row, column, schema)
		paste_val_xl(wb, sheet, row+1, column, table)
	else:		
		paste_val_xl(wb,sheet, row, column, table)


def get_sheet_index(wb, sheet):
	'''
	Retorna el indice e una hoja, dado su nombre.
	'''
	return wb.sheets[sheet].index


def export_sheet_pdf(sheet_index, path_in, path_out):
	'''
	Exporta una hoja de Excel en PDF. 
	'''
	xlApp = client.Dispatch("Excel.Application")
	books = xlApp.Workbooks(1)
	ws = books.Worksheets[sheet_index]
	ws.ExportAsFixedFormat(0, path_out)


def get_value_xl(wb, sheet, row, column):
	'''
	Retorna el valor de una celda dado un rango y la hoja.
	'''
	return wb.sheets(sheet).range(row,column).value


def get_column_xl(wb, sheet, row, column):
	'''
	Retorna la columna en una lista, dado un rango y la hoja.
	'''
	return wb.sheets(sheet).range(row,column).expand('down').value


def get_table_xl(wb, sheet, row, column):
	'''
	Retorna la tabla en un arreglo, dado un rango y la hoja.
	'''
	return wb.sheets[sheet].range(row,column).expand('table').value


def get_frame_xl(wb, sheet, row, column, index_pos):
	'''
	Retorna la tabla en un dataframe, dado un rango y la hoja.
	Ademas se le da la posicion de las columnas a indexar
	'''
	table = get_table_xl(wb, sheet, row, column)
	data = table[1:]



	columns = np.array(table[0])
	table = pd.DataFrame(data, columns=columns)
	table.set_index(columns[index_pos].tolist(), inplace=True)
	return table


def clear_sheet_xl(wb, sheet):
	'''
	Borra todos los contenidos de una hoja de excel.
	'''
	return wb.sheets[sheet].clear_contents()


def setFrameSql(server, database, dataframe, username, password):
	'''
	Actualiza la tabla de SQL a partir de un Dataframe
	'''
	if dataframe.empty:
		print("El dataframe esta vacio")
		return False
	# engine = create_engine('mssql+pymssql://' + server + '/' + database)
	engine = create_engine('mssql+pymssql://'+ username + ':' + password +'@' + server + '/' + database)

	con = engine.connect()


	dataframe.to_sql('Solicitudes', con, if_exists = 'replace', index = False, schema = 'dbo',
					dtype = {
					'[fecha]':     					types.NVARCHAR(),
					'[codigo_fdo]' :				types.NVARCHAR(),
					'[codigo_instrumento]' : 		types.NVARCHAR(),
					'[codigo_ins]': 				types.NVARCHAR(),
					'[codigo_emi]' :				types.NVARCHAR(),
					'[Cantidad]' :					types.NVARCHAR(),
					'[Nominal]': 					types.NVARCHAR(),
					'[duration]': 					types.NVARCHAR(),
					'[MONEDA]': 					types.NVARCHAR(),
					'[Tipo_Instrumento]': 			types.NVARCHAR(),
					'[Nombre_Instrumento]':			types.NVARCHAR(),
					'[Tasa]': 						types.NVARCHAR(),
					'[fec_vcto]':					types.NVARCHAR(),
					'[Weight]': 					types.NVARCHAR(),
					'[Monto]' : 					types.NVARCHAR(),
					'[tasa_compra]' : 				types.NVARCHAR(),
					'[precio_mdo]': 				types.NVARCHAR(),
					'[precio_dirty]': 				types.NVARCHAR(),
					'[NOMBRE_EMISOR]': 				types.NVARCHAR(),
					'[Pais_Emisor]': 				types.NVARCHAR(),
					'[codigo_tipo_instrumento]': 	types.NVARCHAR(),
					'[Riesgo]':						types.NVARCHAR(),
					'[Zona]': 						types.NVARCHAR(),
					'[Renta]': 						types.NVARCHAR(),
					})


def dataframe_join(df1,df2, row):
	df_merge = pd.merge(df1,df2,on='{}'.format(row))
	return df_merge


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

def poblate_fondos(self):
    '''
    Llena la lista con los fondos disponibles en la base de datos
    '''
    self.model1.clear()
    
    df = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "Select * From Fondos")
    for i in df["codigo_fdo"]:
        self.model1.appendRow(QStandardItem(i))

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



def MMFDTUG(fecha):

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

    date = fecha[6:10] +'-'+ fecha[3:5]+'-'+ fecha[0:2]

    caja = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "select * from zhis_carteras_main where fecha = '"+date+"' and codigo_emi = 'CAJA' and codigo_ins = 'MMFDUTG'")
   
    paridades =  get_frame_sql_user("Puyehue", "replicasCredicorp","usrConsultaComercial" , "Comercial1w","Select * from VISTA_PARIDADES_CDP where CODIGO_MONEDA = 'CLP' AND FECHA_PARIDAD ='"+ date +"' ORDER BY CODIGO_MONEDA_ORIGEN ASC")

    paridades['CODIGO_MONEDA_ORIGEN'] = paridades['CODIGO_MONEDA_ORIGEN'].str.strip()

    dolar = paridades['VALOR_PARIDAD'].loc[(paridades['CODIGO_MONEDA_ORIGEN'] == 'USD') & (paridades['GRUPO_COTIZACION'] == 1)]
    dolar = dolar.to_string(index=False)
    

    ## caja printeado es :
 

    #fecha	codigo_fdo	codigo_emi	codigo_ins	Monto	Tipo_Instrumento	Moneda	Sector	Weight	nominal	Cantidad	Precio	Precio_Dirty	duration	Tasa	Riesgo	Nombre_Emisor	Nombre_Instrumento	fec_vcto	Pais_Emisor	Estrategia	Zona	Renta	Riesgo_Internacional	Tasa_Compra	Moneda_Fdo	Spread	Spread_Carry	Riesgo_RA	Tipo_Ra
    #06-01-2020	INV_HORTEN	CAJA      	MMFDUTG   	28.851.090.000	Cuota de Fondo	US$       	Financieras	0.000250	37.380.600	37.380.600	1	99,95983506	0	0.0000000	N/A	Dreyfus Universal US Treasury	Dreyfus Universal Ustreasury Class G	          	CL	NULL	Global	Fija	NULL	0.0000000000	$	0	0	NULL	N/A

    caja['Estado'] = 'T'
    dfMM = pd.DataFrame(columns = l)

    dfMM = dfMM.assign(estado =caja['Estado'])

      
    
    ## agregamos columna de fecha de session 
    dfMM = dfMM.assign(fecha_sesion = date)
    ## agregamos columna con el codigo de secuencia
    dfMM = dfMM.assign(codigo_secuencia = caja['ID']) 
    dfMM = dfMM.assign(apertura_cierre = 'NULL')
   
    #
    #
    dfMM = dfMM.assign(base_calculo =  '3')  #Agregamos columna base calculo segun las condiciones                                                     #Para RV dejaremos fijo "3". Para RFN debe ser "2" (Act/365). 
                                                    #Para IIF sería "14" (Act/30) y si el papel es en moneda distinta de $ debe ser "3" (Act/360).

    dfMM = dfMM.assign(codigo_secuencia = caja.index) 


    dfMM = dfMM.assign(broker = 'CCCAPITAL')

    dfMM = dfMM.assign(cam_divisa = caja['Moneda'])

    #print(cambio_diario['VALOR_PARIDAD'].loc[cambio_diario['Moneda']== 'USD'] )

    dfMM.loc[dfMM['cam_divisa'] == 'US$', 'cam_divisa'] = dolar
    dfMM.loc[dfMM['cam_divisa'] == '$',   'cam_divisa'] =   '1'

    dfMM = dfMM.assign(canones_div = None)


    dfMM = dfMM.assign(cartera = caja['codigo_fdo'] )


    dfMM = dfMM.assign(comision_div = 0)
    dfMM = dfMM.assign(corretaje_div = 0)
    dfMM = dfMM.assign(cot_recompra = 0)


    dfMM = dfMM.assign(cot_rf_cup = 0) # cotizacion renta fija con cupon

    dfMM = dfMM.assign(cotizacion = caja['Precio'])



    dfMM = dfMM.assign(cupon_corrido_div = None)


    dfMM = dfMM.assign(cupon_per_div = 0)
    
    dfMM = dfMM.assign(cupon_per = 0)


    dfMM = dfMM.assign(dias=0)

    dfMM = dfMM.assign(divisa = caja['Moneda'])



    dfMM.loc[dfMM['divisa'] == 'US$', 'divisa'] = '5'
    dfMM.loc[dfMM['divisa'] == '$',   'divisa'] =  '39'
    dfMM.loc[dfMM['divisa'] == 'UF', ' divisa'] =  '84'
    dfMM.loc[dfMM['divisa'] == 'EUR', 'divisa'] =  '4'

    dfMM = dfMM.assign(efectivo_div = caja['Monto']/10000)




    dfMM = dfMM.assign(efectivo_vto_div =0) #  

    dfMM = dfMM.assign(ent_contrapartida = 'CCCAPITAL') 

    dfMM = dfMM.assign(ent_depositaria = 'DCV')     # deposito central de valores 
                                            # sicav sociedad de inversion de capital variable
    dfMM = dfMM.assign(ent_liquidadora = 'PERSHING') # valor por defecto  
    #
    dfMM = dfMM.assign(ent_mediadora = None)  # Valor por defecto
    #
    dfMM = dfMM.assign(fec_comunicacion_be = None)  # valores por defecto 
    dfMM = dfMM.assign(fec_liquidacion =     None ) # valores por defecto 
    dfMM = dfMM.assign(fec_operacion =       caja['fechasess'])  # valores por defecto 


    dfMM = dfMM.assign(fec_recompra =        caja['fechasess'])  # valores por defecto 
    dfMM = dfMM.assign(fec_valor    =        caja['fechasess'])  # valores por defecto 
    dfMM = dfMM.assign(fec_vto      =        caja['fec_vcto'])

    dfMM = dfMM.assign(gasto_extra = 0)
    dfMM = dfMM.assign(gastos_div = '')

    dfMM = dfMM.assign(gestor_ordenante = 2960) # valor por defecto
    dfMM = dfMM.assign(ind_nominal_srn  = 'R')

    dfMM = dfMM.assign(ind_periodificar = 'NULL')

 

    dfMM = dfMM.assign(instrumento = caja['codigo_ins'])

    dfMM['instrumento'] = dfMM['instrumento'].str.strip()    

    


    dfMM = dfMM.assign(mercado ='RVE')

    dfMM = dfMM.assign(minusvalia = 0)
    dfMM = dfMM.assign(minusvalia_div = 0)
    dfMM = dfMM.assign(neto_div = dfMM['efectivo_div'])
    dfMM = dfMM.assign(nominal_div = caja['Cantidad']) # consultar

  

    # venta o compra segun la tasa de compra, si existe es compra si no es venta
    dfMM = dfMM.assign(operacion = 'C')

    dfMM['operacion'] = dfMM['operacion'].str.strip()

    
    dfMM = dfMM.assign(otros_gastos_div = '')
    dfMM = dfMM.assign(previa_compromiso_plazo = 'D') # compromiso diario segun que
    dfMM = dfMM.assign(plusvalia = 0 )
    dfMM = dfMM.assign(plusvalia_div = 0)
    dfMM = dfMM.assign(prima_div = 'NULL')
    dfMM = dfMM.assign(primario_secundario = 'S')# VALORES POR DEFECTO    
    dfMM = dfMM.assign(regularizar_sn = 'S')     # VALORES POR DEFECTO    
    dfMM = dfMM.assign(repo_dia = 'N')# VALORES POR DEFECTO   

    dfMM = dfMM.assign(tipo_interes = 0)

    dfMM = dfMM.assign(titulos =  dfMM['nominal_div']) 

    dfMM = dfMM.assign(ventana_tratamiento = 'RVE')

    dfMM = dfMM.assign(ind_proceso = 'D')
    dfMM = dfMM.assign(enlace_sucursal = '')
    dfMM = dfMM.assign(enlace_tipo_ap = '')
    dfMM = dfMM.assign(enlace_plan_subplan = '')
    dfMM = dfMM.assign(retencion_div = 0)

    # Ind_cambio_asegurado: si la fecha_operacion es distinta a fecha_liquidacion el campo debe ser “S”, si son iguales debe ser “N”
    dfMM = dfMM.assign(ind_cambio_asegurado =  dfMM['nominal_div']) 
    dfMM = dfMM.assign(ind_cambio_asegurado = 'N')




    dfMM = dfMM.assign(ind_depositario = '')
    dfMM = dfMM.assign(simple_compuesto = '')
    dfMM = dfMM.assign(tex_pago = '')
    dfMM = dfMM.assign(tex_pagare_cod = 'NULL')
    dfMM = dfMM.assign(tex_pagares = 'NULL')
    dfMM = dfMM.assign(num_comunica_be = 'NULL')
    dfMM = dfMM.assign(iva = 'NULL')
    dfMM = dfMM.assign(iva_div =' NULL')
    dfMM = dfMM.assign(cuenta_ccc = '')
    dfMM = dfMM.assign(ind_cobertura = 'NULL')
    dfMM = dfMM.assign(fecha_saldo = '')
    dfMM = dfMM.assign(libre_n1 = '')
    dfMM = dfMM.assign(libre_n2 = '')
    dfMM = dfMM.assign(libre_x1 = '')
    dfMM = dfMM.assign(libre_x2 = '')
    dfMM = dfMM.assign(fecha_ejecucion = '')
    dfMM = dfMM.assign(codigo_uti = '')
    dfMM = dfMM.assign(comision_ecc_div = '')
    dfMM = dfMM.assign(tir_mercado = 0) # consultar


    
    dfMM['cam_divisa'] = dfMM['cam_divisa'].astype(str) 
    dfMM['cotizacion'] = dfMM['cotizacion'].astype(str) 
    dfMM['efectivo_div'] = dfMM['efectivo_div'].astype(str) 
    dfMM['liquido_div'] = dfMM['liquido_div'].astype(str) 
    dfMM['neto_div'] = dfMM['neto_div'].astype(str) 
    dfMM['nominal_div'] = dfMM['nominal_div'].astype(str) 
    dfMM['titulos'] = dfMM['titulos'].astype(str) 
    dfMM['mercado'] = dfMM['mercado'].astype(str) 
    dfMM['codigo_secuencia'] = dfMM['codigo_secuencia'].astype(str)


    dfMM['cam_divisa'] =   dfMM['cam_divisa'].str.replace('.',',')
    dfMM['cotizacion'] =   dfMM['cotizacion'].str.replace('.',',')
    dfMM['efectivo_div'] = dfMM['efectivo_div'].str.replace('.',',')
    dfMM['liquido_div'] =  dfMM['liquido_div'].str.replace('.',',')
    dfMM['neto_div'] =     dfMM['neto_div'].str.replace('.',',')
    dfMM['nominal_div'] =  dfMM['nominal_div'].str.replace('.',',')
    dfMM['titulos'] =      dfMM['titulos'].str.replace('.',',')
    dfMM['mercado'] =      dfMM['mercado'].str.replace('.',',')
    dfMM['codigo_secuencia'] = dfMM['codigo_secuencia'].str.replace('.',',')




    dfMM = dfMM.reindex(columns=l)
    

    return dfMM




def resource_path(relative_path):
    '''
    Retorna el path absoluto del path entregado.
    '''
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# se realiza la conexion con la interfaz del pop up
popUp = resource_path('popUp.ui')

window_name_12, base_class_12 = uic.loadUiType(popUp)


class popUp1(window_name_12, base_class_12):
    def __init__(self, texto):
        super().__init__()
        self.setupUi(self)

        self.label_2.setText(texto)




# Interfaz que despliega los fondos y permite realizar la consulta
class Ui(QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('inicio.ui', self)
        # la query para obtener los fondos, y luego ingresarlos a la lista en pantalla
        fondos = get_frame_sql_user("Puyehue", "MesaInversiones", "usuario1", "usuario1", "Select * From Fondos")
        # pobla la lista con los datos
        self.ListFondos.addItems(fondos['codigo_fdo'])
        # le agrega la funcion al botn consulta
  
        self.Consulta.clicked.connect(self.Consultar) 
       
           
        # obtenemos la fecha que se ingresa en la gui
        self.fecha = self.findChild(QDateEdit, 'Fecha')
        # seleccionamos las rutas de acceso a los ficheros .txt
        self.pathap.clicked.connect(self.pathap1)
        self.pathgeo.clicked.connect(self.pathgeo1)
        #barra de progreso solo estetico
        self.progress = QProgressBar(self.progreso)
        self.progress.setTextVisible(False)
      

    # funciones que permiten abrir el directorio donde se encuentra el archivo de carga
    def pathap1(self):
        file2 = str(QFileDialog.getOpenFileName(self,"Aportes y rescates", "","Archivos de texto (*.txt)"))
        file2 = file2[2:-31]
        self.AP.setText(file2)
       
    def pathgeo1(self):
        file1 = str(QFileDialog.getOpenFileName(self,"Operaciones", "","Archivos de texto (*.txt)"))
        file1 = file1[2:-31]
        self.GEO.setText(file1)
    


    def Consultar(self):# funcion que realiza la consulta a la tabla, realmente es solo el boton que llama a la funcion fichero
        print('[INFO]: Iniciando...')
        self.fondo = self.findChild(QComboBox, 'ListFondos')
        path_ap = self.AP.text()
        path_geo = self.GEO.text()
        Codigo_Cartera = self.Codigo_Cartera.text()
        self.progress.setValue(1)
        # el try esta para controlar el error que salta en caso de que tengan abierto el excel y no lo cierren
        l=["estado","fecha_sesion",	"codigo_secuencia",	"apertura_cierre"	,"base_calculo",	"broker",	"cam_divisa",	"canones_div",	"cartera",	"comision_div",	"corretaje_div"	,"cot_recompra",	"cot_rf_cup",	"cotizacion",	"cupon_corrido_div"	,"cupon_per",	"cupon_per_div"	,"dias",	"divisa","efectivo_div",	"efectivo_vto_div"	,"ent_contrapartida",	"ent_depositaria",	"ent_liquidadora","ent_mediadora"	,"fec_comunicacion_be",	"fec_liquidacion",	"fec_operacion"	,"fec_recompra","fec_valor",	"fec_vto",	"gasto_extra",	"gastos_div",	"gestor_ordenante",	"ind_nominal_srn","ind_periodificar",	"instrumento",	"liquido_div",	"mercado",	"minusvalia",	"minusvalia_div","neto_div",	"nominal_div",	"operacion"	,"otros_gastos_div"	,"previa_compromiso_plazo","plusvalia"	,"plusvalia_div",	"prima_div"	,"primario_secundario"	,"regularizar_sn",	"repo_dia","tipo_interes",	"titulos",	"ventana_tratamiento",	"ind_proceso",	"enlace_sucursal",	"enlace_tipo_ap","enlace_plan_subplan",	"retencion_div"	,"ind_cambio_asegurado"	,"ind_depositario",	"simple_compuesto","tex_pago",	"tex_pagare_cod",	"tex_pagares",	"num_comunica_be",	"iva",	"iva_div",	"cuenta_ccc","ind_cobertura",	"fecha_saldo",	"libre_n1",	"libre_n2",	"libre_x1",	"libre_x2",	"fecha_ejecucion","codigo_uti"]
      
        try:
            # se carga el dataframe de transacciones no p*q que corresponden a intermediacion financiera y renta fija 
            nopq = fichero(self.fecha.text())
            
            nopq = nopq.loc[nopq["cartera"].str.strip() == self.fondo.currentText().strip()]
            nopq.to_csv("R:/10 CARTERAS ADMINISTRADAS/CAD RD/Ficheros de Carga/FicherosNO_PQ/"+self.fondo.currentText()+self.fecha.text()+"_NO__PQ.csv", index=False,sep = ';')
            print('[INFO]: Transacciones RF e IIF listo...')
            self.progress.setValue(25)
            try:
                per = Pershing(self.fecha.text())
                per = per.loc[per["cartera"].str.strip() == self.fondo.currentText().strip()]
            except ValueError:

                per = pd.DataFrame(columns = l)
            
            
            dfMM = MMFDTUG(self.fech.text())
            dfMM = dfMM.loc[dfMM['cartera'].str.strip() == self.fondo.currentText().strip()]
            
            print('[INFO]: Transacciones de pershing listo...')
            self.progress.setValue(50)
            # se carga el dataframe de aportes y rescates
            ap = aporteRescate(path_ap)
            ap = ap.loc[ap["cartera"].str.strip() == self.fondo.currentText().strip()]
            print('[INFO]: Transacciones de aportes y rescates listo...')
            self.progress.setValue(75)
            # se carga el dataframe de renta variable
            trans = transaccionesRVN(path_geo,self.fecha.text() )
            trans = trans.loc[trans["cartera"].str.strip() == self.fondo.currentText().strip()]
            print('[INFO]: Transacciones RVN listo...')
            self.progress.setValue(90)
            #se concatenan los dataframes de p*q para obtener un fichero consolidado
            concat = pd.concat([ap,trans,per , dfMM])
            concat.drop(["comision_ecc_div","tir_mercado"], axis=1)
            concat = concat.reindex(columns=l)
            concat.to_csv("R:/10 CARTERAS ADMINISTRADAS/CAD RD/Ficheros de Carga/FicherosPQ/"+self.fondo.currentText()+self.fecha.text()+"PQ.csv", index=False,sep = ';')
            self.progress.setValue(95)
	
            carteras = concat.copy()
	
            fondito = self.fondo.currentText().strip() + self.Codigo_Cartera.text().strip()
          
            carteras = carteras.assign(cartera = fondito )
            carteras.to_csv("R:/10 CARTERAS ADMINISTRADAS/CAD RD/Ficheros de Carga/CARTERAS/CARTERAS_"+self.fondo.currentText()+self.fecha.text()+"PQ.csv", index=False,sep = ';')
		
            


            
            self.progress.setValue(100)
            self.pop_up = popUp1('Ficheros para {} de la fecha {} listos'.format(self.fondo.currentText().strip(),self.fecha.text()))
            self.pop_up.show()
            self.progress.setValue(0)
        
        except FileNotFoundError:
            self.pop_up = popUp1('No se encuentran ficheros' )
            self.pop_up.show()
            self.progress.setValue(0)
        except PermissionError:
            self.pop_up = popUp1('Hay un excel abierto con la fecha solicitada' )
            self.pop_up.show()
            self.progress.setValue(0)
            pass


# se inicia la aplicacion
''' MAIN '''
if __name__ == '__main__':
    app = QApplication(sys.argv)
    #app.setWindowIcon(QtGui.QIcon('icon.png'))
    window = Ui()
    window.show()
    app.exec_()
