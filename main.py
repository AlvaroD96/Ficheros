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
import time
import timeit
from math import isnan
from sqlalchemy import create_engine, types 
import math
#from libreria_fdo import connect_database_user, query_database , get_schema_sql , disconnect_database,get_frame_sql_user, get_ndays_from_date, convert_string_to_date, convert_date_to_string, dataframe_join,get_frame_xl, print_full
from decimal import Decimal
from PyQt5 import uic
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import sys
from datetime import datetime
import Aporteyrescate
import pershing
import transacconesRVN
import NO_PQ

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
            nopq = NO_PQ.fichero(self.fecha.text())
            
            nopq = nopq.loc[nopq["cartera"].str.strip() == self.fondo.currentText().strip()]
            nopq.to_csv("R:/10 CARTERAS ADMINISTRADAS/CAD RD/Ficheros de Carga/FicherosNO_PQ/"+self.fondo.currentText()+self.fecha.text()+"_NO__PQ.csv", index=False,sep = ';')
            print('[INFO]: Transacciones RF e IIF listo...')
            # barra de progreso al 25%
            self.progress.setValue(25)
            try:
                per = pershing.Pershing(self.fecha.text())
                per = per.loc[per["cartera"].str.strip() == self.fondo.currentText().strip()]
            except ValueError:

                per = pd.DataFrame(columns = l)

            print('[INFO]: Transacciones de pershing listo...')
    
            self.progress.setValue(50)
            # se carga el dataframe de aportes y rescates
            ap = Aporteyrescate.aporteRescate(path_ap)
            ap = ap.loc[ap["cartera"].str.strip() == self.fondo.currentText().strip()]
            print('[INFO]: Transacciones de aportes y rescates listo...')
    
            self.progress.setValue(75)
            # se carga el dataframe de renta variable
            trans = transacconesRVN.transaccionesRVN(path_geo,self.fecha.text() )
            trans = trans.loc[trans["cartera"].str.strip() == self.fondo.currentText().strip()]
            print('[INFO]: Transacciones RVN listo...')
            self.progress.setValue(90)
            #se concatenan los dataframes de p*q para obtener un fichero consolidado
            concat = pd.concat([ap,trans,per])
            concat.drop(["comision_ecc_div","tir_mercado"], axis=1)
            concat = concat.reindex(columns=l)
            concat.to_csv("R:/10 CARTERAS ADMINISTRADAS/CAD RD/Ficheros de Carga/FicherosPQ/"+self.fondo.currentText()+self.fecha.text()+"PQ.csv", index=False,sep = ';')
            self.progress.setValue(95)
            carteras = concat.copy()
            fondito = self.fondo.currentText().strip() + self.Codigo_Cartera.text().strip()
            print(fondito)
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
    app.setWindowIcon(QtGui.QIcon('icon.png'))
    window = Ui()
    window.show()
    app.exec_()
