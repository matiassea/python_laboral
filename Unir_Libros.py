# -*- coding: utf-8 -*-
"""
Created on Mon May  4 06:37:43 2020

@author: mvidal2
"""

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import datetime
import os

frames=[]
x = datetime.datetime.now()
ruta='C:/Users/mvidal2/Desktop/data scientist/excel/Peru (xlsx)/'

for n in range(len(os.listdir(ruta))):
#for files in os.listdir(ruta):
#for n in range(len(num)):
    Nombre_Libro="C:/Users/mvidal2/Desktop/data scientist/excel/Peru (xlsx)/"+'A (' + str(n) + ').xlsx'
    print(Nombre_Libro)
    db1=pd.read_excel(Nombre_Libro, skiprows = range(1,2),usecols = "B:C",index_col=False)
    #db2=db1.drop(['Nro', 'Tipo Compra','Razon Social','Fecha Docto','Fecha Recepcion','Fecha Acuse','Monto Exento','Monto Neto','Monto IVA Recuperable','Monto Iva No Recuperable','Codigo IVA No Rec.','Monto Total','Monto Neto Activo Fijo','IVA Activo Fijo','IVA uso Comun','Impto. Sin Derecho a Credito','IVA No Retenido','Tabacos Puros','Tabacos Cigarrillos','Tabacos Elaborados','NCE o NDE sobre Fact. de Compra','Codigo Otro Impuesto','Valor Otro Impuesto','Tasa Otro Impuesto'],axis=1)
    #db2['RUT_Receptor']=[RUT]*len(db2['Tipo Doc'])
    #Solo_Facturas = db2[(db2['Tipo Doc']==33)]
    #Solo_Facturas = db1[(db1['CATALOGO'])]
    frames.append(db1)
    #Definitivo = pd.concat(db1)
Solo_Facturas = pd.concat(frames, ignore_index=True)
Solo_Facturas.to_excel("Resumen.xlsx",index = True) 
#nombre_con_hora=
#Solo_Facturas.to_excel("Resumen_libros_compras"+"-" + x.strftime("%A")+"-"+x.strftime("%H")+"-"+x.strftime("%M")+"-"+x.strftime("%S") +".xlsx",index = False)