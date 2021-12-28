# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 12:00:43 2020
@author: mvidal2
Version Impresion OC V5
Habilitacion OC V1
Producto V1
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import win32com.client
from win32com.client import Dispatch, constants
import os
import socket
#import Envio_email

class MainApplication(tk.Frame):
    #https://stackoverflow.com/questions/17466561/best-way-to-structure-a-tkinter-application
    #https://stackoverflow.com/questions/40128061/tkinter-class-structure-class-per-frame-issue-with-duplicating-widgets
    #https://www.begueradj.com/tkinter-best-practices/
    
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.X = tk.StringVar()
        self.db1=tk.StringVar()
        self.Y=tk.StringVar()
        self.var=tk.StringVar() 
        self.db3=tk.Listbox()
        self.entry_var=tk.StringVar()
        self.mensaje=tk.StringVar()
        password=tk.StringVar()
        usuario=tk.StringVar()
        

        ############################# Frame #############################        
        self.frame = tk.Frame(self)
        self.frame.grid(row=0, column=0)

        #Colocando Gadget
        self.button1 = tk.Button(self.frame,text='Open File',font=('arial',10,'bold'), command=self.mfileopen)
        self.button1.grid(row=0, column=0,padx=10,pady=10)
        

        self.button3 = tk.Button(self.frame,text='Procesar',font=('arial',10,'bold'), command=self.procesar)
        self.button3.grid(row=0, column=2,padx=10,pady=10)
        
        ############################# Frame 2 #############################
        self.frame2 = tk.Frame(self)
        self.frame2.grid(row=1, column=0)        
        
        self.etiqueta_usuario = tk.Label(self.frame2,text="Password",font=('arial',10,'bold'))
        self.etiqueta_usuario.grid(row=2, column=1,padx=10,pady=10)

        self.letrero_password=tk.Entry(self.frame2,fg='blue', width = 20, show="*") 
        self.letrero_password.grid(row=2, column=2,padx=10,pady=10)

        self.etiqueta_password = tk.Label(self.frame2,text="Usuario",font=('arial',10,'bold'))
        self.etiqueta_password.grid(row=1, column=1,padx=10,pady=10)

        self.letrero_usuario=tk.Entry(self.frame2,fg='blue', width = 20)
        self.letrero_usuario.grid(row=1, column=2,padx=10,pady=10)
        
    
        #https://recursospython.com/guias-y-manuales/caja-de-texto-entry-tkinter/
        #https://recursospython.com/guias-y-manuales/posicionar-elementos-en-tkinter/
        
        """
        Label etiqueta
        http://effbot.org/tkinterbook/label.htm
        """
        ############################# Frame 3 ############################# 
        self.frame3=tk.Frame(self)
        self.frame3.grid(row=3, column=0)

        self.observacion1 = tk.Label(self.frame3, text="Instrucciones de uso",font=('arial',10,'bold'))
        self.observacion1.grid(row=3, column=0,padx=5,pady=5)
        
        self.observacion1 = tk.Label(self.frame3, text="Debe cargar plantilla excel, aprentando boton 'Open File'",font=('arial',10))
        self.observacion1.grid(row=4, column=0,padx=5,pady=5)
        #sticky="n" (norte), "s" (sur), "e" (este) o "w" (oeste) 
        
        self.observacion1 = tk.Label(self.frame3, text="Este excel debe tener los siguientes campos:",font=('arial',10))
        self.observacion1.grid(row=5, column=0,padx=5,pady=5)

        self.observacion1 = tk.Label(self.frame3, text="'BU' = Unidad de negocio",font=('arial',10))
        self.observacion1.grid(row=6, column=0,padx=5,pady=5)
        
        self.observacion1 = tk.Label(self.frame3, text="'OC' = N° de OC",font=('arial',10))
        self.observacion1.grid(row=7, column=0,padx=5,pady=5)

        self.observacion1 = tk.Label(self.frame3, text="Primero proceso: Se habilitara cada OC para su impresion",fg='blue',font=('arial',10))
        self.observacion1.grid(row=8, column=0,padx=5,pady=5)

        self.observacion1 = tk.Label(self.frame3, text="Segundo proceso: Se sacara la impresion en PDF de las OC que esten habilitadas",fg='blue',font=('arial',10))
        self.observacion1.grid(row=9, column=0,padx=5,pady=5)

        # self.observacion1 = tk.Label(self.frame2, text="'RUT_Emisor'= RUT del emisor de la factura, se obtiene del libro de compras",font=('arial',10))
        # self.observacion1.grid(row=8, column=0,padx=5,pady=5)
        
        ############################# Frame 4 ############################# 
        self.frame4=tk.Frame(self)
        self.frame4.grid(row=11, column=0)
        
        #Entry            
        self.Link=tk.Entry(self.frame4,width = 80)
        self.Link.grid(row=11, column=1,padx=5,pady=5)

        #Label
        self.etiqueta = tk.Label(self.frame4, text="Ingrese el link donde bajaran OC",fg='green',font=('arial',10,'bold'))
        self.etiqueta.grid(row=11, column=0,padx=5,pady=5)
            
        
        ############################# Frame 5 ############################# 

        self.frame5=tk.Frame(self)
        self.frame5.grid(row=12, column=0)
        
        """
        Text permite mostrar texto con varios estilos y atributos, soporta imagenes y ventanas   
        https://www.python-course.eu/tkinter_text_widget.php
        """
        
        """
        Scroll
        https://stackoverflow.com/questions/19646752/python-scrollbar-on-text-widget/19647325
        """
        scroll = tk.Scrollbar(self.frame5)
        scroll.grid(row=12, column=1,sticky="n"+"s"+"w")

        self.letrero1=tk.Text(self.frame5,wrap='none',padx=10,pady=20,width=70,height=15,yscrollcommand=scroll.set) #width=80,height=10,  
        self.letrero1.config(yscrollcommand=scroll.set) #width=80,height=10,  
        scroll.config(command=self.letrero1.yview)
        self.letrero1.grid(row=12, column=0)            

    
    def procesar(self):
        global db1
        ########################### Habilitacion de la OC #####################
        
        
        
        self.letrero1.insert(tk.END, "Email enviado a director\n")
        self.letrero1.config(fg = 'green',height=12)
        
        start1 = time.time()
        #driver = webdriver.Firefox(fp)
        driver = webdriver.Firefox()
        driver.get("http://www.google.com/")
        
        #######################################################################
        #Apertura del PeopleSoft
        
        #open tab
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't') 
        
        # Load a page 
        driver.get('https://leifs.mycmsc.com/psp/leifsprd/EMPLOYEE/ERP/?cmd=logout')
        
       
        # username.send_keys("311800185")
        # password.send_keys("Danola11.")
        
        #Cambiar a Ingles
        time.sleep(3)
        driver.find_element_by_css_selector(".pslanguageframe > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").click() #CSS Selector
        

        time.sleep(1)
        username = driver.find_element_by_id("userid") #input id o name
        password = driver.find_element_by_id("pwd") #input id o name
        
        time.sleep(1)
        username.send_keys(self.letrero_usuario.get())
        password.send_keys(self.letrero_password.get())
        
        driver.find_element_by_name("Submit").click() #name
        
        #######################################################################
        end1 = time.time()
        #print(end1 - start1, " Tiempo de comando ingreso de password")
        start2 = time.time()
        #######################################################################
        
        #Comienzo del ciclo
        #Main Menu
        driver.implicitly_wait(2)
        driver.find_element_by_id("pthnavbca_PORTAL_ROOT_OBJECT").click() #ID
        
        #Legal Contracts
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPCO_CONTRACT_MANAGEMENT").click() #ID
        
        #Related Links
        driver.implicitly_wait(2)
        driver.find_element_by_id("EPCO_GENERAL_SETUP").click() #ID
        
        #Add/Update PO
        driver.implicitly_wait(5)
        driver.find_element_by_css_selector("#crefli_EP_PURCHASE_ORDER_GBL_01 > a:nth-child(1)").click() #CSS Selector
        
        #Publicacion de informacion
        # self.letrero1.insert(tk.END, "Ingreso a menu\n")
        # self.letrero1.config(fg = 'green',height=12)
        
        #Find an Existing Value
        time.sleep(1)
        driver.switch_to.default_content()
        time.sleep(1)
        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
        driver.switch_to.frame("ptifrmtgtframe")
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.PSHYPERLINK:nth-child(1)"))).click() 
            
        A=db1['BU']
        B=db1['OC']
        self.Observaciones=[]
        self.BU=[]
        self.OC=[]
        try:
            for n in range (len(db1['BU'])):
                start1 = time.time()
                self.a=str(A[n]) #BU
                self.b=str("0")*(10-len(str(B[n])))+str(B[n]) #OC
                
                
                #Ingresando unidad de negocio y OC
                time.sleep(2)
                driver.switch_to.default_content()
                time.sleep(1)
                driver.switch_to.frame("ptifrmtgtframe")
                time.sleep(1)
                #Colocar BU
                driver.find_element_by_id("PO_SRCH_BUSINESS_UNIT").click()
                time.sleep(1)
                driver.find_element_by_id("PO_SRCH_BUSINESS_UNIT").clear()
                time.sleep(1)
                driver.find_element_by_id("PO_SRCH_BUSINESS_UNIT").send_keys(self.a)
                time.sleep(1)
                #Colocar OC
                driver.find_element_by_id("PO_SRCH_PO_ID").click()
                time.sleep(1)
                driver.find_element_by_id("PO_SRCH_PO_ID").clear()
                time.sleep(1)
                driver.find_element_by_id("PO_SRCH_PO_ID").send_keys(self.b)
                #Bajar la pagina
                driver.execute_script("window.scrollTo(0, 600)")
                #Buscar
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "#ICSearch"))).click()
                time.sleep(4)
                
                #Cuando la OC no existe
                Titulo=(By.CLASS_NAME, "PSSRCHINSTRUCTIONS")
                try: 
                    WebDriverWait(driver, 2).until(EC.text_to_be_present_in_element((Titulo),"No matching values were found."))
                    self.mensaje="OC no existe"
                    self.Observaciones.append(self.mensaje)
                    self.BU.append(self.a)
                    self.OC.append(self.b)
                    self.mensaje="OC no existe" + ", BU " + self.a  + ", OC " + self.b
                    print(self.mensaje)
                    self.letrero1.insert(tk.END, self.mensaje+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()
                    continue
                except TimeoutException:
                    pass
                
                #There are 2 distribution lines whose budget status is either error or warning
                try:
                    #Ventana emergente
                    driver.switch_to.default_content()
                    #time.sleep(1)  
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICOK"))).click()
                    self.mensaje="Error de presupuesto"
                    self.Observaciones.append(self.mensaje)
                    self.BU.append(self.a)
                    self.OC.append(self.b)
                    self.mensaje="error de presupuesto" + ", BU " + self.a  + ", OC " + self.b
                    print(self.mensaje)
                    self.letrero1.insert(tk.END, self.mensaje+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()
                except TimeoutException:
                    pass
                    
                try:
                    #Ventana emergente
                    driver.switch_to.default_content()
                    #time.sleep(1)  
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICOK"))).click()
                    self.mensaje="Error de presupuesto"
                    self.Observaciones.append(self.mensaje)
                    self.BU.append(self.a)
                    self.OC.append(self.b)
                    self.mensaje="error de presupuesto" + ", BU " + self.a  + ", OC " + self.b
                    print(self.mensaje)
                    self.letrero1.insert(tk.END, self.mensaje+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()
                    
                except TimeoutException:
                    pass
        
                try:
                    #Ventana emergente
                    driver.switch_to.default_content()
                    #time.sleep(1)  
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICOK"))).click()
                    self.mensaje="Error de presupuesto"
                    self.Observaciones.append(self.mensaje)
                    self.BU.append(self.a)
                    self.OC.append(self.b)
                    self.mensaje="error de presupuesto" + ", BU " + self.a  + ", OC " + self.b
                    print(self.mensaje)
                    self.letrero1.insert(tk.END, self.mensaje+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()

                except TimeoutException:
                    pass
                    
                # Maintain Purchase Order
                time.sleep(1)
                driver.switch_to.default_content()
                time.sleep(1)
                #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                driver.switch_to.frame("ptifrmtgtframe")
                #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                #Listado desplegable Dispatch method
                time.sleep(1)
                List_current_status=Select(driver.find_element_by_id("PO_HDR_DISP_METHOD"))
                time.sleep(1)
                List_current_status.select_by_visible_text('Print')
                time.sleep(1)
                List_future_status=Select(driver.find_element_by_id("PO_HDR_DISP_METHOD"))
                time.sleep(1)
                #Bajar la pagina
                driver.execute_script("window.scrollTo(0, 600)")
                time.sleep(1)
                #Guardar
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "#ICSave"))).click()  
                time.sleep(1)
                #Volver a la busqueda
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()      
                
                #Trabajando con el link
                #https://stackoverflow.com/questions/25146960/python-convert-back-slashes-to-forward-slashes/25147093
                ruta1=self.Link.get().replace(os.sep, '/')
                ruta=str("")+str(ruta1)+str("")
                #print(str(self.Link.get()))
                #print(ruta)
                
                #Asegurando salida con Purchase Order
                Titulo=(By.CLASS_NAME, "PSSRCHTITLE")
                try: 
                    WebDriverWait(driver, 5).until(EC.text_to_be_present_in_element((Titulo),"Purchase Order"))
                except TimeoutException:
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(1)
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                    time.sleep(2)
                    
                Titulo=(By.CLASS_NAME, "PSSRCHTITLE")
                try: 
                    WebDriverWait(driver, 5).until(EC.text_to_be_present_in_element((Titulo),"Purchase Order"))
                except TimeoutException:
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(1)
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                    
                end1 = time.time()    
                self.mensaje="Proceso OK"
                self.Observaciones.append(self.mensaje)
                self.BU.append(self.a)
                self.OC.append(self.b)
                self.mensaje="Termine de habilitar la OC " + self.b +" , me demore " + str(int(end1-start1))
                print(self.mensaje)
                self.letrero1.insert(tk.END, self.mensaje+"\n")
                self.letrero1.config(fg = 'green',height=12)
                self.letrero1.update()

                
            #db2=pd.DataFrame(columns=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'])
            df1 = pd.DataFrame({'BU':self.BU})
            df2 = pd.DataFrame({'OC':self.OC})
            df3 = pd.DataFrame({'Observaciones':self.Observaciones})
                
            #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
            db1=pd.concat([df1,df2,df3], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
            db1.columns=["BU","OC","Observaciones"]
            db1.to_excel("Resumen_del_proceso_habilitacion.xlsx",index = False)
            driver.close()
        
        except TimeoutException:
        
                self.mensaje= "me quede pegado en Psoft"
                self.Observaciones.append(self.mensaje)
                self.mensaje= "me quede pegado en Psoft, OC " + self.b
                print(self.mensaje)
                self.letrero1.insert(tk.END, self.mensaje+"\n")
                self.letrero1.config(fg = 'red',height=12)
                self.letrero1.update()

                    
                #db2=pd.DataFrame(columns=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'])
                df1 = pd.DataFrame({'BU':self.BU})
                df2 = pd.DataFrame({'OC':self.OC})
                df3 = pd.DataFrame({'Observaciones':self.Observaciones})
                    
                #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
                db1=pd.concat([df1,df2,df3], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
                db1.to_excel("Resumen_del_proceso_habilitacion.xlsx",index = False)
                
        
        
        ############################### OC con habilitacion completa #############################
        # El mensaje es print("Proceso OK")
        # En la columna Observaciones o el libro "Resumen_del_proceso_habilitacion"
        
        
        self.letrero1.insert(tk.END, "se bajaran los archivo en la carpeta " + self.Link.get() + "\n")
        self.letrero1.config(fg = 'green',height=12)
        self.letrero1.update()
        
        Link=str(self.Link.get()).replace('\\','\\\\')
        
        
        ########################### Impresion de la OC #############################################
        
        fp = webdriver.FirefoxProfile()
        fp.set_preference("browser.download.folderList", 2)
        #browser.download.folderList tells it not to use default Downloads directory and use directory whatever we want to give.
        fp.set_preference("browser.helperApps.alwaysAsk.force", False);
        fp.set_preference("browser.download.manager.showWhenStarting", False)
        #(“browser.download.manager.showWhenStarting”, False) – disabling Download Manager window when a download begins i.e. turns of showing download progress.
        fp.set_preference("browser.download.manager.showAlertOnComplete", False)
        #(“browser.download.manager.showAlertOnComplete”, False) – popup window at bottom right corner of the screen will not appear once all downloads are finished
        fp.set_preference('browser.helperApps.neverAsk.saveToDisk','application/pdf') #'application/pdf', 'text/csv','application/xls'
        #(“browser.helperApps.neverAsk.saveToDisk”, “text/csv”) – list of MIME types to save to disk without asking what to use to open the file – (in different file types like CSV,XLSB,XLSX,.RTF etc ). This setting is actually disabling download dialog box i.e. tells Firefox to automatically download the files of the selected mime-types(I will tell what is MIME type)
        fp.set_preference("browser.download.dir", Link) #'C:/Users/mvidal2/Downloads'
        #fp.set_preference("browser.download.dir", download_dir)
        
        #https://stackoverflow.com/questions/23800195/auto-download-pdf-in-firefox
        #https://yizeng.me/2014/05/23/download-pdf-files-automatically-in-firefox-using-selenium-webdriver/
        #https://yizeng.me/2014/05/23/download-pdf-files-automatically-in-firefox-using-selenium-webdriver/
        
        
        #https://stackoverflow.com/questions/45589571/how-to-auto-download-through-firefox-browser-using-firefoxprofile
        #https://stackoverflow.com/questions/30452395/selenium-pdf-automatic-download-not-working
        #https://stackoverflow.com/questions/52208798/selenium-problems-with-pdf-download-in-firefox
        
        fp.set_preference("pdfjs.disabled", True)
        
        #Use this to disable Acrobat plugin for previewing PDFs in Firefox (if you have Adobe reader installed on your computer)
        fp.set_preference("plugin.scan.Acrobat", "99.0");
        fp.set_preference("plugin.scan.plid.all", False);
        
        start1 = time.time()
        driver = webdriver.Firefox(fp)
        driver.get("http://www.google.com/")
        
        #######################################################################
        #Apertura del PeopleSoft
        
        #open tab
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't') 
        
        # Load a page 
        driver.get('https://leifs.mycmsc.com/psp/leifsprd/EMPLOYEE/ERP/?cmd=logout')
        
        #Cambiar a Ingles
        time.sleep(3)
        driver.find_element_by_css_selector(".pslanguageframe > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").click() #CSS Selector
        

        time.sleep(1)
        username = driver.find_element_by_id("userid") #input id o name
        password = driver.find_element_by_id("pwd") #input id o name
        
        # username.send_keys("311800185")
        # password.send_keys("Danola11.")
        
        time.sleep(1)        
        username.send_keys(self.letrero_usuario.get())
        password.send_keys(self.letrero_password.get())

        # para llevarlo a español
        # time.sleep(1)
        # driver.find_element_by_css_selector(".pslanguageframe > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > a:nth-child(1)").click()
        # #para llevarlo a ingles
        # driver.find_element_by_css_selector(".pslanguageframe > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").click()
        
        
        driver.find_element_by_name("Submit").click() #name
        
        #######################################################################
        end1 = time.time()
        #######################################################################
        
        #Comienzo del ciclo
        #Main Menu
        driver.implicitly_wait(2)
        driver.find_element_by_id("pthnavbca_PORTAL_ROOT_OBJECT").click() #ID
        
        #Purchasing
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPPO_PURCHASING").click() #ID
        
        #Purchase Order
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPCO_PURCHASE_ORDERS1").click() #ID
        
        #Dispatch PO
        driver.implicitly_wait(5)
        driver.find_element_by_css_selector("#crefli_EP_PO_DISPATCH_GBL > a:nth-child(1)").click() #CSS Selector
        
        #Publicacion de informacion
        # self.letrero1.insert(tk.END, "Ingreso a menu\n")
        # self.letrero1.config(fg = 'green',height=12)
        
                
        #db1 = pd.read_excel("Base(1).xlsx")
        db1 = pd.read_excel("Resumen_del_proceso_habilitacion.xlsx")
        A=db1['BU']
        B=db1['OC']
        C=db1['Observaciones']
        self.Observaciones2=[]
        self.BU=[]
        self.OC=[]
        
        try:
                for n in range (len(db1['BU'])):
                    start2 = time.time()
                    self.a=str(A[n]) #BU
                    self.b=str("0")*(10-len(str(B[n])))+str(B[n]) #OC
                    self.c=str(C[n]) #Observaciones
                    if self.c.find("Proceso OK") != -1: #Si encuentra la palabra Chile               
                        #Colocar Unidad de negocio
                        time.sleep(2)
                        driver.switch_to.default_content()
                        time.sleep(2)
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        driver.switch_to.frame("ptifrmtgtframe")
                        time.sleep(2)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICSearch")))
                        driver.find_element_by_id("#ICSearch").click()
                        time.sleep(2)
                    else:
                        continue            
                    try:
                        if driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(4) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(4) > td:nth-child(1) > a:nth-child(1)").click()
                                
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(1) > a:nth-child(1)").click()
                                    
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(6) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(6) > td:nth-child(1) > a:nth-child(1)").click()
                                        
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(7) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(7) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(8) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(8) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(9) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(9) > td:nth-child(1) > a:nth-child(1)").click()

                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(10) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(10) > td:nth-child(1) > a:nth-child(1)").click()

                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(11) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(11) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(12) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(12) > td:nth-child(1) > a:nth-child(1)").click()

                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(13) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(13) > td:nth-child(1) > a:nth-child(1)").click()

                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(14) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(14) > td:nth-child(1) > a:nth-child(1)").click()

                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(15) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(15) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        else:
                            self.mensaje= "No encontre la unidad de negocio " + self.a
                            self.letrero1.insert(tk.END, self.mensaje+"\n")
                            self.letrero1.config(fg = 'red',height=12)
                            self.letrero1.update()
                            continue
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(1) > a:nth-child(1)
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(4) > td:nth-child(1) > a:nth-child(1)
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(1) > a:nth-child(1)
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(1) > a:nth-child(1)
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(6) > td:nth-child(1) > a:nth-child(1)
                        #PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(7) > td:nth-child(1) > a:nth-child(1)
                        
                        #En caso de no encontrar la unidad de negocio
                    except NoSuchElementException:
                        self.mensaje= "me quede pegado en Psoft, OC " + self.a
                        print(self.mensaje)
                        self.letrero1.insert(tk.END, self.mensaje+"\n")
                        self.letrero1.config(fg = 'red',height=12)
                        self.letrero1.update()
                        continue
                
                    #Colocar OC
                    time.sleep(3)
                    driver.find_element_by_css_selector("#RUN_CNTL_PUR_PO_ID").click()
                    time.sleep(2)
                    driver.find_element_by_css_selector("#RUN_CNTL_PUR_PO_ID").clear()
                    time.sleep(2)
                    driver.find_element_by_css_selector("#RUN_CNTL_PUR_PO_ID").send_keys(self.b)
                                    
                    #Apretar RUN
                    time.sleep(2)
                    driver.find_element_by_css_selector("#PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB").click()
            
                    if self.a.find("CHL") != -1: #Si encuentra la palabra Chile
                        #Impresión de PO para Chile
                        time.sleep(2)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PRCSRQSTDLG_WRK_SELECT_FLAG$0")))
                        time.sleep(1)
                        
                        if driver.find_element_by_id("PRCSRQSTDLG_WRK_SELECT_FLAG$0").is_selected() == False:
                            driver.find_element_by_id("PRCSRQSTDLG_WRK_SELECT_FLAG$0").click()
                            
                        #Seleccionar el tipo, window
                        time.sleep(2)
                        List_Type=driver.find_element_by_css_selector("#PRCSRQSTDLG_WRK_OUTDESTTYPE\$0")
                        List_Type_elegir=Select(List_Type)
                        time.sleep(0.5)
                        List_Type_elegir.select_by_visible_text('Window')
                        
                        #Apretar OK
                        time.sleep(2)
                        driver.find_element_by_id("#ICSave").click()
                    
                    elif self.a.find("PER") != -1: #Si encuentra la palabra PER
                        #Impresión de PO para Peru
                        time.sleep(2) 
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PRCSRQSTDLG_WRK_SELECT_FLAG$2")))
                        time.sleep(1)
                        
                        if driver.find_element_by_id("PRCSRQSTDLG_WRK_SELECT_FLAG$2").is_selected() == False:
                            driver.find_element_by_id("PRCSRQSTDLG_WRK_SELECT_FLAG$2").click()
                        
                        #Seleccionar el tipo, window
                        time.sleep(2)
                        List_Type=driver.find_element_by_css_selector("#PRCSRQSTDLG_WRK_OUTDESTTYPE\$2")
                        List_Type_elegir=Select(List_Type)
                        time.sleep(0.5)
                        List_Type_elegir.select_by_visible_text('Window')
                        
                        #Apretar OK
                        time.sleep(2)
                        driver.find_element_by_id("#ICSave").click()
                        
                    #ventana para bajar PDF en Chile y Peru similar
                    #STATUS = Queued
                    #https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python
                    time.sleep(4)
                
                    
                    #file_path='/Users/mvidal2/Desktop/data scientist/Impresion OC/'
                    #file_path=self.Link.get()
                    #print("Comienza la espera por donwload pdf")
                    time.sleep(40)
                    #print("Comienza la busqueda")
                    #Esta es la ruta que funciona en Listdir
                    #ruta='C:/Users/mvidal2/Desktop/data scientist/Impresion OC/'
                    #ruta=str(self.Link.get())
                    #ruta=str(self.Link.get()).replace('\\', '\\\\')
                    #print(ruta)
                    time.sleep(1)
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(1)
                    driver.get("about:downloads")
                    time.sleep(1)
                    fileName = driver.execute_script("return document.querySelector('#contentAreaDownloadsView .downloadMainArea .downloadContainer description:nth-of-type(1)').value")
                    time.sleep(1)
                    #print(filename)
                    driver.close()
                    
                    """
                    Documentacion por link
                    https://stackoverflow.com/questions/25146960/python-convert-back-slashes-to-forward-slashes/25147093%7D
                    https://stackoverflow.com/questions/22207936/python-how-to-find-files-and-skip-directories-in-os-listdir
                    https://docs.python.org/3/library/os.html#os.supports_fd
                    https://stackoverflow.com/questions/431684/how-do-i-change-the-working-directory-in-python
                    
                    
                    """
                    #Se cambia de ruta a Link
                    for file in os.listdir(ruta):
                        try:
                            if file.endswith(fileName):
                                end2=time.time()
                                self.mensaje="Fue bajada la OC" + self.b + ", de la BU " + self.a + ", me demore " + str(int(end2-start2)) + " seg."
                                #os.rename(fileName,self.a+self.b+".PDF")
                                os.rename(os.path.join(ruta,fileName),os.path.join(ruta,self.a+self.b+".PDF"))
                                print(self.mensaje)
                                self.Observaciones2.append(self.mensaje)
                                self.BU.append(self.a)
                                self.OC.append(self.b)
                                self.letrero1.insert(tk.END, self.mensaje+"\n")
                                self.letrero1.config(fg = 'green',height=12)
                                self.letrero1.update()
                                # time.sleep(2)
                                driver.switch_to.window(driver.window_handles[0])
                                time.sleep(2)
                                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB")))
                                #Bajar pantalla
                                driver.execute_script("window.scrollTo(0, 600)")
                                #Hacer click en retornar a la busqueda
                                time.sleep(1)
                                #WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                                driver.find_element_by_id("#ICList").click() 
                                time.sleep(3)
                                #Hacer click en la unidad de negocio
                                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).click()
                                #Borrar la unidad de negocio
                                time.sleep(1)
                                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).clear()
                                #driver.close()
                                break
                            #En el ciclo de verificiacion de bajada del archivo
                            else:
                                self.mensaje="No baje la OC " + self.b + ", BU " + self.a
                                #Observaciones.append(mensaje)
                                #print(mensaje)
                                #driver.switch_to.window(driver.window_handles[0])
                                continue
                                # time.sleep(2)
                                # driver.execute_script("window.scrollTo(0, 600)")
                                # #Hacer click en retornar a la busqueda
                                # time.sleep(1)
                                # #WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                                # driver.find_element_by_id("#ICList").click() 
                                # time.sleep(3)
                                # #Hacer click en la unidad de negocio
                                # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).click()
                                # #Borrar la unidad de negocio
                                # time.sleep(1)
                                # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).clear()
                                # #driver.close()
                        except FileExistsError:
                            self.mensaje="OC ya bajada, con nombre " + self.a + ", BU " + self.b 
                            print(self.mensaje)
                            self.Observaciones2.append(self.mensaje)
                            self.BU.append(self.a)
                            self.OC.append(self.b)
                            self.letrero1.insert(tk.END, self.mensaje+"\n")
                            self.letrero1.config(fg = 'blue',height=12)
                            self.letrero1.update()
                            continue
                
                    if self.mensaje.find("OC ya bajada, con nombre") != -1 or self.mensaje.find("No baje la OC") != -1:  #mensaje=="No baje el archivo":
                            #print("OC " + b +", no bajada")
                            time.sleep(2)
                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(2)
                            #Detectar si esta el boton run
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB")))
                            time.sleep(2)
                            driver.execute_script("window.scrollTo(0, 600)")
                            #Hacer click en retornar a la busqueda
                            time.sleep(1)
                            #WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                            driver.find_element_by_id("#ICList").click() 
                            time.sleep(3)
                            #Hacer click en la unidad de negocio
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).click()
                            #Borrar la unidad de negocio
                            time.sleep(1)
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "RUN_CNTL_PUR_RUN_CNTL_ID"))).clear()
                            #driver.close()
                    else:    
                        #print("Ya termine un ciclo")
                        continue
                                
                db1.reset_index(drop=True)
                df1 = pd.DataFrame({'BU':self.BU})
                df2 = pd.DataFrame({'OC':self.OC})
                df3 = pd.DataFrame({'Observaciones':self.Observaciones2})
                db1=pd.concat([df1,df2,df3], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
        
                #self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx\n")
                #self.letrero1.update()
                #self.letrero1.config(fg = 'blue',height=12)
                db1.columns=["BU","OC","Observaciones"]
                db1.to_excel("Resumen_del_proceso_impresion_OC.xlsx",index = False)
                self.mensaje="Favor revisar archivo Resumen_del_proceso_impresion_OC"
                self.letrero1.insert(tk.END, self.mensaje+"\n")
                self.letrero1.config(fg = 'blue',height=12)
                self.letrero1.update()
                driver.close()
                        
        except NoSuchElementException:
            self.Observaciones2.append(self.mensaje)
            df1 = pd.DataFrame({'BU':self.BU})
            df2 = pd.DataFrame({'OC':self.OC})
            df3 = pd.DataFrame({'Observaciones':self.Observaciones2})
        
            #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
            db1=pd.concat([df1,df2,df3], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
            db1.columns=["BU","OC","Observaciones"]
            db1.to_excel("Resumen_del_proceso_impresion_OC.xlsx",index = False)         
    
    def mfileopen(self):
        global filename
        global db1
        global Y
        #global var
        filename = filedialog.askopenfilename()
        db1 = pd.read_excel(filename)
        self.mensaje="Import successfully!"
        self.letrero1.insert(tk.END, "Archivo cargado!\n")
        self.letrero1.config(fg = 'green',height=12)
        #print(self.mensaje)
        
    def destroy(self):
        self.quit()
        #self.parent.destroy()
        #self.destroy()
        
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Robot Habilitacion e impresion de OC")
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.resizable(0,0)
    app=MainApplication(root)
    root.mainloop()