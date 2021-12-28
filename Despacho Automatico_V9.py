# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 12:00:43 2020
Los email en HYML, bajar el email, abrir con navegador apretar ctrl+U, copiar y pegar
https://developer.mozilla.org/en-US/docs/Web/HTML/Element/input/file

@author: mvidal2
https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send
https://stackoverflow.com/questions/24650518/python-send-html-formatted-email-via-outlook-2007-2010-and-win32com
https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send
https://gist.github.com/ITSecMedia/b45d21224c4ea16bf4a72e2a03f741af
https://stackoverflow.com/questions/50926514/send-email-through-python-using-outlook-2016-without-opening-it
https://automatetheboringstuff.com/chapter16/
http://eyana.me/send-emails-in-outlook-using-python/

HTML
https://www.w3schools.com/jsref/dom_obj_fileupload.asp
https://www.w3schools.com/tags/default.asp
https://www.w3schools.com/html/html_links.asp
https://www.w3schools.com/tags/tag_td.asp
https://developer.mozilla.org/en-US/docs/Web/HTML/Element/input/file


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
import socket
import os
from Texto_email.Textos_email import *
from Envio_email.Email import *

class MainApplication(tk.Frame):
    
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
        
        #https://recursospython.com/guias-y-manuales/caja-de-texto-entry-tkinter/
        #https://recursospython.com/guias-y-manuales/posicionar-elementos-en-tkinter/

        
        ############################# Frame #############################        
        self.frame = tk.Frame(self)
        self.frame.grid(row=0, column=0)

        #Colocando Gadget
        self.button1 = tk.Button(self.frame,text='Open File',font=('arial',10,'bold'), command=self.mfileopen)
        self.button1.grid(row=0, column=0,padx=10,pady=5)
        
        self.button2 = tk.Button(self.frame,text='Procesar',font=('arial',10,'bold'), command=self.procesar_habilitacion_OC)
        self.button2.grid(row=0, column=1,padx=10,pady=5)
        
        ############################# Frame 2 #############################
        self.frame2 = tk.Frame(self)
        self.frame2.grid(row=1, column=0)        
        
        self.etiqueta_usuario = tk.Label(self.frame2,text="Password",font=('arial',10,'bold'))
        self.etiqueta_usuario.grid(row=2, column=1,padx=10,pady=5)

        self.letrero_password=tk.Entry(self.frame2,fg='blue', width = 20, show="*") 
        self.letrero_password.grid(row=2, column=2,padx=10,pady=5)

        self.etiqueta_password = tk.Label(self.frame2,text="Usuario",font=('arial',10,'bold'))
        self.etiqueta_password.grid(row=1, column=1,padx=10,pady=5)

        self.letrero_usuario=tk.Entry(self.frame2,fg='blue', width = 20)
        self.letrero_usuario.grid(row=1, column=2,padx=10,pady=5)
        
    
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
        self.observacion1.grid(row=3, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Debe cargar plantilla excel, aprentando boton 'Open File'",font=('arial',10))
        self.observacion1.grid(row=4, column=0,padx=2,pady=2)
        #sticky="n" (norte), "s" (sur), "e" (este) o "w" (oeste) 

        self.observacion1 = tk.Label(self.frame3, text="Cada vez que presiona Procesar, se envia email a director solicitando autorizacion de uso",font=('arial',10))
        self.observacion1.grid(row=5, column=0,padx=2,pady=2)
               
        self.observacion1 = tk.Label(self.frame3, text="En el email de autorizacion, se adjunta excel de carga y nombre del usuario que opera robot.",font=('arial',10))
        self.observacion1.grid(row=6, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Se utiliza en PC " + socket.gethostname() + ". Numero IP " + socket.gethostbyname(socket.gethostname()),font=('arial',10))
        self.observacion1.grid(row=7, column=0,padx=2,pady=2)
               
        self.observacion1 = tk.Label(self.frame3, text="Primero proceso: Se habilitara las OC para su impresion",fg='blue',font=('arial',10))
        self.observacion1.grid(row=8, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Segundo proceso: Se obtiene la OC en PDF, solo OC habilitadas",fg='blue',font=('arial',10))
        self.observacion1.grid(row=9, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Tercer proceso: Se enviara las OC via email",fg='blue',font=('arial',10))
        self.observacion1.grid(row=10, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Cada termino de operacion debe revisar el excel Resumen_de_proceso",font=('arial',10),fg='red')
        self.observacion1.grid(row=11, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Su uso inapropiado sera sancionado",font=('arial',10),fg='red')
        self.observacion1.grid(row=12, column=0,padx=2,pady=2)
        

        ############################# Frame 4 ############################# 
        self.frame4=tk.Frame(self)
        self.frame4.grid(row=16, column=0)
        
        #Entry            
        self.Link=tk.Entry(self.frame4,width = 80)
        self.Link.grid(row=16, column=1,padx=5,pady=5)

        #Label
        self.etiqueta = tk.Label(self.frame4, text="Ingrese el link donde bajaran OC",fg='green',font=('arial',10,'bold'))
        self.etiqueta.grid(row=16, column=0,padx=5,pady=5)
            
        
        ############################# Frame 5 ############################# 

        self.frame5=tk.Frame(self)
        self.frame5.grid(row=17, column=0)
        
        """
        Text permite mostrar texto con varios estilos y atributos, soporta imagenes y ventanas   
        https://www.python-course.eu/tkinter_text_widget.php
        """
        
        """
        Scroll
        https://stackoverflow.com/questions/19646752/python-scrollbar-on-text-widget/19647325
        """
        scroll = tk.Scrollbar(self.frame5)
        scroll.grid(row=17, column=1,sticky="n"+"s"+"w")

        self.letrero1=tk.Text(self.frame5,wrap='none',padx=10,pady=20,width=70,height=15,yscrollcommand=scroll.set) #width=80,height=10,  
        self.letrero1.config(yscrollcommand=scroll.set) #width=80,height=10,  
        scroll.config(command=self.letrero1.yview)
        self.letrero1.grid(row=17, column=0)
           
           
    def procesar_habilitacion_OC(self):
        global db1
        
        #Envio de email
        #envio_email(filename)

        ########################### Impresion de la OC #############################################
        Link=str(self.Link.get()).replace('\\','\\\\')        
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

        #########################################################################

        driver = webdriver.Firefox(fp)
        #driver = webdriver.Firefox()
        driver.get("http://www.google.com/")
        
        #######################################################################
        #Apertura del PeopleSoft
        
        #open tab
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't') 
        
        # Load a page 
        driver.get('https://leifs.mycmsc.com/psp/leifsprd/EMPLOYEE/ERP/?cmd=logout')
        
       
        #username.send_keys("311800185")
        #password.send_keys("Danola11.")
        
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


        #db1 = pd.read_excel("Excel_de_envio.xlsx")
        A=db1['UN']
        B=db1['OC']
        C=db1['SC']
        D=db1['Contacto']
        E=db1['Despachar en']
        F=db1['Fecha de Entrega']
        G=db1['Ubicación OC']
        H=db1['Mail Solicitante']
        I=db1['Mail Proveedor']
        J=db1['Mail implant']
        K=db1['Causal']
        L=db1['Otros']
        self.Observaciones_habilitacion=[]
        self.Observaciones_OC_no_existe=[]
        self.Observaciones_impresion=[]
        self.Observaciones_impresion2=[]
        self.Observaciones_bajada=[]
        self.Observaciones_existencia=[]
        self.Observaciones_envio=[]
        self.Observaciones_error=[]
        self.BU=[]
        self.OC=[]
        ########################### Habilitacion de la OC #############################################
        try:
            for n in range (len(db1['UN'])):
                start1 = time.time()
                self.a=str(A[n]) #BU
                self.b=str("0")*(10-len(str(B[n])))+str(B[n]) #OC
                        
                #Ingresando unidad de negocio y OC
                time.sleep(2)
                driver.switch_to.default_content()
                time.sleep(0.5)
                driver.switch_to.frame("ptifrmtgtframe")
                time.sleep(0.5)
                #Colocar BU
                #WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PO_SRCH_BUSINESS_UNIT")))
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
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICSearch"))).click()
                time.sleep(2)
                
                ########################### Errores #############################################
                #Cuando la OC no existe
                Titulo=(By.CLASS_NAME, "PSSRCHINSTRUCTIONS")
                try:
                    WebDriverWait(driver, 1).until(EC.text_to_be_present_in_element((Titulo),"No matching values were found."))
                    self.mensaje_OC_no_existe="OC no existe" + ", BU " + self.a  + ", OC " + self.b
                    driver.find_element_by_id("#ICClear").click()
                    self.Observaciones_OC_no_existe.append(self.mensaje_OC_no_existe)
                    self.BU.append(self.a)
                    self.OC.append(self.b)
                    print(self.mensaje_OC_no_existe)
                    self.letrero1.insert(tk.END, self.mensaje_OC_no_existe+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()
                    continue
                except TimeoutException:
                    pass
                    
                ########################### Aseguramiento de Entrada ###########################
                #Cuando la OC no existe
                #By.ID PO_PNLS_PB_PAGE_TITLE_PO
                #By.CLASS_NAME PAPAGETITLE
                Titulo=(By.ID, "PO_PNLS_PB_PAGE_TITLE_PO")
                try:
                    WebDriverWait(driver, 1).until(EC.text_to_be_present_in_element((Titulo),"Purchase Order"))
                except TimeoutException:
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICSearch"))).click()
                    pass

                # Maintain Purchase Order
                ########################### Habilitacion de la OC #############################
                time.sleep(1)
                driver.switch_to.default_content()
                time.sleep(0.5)
                #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                driver.switch_to.frame("ptifrmtgtframe")
                #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                #Listado desplegable Dispatch method
                time.sleep(0.5)
                List_current_status=Select(driver.find_element_by_id("PO_HDR_DISP_METHOD"))
                time.sleep(0.5)
                List_current_status.select_by_visible_text('Print')
                time.sleep(0.5)
                List_future_status=Select(driver.find_element_by_id("PO_HDR_DISP_METHOD"))
                # time.sleep(0.5)
                #Bajar la pagina
                driver.execute_script("window.scrollTo(0, 600)")
                time.sleep(1)
                #Guardar
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "#ICSave"))).click()  
                # time.sleep(0.5)
                #Volver a la busqueda
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()      
                
                
                #Se trabaja el link declarado en la mascara para bajar las OC
                #https://stackoverflow.com/questions/25146960/python-convert-back-slashes-to-forward-slashes/25147093
                ruta1=self.Link.get().replace(os.sep, '/')
                ruta=str("")+str(ruta1)+str("")
                #print(str(self.Link.get()))
                #print(ruta)

                
                #Asegurando dos veces la salida con Purchase Order
                Titulo=(By.CLASS_NAME, "PSSRCHTITLE")
                try: 
                    WebDriverWait(driver, 0.5).until(EC.text_to_be_present_in_element((Titulo),"Purchase Order"))
                except TimeoutException:
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                    time.sleep(1)
                
                end1 = time.time()
                
                Titulo=(By.CLASS_NAME, "PSSRCHTITLE")
                try: 
                    WebDriverWait(driver, 0.5).until(EC.text_to_be_present_in_element((Titulo),"Purchase Order"))
                except TimeoutException:
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "#ICList"))).click()
                        
                self.mensaje_habilitacion2="Proceso OK"
                self.mensaje_habilitacion1="Termine de habilitar la OC " + self.a + self.b +" , me demore " + str(int(end1-start1)) + " seg."
                print(self.mensaje_habilitacion1)
                self.Observaciones_habilitacion.append(self.mensaje_habilitacion1)
                self.BU.append(self.a)
                self.OC.append(self.b)
                self.letrero1.insert(tk.END, self.mensaje_habilitacion1+"\n")
                self.letrero1.config(fg = 'green',height=12)
                self.letrero1.update()
                
                
                self.letrero1.insert(tk.END, "Se bajaran en: " + self.Link.get() + "\n")
                self.letrero1.config(fg = 'green',height=12)
                self.letrero1.update()
                
                
                #Comienzo del ciclo        
                driver.implicitly_wait(1)
                driver.switch_to.default_content()
                
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

############################## Impresion OC ###################################
                try:
                    start2 = time.time()
                    
                    #Solamente realizara la impresion de las OC que esten con el proceso OK
                    if self.mensaje_habilitacion2.find("Proceso OK") != -1:
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
                    else:#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(1) > a:nth-child(1)
                        continue
                    try:
                        if driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(1) > a:nth-child(1)").text == self.a:
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
    
                        elif driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(16) > td:nth-child(1) > a:nth-child(1)").text == self.a:
                            driver.find_element_by_css_selector("#PTSRCHRESULTS > tbody:nth-child(1) > tr:nth-child(16) > td:nth-child(1) > a:nth-child(1)").click()
                            
                        else:
                            self.mensaje_impresion1= "No encontre la unidad de negocio " + self.a
                            self.letrero1.insert(tk.END, self.mensaje_impresion1+"\n")
                            self.letrero1.config(fg = 'red',height=12)
                            self.letrero1.update()
                            self.BU.append(self.a)
                            self.OC.append(self.b)
                            self.Observaciones_impresion.append(self.mensaje_impresion1)
                            continue
                        
                        #En caso de no encontrar la unidad de negocio
                    except NoSuchElementException:
                        self.mensaje_impresion2= "Me quede pegado en Psoft, OC " + self.a
                        print(self.mensaje)
                        self.letrero1.insert(tk.END, self.mensaje+"\n")
                        self.letrero1.config(fg = 'red',height=12)
                        self.letrero1.update()
                        self.BU.append(self.a)
                        self.OC.append(self.b)
                        self.Observaciones_impresion2.append(self.mensaje_impresion2)
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
                    
                    #si corresponde a una unidad de negocio CHL
                    if self.a.find("CHL") != -1: 
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
                            
                            #Si corresponde a una unidad de negocio PER
                    elif self.a.find("PER") != -1:
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
                                self.mensaje_bajada1="Fue bajada la OC " + self.a + self.b + ", me demore " + str(int(end2-start2)) + " seg."
                                #os.rename(fileName,self.a+self.b+".PDF")
                                os.rename(os.path.join(ruta,fileName),os.path.join(ruta,self.a+self.b+".PDF"))
                                print(self.mensaje_bajada1)
                                self.Observaciones_bajada.append(self.mensaje_bajada1)
                                self.BU.append(self.a)
                                self.OC.append(self.b)
                                self.letrero1.insert(tk.END, self.mensaje_bajada1+"\n")
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
                                self.mensaje_bajada1="No baje la OC " + self.b + ", BU " + self.a
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
                            self.mensaje_bajada1="OC ya bajada, con nombre " + self.a + self.b
                            print(self.mensaje_bajada1)
                            self.Observaciones_bajada.append(self.mensaje_bajada1)
                            self.BU.append(self.a)
                            self.OC.append(self.b)
                            self.letrero1.insert(tk.END, self.mensaje_bajada1+"\n")
                            self.letrero1.config(fg = 'blue',height=12)
                            self.letrero1.update()
                            continue
                    
                    if self.mensaje_bajada1.find("OC ya bajada, con nombre") != -1 or self.mensaje_bajada1.find("No baje la OC") != -1:  #mensaje=="No baje el archivo":
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
                        
                        #Comienzo del ciclo
                        #Main Menu
                        #print("vamos a comenzar a salir 2")
                        time.sleep(1)
                        driver.switch_to.default_content()
                        #time.sleep(1)
                        #driver.switch_to.frame("ptifrmtgtframe")

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
    
                        #Find an Existing Value
                        time.sleep(1)
                        driver.switch_to.default_content()
                        time.sleep(1)
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        driver.switch_to.frame("ptifrmtgtframe")
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.PSHYPERLINK:nth-child(1)"))).click()
                        #break
                    else:    
                        #continue
                        #print("vamos a comenzar a salir")
                        time.sleep(1)
                        driver.switch_to.default_content()
                        #time.sleep(1)
                        #driver.switch_to.frame("ptifrmtgtframe")
                    
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
    
                        #Find an Existing Value
                        time.sleep(1)
                        driver.switch_to.default_content()
                        time.sleep(1)
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        driver.switch_to.frame("ptifrmtgtframe")
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.PSHYPERLINK:nth-child(1)"))).click() 
                        
                        # self.Observaciones_bajada.append(self.mensaje_bajada1)
                        # self.letrero1.insert(tk.END, self.mensaje_bajada1+"\n")
                        # self.letrero1.config(fg = 'blue',height=12)
                        # self.letrero1.update()

                except NoSuchElementException:
                    self.mensaje_bajada1="No pude imprimir la OC " + self.a + self.b
                    self.Observaciones_bajada.append(self.mensaje_bajada1)
                    self.letrero1.insert(tk.END, self.mensaje_bajada1+"\n")
                    self.letrero1.config(fg = 'red',height=12)
                    self.letrero1.update()



##########################Envio OC via email###################################

                self.c=str(C[n]) #SC
                self.d=str(D[n]) #Contacto
                e=str(E[n]) #Despachar en
                f=str(F[n]) #Fecha de Entrega
                self.g=str(G[n]) #Ubicación OC
                h=str(H[n]) #Mail Solicitante
                self.i=str(I[n]) #Mail Proveedor
                self.j=str(J[n]) #Mail implant
                self.k=str(K[n]) #Causal
                self.l=str(L[n]) #Otros

###############################################################################
                #Verificacion de existencia del archivo
                #https://stackoverflow.com/questions/3964681/find-all-files-in-a-directory-with-extension-txt-in-python
                #https://stackoverflow.com/questions/1724693/find-a-file-in-python
            
                archivo=self.a+self.b+".pdf"
                archivo2=self.a+self.b+".PDF"
                try:
                    for file in os.listdir(self.g):
                        if file.endswith(archivo) or file.endswith(archivo2):
                            #Si encuentra el archivo sale del proceso
                            #print(os.path.join(path, file)) 
                            #print(bool(os.path.join(g,archivo)))
                            #print("existe")
                            self.mensaje_existencia="Existe"
                            break
                        else:
                            self.mensaje_existencia="No existe el archivo OC " + self.a+self.b
                    #print(mensaje)
                    if self.mensaje_existencia.find("No existe el archivo") != -1:
                        raise TypeError("No encontre archivo")
                except:
                    self.Observaciones_envio.append(self.mensaje_existencia)
                    self.letrero1.insert(tk.END, self.mensaje_existencia+"\n")
                    self.letrero1.update()
                    continue
            
###############################################################################
                #Redaccion de email
                const=win32com.client.constants
                olMailItem = 0x0
                #obj = win32com.client.Dispatch("Outlook.Application")
                obj = win32com.client.Dispatch("Outlook.Application")
                newMail = obj.CreateItem(olMailItem)
                
                if self.a=='PER03':
                    newMail.Subject = "Envio OC "+self.b+" UPN "+"SC "+self.c
                elif self.a=='PER05':
                    newMail.Subject = "Envio OC "+self.b+" UPC "+"SC "+self.c
                elif self.a=='PER07':
                    newMail.Subject = "Envio OC "+self.b+" CIBERTEC "+"SC "+self.c
                elif self.a=='CHL01' or self.a=='CHL05' or self.a=='CHL08'or self.a=='CHL06':
                    newMail.Subject = "Envio OC "+self.b+" UNAB "+"SC "+self.c
                elif self.a=='CHL18' or self.a=='CHL25' or self.a=='CHL28'or self.a=='CHL31':
                    newMail.Subject = "Envio OC "+self.b+" ARO "+"SC "+self.c
                elif self.a=='CHL02':
                    newMail.Subject = "Envio OC "+self.b+" UDLA "+"SC "+self.c
                elif self.a=='CHL04':
                    newMail.Subject = "Envio OC "+self.b+" AIEP "+"SC "+self.c
                elif self.a=='CHL32':
                    newMail.Subject = "Envio OC "+self.b+" UVM "+"SC "+self.c
                
                #Cuerpo del texto segun institucion
                # newMail.Body = "I AM\nTHE BODY MESSAGE!"
                newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
                #newMail.HTMLBody = "<HTML><BODY>Estimado Marco <span style='color:red'><br>Procesaremos la finalizacion de OC</span> <br>text here.</BODY></HTML>"
                if self.a=='PER03' and self.k=='Regularizacion Correo':
                    newMail.HTMLBody = Regularizacion_Correo_UPN(h,e,f) #Peru_Regularizacion_Correo, segun causal   
                elif  self.a=='PER03' and self.k=='Corrección OC':
                    newMail.HTMLBody = Correccion_OC_UPN(h,e,f) #Peru_Corrección_OC, segun causal   
                elif  self.a=='PER03' and self.k=='Importe valor 1':    
                    newMail.HTMLBody = Importe_valor_1_UPN(h,e,f) #Peru_Importe_valor_1, segun causal
                elif self.a=='PER05' and self.k=='Regularizacion Correo':
                    newMail.HTMLBody = Regularizacion_Correo_UPC(h,e,f) #Peru_Regularizacion_Correo, segun causal   
                elif  self.a=='PER05' and self.k=='Corrección OC':
                    newMail.HTMLBody = Correccion_OC_UPC(h,e,f) #Peru_Corrección_OC, segun causal   
                elif  self.a=='PER05' and self.k=='Importe valor 1':    
                    newMail.HTMLBody = Importe_valor_1_UPC(h,e,f) #Peru_Importe_valor_1, segun causal        
                elif self.a=='PER07' and self.k=='Regularizacion Correo':
                    newMail.HTMLBody = Regularizacion_Correo_PER07(h,e,f) #Peru_Regularizacion_Correo, segun causal   
                elif  self.a=='PER07' and self.k=='Corrección OC':
                    newMail.HTMLBody = Correccion_OC_PER07(h,e,f) #Peru_Corrección_OC, segun causal   
                elif  self.a=='PER07' and self.k=='Importe valor 1':    
                    newMail.HTMLBody = Importe_valor_1_PER07(h,e,f) #Peru_Importe_valor_1, segun causal                
                elif  self.a=='CHL04' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='CHL08' or self.a=='CHL05'or self.a=='CHL06'and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='CHL18' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='CHL25' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                            
                elif  self.a=='CHL28' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal 
                elif  self.a=='CHL31' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                                
                elif  self.a=='CHL32' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='CHL02' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='CHL01' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_CHL(h,e,f) #, segun causal                
                elif  self.a=='PER07' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_PER07(h,e,f) #, segun causal                
                elif  self.a=='PER05' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_PER05(h,e,f) #, segun causal                
                elif  self.a=='PER03' and self.k=='Normal':    
                    newMail.HTMLBody = Normal_PER03(h,e,f) #, segun causal                
                elif  self.k=='Regularizar envio de OC':
                    newMail.HTMLBody = Regularizar_envio_orden_compra(h,e,f) #Regularización_Correo, segun causal
                elif  self.k=='Regularización Correo':
                    newMail.HTMLBody = Regularización_Correo(h,e,f) #Regularización_Correo, segun causal
                elif  self.k=='Corrección OC':
                    newMail.HTMLBody = Corrección_OC(h,e,f) #Corrección_OC, segun causal
                elif  self.k=='Importe valor 1':
                    newMail.HTMLBody = Importe_valor_1(h,e,f) #Importe_valor_1, segun causal
                elif  self.k=='Regularización Correo COV':
                    newMail.HTMLBody = Regularización_Correo_COV(h,e,f) #Regularización_Correo_COV, segun causal
                elif  self.k=='Corrección OC COV':
                    newMail.HTMLBody = Corrección_OC_COV(h,e,f) #Corrección_OC_COV, segun causal
                elif  self.k=='Importe valor 1 COV':
                    newMail.HTMLBody = Importe_valor_1_COV(h,e,f) #Importe_valor_1_COV, segun causal
                else :
                    continue
                
                #Aqui debemos sumar la columna otros mas los liaison
                """
                H=db1['Mail Solicitante']
                I=db1['Mail Proveedor']
                J=db1['Mail implant']
                L=db1['Otros']
                """
                #Email de Liaison
                if self.a=='CHL01' or self.a=='CHL05' or self.a=='CHL08'or self.a=='CHL06':
                    liaison="arturo.gonzalez@serviciosandinos.net"
                elif  self.a=='CHL02':
                    liaison="cinthia.gonzalez@serviciosandinos.net"
                elif  self.a=='CHL04':
                    liaison="berenise.balbontin@serviciosandinos.net"
                elif self.a=='PER03':
                    liaison="clara.rivera@upc.pe;blanky.gayoso@upc.pe"
                elif self.a=='PER05':
                    liaison="eiker.briceno@upc.pe;karla.barbie@upc.pe;blanky.gayoso@upc.pe"
                elif self.a=='PER07':
                    liaison="clara.rivera@upc.pe;blanky.gayoso@upc.pe"
                else:
                    liaison='nan'
                    
                #Email
                email = [self.i,self.j,h,self.l,liaison]
                
                #https://stackoverflow.com/questions/21011777/how-can-i-remove-nan-from-list-python-numpy
                #https://stackoverflow.com/questions/45695373/removing-a-nan-from-a-list?rq=1
                
                cleanedList = [x for x in email if str(x) != 'nan'] 
                
                #https://www.geeksforgeeks.org/join-function-python/
                #https://www.programiz.com/python-programming/methods/string/join
                
                direccion_email=';'.join(cleanedList)
                newMail.CC = direccion_email
                
                    
                newMail.To = self.d
                    
                #attachment1 = g+a+b #link segun hoja excel
                #attachment1 = r"C:\Users\mvidal2\Desktop\data scientist\Despacho automatico\CHL010000000002.xlsx"
                #attachment1 = "C:\\Users\\mvidal2\\Desktop\\data scientist\\Despacho automatico\\CHL010000000002.xlsx"
                              
                
                #Para confirmar la existencia del archivo
                #https://stackoverflow.com/questions/82831/how-do-i-check-whether-a-file-exists-without-exceptions
                
                attachment1 = self.g+self.a+self.b+".pdf"
                attachment2 = self.g+self.a+self.b+".PDF"
                if os.path.exists(attachment1):
                    newMail.Attachments.Add(Source=attachment1)
                    print(attachment1)
                elif  os.path.exists(attachment2):
                    newMail.Attachments.Add(Source=attachment2)
                    print(attachment2)
                    
                """
                https://stackoverflow.com/questions/36400683/adding-attachment-to-email-through-outlook-python
                https://stackoverflow.com/questions/15494911/python-win32com-outlook-attach-file-with-insert-as-text-method
                http://timgolden.me.uk/python/win32_how_do_i/replace-outlook-attachments-with-links.html
                """
                #attachment1 = r+\"+C:\Users\mvidal2\Desktop\data scientist\Despacho automatico\+CHL010000000002+.xlsx+\"
                #newMail.Attachments.Add(Source=attachment1)
                newMail.Display()
                #newMail.send()
                newMail.Send()   
                #print('Termine un email'+ ". Me demore " + str(int(end1-start1)))
                self.mensaje_enviado= "Enviada, OC " + self.a + self.b
                self.letrero1.insert(tk.END, self.mensaje_enviado+"\n")
                self.letrero1.update()
                self.Observaciones_envio.append(self.mensaje_enviado)
            
            db1.reset_index(drop=True)
            df1 = pd.DataFrame({'BU':self.BU})
            df2 = pd.DataFrame({'OC':self.OC})
            df3 = pd.DataFrame({'Observaciones_OC_no_existe':self.Observaciones_OC_no_existe})
            df4 = pd.DataFrame({'Observaciones_habilitacion':self.Observaciones_habilitacion})
            #df5 = pd.DataFrame({'Observaciones_impresion1':self.Observaciones_impresion1})
            df5 = pd.DataFrame({'Observaciones_impresion2':self.Observaciones_impresion2})
            df6 = pd.DataFrame({'Observaciones_bajada':self.Observaciones_bajada})
            df7 = pd.DataFrame({'Observaciones_ existencia':self.Observaciones_existencia })
            df8 = pd.DataFrame({'Observaciones_envio':self.Observaciones_envio})
            
            #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
            db1=pd.concat([df1,df2,df3,df4,df5,df6,df7,df8], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
            db1.columns=['BU','OC','Observaciones_OC_no_existe','Observaciones_hablitacion','Observaciones_impresion2','Observaciones_bajada','Observaciones_ existencia','Observaciones_envio']
            db1.to_excel("Resumen_del_proceso_habilitacion.xlsx",index = False)

            self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx\n")
            self.letrero1.update()
            self.letrero1.config(fg = 'blue',height=12)
    
        except TimeoutException:
                self.mensaje_error= "me quede pegado en Psoft, OC " + self.a+self.b
                self.Observaciones_error.append(self.mensaje_error)
                self.BU.append(self.a)
                self.OC.append(self.b)
                print(self.mensaje_error)
                self.letrero1.insert(tk.END, self.mensaje_error+"\n")
                self.letrero1.config(fg = 'red',height=12)
                self.letrero1.update()

                df1 = pd.DataFrame({'BU':self.BU})
                df2 = pd.DataFrame({'OC':self.OC})
                df3 = pd.DataFrame({'Observaciones_OC_no_existe':self.Observaciones_OC_no_existe})
                df4 = pd.DataFrame({'Observaciones_hablitacion':self.Observaciones_hablitacion})
                df5 = pd.DataFrame({'Observaciones_impresion1':self.Observaciones_impresion1})
                df6 = pd.DataFrame({'Observaciones_impresion2':self.Observaciones_impresion2})
                df7 = pd.DataFrame({'Observaciones_bajada':self.Observaciones_bajada})
                df8 = pd.DataFrame({'Observaciones_ existencia':self.Observaciones_existencia })
                df9 = pd.DataFrame({'Observaciones_envio':self.Observaciones_envio})

                   
                #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
                db2=pd.concat([df1,df2,df3,df4,df5,df6,df7,df8,df9], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
                db2.columns=['BU','OC','Observaciones_OC_no_existe','Observaciones_hablitacion','Observaciones_impresion1','Observaciones_impresion2','Observaciones_bajada','Observaciones_ existencia','Observaciones_envio']
                db2.to_excel("Resumen_del_proceso_habilitacion.xlsx",index = False)


        # except NoSuchElementException:
        #         self.mensaje_error= "me quede pegado en Psoft, OC " + self.a+self.b
        #         self.Observaciones_error.append(self.mensaje_error)
        #         self.BU.append(self.a)
        #         self.OC.append(self.b)
        #         print(self.mensaje_error)
        #         self.letrero1.insert(tk.END, self.mensaje_error+"\n")
        #         self.letrero1.config(fg = 'red',height=12)
        #         self.letrero1.update()

                df1 = pd.DataFrame({'BU':self.BU})
                df2 = pd.DataFrame({'OC':self.OC})
                df3 = pd.DataFrame({'Observaciones_OC_no_existe':self.Observaciones_OC_no_existe})
                df4 = pd.DataFrame({'Observaciones_hablitacion':self.Observaciones_hablitacion})
                df5 = pd.DataFrame({'Observaciones_impresion1':self.Observaciones_impresion1})
                df6 = pd.DataFrame({'Observaciones_impresion2':self.Observaciones_impresion2})
                df7 = pd.DataFrame({'Observaciones_bajada':self.Observaciones_bajada})
                df8 = pd.DataFrame({'Observaciones_ existencia':self.Observaciones_existencia })
                df9 = pd.DataFrame({'Observaciones_envio':self.Observaciones_envio})

                   
                #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
                db2=pd.concat([df1,df2,df3,df4,df5,df6,df7,df8,df9], ignore_index=True, names=['BU','OC','Observaciones'],axis=1).fillna("-")
                db2.columns=['BU','OC','Observaciones_OC_no_existe','Observaciones_hablitacion','Observaciones_impresion1','Observaciones_impresion2','Observaciones_bajada','Observaciones_ existencia','Observaciones_envio']
                db2.to_excel("Resumen_del_proceso_habilitacion.xlsx",index = False)
                

    def mfileopen(self):
        global filename
        global db1
        global Y
        filename = filedialog.askopenfilename()
        db1 = pd.read_excel(filename)
        self.letrero1.insert(tk.END, "Archivo cargado!\n")
        self.letrero1.config(fg = 'green',height=12)
        #print(self.mensaje)
        
    def destroy(self):
        self.quit()
        
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Despacho automatico de OC")
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.resizable(0,0)
    #root.iconbitmap("Let's put smart to work.ico")
    app=MainApplication(root)
    root.mainloop()