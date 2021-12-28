# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 10:49:48 2020
@author: mvidal2
Version 1 
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import win32com.client
from win32com.client import Dispatch, constants
import socket

"""
Ventana de seguridad
https://stackoverflow.com/questions/16115378/tkinter-example-code-for-multiple-windows-why-wont-buttons-load-correctly
https://realpython.com/python-gui-tkinter/

"""
class MainApplication:
    
    def __init__(self, master):
        self.master = master
        self.master.geometry("800x800")
        self.frame = tk.Frame(self.master)
        
        # here is the application variable, variable contents
        self.X = tk.StringVar()
        self.db1=tk.StringVar()
        self.Y=tk.StringVar()
        self.var=tk.StringVar() 
        self.db3=tk.Listbox()
        self.entry_var=tk.StringVar()
        self.mensaje=tk.StringVar()
        password=tk.StringVar()
        usuario=tk.StringVar()
        
        #Colocando Gadget
        self.button2 = tk.Button(self.frame,text='Open File',font=('arial',10,'bold'), command=self.mfileopen)
        self.button2.place(x=10,y=10)
        #self.button2.pack(side='left')
        self.frame.pack()
        
        self.button3 = tk.Button(self.frame,text='Salir',command= self.destroy,font=('arial',10,'bold')) #Fg = letter, bg = Fondo
        self.button3.place(x=90,y=10)
        #self.button3.pack(side='bottom')
        self.frame.pack()

        self.button4 = tk.Button(self.frame,text='Procesar',font=('arial',10,'bold'), command=self.procesar)
        self.button4.place(x=140,y=10)
        #self.button4.pack(padx=10, pady=125)
        self.frame.pack()
        
        #https://recursospython.com/guias-y-manuales/caja-de-texto-entry-tkinter/
        #https://recursospython.com/guias-y-manuales/posicionar-elementos-en-tkinter/
        
        """
        Entry es solamente para una linea de texto
        https://effbot.org/tkinterbook/entry.htm
        self.letrero1=tk.Entry(self.frame,textvariable=self.var, fg='red', bd=3,width = 90, height = 1)        
        """
        self.letrero_usuario=tk.Entry(self.frame,fg='blue', width = 20)
        self.letrero_usuario.place(x=100,y=50)
        self.frame.pack()


        self.letrero_password=tk.Entry(self.frame,fg='blue', width = 20, show="*") 
        self.letrero_password.place(x=100,y=90)
        self.frame.pack()

        """
        Label etiqueta
        http://effbot.org/tkinterbook/label.htm        
        """        
        
        self.observacion1 = tk.Label(self.frame, text="Instrucciones de uso",font=('arial',10,'bold'))
        self.observacion1.place(x=10,y=120)
        self.frame.pack()
        
        self.observacion1 = tk.Label(self.frame, text="Debe cargar plantilla excel ID set, ID articulo, Descripcion, Precio, Moneda, Vigencia, Cantidad minima, Dias de plazo.",font=('arial',10))
        self.observacion1.place(x=10,y=140)
        self.frame.pack()
              
        self.observacion1 = tk.Label(self.frame, text="Creado para PeopleSoft en Ingles.",font=('arial',10))
        self.observacion1.place(x=10,y=160)
        self.frame.pack()
        
        self.observacion1 = tk.Label(self.frame, text="Cada vez que presiona Procesar, se envia email al director solicitando autorizacion de uso.",font=('arial',10))
        self.observacion1.place(x=10,y=180)
        self.frame.pack()
        
        self.observacion1 = tk.Label(self.frame, text="En el email de autorizacion, se adjunta excel de carga y nombre del usuario que opera robot.",font=('arial',10))
        self.observacion1.place(x=10,y=200)
        self.frame.pack()
        
        self.observacion1 = tk.Label(self.frame, text="Se utiliza en PC " + socket.gethostname() + ". Numero IP " + socket.gethostbyname(socket.gethostname()),font=('arial',10))
        self.observacion1.place(x=10,y=220)
        self.frame.pack()

        self.observacion1 = tk.Label(self.frame, text="Su uso inapropiado sera sancionado",font=('arial',10),fg='red')
        self.observacion1.place(x=10,y=240)
        self.frame.pack()

       
        """
        Scroll
        https://stackoverflow.com/questions/19646752/python-scrollbar-on-text-widget/19647325
        """
        # Vertical (y) Scroll Bar
        scroll = tk.Scrollbar(self.frame)
        scroll.pack(side='right', fill='y')


        """
        Text permite mostrar texto con varios estilos y atributos, soporta imagenes y ventanas   
        https://www.python-course.eu/tkinter_text_widget.php
        """
        self.letrero1=tk.Text(self.frame,width=100,height=200,wrap='none', yscrollcommand=scroll.set)        
        #self.letrero1.insert(tk.END, self.var.get())
        scroll.config(command=self.letrero1.yview)
        """
        pack
        https://recursospython.com/guias-y-manuales/posicionar-elementos-en-tkinter/
        """
        #self.letrero1.place(x=200,y=40) #padx=100,pady=200
        self.letrero1.pack(expand=True,fill='both',padx=10, pady=300)
        #self.letrero1.pack(fill='both',expand=1)
        self.frame.pack()
        
        """
        Label
        http://effbot.org/tkinterbook/label.htm
        """
        self.etiqueta_password = tk.Label(self.frame, text="Usuario",font=('arial',10,'bold'))
        self.etiqueta_password.place(x=10,y=50)
        self.frame.pack()
        
        self.etiqueta_usuario = tk.Label(self.frame, text="Password",font=('arial',10,'bold'))
        self.etiqueta_usuario.place(x=10,y=90)
        self.frame.pack()
         
    def procesar(self):
        global driver
        """
        https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send
        https://gist.github.com/ITSecMedia/b45d21224c4ea16bf4a72e2a03f741af
        https://stackoverflow.com/questions/50926514/send-email-through-python-using-outlook-2016-without-opening-it
        """
        const=win32com.client.constants
        olMailItem = 0x0
        #obj = win32com.client.Dispatch("Outlook.Application")
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Autorizacion operacion de robot mantencion de ID articulos"
        # newMail.Body = "I AM\nTHE BODY MESSAGE!"
        newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
        #newMail.HTMLBody = "<HTML><BODY>Estimado Marco <span style='color:red'><br>Procesaremos la finalizacion de OC</span> <br>text here.</BODY></HTML>"
        #newMail.HTMLBody = "<HTML><BODY><p>Estimado Marco</p><p>Comenzara el proceso de mantencion de ID articulos</p><p>Adjunto listado de ID de arituclo a mantener</p><br>Favor enviar su aprobacion</BODY></HTML>" + "Se utilizara en PC " + socket.gethostname()
        newMail.HTMLBody = "<HTML><BODY><p>Estimado Marco</p><p>Comenzara el proceso robot de mantencion de ID articulos</p><p>El proceso se realizara en PC "+ socket.gethostname() + "</p></BODY></HTML>" +"<HTML><BODY><p>Adjunto listado de ID de articulos a mantener</p><br>Favor enviar su aprobacion</BODY></HTML>"
        newMail.To = "matias.vidal@serviciosandinos.net"
        #newMail.To = "marco.vera@serviciosandinos.net"
        #attachment1 = r"C:\Temp\example.pdf"
        #attachment1 = r"C:\Users\mvidal2\Desktop\data scientist\Finalizacion de OC\Finalizacion.xlsx"
        attachment1 = filename
        newMail.Attachments.Add(Source=attachment1)
        newMail.display()
        #newMail.send()
        newMail.send

        self.letrero1.insert(tk.END, "Email enviado a director\n")
        self.letrero1.config(fg = 'green',height=12)

        start1 = time.time()
        driver = webdriver.Firefox()
        driver.get("http://www.google.com/")
        
        #######################################################################
        #Apertura del PeopleSoft
        
        #open tab
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't') 
        
        # Load a page 
        driver.get('https://leifs.mycmsc.com/psp/leifsprd/EMPLOYEE/ERP/?cmd=logout')
        
        #para llevarlo a español
        driver.find_element_by_css_selector(".pslanguageframe > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(1)").click()
        time.sleep(2)
        
        username = driver.find_element_by_id("userid") #input id o name
        password = driver.find_element_by_id("pwd") #input id o name
        
        username.send_keys("311800185")
        password.send_keys("Danola10.")
        
        #username.send_keys(self.letrero_usuario.get())
        #password.send_keys(self.letrero_password.get())
        
        driver.find_element_by_name("Submit").click() #name
        
        #######################################################################
        end1 = time.time()
        #print(end1 - start1, " Tiempo de comando ingreso de password")
        start2 = time.time()
        #######################################################################
        
        #Comienzo del ciclo
        #Main Menu
        driver.implicitly_wait(1)
        driver.find_element_by_id("pthnavbca_PORTAL_ROOT_OBJECT").click() #ID
        
        #Items
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPCO_ITEMS").click() #ID
        
        #Define Items And Attributes
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPIN_DEFINE_ITEMS").click() #ID
        
        #Define Item
        driver.implicitly_wait(5)
        driver.find_element_by_css_selector("#crefli_EP_ITEM_DEFIN_GBL > a:nth-child(1)").click() #CSS Selector
        
        #Publicacion de informacion
        self.letrero1.insert(tk.END, "Ingreso a menu\n")
        self.letrero1.config(fg = 'green',height=12)
       
        #######################################################################
        end2 = time.time()
        #print(end2 - start2, " Tiempo de comando para ingresar a menu general")
        #######################################################################
        #Preparando el archivo y columna de observaciones
        #db1 = pd.read_excel('Input.xlsx')
        S=db1['Dias de plazo']
        T=db1['Cantidad minima']
        U=db1['Vigencia']
        V=db1['Moneda']
        W=db1['Precio']
        X=db1['Descripcion']
        Y=db1['ID articulo']
        Z=db1['ID Set']
        Observaciones_descripcion = []
        Observaciones_precio = []
        Observaciones_vigencia = []
        Observaciones_cantidad_minima = []
        Observaciones_dias_plazo = []
        
        for n in range (len(db1['ID articulo'])):
            start3 = time.time()
            z=str(Z[n]) #ID Set
            y=str(Y[n]) #ID Articulo
            x=str(X[n]) #Descripcion
            w=str(W[n]) #Precio
            v=str(V[n]) #Moneda
            u=str(U[n]) #Vigencia
            t=str(T[n]) #Cantidad minima
            s=str(S[n]) #Dias de plazo
            ###################################################################
            #Si ID Set no esta vacio, ejecutar ciclo
            if z!='nan':
                #Item Definition
                driver.switch_to.default_content()
                time.sleep(1)
                driver.switch_to.frame("ptifrmtgtframe")      
                time.sleep(1)
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#MST_ITM_INV_VW_SETID")))
                driver.implicitly_wait(1)
                driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").click()
                driver.implicitly_wait(1)
                driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").clear()
                driver.implicitly_wait(1)
                #Colocar unidad de negocio
                driver.find_element_by_id("MST_ITM_INV_VW_SETID").send_keys(z)
                driver.implicitly_wait(1)
            
                driver.implicitly_wait(1)
                driver.find_element_by_css_selector("#MST_ITM_INV_VW_INV_ITEM_ID").click()
                driver.implicitly_wait(1)
                driver.find_element_by_id("MST_ITM_INV_VW_INV_ITEM_ID").clear()
                #Colocar ID articulo                    
                driver.find_element_by_id("MST_ITM_INV_VW_INV_ITEM_ID").send_keys(y)
                driver.implicitly_wait(1)
                
                #Hacer Click
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"#ICSearch"))).click()                            
                ###############################################################################
                #https://www.geeksforgeeks.org/switch-case-in-python-replacement/
                end3 = time.time()
                
                
                if x!='nan': #cambio de descripcion
                    start4 = time.time()
                    #Cambio de descripcion
                    driver.switch_to.default_content()
                    #driver.switch_to.frame("TargetContent")  
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_TBL_DESCR60"))).clear()
                    #driver.find_element_by_id("REQ_RC_WB_FROM_REQ").clear()
                    driver.implicitly_wait(1)
                    driver.find_element_by_id("MASTER_ITEM_TBL_DESCR60").send_keys(x)
                    
                    #Purchasing Item Attributes
                    driver.implicitly_wait(1)       
                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(1)
                    driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    
                    
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR254_MIXED"))).click()
                    driver.implicitly_wait(1)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR254_MIXED"))).clear()
                    driver.implicitly_wait(1)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_DESCR254_MIXED").send_keys(x)
                    driver.implicitly_wait(2)
                        
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR"))).click()
                    driver.implicitly_wait(1)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR"))).clear()
                    driver.implicitly_wait(1)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_DESCR").send_keys(x)
                    driver.implicitly_wait(2)
                    
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCRSHORT"))).click()
                    driver.implicitly_wait(1)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCRSHORT"))).clear()
                    driver.implicitly_wait(1)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_DESCRSHORT").send_keys(x)
                    driver.implicitly_wait(2)
                    
                    #OK
                    driver.find_element_by_name("#ICSave").click()
                    time.sleep(1)
                    #Save
                    driver.execute_script("window.scrollTo(0, 600)")
                    driver.find_element_by_name("#ICSave").click()
                    time.sleep(1)
                    end4 = time.time()
                    mensaje="en cambiar la descripcion me demore " + str(int(end4-start4)) + " seg. Al articulo " + y
                    Observaciones_descripcion.append(mensaje)
                    print(mensaje)
                else:
                    mensaje="no cambie descripcion"
                    Observaciones_descripcion.append(mensaje)
                    print(mensaje)
                
                if w!='nan': #cambio de precio
                    start5=time.time()
                    
                    #Correct History
                    time.sleep(1)     
                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(1)
                    driver.find_element_by_id("#ICCorrection").click()                    
                    
                    #Purchasing Item Attributes
                    time.sleep(1)     
                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(1)
                    driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    #driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    
                    #Cambio de precio en Purchasing Attributes
                    #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    time.sleep(0.5)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_PRICE_LIST").click()
                    time.sleep(0.5)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRICE_LIST"))).clear()
                    time.sleep(0.5)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_PRICE_LIST").send_keys(w)
                    
                    #Cambio de moneda en Purchasing Attributes
                    time.sleep(0.5)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_CURRENCY_CD").click() #ID
                    time.sleep(0.5)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_CURRENCY_CD"))).clear()
                    time.sleep(0.5)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_CURRENCY_CD").send_keys(v)
                    
                    #Ventana "Item Vendor"
                    time.sleep(0.5)
                    driver.switch_to.default_content()
                    time.sleep(0.5)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(0.5)
                    driver.find_element_by_css_selector("#PSTAB > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1)").click()
                    
                    #Item Vendor UOM 
                    #time.sleep(1)
                    #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    time.sleep(1)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_WRK_VNDR_UOM_PB$0"))).click()
                    
                    # apretar "+"
                    #time.sleep(1)
                    #driver.execute_script("window.scrollTo(0, 600)")
                    #time.sleep(1)
                    #driver.switch_to.default_content()
                    #time.sleep(1)
                    #driver.switch_to.frame("ptifrmtgtframe")  
                    #time.sleep(1)
                    #driver.find_element_by_css_selector("#\$ICField39\$new\$0\$\$0 > img:nth-child(1)").click()
                    
                    #En pantalla "Vendor's UOM and Pricing Information"
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").click()
                    time.sleep(0.5)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").clear()
                    time.sleep(0.5)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").send_keys(w)
                    time.sleep(0.5)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").click()
                    time.sleep(0.5)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").clear()
                    time.sleep(0.5)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").send_keys(v)
                    
                    #OK
                    time.sleep(0.5)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()
                    driver.implicitly_wait(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()
                    driver.implicitly_wait(1)
                    #Save
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()           
                    end5=time.time()
                    mensaje="en cambiar la precios me demore " + str(int(end5-start5)) + " seg. Al Articulo " + y
                    Observaciones_precio.append(mensaje)
                    print(mensaje)
                else:
                    mensaje="no cambie precios"
                    Observaciones_precio.append(mensaje)
                    print(mensaje)
                
                if u!='nan': #cambio de vigencia
                    start6=time.time()
                    driver.switch_to.default_content()
                    #driver.switch_to.frame("TargetContent")
                    time.sleep(0.5)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    time.sleep(0.5)
                    List_current_status=Select(driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STATUS_CURRENT"))
                    time.sleep(0.5)
                    List_current_status.select_by_visible_text('Active')
                    time.sleep(0.5)
                    List_future_status=Select(driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STATUS_FUTURE"))
                    time.sleep(0.5)
                    List_future_status.select_by_visible_text('Inactive')
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_TBL_ITM_STAT_DT_FUTURE"))).clear()
                    time.sleep(0.5)
                    driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STAT_DT_FUTURE").send_keys(u)

                    #Ok
                    time.sleep(1)
                    driver.find_element_by_name("#ICSave").click()
                    end6=time.time()
                    mensaje="en cambiar la fecha vigencia me demore " + str(int(end6-start6)) + " seg. Al Articulo " + y
                    Observaciones_vigencia.append(mensaje)
                    print(mensaje)
                else:
                    mensaje="no cambie vigencia"
                    Observaciones_vigencia.append(mensaje)
                    print(mensaje)
                
                if t!='nan': #cambio de cantidad minima
                    start7=time.time()
                    
                    #Correct History
                    time.sleep(1)     
                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(1)
                    driver.find_element_by_id("#ICCorrection").click()
                    
                    #Purchasing Item Attributes
                    time.sleep(0.5)     
                    driver.switch_to.default_content()
                    time.sleep(0.5)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(0.5)
                    driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    
                    #Item vendor
                    time.sleep(0.5)     
                    driver.switch_to.default_content()
                    time.sleep(0.5)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(0.5)
                    driver.find_element_by_css_selector("#PSTAB > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1)").click()
                        
                    #Item Vendor UOM 
                    time.sleep(1)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_WRK_VNDR_UOM_PB$0"))).click()
                    
                    #Estos 2 se cancelan debido a que se activa el Correct History
                    # apretar "+"
                    #time.sleep(1)
                    #driver.switch_to.default_content()
                    #time.sleep(1)
                    #driver.switch_to.frame("ptifrmtgtframe")  
                    #time.sleep(1)
                    #driver.find_element_by_css_selector("#\$ICField9\$new\$0\$\$0 > img:nth-child(1)").click()
                    
                    #UOM
                    #time.sleep(1)
                    #driver.find_element_by_id("ITM_VNDR_UOM_UNIT_OF_MEASURE$0").click()
                    #time.sleep(1)
                    #driver.find_element_by_id("ITM_VNDR_UOM_UNIT_OF_MEASURE$0").send_keys('UN')
                    
                    #Minimum Quantity superior
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_QTY_MIN$0").click()
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_QTY_MIN$0").clear()
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_QTY_MIN$0").send_keys(t)

                    # apretar "+"
                    #time.sleep(1)
                    #driver.execute_script("window.scrollTo(0, 600)")
                    #time.sleep(1)
                    #driver.switch_to.default_content()
                    #time.sleep(1)
                    #driver.switch_to.frame("ptifrmtgtframe")  
                    #time.sleep(1)
                    #driver.find_element_by_css_selector("#\$ICField39\$new\$0\$\$0 > img:nth-child(1)").click()
                    
                    
                    #Minimum Quantity inferior
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_QTY_MIN$0").click()
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_QTY_MIN$0").clear()
                    time.sleep(1)
                    driver.find_element_by_id("ITM_VNDR_UOM_PR_QTY_MIN$0").send_keys(t)
                
                    #OK
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()
                    driver.implicitly_wait(1)
                    #Save
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()           
                    time.sleep(0.5)
                    end7=time.time()
                    mensaje="en cambiar las cantidades minimas me demore " + str(int(end7-start7)) + " seg. Al Articulo " + y
                    Observaciones_cantidad_minima.append(mensaje)
                    print(mensaje)
                else:
                    mensaje="no cambie cantidad minima"
                    Observaciones_cantidad_minima.append(mensaje)
                    print(mensaje)
                
                if s!='nan': #cambio dias de plazo
                    start8=time.time()
                    #Purchasing Item Attributes
                    time.sleep(1)       
                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.frame("ptifrmtgtframe")  
                    time.sleep(1)
                    driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    #driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                    
                    #Cambio de dias de plazo
                    time.sleep(1)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_STD_LEAD").click() #ID
                    time.sleep(1)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_STD_LEAD"))).clear()
                    time.sleep(1)
                    driver.find_element_by_id("PURCH_ITEM_ATTR_STD_LEAD").send_keys(s)
                    
                    #Ok
                    time.sleep(0.5)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(0.5)
                    driver.find_element_by_name("#ICSave").click()
                    end8=time.time()
                    mensaje="en cambiar los dias de plazo me demore " + str(int(end8-start8)) + " seg. Al Articulo " + y
                    Observaciones_dias_plazo.append(mensaje)
                    print(mensaje)
                else:
                    mensaje="no cambie dias de plazos"
                    Observaciones_dias_plazo.append(mensaje)
                    print(mensaje)
                    
                #Return to search
                driver.execute_script("window.scrollTo(0, 600)")
                time.sleep(1)
                driver.find_element_by_name("#ICList").click()              
        
        db1.reset_index(drop=True)       
        db1['Observaciones_descripcion']=Observaciones_descripcion
        db1['Observaciones_precio']=Observaciones_precio
        db1['Observaciones_vigencia']=Observaciones_vigencia
        db1['Observaciones_cantidad_minima']=Observaciones_cantidad_minima
        db1['Observaciones_dias_plazo']=Observaciones_dias_plazo
        
        self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx \n")
        self.letrero1.config(fg = 'blue',height=12)
        db1.to_excel("Resumen_del_proceso.xlsx",index = False)
        
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
        print(self.mensaje)
        
    def destroy(self):
        self.master.destroy()
        
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Roboter mantencion de ID de articulos de catalogo y contrato")
    app=MainApplication(root)
    #root.geometry("800x500")
    root.mainloop()