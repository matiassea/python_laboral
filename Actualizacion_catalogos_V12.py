# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 10:49:48 2020
@author: mvidal2
Version 6
se agrega el Try para el tratamiento de errores
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
import winsound


"""
Ventana de seguridad
https://stackoverflow.com/questions/16115378/tkinter-example-code-for-multiple-windows-why-wont-buttons-load-correctly
https://realpython.com/python-gui-tkinter/

"""
class MainApplication(tk.Frame):
        
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        #self.master.geometry("830x800")
        
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
        
        ############################# Frame #############################
        self.frame = tk.Frame(self)
        self.frame.grid(row=0, column=0)
        
        #Colocando Gadget
        self.button2 = tk.Button(self.frame,text='Open File',font=('arial',10,'bold'), command=self.mfileopen)
        self.button2.grid(row=0, column=0,padx=10,pady=5)
        
        self.button4 = tk.Button(self.frame,text='Procesar',font=('arial',10,'bold'), command=self.procesar)
        self.button4.grid(row=0, column=1,padx=10,pady=5)
        
        #https://recursospython.com/guias-y-manuales/caja-de-texto-entry-tkinter/
        #https://recursospython.com/guias-y-manuales/posicionar-elementos-en-tkinter/
        
        """
        Entry es solamente para una linea de texto
        https://effbot.org/tkinterbook/entry.htm
        self.letrero1=tk.Entry(self.frame,textvariable=self.var, fg='red', bd=3,width = 90, height = 1)
        """
        ############################# Frame 2 #############################
        self.frame2 = tk.Frame(self)
        self.frame2.grid(row=1, column=0)
        
        self.letrero_usuario=tk.Entry(self.frame2,fg='blue', width = 20)
        self.letrero_usuario.grid(row=1, column=2,padx=10,pady=5)

        self.letrero_password=tk.Entry(self.frame2,fg='blue', width = 20, show="*") 
        self.letrero_password.grid(row=2, column=2,padx=10,pady=5)
        
        self.etiqueta_password = tk.Label(self.frame2, text="Usuario",font=('arial',10,'bold'))
        self.etiqueta_password.grid(row=1, column=1,padx=10,pady=5)
        
        self.etiqueta_usuario = tk.Label(self.frame2, text="Password",font=('arial',10,'bold'))
        self.etiqueta_usuario.grid(row=2, column=1,padx=10,pady=5)

        """
        Label etiqueta
        http://effbot.org/tkinterbook/label.htm
        """
        ############################# Frame 3 ############################# 
        self.frame3=tk.Frame(self)
        self.frame3.grid(row=3, column=0)
        
        self.observacion1 = tk.Label(self.frame3, text="Instrucciones de uso",font=('arial',10,'bold'))
        self.observacion1.grid(row=3, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="En plantilla excel debe cargar los siguientes campos",font=('arial',10))
        self.observacion1.grid(row=4, column=0,padx=2,pady=2)

        self.observacion1 = tk.Label(self.frame3, text="ID set, ID articulo, Descripcion, Precio, Moneda, Vigencia, Cantidad minima, Dias de plazo y Proveedor.",font=('arial',10,'bold'))
        self.observacion1.grid(row=5, column=0,padx=2,pady=2)

        self.observacion1 = tk.Label(self.frame3, text="En caso de no cargar nada, solamente se colocara el check a Inventory item",font=('arial',10))
        self.observacion1.grid(row=6, column=0,padx=2,pady=2)        
              
        self.observacion1 = tk.Label(self.frame3, text="Creado para PeopleSoft en Ingles.",font=('arial',10))
        self.observacion1.grid(row=7, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Cada vez que presiona Procesar, se envia email al director solicitando autorizacion de uso.",font=('arial',10))
        self.observacion1.grid(row=8, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="En el email de autorizacion, se adjunta excel de carga y nombre del usuario que opera robot.",font=('arial',10))
        self.observacion1.grid(row=9, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="El Robot se utiliza en su PC con N° " + socket.gethostname() + " e IP " + socket.gethostbyname(socket.gethostname()),font=('arial',10))
        self.observacion1.grid(row=10, column=0,padx=2,pady=2)

        self.observacion1 = tk.Label(self.frame3, text="Seguir recomendaciones para actualizacion del precio en formato decimal",font=('arial',10,'bold'),fg='blue')
        self.observacion1.grid(row=11, column=0,padx=2,pady=2)

        self.observacion1 = tk.Label(self.frame3, text="Configuracion excel: Archivo=>Opciones=>Avanzada=>Separador decimal “.” Separador en miles “,”",font=('arial',10))
        self.observacion1.grid(row=12, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Se recomienda focalizar uso de internet al robot",font=('arial',10,'bold'),fg='red')
        self.observacion1.grid(row=13, column=0,padx=2,pady=2)

       
        ############################# Frame 4 ############################# 

        self.frame4=tk.Frame(self)
        self.frame4.grid(row=12, column=0)
        
        scroll = tk.Scrollbar(self.frame4)
        scroll.grid(row=12, column=1,sticky="n"+"s"+"w")

        self.letrero1=tk.Text(self.frame4,wrap='none',padx=10,pady=20,width=70,height=15,yscrollcommand=scroll.set) #width=80,height=10,  
        self.letrero1.config(yscrollcommand=scroll.set) #width=80,height=10,  
        scroll.config(command=self.letrero1.yview)
        self.letrero1.grid(row=12, column=0)
        
         
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
        newMail.To = "marco.vera@serviciosandinos.net"
        #newMail.To = "matias.vidal@serviciosandinos.net"
        newMail.CC = "matias.vidal@serviciosandinos.net"
        #attachment1 = r"C:\Temp\example.pdf"
        #attachment1 = r"C:\Users\mvidal2\Desktop\data scientist\Finalizacion de OC\Finalizacion.xlsx"
        attachment1 = filename
        newMail.Attachments.Add(Source=attachment1)
        newMail.Display()
        #newMail.send()
        newMail.Send()

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
        
        #username.send_keys("311800185")
        #password.send_keys("Danola10.")
        
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
        
        #Items
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPCO_ITEMS").click() #ID
        
        #Define Items And Attributes
        driver.implicitly_wait(2)
        driver.find_element_by_id("fldra_EPIN_DEFINE_ITEMS").click() #ID
        
        #Define Item
        driver.implicitly_wait(4)
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
        R=db1['Proveedor']
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
        Observaciones_proveedor = []
        ID_articulo =[]
        #self.letrero1.update()
        
        try:
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
                r=str(R[n]) #Proveedor
                ###################################################################
                #Si ID Set no esta vacio, ejecutar
                if z!='nan':
                    ID_articulo.append(y)
                    mensaje="Modificando ID " + y
                    self.letrero1.insert(tk.END, mensaje+"\n")
                    self.letrero1.config(fg = 'green',height=12)
                    self.letrero1.update()
                    #Item Definition
                    driver.switch_to.default_content()
                    time.sleep(2)
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    #driver.switch_to.frame("ptifrmtgtframe")      
                    #time.sleep(2)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#MST_ITM_INV_VW_SETID")))
                    #time.sleep(2)
                    driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").click()
                    #time.sleep(1)
                    driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").clear()
                    #time.sleep(1)
                    #Colocar unidad de negocio
                    driver.find_element_by_id("MST_ITM_INV_VW_SETID").send_keys(z)
                    #time.sleep(1)
                    driver.find_element_by_css_selector("#MST_ITM_INV_VW_INV_ITEM_ID").click()
                    #time.sleep(1)
                    driver.find_element_by_id("MST_ITM_INV_VW_INV_ITEM_ID").clear()
                    #Colocar ID articulo
                    time.sleep(1)                    
                    driver.find_element_by_id("MST_ITM_INV_VW_INV_ITEM_ID").send_keys(y)
                    time.sleep(1)
                    #Asegurar click en Correct History
                    if driver.find_element_by_id("#ICCorrectHistory").is_selected() == False:
                        driver.find_element_by_id("#ICCorrectHistory").click()
                    #Hacer Click
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"#ICSearch"))).click()                            
                    ###############################################################################
                    #https://www.geeksforgeeks.org/switch-case-in-python-replacement/
                    #Hacer click en Item Type,Inventory Item
                    if driver.find_element_by_id("MASTER_ITEM_TBL_INVENTORY_ITEM").is_selected() == True:
                        driver.find_element_by_id("MASTER_ITEM_TBL_INVENTORY_ITEM").click()
                    if x!='nan': #cambio de descripcion
                        start4 = time.time()
                        #Cambio de descripcion
                        driver.switch_to.default_content()
                        #driver.switch_to.frame("TargetContent")  
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_TBL_DESCR60"))).clear()
                        #driver.find_element_by_id("REQ_RC_WB_FROM_REQ").clear()
                        driver.implicitly_wait(1)
                        driver.find_element_by_id("MASTER_ITEM_TBL_DESCR60").send_keys(x)
                        #Purchasing Item Attributes
                        driver.implicitly_wait(1)       
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        try:
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_WRK_PO_ITEM_ATTR_PB")))
                            driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        except TimeoutException:
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_WRK_PO_ITEM_ATTR_PB")))
                            driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR254_MIXED"))).click()
                        driver.implicitly_wait(2)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR254_MIXED"))).clear()
                        driver.implicitly_wait(2)
                        driver.find_element_by_id("PURCH_ITEM_ATTR_DESCR254_MIXED").send_keys(x)
                        driver.implicitly_wait(4)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR"))).click()
                        driver.implicitly_wait(2)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCR"))).clear()
                        driver.implicitly_wait(2)
                        driver.find_element_by_id("PURCH_ITEM_ATTR_DESCR").send_keys(x)
                        driver.implicitly_wait(4)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCRSHORT"))).click()
                        driver.implicitly_wait(2)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_DESCRSHORT"))).clear()
                        driver.implicitly_wait(2)
                        driver.find_element_by_id("PURCH_ITEM_ATTR_DESCRSHORT").send_keys(x)
                        driver.implicitly_wait(2)
                        #OK
                        #driver.find_element_by_name("#ICSave").click()
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME,"#ICSave"))).click()
                        time.sleep(2)
                        #Save
                        driver.execute_script("window.scrollTo(0, 600)")
                        #driver.find_element_by_name("#ICSave").click()
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME,"#ICSave"))).click()
                        time.sleep(2)
                        end4 = time.time()
                        mensaje="en cambiar la descripcion me demore " + str(int(end4-start4)) + " seg. Al articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
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
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        #driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_id("#ICCorrection").click()                    
                        #Purchasing Item Attributes
                        time.sleep(1)     
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_WRK_PO_ITEM_ATTR_PB"))).click()
                        #driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        #Cambio de precio en Purchasing Attributes
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRICE_LIST"))).click()
                        #driver.find_element_by_id("PURCH_ITEM_ATTR_PRICE_LIST").click()
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRICE_LIST"))).clear()
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRICE_LIST"))).send_keys(w)
                        #Cambio de moneda en Purchasing Attributes
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_CURRENCY_CD"))).click() #ID
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_CURRENCY_CD"))).clear()
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_CURRENCY_CD"))).send_keys(v)
                        #Ventana "Item Vendor"
                        time.sleep(1)
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        item_vendor=WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#PSTAB > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1)")))
                        item_vendor.click()
                        #Item Vendor UOM 
                        #time.sleep(1)
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_WRK_VNDR_UOM_PB$0"))).click()
                        # apretar "+"
                        #time.sleep(1)
                        #driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_css_selector("#\$ICField39\$new\$0\$\$0 > img:nth-child(1)").click()
                        #En pantalla "Vendor's UOM and Pricing Information"
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,"ITM_VNDR_UOM_PR_PRICE_VNDR$0"))).click()
                        #driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").clear()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VNDR_UOM_PR_PRICE_VNDR$0").send_keys(w)
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").clear()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VNDR_UOM_PR_CURRENCY_CD$0").send_keys(v)
                        #OK
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICSave"))).click()
                        #driver.find_element_by_name("#ICSave").click()
                        #Ok
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICSave"))).click()
                        #driver.find_element_by_name("#ICSave").click()
                        #Save
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICSave"))).click()
                        #driver.find_element_by_name("#ICSave").click()           
                        end5=time.time()
                        mensaje="en cambiar la precios me demore " + str(int(end5-start5)) + " seg. Al Articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
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
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        #driver.switch_to.frame("ptifrmtgtframe")  
                        #WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        time.sleep(1)
                        List_current_status=Select(driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STATUS_CURRENT"))
                        time.sleep(1)
                        List_current_status.select_by_visible_text('Active')
                        time.sleep(1)
                        List_future_status=Select(driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STATUS_FUTURE"))
                        time.sleep(1)
                        List_future_status.select_by_visible_text('Inactive')
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_TBL_ITM_STAT_DT_FUTURE"))).clear()
                        time.sleep(1)
                        driver.find_element_by_id("MASTER_ITEM_TBL_ITM_STAT_DT_FUTURE").send_keys(u)
                        #Ok
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        end6=time.time()
                        mensaje="en cambiar la fecha vigencia me demore " + str(int(end6-start6)) + " seg. Al Articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
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
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        #driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_id("#ICCorrection").click() 
                        #Purchasing Item Attributes
                        time.sleep(1)     
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        #Item vendor
                        time.sleep(1)     
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
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
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        driver.implicitly_wait(1)
                        #Save
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()           
                        time.sleep(1)
                        end7=time.time()
                        mensaje="en cambiar las cantidades minimas me demore " + str(int(end7-start7)) + " seg. Al Articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
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
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        #driver.switch_to.frame("ptifrmtgtframe")  
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
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        end8=time.time()
                        mensaje="en cambiar los dias de plazo me demore " + str(int(end8-start8)) + " seg. Al Articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
                        Observaciones_dias_plazo.append(mensaje)
                        print(mensaje)
                    else:
                        mensaje="no cambie dias de plazos"
                        Observaciones_dias_plazo.append(mensaje)
                        print(mensaje)
                    if r!='nan': #Proveedor
                        start9=time.time()
                        #Purchasing Item Attributes
                        time.sleep(1)       
                        driver.switch_to.default_content()
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                        #driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        #driver.find_element_by_id("MASTER_ITEM_WRK_PO_ITEM_ATTR_PB").click()
                        #Item vendor
                        time.sleep(1)     
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        #driver.find_element_by_css_selector("")
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#PSTAB > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1)"))).click()
                        #Status inactive al viejo proveedor
                        time.sleep(1)
                        List_current_status=Select(driver.find_element_by_id("ITM_VENDOR_ITM_STATUS$0"))
                        time.sleep(1)
                        List_current_status.select_by_visible_text('Inactive')           
                        #Priority 2 al viejo proveedor
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").clear()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").send_keys("2")                   
                        # apretar "+"
                        time.sleep(1)
                        driver.switch_to.default_content()
                        time.sleep(1)
                        driver.switch_to.frame("ptifrmtgtframe")  
                        time.sleep(1)
                        driver.find_element_by_css_selector("#\$ICField9\$new\$0\$\$0 > img:nth-child(1)").click()                   
                        #colocar el RUT de proveedor
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_VENDOR_ID$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_VENDOR_ID$0").clear()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_VENDOR_ID$0").send_keys(r)                    
                        #Status active
                        time.sleep(1)
                        List_current_status=Select(driver.find_element_by_id("ITM_VENDOR_ITM_STATUS$0"))
                        time.sleep(1)
                        List_current_status.select_by_visible_text('Active')                       
                        #Priority 1 al nuevo proveedor                   
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").clear()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_VNDR_PRIORITY$0").send_keys("1")           
                        #Colocando ID de articulo en Vendor Item ID                 
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_ID_VNDR$0").click()
                        time.sleep(1)
                        driver.find_element_by_id("ITM_VENDOR_ITM_ID_VNDR$0").send_keys(y)       
                        #Ok
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        #Save
                        time.sleep(1)
                        driver.execute_script("window.scrollTo(0, 600)")
                        time.sleep(1)
                        driver.find_element_by_name("#ICSave").click()
                        end9=time.time()
                        mensaje="en cambiar el proveedor me demore " + str(int(end9-start9)) + " seg. Al Articulo " + y
                        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
                        Observaciones_proveedor.append(mensaje)
                        print(mensaje)
                    else:
                        mensaje="no cambie de proveedor"
                        Observaciones_proveedor.append(mensaje)
                        print(mensaje)            
                    #Return to search
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,'#ICSave'))).click()
                    time.sleep(3)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,'#ICList'))).click() 
                    #driver.find_element_by_name("#ICList").click()
                    time.sleep(3) #Para asegurar la salida, lo ideal es un try
                    
            #db1.reset_index(drop=True)
            db2=pd.DataFrame(columns=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'])
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            driver.close()
            db2['Articulo']=ID_articulo
            db2['Observaciones_descripcion']=Observaciones_descripcion
            db2['Observaciones_precio']=Observaciones_precio
            db2['Observaciones_vigencia']=Observaciones_vigencia
            db2['Observaciones_cantidad_minima']=Observaciones_cantidad_minima
            db2['Observaciones_dias_plazo']=Observaciones_dias_plazo
            db2['Observaciones_proveedor']=Observaciones_proveedor
        
            self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg = 'blue',height=12)
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
        
        except TimeoutException:
            
            self.letrero1.insert(tk.END, "Proceso detenido debido a exceso tiempo de espera, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg = 'red',height=12)
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)

            mensaje= "me quede pegado en Psoft"
            Observaciones_descripcion.append(mensaje)
            Observaciones_precio.append(mensaje)
            Observaciones_vigencia.append(mensaje)
            Observaciones_cantidad_minima.append(mensaje)
            Observaciones_dias_plazo.append(mensaje)
            Observaciones_proveedor.append(mensaje)
            
            #db2=pd.DataFrame(columns=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'])
            df1 = pd.DataFrame({'Articulo':ID_articulo})
            df2 = pd.DataFrame({'Observaciones_descripcion':Observaciones_descripcion})
            df3 = pd.DataFrame({'Observaciones_precio':Observaciones_precio})
            df4 = pd.DataFrame({'Observaciones_vigencia':Observaciones_vigencia})
            df5 = pd.DataFrame({'Observaciones_cantidad_minima':Observaciones_cantidad_minima})
            df6 = pd.DataFrame({'Observaciones_dias_plazo':Observaciones_dias_plazo})
            df7 = pd.DataFrame({'Observaciones_proveedor':Observaciones_proveedor})
            
            #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
            db2=pd.concat([df1,df2,df3,df4,df5,df6,df7], ignore_index=True, names=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'],axis=1).fillna("-")
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
                   

        except NoSuchElementException:  
            self.letrero1.insert(tk.END, "Proceso detenido debido que no encontro link, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg ='red',height=12)
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
            
            mensaje= "no encontre link"
            Observaciones_descripcion.append(mensaje)
            Observaciones_precio.append(mensaje)
            Observaciones_vigencia.append(mensaje)
            Observaciones_cantidad_minima.append(mensaje)
            Observaciones_dias_plazo.append(mensaje)
            Observaciones_proveedor.append(mensaje)
            
            #db2=pd.DataFrame(columns=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'])
            df1 = pd.DataFrame({'Articulo':ID_articulo})
            df2 = pd.DataFrame({'Observaciones_descripcion':Observaciones_descripcion})
            df3 = pd.DataFrame({'Observaciones_precio':Observaciones_precio})
            df4 = pd.DataFrame({'Observaciones_vigencia':Observaciones_vigencia})
            df5 = pd.DataFrame({'Observaciones_cantidad_minima':Observaciones_cantidad_minima})
            df6 = pd.DataFrame({'Observaciones_dias_plazo':Observaciones_dias_plazo})
            df7 = pd.DataFrame({'Observaciones_proveedor':Observaciones_proveedor})
            
            #https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243
            db2=pd.concat([df1,df2,df3,df4,df5,df6,df7], ignore_index=True, names=['Articulo','Observaciones_descripcion','Observaciones_precio','Observaciones_vigencia','Observaciones_cantidad_minima','Observaciones_dias_plazo','Observaciones_proveedor'],axis=1).fillna("-")
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
            
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
        self.quit()
        
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Mantencion de ID de articulos de catalogo y contrato")
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.resizable(0,0)
    app=MainApplication(root)
    root.mainloop()