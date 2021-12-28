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


class MainApplication:
    
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)

        # here is the application variable, variable contents
        self.X = tk.StringVar()
        self.db1=tk.StringVar()
        self.Y=tk.StringVar()
        self.var=tk.StringVar() 
        self.db3=tk.Listbox()
        self.Observaciones=tk.StringVar()
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
        
        self.observacion1 = tk.Label(self.frame, text="No es apto para uso inapropiado",font=('arial',10),fg='red')
        self.observacion1.place(x=10,y=220)
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
        newMail.HTMLBody = "<HTML><BODY><p>Estimado Marco</p><p>Comenzara el proceso de mantencion de ID articulos</p><p>Adjunto listado de ID de arituclo a mantener</p><br>Favor enviar su aprobacion</BODY></HTML>"
        newMail.To = "matias.vidal@serviciosandinos.net"
        #attachment1 = r"C:\Temp\example.pdf"
        #attachment1 = r"C:\Users\mvidal2\Desktop\data scientist\Finalizacion de OC\Finalizacion.xlsx"
        attachment1 = filename
        newMail.Attachments.Add(Source=attachment1)
        newMail.display()
        #newMail.send()
        newMail.send
        
        start1 = time.time()
        driver = webdriver.Firefox()
        driver.get("http://www.google.com/")
        
        #######################################################################
        #Apertura del PeopleSoft
        
        #open tab
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't') 
        
        # Load a page 
        driver.get('https://leifs.mycmsc.com/psp/leifsprd/EMPLOYEE/ERP/?cmd=logout')
        
        username = driver.find_element_by_id("userid") #input id o name
        password = driver.find_element_by_id("pwd") #input id o name
        
        #username.send_keys("311800185")
        #password.send_keys("Danola9.")
        
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
        W=db1['Precio']
        X=db1['Descripcion']
        Y=db1['ID articulo']
        Z=db1['ID Set']
        Observaciones = []
        
        for n in range (len(db1['ID articulo'])):
            z=str(Z[n]) #ID Set
            y=str(Y[n]) #ID Articulo
            x=str(X[n]) #Descripcion
            w=str(W[n]) #Precio
            #w=str("0")*(10-z)+str(X[n])#OC  
            #t="4" #linea
            ###################################################################
            #Item Definition
            start3 = time.time()
            driver.switch_to.default_content()
            time.sleep(1)
            driver.switch_to.frame("ptifrmtgtframe")      
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#MST_ITM_INV_VW_SETID")))
            #driver.find_element_by_id("PO_RC_WB_BUSINESS_UNIT").click()
            driver.implicitly_wait(1)
            driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").click()
            driver.implicitly_wait(1)
            #driver.find_element_by_id("PO_RC_WB_BUSINESS_UNIT").clear()
            driver.find_element_by_css_selector("#MST_ITM_INV_VW_SETID").clear()
            driver.implicitly_wait(1)
            #Colocar unidad de negocio
            driver.find_element_by_id("MST_ITM_INV_VW_SETID").send_keys(z)
            driver.implicitly_wait(1)
            
            driver.implicitly_wait(1)
            driver.find_element_by_css_selector("#MST_ITM_INV_VW_INV_ITEM_ID").click()
            driver.implicitly_wait(1)
            #Colocar unidad de negocio
            driver.find_element_by_id("MST_ITM_INV_VW_INV_ITEM_ID").send_keys(y)
            driver.implicitly_wait(1)

            #Hacer Click
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"#ICSearch"))).click()
            #WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"#ICSearch"))).click()                                 
            ###############################################################################
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
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "MASTER_ITEM_WRK_PO_ITEM_ATTR_PB"))).click()
            
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
            
            #Search
            driver.find_element_by_name("#ICSave").click()
            driver.implicitly_wait(1)
            
            ##################################################################
            
            #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#REQ_RC_WB_WRK_SEARCH"))).click()
            #driver.find_element_by_name("REQ_RC_WB_WRK_SEARCH").click()
            #driver.implicitly_wait(1)
            #driver.find_element_by_css_selector("#REQ_RC_WB_WRK_SEARCH").click() #CSS Selector
            #driver.implicitly_wait(1)
            #driver.find_element_by_name("REQ_RC_WB_WRK_SEARCH").click()
            #driver.find_element_by_class_name("PSPUSHBUTTON").click()
            
            #Include Closed check
            driver.find_element_by_id("PO_RC_WB_WRK_FILTER_OPTIONS_PB").click()
                        
            #Bajar pagina
            #https://stackoverflow.com/questions/20986631/how-can-i-scroll-a-web-page-using-selenium-webdriver-in-python
            driver.execute_script("window.scrollTo(0, 600)")
            
            #repetido=(By.CLASS_NAME, "PSPUSHBUTTON")
            #repetido=(By.ID, "REQ_RC_WB_WRK_SEARCH")
            #https://selenium-python.readthedocs.io/waits.html
            #WebDriverWait(driver, 5).until(EC.text_to_be_present_in_element((repetido),"No matching values were found."))
            
            #WebDriverWait(driver, 5).until(EC.element_to_be_clickable(repetido)).click()
            #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_SEARCH")))
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_SEARCH"))).click()
            #driver.find_element_by_id("PO_RC_WB_WRK_SEARCH").click()
            #driver.find_element_by_id("PO_RC_WB_WRK_SEARCH").click()
            
            Titulo=(By.CLASS_NAME, "PAPAGETITLE")
            try: 
                WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((Titulo),"Buyer's WorkBench"))
                pass
            except TimeoutException:
                driver.execute_script("window.scrollTo(0, 600)")
                driver.find_element_by_id("PO_RC_WB_WRK_SEARCH").click()
                
                
            ###############################################################################
            #Segunda ventana, Requester's Workbench
                
            driver.switch_to.default_content()
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
            
            #https://selenium-python.readthedocs.io/locating-elements.html   
            #http://docs.python.org.ar/tutorial/3/errors.html
            
            #Marcando todas las lineas
            try: 
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_SCR$selm$0$$0")))
                driver.find_element_by_id("PO_RC_WB_SCR$selm$0$$0").click()
            except TimeoutException:
                break

            end3 = time.time()
            #print(end3 - start3, " Tiempo de comando antes de seleccionar lineas")
            start4 = time.time()
            
            ###############################################################################
            #Cuarta ventana
            
            #Close
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_CLOSED_FLG$0"))).click()
            
            ###############################################################################
            #Quinta ventana,Processing Results 
            
            #Cuadro emergente "No Purchase Orders found to process
            #Titulo=(By.CLASS_NAME, "PSEDITBOX_DISPONLY")
            #Verificando que el N°de OC sea el correcto
            Titulo=(By.ID,"PO_ID_EXP$0")
            try:
                WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((Titulo),w))
            except TimeoutException:
                #driver.switch_to.default_content()
                #En caso que la OC este cerrada, salir del Frame y volver a la pagina inicial
                mensaje="OC "+w+" ya cerrada."
                Observaciones.append(mensaje)
                print(mensaje)
                driver.switch_to.default_content()
                driver.find_element_by_css_selector("#pthnavbccref_EP_PO_RC_WB > a:nth-child(1)").click()
                ###############################################################################
                #Asegurando la salida a la pagina principal Buyer's WorkBench
                driver.switch_to.frame("ptifrmtgtframe")
                time.sleep(0.5)
                Titulo=(By.ID, "PO_RC_WB_BUSINESS_UNIT_LBL")
                WebDriverWait(driver, 240).until(EC.text_to_be_present_in_element((Titulo),"Business Unit:"))
                time.sleep(0.5)
                driver.switch_to.default_content()    
                continue
            
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "SELECTED_FLAG$0"))).click()
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.NAME,"PO_RC_WB_WRK_OVERRIDE_PB$IMG"))).click()
            time.sleep(10)

            #Proceed, yes
            driver.switch_to.default_content()
            driver.implicitly_wait(15)
            driver.switch_to.frame("ptifrmtgtframe")   
            driver.implicitly_wait(15)
            WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.ID,"PO_RC_WB_WRK_CONTINUE_PB")))

            driver.find_element_by_id("PO_RC_WB_WRK_CONTINUE_PB").click()

            #Ventana emergente
            driver.switch_to.default_content()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "#ICYes"))).click()
            
            
            end4 = time.time()
            #print(end4 - start4, " Tiempo de comando antes de terminar")
            start5 = time.time()
            
            ###############################################################################
            # Volver a Buyer Workbench, para hacer control de presupuesto
            driver.switch_to.frame("ptifrmtgtframe") 
            WebDriverWait(driver, 300).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_PB_BUDGET_CHECK$0"))).click()
            WebDriverWait(driver, 300).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_CONTINUE_PB"))).click()
            driver.switch_to.default_content()
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "#ICYes"))).click()
            
            end5 = time.time()
            #print(end6 - start6, "Fin del proceso Psoft " + w )
            start6=time.time()
            ###############################################################################
            #Volver a Buyer Workbench
            #driver.execute_script("window.scrollTo(0,0)")
            #Titulo=(By.CLASS_NAME, "PAPAGETITLE")
            #http://docs.python.org.ar/tutorial/3/errors.html
            try: 
                #WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((Titulo),"Buyer's WorkBench"))
                driver.switch_to.default_content()
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#pthnavbccref_EP_PO_RC_WB > a:nth-child(1)")))
                driver.find_element_by_css_selector("#pthnavbccref_EP_PO_RC_WB > a:nth-child(1)").click() #CSS Selector
            except TimeoutException:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PO_RC_WB_WRK_CONTINUE_PB"))).click()
                print("No pudo salir de la 'Buyer Workbench'")

            end6 = time.time()
            #print(end6 - start6, "Fin del proceso OC " + w + ",closed")
            start7=time.time()
    
            ###############################################################################
            #Asegurando la salida a la pagina principal Buyer's WorkBench
            driver.switch_to.frame("ptifrmtgtframe")
            time.sleep(0.5)
            Titulo=(By.ID, "PO_RC_WB_BUSINESS_UNIT_LBL")
            WebDriverWait(driver, 240).until(EC.text_to_be_present_in_element((Titulo),"Business Unit:"))
            time.sleep(0.5)
            driver.switch_to.default_content()
    
            end7 = time.time()
            mensaje= "Fin del proceso Psoft OC " + w +", tiempo total " + str(int(end7-start3)) + " seg."
            Observaciones.append(mensaje)
            print(mensaje)
    
        db1.reset_index(drop=True)
        db1['Observaciones']=Observaciones
        self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx")
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
    root.geometry("800x800")
    app=MainApplication(root)
    root.mainloop()