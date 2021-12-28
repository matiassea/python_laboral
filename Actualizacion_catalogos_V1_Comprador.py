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
import datetime
from datetime import date
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

"""
Colocar PER00 en ITM_VENDOR_VENDOR_SETID$0
Coloar formator peroveedor 0000009302 ITM_VENDOR_VENDOR_ID$0, linea 664
"""

"""
Ventana de seguridad
https://stackoverflow.com/questions/16115378/tkinter-example-code-for-multiple-windows-why-wont-buttons-load-correctly
https://realpython.com/python-gui-tkinter/

"""
class MainApplication(tk.Frame):
        
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        
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

        hoy = datetime.date.today()
        self.final = datetime.date(2021,5,27) #YY//MM//DD
        
        #Se captura el nombre del equipo para que trabaje en solo un equipo
        hostname = socket.gethostname()
        
        if hoy < self.final and (hostname == "CA-HP5CG01860Y3" or hostname == "SSC-SD307"): # or hostname == "SSC-SD307"
            self.grafica_completa()
        elif hoy < self.final and (hostname != "CA-HP5CG01860Y3"):
            self.grafica_incompleta_computador_no_autorizado()
        elif hoy > self.final and (hostname == "CA-HP5CG01860Y3"):
            self.grafica_incompleta()
        else:
            self.grafica_incompleta()
            
        
    def grafica_completa(self):
        
        self.frame = tk.Frame(self)
        self.frame.grid(row=0, column=0)
        
        #Colocando Gadget
        self.button2 = tk.Button(self.frame,text='Open File',font=('arial',10,'bold'), command=self.mfileopen)
        self.button2.grid(row=0, column=0,padx=10,pady=5)
        
        self.button4 = tk.Button(self.frame,text='Procesar',font=('arial',10,'bold'), command=self.procesar)
        self.button4.grid(row=0, column=1,padx=10,pady=5)

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
        
        self.observacion1 = tk.Label(self.frame3, text="Debe cargar plantilla excel ID set y ID articulo",font=('arial',10))
        self.observacion1.grid(row=4, column=0,padx=2,pady=2)
              
        self.observacion1 = tk.Label(self.frame3, text="Creado para PeopleSoft en Ingles.",font=('arial',10))
        self.observacion1.grid(row=5, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Cada vez que presiona Procesar, se envia email para registro del proceso.",font=('arial',10))
        self.observacion1.grid(row=6, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="En el email de autorizacion, se adjunta excel de carga y nombre del usuario que opera robot.",font=('arial',10))
        self.observacion1.grid(row=7, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Se utiliza en PC " + socket.gethostname() + ". Numero IP " + socket.gethostbyname(socket.gethostname()),font=('arial',10))
        self.observacion1.grid(row=8, column=0,padx=2,pady=2)

        self.observacion1 = tk.Label(self.frame3, text="Configuracion excel: Archivo=>Opciones=>Avanzada=>Separador decimal “.” Separador en miles “,”",font=('arial',10))
        self.observacion1.grid(row=9, column=0,padx=2,pady=2)
        
        self.observacion1 = tk.Label(self.frame3, text="Debe optimizar el uso de internet",font=('arial',10),fg='red')
        self.observacion1.grid(row=10, column=0,padx=2,pady=2)

        """
        Scroll
        https://stackoverflow.com/questions/19646752/python-scrollbar-on-text-widget/19647325
        """
        ############################# Frame 4 ############################# 

        self.frame4=tk.Frame(self)
        self.frame4.grid(row=12, column=0)
        
        scroll = tk.Scrollbar(self.frame4)
        scroll.grid(row=12, column=1,sticky="n"+"s"+"w")

        self.letrero1=tk.Text(self.frame4,wrap='none',padx=10,pady=20,width=70,height=15,yscrollcommand=scroll.set) #width=80,height=10,  
        self.letrero1.config(yscrollcommand=scroll.set) #width=80,height=10,  
        scroll.config(command=self.letrero1.yview)
        self.letrero1.grid(row=12, column=0)
        
    def grafica_incompleta(self):
        
        """
        Entry es solamente para una linea de texto
        https://effbot.org/tkinterbook/entry.htm
        self.letrero1=tk.Entry(self.frame,textvariable=self.var, fg='red', bd=3,width = 90, height = 1)
        """
        
        ############################# Frame 2 #############################
        self.frame2 = tk.Frame(self)
        self.frame2.grid(row=1, column=0)
        
        self.observacion1A = tk.Label(self.frame2, text="Observaciones",font=('arial',10,'bold'))
        self.observacion1A.grid(row=1, column=2,padx=10,pady=5)

        self.observacion1A = tk.Label(self.frame2, text="Su uso ha expirado",font=('arial',10,'bold'))
        self.observacion1A.grid(row=2, column=2,padx=10,pady=5)
        
        self.observacion1A = tk.Label(self.frame2, text="Favor contactar al celular 957198751",font=('arial',10))
        self.observacion1A.grid(row=3, column=2,padx=10,pady=5)
        
        self.observacion1A = tk.Label(self.frame2, text="Este programa expira el dia " + str(self.final.strftime("%d")) + " de " + str(self.final.strftime("%B")),font=('arial',10),fg='red')
        self.observacion1A.grid(row=4, column=2,padx=10,pady=5)

    def grafica_incompleta_computador_no_autorizado(self):
        
        """
        Entry es solamente para una linea de texto
        https://effbot.org/tkinterbook/entry.htm
        self.letrero1=tk.Entry(self.frame,textvariable=self.var, fg='red', bd=3,width = 90, height = 1)
        """
        
        ############################# Frame 2 #############################
        self.frame2 = tk.Frame(self)
        self.frame2.grid(row=1, column=0)
        
        self.observacion1A = tk.Label(self.frame2, text="Observaciones",font=('arial',10,'bold'))
        self.observacion1A.grid(row=1, column=2,padx=10,pady=5)

        self.observacion1A = tk.Label(self.frame2, text="Su uso no esta autorizado para este PC",font=('arial',10,'bold'))
        self.observacion1A.grid(row=2, column=2,padx=10,pady=5)
        
        self.observacion1A = tk.Label(self.frame2, text="Favor contactar al celular 957198751",font=('arial',10))
        self.observacion1A.grid(row=3, column=2,padx=10,pady=5)


    def procesar(self):
        global driver
        """
        https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send
        https://gist.github.com/ITSecMedia/b45d21224c4ea16bf4a72e2a03f741af
        https://stackoverflow.com/questions/50926514/send-email-through-python-using-outlook-2016-without-opening-it
        """
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Operacion de robot de limpieza comprador primario en ID articulos"
        newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
        newMail.HTMLBody = "<HTML><BODY><p>Estimada</p><p>Comenzara el proceso robot de limpieza comprador por ID articulos</p><p>El proceso se realizara en PC "+ socket.gethostname() + "</p></BODY></HTML>" +"<HTML><BODY><p>Adjunto listado de ID de articulos a borrar comprador, robot V1</p><br>Favor considerar</BODY></HTML>"
        newMail.To = "sofia.fuentes@unab.cl"
        newMail.CC = "matias.vidal@serviciosandinos.net"
        attachment1 = filename
        newMail.Attachments.Add(Source=attachment1)
        newMail.Display()
        newMail.Send()

        self.letrero1.insert(tk.END, "Email enviado\n")
        self.letrero1.config(fg = 'green',height=12)

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
        #Preparando el archivo y columna de observaciones
        #db1 = pd.read_excel('Input.xlsx')
        Y=db1['ID articulo']
        Z=db1['ID Set']

        ID_articulo =[]
        
        try:
            for n in range (len(db1['ID articulo'])):
                start1 = time.time()
                z=str(Z[n]) #ID Set
                y=str(Y[n]) #ID Articulo
                ###################################################################
                #Si ID Set no esta vacio, ejecutar
                if z!='nan':
                    ID_articulo.append(y)
                    #Item Definition
                    driver.switch_to.default_content()
                    time.sleep(2)
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'TargetContent')))
                    #driver.switch_to.frame("ptifrmtgtframe")
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#MST_ITM_INV_VW_SETID")))
                    #Colocar unidad de negocio
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_SETID"))).click()
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_SETID"))).clear()
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_SETID"))).send_keys(z)
                    
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform() #se aprieta escape para evitar el menu desplegable

                    #Colocar ID articulo
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_INV_ITEM_ID"))).click()
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_INV_ITEM_ID"))).clear()
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"MST_ITM_INV_VW_INV_ITEM_ID"))).send_keys(y)

                    #Asegurar click en Correct History
                    if driver.find_element_by_id("#ICCorrectHistory").is_selected() == False:
                        driver.find_element_by_id("#ICCorrectHistory").click()
                    #Hacer Click
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID,"#ICSearch"))).click()
                    
                    ###############################################################################
                    #https://www.geeksforgeeks.org/switch-case-in-python-replacement/
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

                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRIMARY_BUYER"))).click()
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PURCH_ITEM_ATTR_PRIMARY_BUYER"))).clear()

                    #Save
                    driver.execute_script("window.scrollTo(0, 600)")
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME,"#ICSave"))).click()
                    time.sleep(1)
                    end1 = time.time()
                    mensaje="He borrado comprador primario, me demore " + str(int(end1-start1)) + " seg. Al articulo " + y
                    print(mensaje)
                    self.letrero1.insert(tk.END, mensaje+"\n")
                    self.letrero1.update()

                    #Save
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(1)
                    driver.find_element_by_name("#ICSave").click()

                    #Return to search
                    time.sleep(1)
                    driver.execute_script("window.scrollTo(0, 600)")
                    time.sleep(1)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID,'#ICList'))).click()
                    #driver.find_element_by_name("#ICList").click()
                    
                    Titulo=(By.CLASS_NAME, "PSSRCHEDITBOXLABEL")
                    try:
                        WebDriverWait(driver, 5).until(EC.text_to_be_present_in_element((Titulo),"SetID:"))
                        pass
                    except TimeoutException:
                        driver.execute_script("window.scrollTo(0, 600)")
                        driver.find_element_by_name("#ICList").click()
                    time.sleep(2)

            driver.close()
            db2=pd.DataFrame(columns=['Articulo'])
            db2['Articulo']=ID_articulo
            self.letrero1.insert(tk.END, "Proceso terminado, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg = 'blue',height=12)
            df1 = pd.DataFrame({'Articulo':ID_articulo})
            #db2=pd.concat([df1], ignore_index=True, names=['Articulo'],axis=1).fillna("-")
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
        
        except TimeoutException:
            self.letrero1.insert(tk.END, "Proceso detenido debido a exceso tiempo de espera, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg = 'red',height=12)
            df1 = pd.DataFrame({'Articulo':ID_articulo})
            db2=pd.concat([df1], ignore_index=True, names=['Articulo'],axis=1).fillna("-")
            db2=pd.DataFrame(columns=['Articulo'])
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
                   
        except NoSuchElementException:  
            self.letrero1.insert(tk.END, "Proceso detenido debido que no encontro link, revisar Resumen_del_proceso.xlsx \n")
            self.letrero1.config(fg ='red',height=12)
            df1 = pd.DataFrame({'Articulo':ID_articulo})
            db2=pd.concat([df1], ignore_index=True,axis=1).fillna("-")
            db2=pd.DataFrame(columns=['Articulo'])
            db2.to_excel("Resumen_del_proceso.xlsx",index = False)
            
    def mfileopen(self):
        global filename
        global db1
        global Y
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
    root.title("Robot limpiador comprador primario en ID de articulos")
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.resizable(0,0)
    app=MainApplication(root)
    root.mainloop()