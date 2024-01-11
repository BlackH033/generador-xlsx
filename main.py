
import customtkinter
import os
from PIL import Image, ImageTk
from tkinter import filedialog
from windows import *
import json
import pandas as pd
import numpy
import openpyxl
#----------------------------
customtkinter.set_appearance_mode("Light")    
customtkinter.set_default_color_theme("blue") 
#----------------------------

class App(customtkinter.CTk):
    carpeta_raiz=os.path.dirname(__file__)          #guarda la ruta donde se encuentra este archivo .py
    carpeta_img=os.path.join(carpeta_raiz,"img")    #crea la ruta relativa a la carpeta /img - la cual se guarda en la misma ruta del archivo .py
    route=""
    folder=""
    conten=[]
    def __init__(self):
        super().__init__()
        self.title("Generador V1.0")
        self.geometry(f"{350}x{540}")  
        self.resizable(width=False, height=False)
        self.iconbitmap(os.path.join(App.carpeta_img,"icono.ico"))   

        self.back=customtkinter.CTkFrame(self, width=310, corner_radius=0,fg_color="transparent")
        self.back.grid(row=0,column=0,rowspan=4,sticky="nsew")
        self.back.grid_rowconfigure(4, weight=1)

        self.f1=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.f1.grid(row=0,column=0,sticky="nswe",pady=20)
        self.texto1=customtkinter.CTkLabel(self.f1,text="Generar XLSX",font=customtkinter.CTkFont(size=20,weight="bold"))
        self.texto1.grid(row=0,column=0,sticky="nswe",padx=100)

        self.f2=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.f2.grid(row=1,column=0,sticky="nswe",pady=10)
        self.img1=customtkinter.CTkImage(Image.open(os.path.join(App.carpeta_img,"carpeta.png")),size=(120,120))
        self.img1_insert=customtkinter.CTkLabel(self.f2,image=self.img1,text="")
        self.img1_insert.grid(row=0,column=0,sticky="nswe",padx=110)

        
        self.textcorrecto=customtkinter.CTkLabel(self.back,text="Carpeta:",font=customtkinter.CTkFont(size=18,weight="bold"))
        self.textcorrecto.grid(row=2,column=0,padx=40,pady=(0,5),sticky="nsew")

        self.fl=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.fl.grid(row=3,column=0,pady=(0,5),sticky="nsew")
        self.textbox = customtkinter.CTkTextbox(self.fl,width=290,height=30)
        self.textbox.insert("0.0","/")
        self.textbox.grid(row=0, column=0, sticky="nsew",padx=(10,5))
        self.boton_dl=customtkinter.CTkButton(self.fl, text="",image=customtkinter.CTkImage(Image.open(os.path.join(App.carpeta_img,"image.png")),size=(20,20)),font=customtkinter.CTkFont(size=15,weight="bold"),height=20,width=20,fg_color="#56524D",state="disable")
        self.boton_dl.grid(row=0,column=1,padx=(0,10))

        self.ft=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.ft.grid(row=4,column=0,pady=(5,0))
        self.textcorrecto2=customtkinter.CTkLabel(self.ft,text="# archivos txt:",font=customtkinter.CTkFont(size=14,weight="bold"))
        self.textcorrecto2.grid(row=0,column=0,sticky="nsew")
        self.textcorrecto3=customtkinter.CTkLabel(self.ft,text=" 0",font=customtkinter.CTkFont(size=15,weight="bold",),text_color="red")
        self.textcorrecto3.grid(row=0,column=1,sticky="nsew")

        self.ft1=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.ft1.grid(row=5,column=0,pady=(5,2))
        self.textcorrectob1=customtkinter.CTkLabel(self.ft1,text=" ",font=customtkinter.CTkFont(size=14,weight="bold"))
        self.textcorrectob1.grid(row=0,column=0,sticky="nsew")
        self.textcorrectob5=customtkinter.CTkLabel(self.ft1,text=" ",font=customtkinter.CTkFont(size=14,weight="bold"))
        self.textcorrectob5.grid(row=0,column=1,sticky="nsew",padx=5)
        self.ft3=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.ft3.grid(row=6,column=0)
        self.textcorrectob2=customtkinter.CTkLabel(self.ft3,text=" ",font=customtkinter.CTkFont(size=14,weight="bold"))
        self.textcorrectob2.grid(row=0,column=0,sticky="nsew",pady=(5,10))

        self.f4=customtkinter.CTkFrame(self.back,fg_color="transparent")
        self.f4.grid(row=7,column=0)
        self.boton1=customtkinter.CTkButton(self.f4, text="Agregar carpeta",command=self.carpeta,font=customtkinter.CTkFont(size=15,weight="bold"),height=40,width=100)
        self.boton1.grid(row=0,column=0,padx=(0,20))
        self.boton1_1=customtkinter.CTkButton(self.f4, text=" Generar .xlsx  ",fg_color="#56524D",state="disable",font=customtkinter.CTkFont(size=15,weight="bold"),height=40,width=100)
        self.boton1_1.grid(row=0,column=1)

        self.f4=customtkinter.CTkFrame(self.back,width=350,corner_radius=0)
        self.f4.grid(row=8,column=0,padx=0,pady=(30,0))
        self.img2=customtkinter.CTkImage(Image.open(os.path.join(App.carpeta_img,"logo_isa.png")),size=(118,63))
        self.img2_insert=customtkinter.CTkLabel(self.f4,image=self.img2,text="")
        self.img2_insert.grid(row=0,column=0,padx=(10,10))
        self.boton2=customtkinter.CTkButton(self.f4, text ="CERRAR", command = self.destroy,fg_color="red",hover_color="#A50000",font=customtkinter.CTkFont(weight="bold"),width=70)
        self.boton2.grid(row=0,column=1,padx=(120,20),pady=30)

    def carpeta(self):
        filename = filedialog.askdirectory(
        parent=self,
        title="Agregar carpeta con los .txt"
        )
        App.route=filename
        if App.route!="":
            App.folder=os.listdir(App.route)
            App.conten=[i for i in App.folder if i[-4:]==".txt"]
            if len(App.conten)>0:
                self.textcorrecto3.configure(text=f" {len(App.conten)}",text_color="green")
                self.textbox.insert("0.0",App.route)
                self.boton1_1.configure(state="normal",fg_color="#3B8ED0",command=self.generar)
                self.boton_dl.configure(fg_color="red",hover_color="#A50000",state="normal",command=self.delete)                
            else:
                self.textbox.delete("0.0", "end")
                self.textbox.insert("0.0","/")
                self.boton1_1.configure(state="disable",fg_color="#56524D",command=None)
                self.textcorrecto3.configure(text=f" 0",text_color="red")
                self.boton_dl.configure(state="disable",fg_color="#56524D",command=None)
                print("!!No hay archivos!!")
                self.ventana_error=ventana_secundaria()
                self.ventana_error.error_carpeta_formato()
            print(App.conten)
    def delete(self):
        self.textcorrectob1=customtkinter.CTkLabel(self.ft1,text=" ",font=customtkinter.CTkFont(size=14,weight="bold"))
        self.textcorrectob1.grid(row=0,column=0,sticky="nsew")
        self.textcorrectob5.configure(text=f" ",text_color="red")
        App.route=""
        self.textbox.delete("0.0", "end")
        self.textbox.insert("0.0","/")
        self.boton1_1.configure(state="disable",fg_color="#56524D",command=None)
        self.textcorrecto3.configure(text=f" 0",text_color="red")
        self.textcorrectob2.configure(text=f" ",text_color="red")
        self.boton_dl.configure(state="disable",fg_color="#56524D",command=None)
        print("bt delete")

    def generar(self):
        print("Generando archivo")
        #----
        self.barr= customtkinter.CTkProgressBar(self.ft1, orientation="horizontal",mode='determinate',progress_color="red")
        self.barr.set(0)
        self.barr.grid(row=0,column=0)
        incremento=1/(len(App.conten)+(int(len(App.conten)*0.1)))
        conteo=1
        suma=incremento
        self.textcorrectob5.configure(text=f"0%",text_color="red") 
        self.barr.start()       
        #----
        dt=pd.DataFrame()
        for i in App.conten:
            self.textcorrectob5.configure(text=f"{min(round(suma*100,0),100)}%",text_color="red") 
            self.barr.set(suma)
            self.textcorrectob2.configure(text=f"Procesando txt {conteo}/{len(App.conten)}",text_color="red")
            suma += incremento
            conteo+=1
            self.update_idletasks()
            print(i)
            archive=open(os.path.join(App.route,i),"r",encoding="utf8").read()
            data_json=json.loads(archive)
            dt2=pd.DataFrame(data_json["Datos"])
            dt=pd.concat([dt,dt2],ignore_index=True)
        name="resultado"
        conteo=1
        self.textcorrectob2.configure(text=f"Configurando xlsx",text_color="red")
        if "resultado.xlsx" in App.folder:
            while True:
                if f"resultado_{conteo}.xlsx" not in App.folder:
                    name=f"resultado_{conteo}"
                    break
                conteo+=1
        self.textcorrectob5.configure(text=f"{min(round(suma*100,0),100)}%",text_color="red") 
        self.barr.set(suma)
        suma += incremento
        self.update_idletasks()
        self.textcorrectob2.configure(text=f"Creando xlsx",text_color="red")
        dt.to_excel(f"{os.path.join(App.route,name)}.xlsx",index=False)
        self.barr.configure(progress_color="green")
        self.textcorrectob5.configure(text=f"{100}%",text_color="green") 
        self.textcorrectob2.configure(text=f"finalizado",text_color="green")
        self.barr.set(1)
        self.update_idletasks()
        self.barr.stop()
        App.folder=os.listdir(App.route)
        self.ventana=ventana_secundaria()
        self.ventana.generado_correcto(f"{os.path.join(App.route,name)}.xlsx")   
    
    
    

if __name__ == "__main__":
        app = App()
        app.mainloop()