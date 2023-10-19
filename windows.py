#--------------librerias --------------
import customtkinter
import os
from PIL import Image
#--------------------------------------

class ventana_secundaria(customtkinter.CTkToplevel):
    carpeta_raiz=os.path.dirname(__file__)                 #guarda la ruta donde se encuentra este archivo .py
    carpeta_img=os.path.join(carpeta_raiz,"img")           #crea la ruta relativa a la carpeta /img - la cual se guarda en la misma ruta del archivo .py
    
    def __init__(self):
        super().__init__()
        self.grab_set()
        self.resizable(width=False, height=False)          #no permite cambiar el tamaño de la ventana
        self.iconbitmap(os.path.join(ventana_secundaria.carpeta_img,"icono.ico"))

    def boton_cerrar_rojo(self,index=int):
        """
        boton_cerrar_rojo(index) -> crea un boton con el texto "CERRAR" en color rojo   
                                    recibe como parametro la posicion fila donde se ubicará
        """
        self.btn_cerrar = customtkinter.CTkButton(self, text ="CERRAR", command = self.destroy,fg_color="red",hover_color="#A50000")
        self.btn_cerrar.grid(row=index,column=0,sticky="ew",padx=60,pady=(20, 30))
    
    
    def icono(self,name=str):
        """
        crea el icono en la ventana de aviso 
        recibe como parametro el nombre del icono a mostrar
        """
        self.iconoerror=customtkinter.CTkImage(Image.open(os.path.join(ventana_secundaria.carpeta_img,name+".png")),size=(90,90))
        self.iconoerror=customtkinter.CTkLabel(self, image = self.iconoerror,text="")
        self.iconoerror.grid(row=0,column=0,pady=(40,20),sticky="nsew") 
    
    def error_carpeta_formato(self):
        """
        rellena la ventana secundaria con la información necesaria 
        para indicar que hubo error generando el formato .xlsx
        """
        self.title("!!ERROR!!")
        self.icono("cancelar")
        self.texterror=customtkinter.CTkLabel(self,text="No hay archivos compatibles en la carpeta",font=customtkinter.CTkFont(size=18,weight="bold"))
        self.texterror.grid(row=1,column=0,padx=40,sticky="nsew")
        self.boton_cerrar_rojo(2)
    
    def generado_correcto(self,ruta=str):
        """
        rellena la ventana secundaria con la información necesaria para 
        indicar que hubo error generando el formato .xlsx .
        recibe como parametro la ruta donde se generó el formato
        """
        self.title(".xlsx generado")
        self.icono("image1")
        self.textcorrecto=customtkinter.CTkLabel(self,text=".xlsx generado correctamente en",font=customtkinter.CTkFont(size=18,weight="bold"))
        self.textcorrecto.grid(row=1,column=0,padx=40,pady=(0,20),sticky="nsew")
        self.textbox = customtkinter.CTkTextbox(self,width=80,height=20)
        self.textbox.insert("0.0",ruta.replace("/","\\"))
        self.textbox.grid(row=2, column=0, sticky="nsew",padx=20)
 
        self.btn_abrir = customtkinter.CTkButton(self, text ="Abrir archivo", command = lambda:self.abrir_carpeta(ruta),fg_color="green",hover_color="#0DAF0A",font=customtkinter.CTkFont(size=14,weight="bold"),width=100)
        self.btn_abrir.grid(row=3,column=0,sticky="nsew",pady=(20, 30),padx=200)

        self.f4=customtkinter.CTkFrame(self,corner_radius=0)
        self.f4.grid(row=5,column=0,padx=0,pady=(60,0))
        self.img2=customtkinter.CTkImage(Image.open(os.path.join(ventana_secundaria.carpeta_img,"logo_isa.png")),size=(118,63))
        self.img2_insert=customtkinter.CTkLabel(self.f4,image=self.img2,text="")
        self.img2_insert.grid(row=0,column=0,padx=(10,10))
        self.boton2=customtkinter.CTkButton(self.f4, text ="CERRAR", command = self.destroy,fg_color="red",hover_color="#A50000",font=customtkinter.CTkFont(weight="bold"),width=70)
        self.boton2.grid(row=0,column=1,padx=(390,20),pady=30)
    
    def abrir_carpeta(self,ruta):
        print(ruta)
        os.system(f'start "excel" "{os.path.realpath(ruta)}"')