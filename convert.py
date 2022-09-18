from cProfile import label
from tkinter import*
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import os 
# import win32com.client
from pdf2docx import Converter

class ventana(Tk):
    # Constructor 
    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs)
        self.title("Convertidor de pdf a word")
        self.geometry("300x200")
        self.resizable(0,0)
        self.widgest()

    def salir(self):
        valor = messagebox.askquestion(title="Convertidor de pdf a word",message="¿Esta seguro que deseas salir?")
        if valor == "yes":
            self.destroy()

    def archivo(self):
        valor = askopenfilename(defaultextension=".pdf",filetypes=[("Selecciona solo archivos .pdf","*.pdf")])
        if valor:
            self.generate_word2(valor)

    def generate_word2(self,valor):
        try:
            labelFile = Label(self,text="Cargando...")
            labelFile.place(x=30,y=39)
            ruta = os.path.abspath(valor)
            ruta_final = os.path.abspath(valor[0:-4] +".docx".format())
            cv = Converter(ruta)
            cv.convert(ruta_final,start=0,end=None)
            cv.close()
            labelFile.destroy()
            messagebox.showinfo(title="Convertidor de pdf a word",message='El archivo se genero correctamente en la siguiente ruta: '+ruta_final)
        except Exception as e :
            messagebox.showerror(title="Convertidor de pdf a word",message=str(e))


    # def generate_word(self,valor):
    #     try:
    #         word = win32com.client.Dispatch("word.Application")
    #         word.visible = 0 
    #         ruta = os.path.abspath(valor)
    #         wb=word.Documents.Open(ruta)
    #         ruta_final = os.path.abspath(valor[0:-4] +"docx".format())
    #         wb.SaveAs2(ruta_final,FileFormat=16)
    #         wb.Close()
    #         word.Quit()
    #         # labelFile.destroy()
    #         messagebox.showinfo(title="Convertidor de pdf a word",message='El archivo se genero correctamente en la siguiente ruta: '+ruta_final)
    #     except Exception as e :
    #         messagebox.showerror(title="Convertidor de pdf a word",message=str(e))

    def widgest(self):
        labelFile=Label(self,text="Selecciona el pdf que deseas convertir.")
        labelFile.place(x=30,y=15)
        # Botones
        btnquitar=Button(self, text="Seleccionar Archivo", command=self.archivo, width=35)
        btnquitar.place(x=30, y=90)

        btnquitar=Button(self,text="Salir", command=self.salir, width=35)
        btnquitar.place(x=30, y=130)

if __name__ == "__main__":
    app=ventana()
    app.mainloop()