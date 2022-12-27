# importadores
from tkinter import *
import time

# VAMOS CRIAR Uma FUNÇÂO PARA CHAMAR A TELA DE AVISO SEMPRE QUE PRECISARMOS
def aviso(texto,tipo,timeout=1500):
 
    tela_aviso = Tk()
    tela_aviso.title(tipo) #titulo
    tela_aviso.geometry("200x100")
    tela_aviso.resizable(0,0)
    tela_aviso.config(bg="white")
    tela_aviso.after(timeout,tela_aviso.destroy)

    Label(tela_aviso,bg="white",highlightthickness=1,highlightbackground="black",width=25,height=5).place(x=10,y=10)
    Label(tela_aviso,text=tipo,bg="white",fg="navy blue").place(x=25)

    texto_erro = Label(tela_aviso,fg="navy blue",bg="white", text=texto)
    texto_erro.pack(pady=35)

    tela_aviso.mainloop()