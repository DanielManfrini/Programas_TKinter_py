#Importadores GLOBAL
from datetime import datetime

#importadores SQL
import pyodbc

#variaveis Tkinter
from tkinter import *
from tkinter import ttk 
from tkinterdnd2 import *
import tkcalendar

#importadores caminhos
import os
import sys

#DEFININDO DIA.
hoje = datetime.today()
dia = hoje.date()
hora = hoje.time()

# Importando telas
import telas

#CAMINHOS
#Caminho da logo.
logo_name = 'logo.png'
icon_name = 'icon.ico'
log_name = str(f"log_{dia}.txt")
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
logo_path = os.path.join(application_path, logo_name)
log_path = os.path.join(application_path, log_name)
icon_path = os.path.join(application_path, icon_name)

# TELA DE LOGIN

def tela_login():
    
    def fazer_login():

        login = campo_login.get()
        senha = campo_senha.get()
        
        dados = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao = pyodbc.connect(dados)
        cursor = conexao.cursor()

        try:
            auth = cursor.execute("SELECT acesso.NivelAcesso, acesso.Nome FROM acesso WHERE acesso.usuario = ? AND acesso.senha = ? ",login,senha).fetchall()
            conexao.commit()
            auth = str(auth[0]).split(",")
            nivel = str(auth[0]).strip("\'").strip("(")
            nome = str(auth[1]).strip("\'").strip(")")
            
            if str(nivel) == "1":

                telas.nivel1(janela_login,login,nome)
            
            elif str(nivel) == "2":
                
                telas.nivel2(janela_login,login,nome)

            elif str(nivel) == "3":
                
                telas.nivel3(janela_login,login,nome)
            
            else:
        
                def fechar():
                    tela_erro.destroy()

                tela_erro = TkinterDnD.Tk()
                tela_erro.title("ERRO") #titulo
                tela_erro.geometry("200x100")
                tela_erro.config(bg = "white")

                Label(tela_erro,bg="white",highlightthickness=1,highlightbackground="black",width=25,height=5).place(x=9,y=12)

                texto_erro = Label(tela_erro, text="Login incorreto.",bg="white")
                texto_erro.place(x=60,y=20)

                botao_erro = Button(tela_erro, text="fechar", command=fechar,fg="navy blue")
                botao_erro.place(x=80,y=60)

                tela_erro.mainloop()
                
        except pyodbc.Error as e:
            print(e)

    janela_login = TkinterDnD.Tk() #criando janela
    janela_login.title("Associa Provisório") #titulo
    janela_login.geometry("300x150") #quadro
    janela_login.config(bg = "white")
    janela_login.iconbitmap(icon_path)
    janela_login.resizable(0,0)

    Label(janela_login,bg="white",highlightthickness=1,highlightbackground="black",width=39,height=8).place(x=10,y=15)
    Label(janela_login,text="AUTENTICAÇÃO",bg="white").place(x=30,y=5)

    Label(janela_login,text="Usuário:",bg="white").place(x=30,y=40)
    campo_login = Entry(janela_login,)
    campo_login.place(x=80,y=40)

    Label(janela_login,text="Senha:",bg="white").place(x=30,y=70)
    campo_senha = Entry(janela_login,show="*")
    campo_senha.place(x=80,y=70)

    auth = Button(janela_login,text="Login",command=fazer_login,fg="navy blue").place(x=30,y=105)

    janela_login.mainloop()

if __name__ == '__main__':
    tela_login()