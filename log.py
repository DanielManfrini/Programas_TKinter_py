#Importadores GLOBAL
from datetime import datetime

#importadores SQL
import pyodbc

#variaveis Tkinter
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import tkcalendar

#DEFININDO DIA.
import arrow
hoje = datetime.today()
dia = hoje.date()
hora = hoje.time()

# Importando telas
import telas

def inicia_log():
    def buscar_log():
        texto_log.delete(1.0,'end')
        dados_log = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao = pyodbc.connect(dados_log)
        cursor = conexao.cursor()
        log = cursor.execute("SELECT * FROM RegistroAcoes ORDER BY Data").fetchall()
        conexao.commit()
        for registro in log:
            print(registro)
            data_hora = str(registro[1]).strip("\'")
            data_hora = data_hora[0:16]
            registro = str(registro).split(",")
            usuario = str(registro[8]).strip("\'")
            try:
                acao = str(registro[10]).replace("'","").strip(")")
            except:
                acao = str(registro[9]).replace("'","").strip(")")
            registro = str(data_hora) + ": " + str(usuario) +" "+ str(acao)+ "\n"
            texto_log.insert(0.0,registro)#aqui inserimos o texto dentro do widget Text.
        conexao.close()    
        
    def buscar_log_filtro():
        texto_log.delete(1.0,'end')
        matricula = campo_matricula.get()
        data_inicio = campo_data_inicial_bloqueio.get_date()
        data_fim = campo_data_final_bloqueio.get_date()
        tipo = caixa_tipo.get()
        
        data_inicio = arrow.get(str(data_inicio),'YYYY-MM-DD')
        data_inicio = data_inicio.format('YYYYMMDD')
        data_fim = arrow.get(str(data_fim),'YYYY-MM-DD')
        data_fim = data_fim.format('YYYYMMDD')

        dados_log = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao = pyodbc.connect(dados_log)
        cursor = conexao.cursor()
        
        #Vamos buscar o tipo primeiro.
        id = cursor.execute("SELECT id FROM Motivos WHERE Motivo LIKE ?",tipo).fetchall()
        conexao.commit()
        id = id[0][0]
        print(id)

        
        comando = "SELECT * FROM RegistroAcoes"

        if  tipo != "": # Caso as informações estejam vazias
            if comando.count("WHERE"):
                comando = comando + f" AND Tipo = {id}"
            else:
                comando = comando + f" WHERE Tipo = {id}"
        
        # tipo
        if matricula != "":
            if comando.count("WHERE"):
                comando = comando + f" AND Usuario = {matricula}"
            else:
                comando = comando + f" WHERE Usuario = {matricula}"
            

        # tipo com data
        if data_inicio != data_fim:
            if comando.count("WHERE"):
                comando = comando + f" AND Data BETWEEN '{data_inicio}' AND '{data_fim}'"
            else:
                comando = comando + f" WHERE Data BETWEEN '{data_inicio}' AND '{data_fim}'"

        log = cursor.execute(f"{comando} ORDER BY Data").fetchall()
        for registro in log:
            data_hora = str(registro[1]).strip("\'")
            data_hora = data_hora[0:16]
            registro = str(registro).split(",")
            usuario = str(registro[8]).strip("\'")
            acao = str(registro[10]).replace("'","").strip(")")
            registro = str(data_hora) + ": " + str(usuario) +" "+ str(acao)+ "\n"
            texto_log.insert(0.0,registro)#aqui inserimos o texto dentro do widget Text.
        conexao.commit()

    def exportar():
        pasta = filedialog.askdirectory()
        caminho = f"{pasta}\Export_{dia}.txt"
        texto = texto_log.get(1.0,'end')
        
        with open(caminho,"a") as arquivo:
            arquivo.write(texto)

    # Preencher a caixa de motivos
    # criamos uma lista coms os nomes
    dados_log = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BANCO>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
    )
    conexao = pyodbc.connect(dados_log)
    cursor = conexao.cursor()
    motivos = cursor.execute("SELECT Motivo FROM Motivos").fetchall()
    conexao.commit()
    lista_motivos = [motivos[0] for motivos in motivos]


    # JANELA
    janela_log =Tk()#criamos a janela
    janela_log.title("log de eventos") #titulo
    janela_log.geometry("1070x550") #quadro
    janela_log.config(bg = "white")
    janela_log.resizable(0,0)

    # DECORACAO
    Label(janela_log,bg="White",highlightthickness=1,highlightbackground="black",height=3,width=149).place(x=10,y=20)
    Label(janela_log,text="FILTROS",bg="white").place(x=100,y=10)

    Label(janela_log,bg="White",highlightthickness=1,highlightbackground="black",height=30,width=149).place(x=10,y=85)
    Label(janela_log,text="REGISTO DE AÇÕES",bg="white").place(x=30,y=75)

    # FILTRO
    Label(janela_log,text="Matrícula:",bg="white").place(x=100,y=35)
    campo_matricula = Entry(janela_log,width=10)
    campo_matricula.place(x=160,y=35)

    # CALENDARIO 
    texto_data_inicial_bloqueio = Label(janela_log,text="Data inicial:",bg="white")
    texto_data_inicial_bloqueio.place(x=230,y=35)
    campo_data_inicial_bloqueio = tkcalendar.DateEntry(janela_log,date_pattern='dd/mm/yyy')
    campo_data_inicial_bloqueio.place(x=300,y=35)

    texto_data_final_bloqueio = Label(janela_log,text="Data final:",bg="white")
    texto_data_final_bloqueio.place(x=400,y=35)
    campo_data_final_bloqueio = tkcalendar.DateEntry(janela_log,date_pattern='dd/mm/yyy')
    campo_data_final_bloqueio.place(x=460,y=35)

    tipo = StringVar()
    texto_tipo = Label(janela_log,bg="white",text="Tipo do registro:")
    texto_tipo.place(x=560,y=35)
    caixa_tipo = Combobox(janela_log, textvariable=tipo, values=lista_motivos)
    caixa_tipo.place(x=660,y=35,width=180)


    # BOTOES.
    Label(janela_log,bg="White",highlightthickness=1,highlightbackground="black",height=3,width=28).place(x=857,y=20)
    Button(janela_log,text="Aplicar",fg="navy blue",width=10,command=buscar_log_filtro).place(x=870,y=32)
    Button(janela_log,text="Exportar",fg="navy blue",width=10,command=exportar).place(x=969,y=32)

    # CAMPO DE LOG
    texto_log = Text(janela_log,borderwidth=3,highlightbackground="black",height=27,width=129)#criamos o widget Text.
    texto_log.place(x=15,y=95)#empacotamos o widget Text

    Label(janela_log,bg="White",highlightthickness=1,highlightbackground="black",height=3,width=10).place(x=10,y=20)
    Button(janela_log,text="Atualizar",fg="navy blue",command=buscar_log).place(x=20,y=32)

    janela_log.mainloop()#fazemos o loop da janela.