#Importadores GLOBAL
from datetime import datetime
import re
import log
import aviso
import provisorio
import heads

#importadores SQL
import pyodbc

#variaveis Tkinter
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox 
from tkinterdnd2 import *
import tkcalendar
from functools import partial

#importadores caminhos
import os
import sys

# importadores json
import requests
import json

#DEFININDO DIA.
import arrow
hoje = datetime.today()
dia = hoje.date()
hora = hoje.time()

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

# Funcção para logoff
def logoff(janela):
    janela.destroy()
    provisorio.tela_login()
    

def nivel1(janela_login,login,usuario):

    
    janela_login.destroy()

    #FUNÇÕES DE LOGS E USUARIOS
    def abrir_log():#Função para abrirmos o log para verificação.
        log.inicia_log()

    def trocar_senha(): # Realiza a troca de senha

        def trocar():

            senha = campo_senha.get()
            confirmacao = campo_confirmacao.get()

            dados = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            if senha == confirmacao:
                cursor.execute("UPDATE acesso SET senha = ? WHERE usuario = ?",senha)
                conexao.commit()
                #vamos registrar a alteração no log.
                registro = f"Realizou a troca de senha"
                cursor.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(usuario,"5",registro))
                conexao.commit()
                conexao.close
                tela_senha.destroy()
                
            else:
                aviso.aviso("Senhas não conferem","ERRO",timeout=1500)
            
            # TELA    
        tela_senha = TkinterDnD.Tk() #criando janela
        tela_senha.title("Associa Provisório") #titulo
        tela_senha.geometry("300x150") #quadro
        tela_senha.config(bg = "white")
        tela_senha.iconbitmap(icon_path)
        tela_senha.resizable(0,0)

        Label(tela_senha,bg="white",highlightthickness=1,highlightbackground="black",width=39,height=8).place(x=10,y=15)
        Label(tela_senha,text="TROCA DE SENHA",bg="white").place(x=30,y=5)

        Label(tela_senha,text="Senha:",bg="white").place(x=30,y=40)
        campo_senha = Entry(tela_senha,show="*")
        campo_senha.place(x=140,y=40)

        Label(tela_senha,text="Confirmar senha:",bg="white").place(x=30,y=70)
        campo_confirmacao = Entry(tela_senha,show="*")
        campo_confirmacao.place(x=140,y=70)

        Button(tela_senha,text="Trocar",command=trocar,fg="navy blue").place(x=30,y=105)

        tela_senha.mainloop()()

    def busca(): # Função para buscar o funcionário.

        # o nome atual no campo, mesmo se vazio, serve para em caso de uma nova consulta.
        campo_nome.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_bloqueio.delete(0, 'end')
        campo_motivo.delete(0,'end')

        # Coletamos a matrícula
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.

        if matricula != '':
            bloqueio()
            busca_cartao()
            # dados da conexão
            dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()


            # buscamos o nome do operador
            try:
                dados_usuario = cursor_geral.execute("SELECT Nome,Cartao,Pis,Departamento,Situacao,DataDemissao FROM Funcionarios WHERE Funcionarios.Matricula = ?", matricula ).fetchall()
                conexao_geral.commit()
                print(dados_usuario)
                nome = dados_usuario[0][0]
                cracha = dados_usuario[0][1]                
                pis = dados_usuario[0][2]
                setor = dados_usuario[0][3]
                situacao = dados_usuario[0][4]
                demissao = dados_usuario[0][5]
                
                if cracha == None:
                    cracha= ""

                if situacao == False:
                    situacao = "TRABALHANDO"
                else:
                    situacao = "DEMITIDO"

                if demissao == None:
                    demissao = ""

                campo_nome.insert(END,nome)
                
            except pyodbc.Error as erro:
                print (erro)
            conexao_geral.close()

        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
            
    def limpa(): # basicamente, limpa os campos.
    
        campo_nome.delete(0, 'end')
        campo_matricula.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')

    # FUNÇÕES DE CARTÕES PROVISÓRIOS
    def busca_cartao():
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.
        if matricula != '':
            
            # dados da conexão
            dados = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            # buscamos o nome do operador
            try:
                linha = cursor.execute("SELECT Nome, COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
                conexao.commit()
                print(linha)
                linha = str(linha[0]).split(",")
                pessoa = str(linha[1]).strip(")").strip("\'")

                #vamos buscar o cartão
                cartao = cursor.execute("SELECT NumeroCartao FROM CartoesProvisorios WHERE COD_PESSOA = ?",pessoa).fetchall()
                conexao.commit()
                if str(cartao) != "[]":
                    cartao = str(cartao[0]).split(",")
                    cartao = str(cartao[0]).strip("(").strip("\'")
                    campo_cartao_provisorio.insert(END,cartao)
            except: 
                pass

    def associados(): # Busca todos os provisórios associados
        # limpar a tabela atual

        campo_quadro.delete(*campo_quadro.get_children())

        # Vamos buscar os cartões assoiados a funcionários.
        dados_catraca =(
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_catraca = pyodbc.connect(dados_catraca)
        cursor = conexao_catraca.cursor()

        # Verificar se o cartão existe.
        assoc = cursor.execute("SELECT	CartoesProvisorios.NumeroCartao,Funcionarios.Matricula,Funcionarios.Nome, CONVERT (VARCHAR(10),CartoesProvisorios.Inicio,103), CONVERT (VARCHAR(10),CartoesProvisorios.Fim,103) FROM dbo.Funcionarios, dbo.CartoesProvisorios WHERE Funcionarios.COD_PESSOA = CartoesProvisorios.COD_PESSOA").fetchall()
        conexao_catraca.commit()

        for linha in assoc:
            campo_quadro.insert("",END, values=(linha[0],linha[1],linha[2],linha[3],linha[4])) 

        conexao_catraca.close()

    def associar(): # Associa o cartão.

        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() #pegar cartao.
        nome = campo_nome.get() # Pegar o nome para log.
        inicio_data = data_inicial.get_date() # pegar data inicial.
        fim_data = data_final.get_date() #pegar data final.

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Adicionamos uma série de comparadores para verificar se os dados foream inseridos corretamente.
        if matricula != '':
            if cartao != '':
                if inicio_data != fim_data:
                    # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
                    # data inicial
                    inicio_data = inicio_data.strftime("%Y-%m-%d")
                    
                    # data final
                    fim_data = fim_data.strftime("%Y-%m-%d")
                    
                    # Dados.
                    dados_catraca =(
                        "driver={SQL Server Native Client 11.0};"
                        "server=<SERVIDOR>;"
                        "Database=<BANCO>;"
                        "UID=<USUARIO>;"
                        "PWD=<SENHA>;"
                    )
                    conexao_catraca = pyodbc.connect(dados_catraca)
                    cursor = conexao_catraca.cursor()

                    # Verificar se o cartão existe.
                    existe = cursor.execute("SELECT * FROM Cartoes WHERE NUM_CARTAO = ?",cartao).fetchall()
                    conexao_catraca.commit()
                    # se existir continua, se não, reporta erro.
                    if str(existe) != "[]":

                        # coletar o código da pessoa da base de funcionarios
                        pessoa = cursor.execute("SELECT COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula).fetchall()
                        conexao_catraca.commit()
                        pessoa = str(pessoa[0]).split(",")
                        pessoa = str(pessoa[0]).strip("(").strip(")")
                        
                        try:
                            cursor.execute("INSERT INTO CartoesProvisorios (NumeroCartao,COD_PESSOA,Inicio,Fim) VALUES (?,?,?,?)",(cartao,pessoa,inicio_data,fim_data))
                            conexao_catraca.commit()
                            # Limpando os dados para nova consulta.
                            busca()

                            #vamos registrar a alteração no log.
                            dados_geral = (
                            "driver={SQL Server Native Client 11.0};"
                            "server=<SERVIDOR>;"
                            "Database=<BASE>;"
                            "UID=<USUARIO>;"
                            "PWD=<SENHA>;"
                            )
                            conexao_geral = pyodbc.connect(dados_geral)
                            cursor_geral = conexao_geral.cursor()
                            registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"1",registro))
                            conexao_geral.close()
                            # tela pra avisar a conclusão.
                            
                            busca()
                            aviso.aviso("Associado com sucesso","AVISO",timeout=1500)
                        except pyodbc.Error as erro:
                            print(erro)

                        conexao_catraca.close()
                    
                    else:
                        aviso.aviso("Cartão não existe","ERRO",timeout=1500)                
                else:
                    aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)
            else:
                aviso.aviso("Insira um cartão","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    def desassociar(): # Dessassocia o cartão. 
        
        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() # Pegar cartao.
        nome = campo_nome.get() # Pegar nome para log

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Se o cartão existir vai excluir da tabela, se não, vai reportar erro.
        if matricula != "":
            if cartao != "":

                # Dados.
                dados_catraca =(
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor = conexao_catraca.cursor()
                
                # Tentar realizar a exclusão.
                try:
                    cursor.execute("DELETE FROM CartoesProvisorios WHERE NumeroCartao = ?",cartao)
                    conexao_catraca.commit()

                    #vamos registrar a alteração no log.
                    dados_geral = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BASE>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                    )
                    conexao_geral = pyodbc.connect(dados_geral)
                    cursor_geral = conexao_geral.cursor()
                    registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"2",registro))
                    conexao_geral.close()

                    # Tela Sucesso.
                    busca()
                    aviso.aviso("Dessassociado com sucesso","AVISO",timeout=1500)

                except pyodbc.Error as erro:
                    print(erro)

                conexao_catraca.close()

            else:
                aviso.aviso("Não há um cartão associado","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    # FUNÇÕES DE BLOQUEIO
    def bloqueio(): # Busca se o funcionário está bloqueado ou não.
        
        matricula = campo_matricula.get()
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
            situacao_bloqueio = cursor.execute("SELECT Bloqueado FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
            conexao.commit()

        except:
            pass

        if str(situacao_bloqueio) != "[]":
            situacao_bloqueio = str(situacao_bloqueio[0]).split(',')
            situacao_bloqueio = str(situacao_bloqueio[0]).strip('(').strip('\'')

            if situacao_bloqueio != "False":
                campo_bloqueio.insert(END,"BLOQUEADO")
            else:
                campo_bloqueio.insert(END,"LIBERADO")

            conexao.close()

    def bloquear_funcionario():#Função para bloquear o funcionario.

        matricula = campo_matricula.get()
        motivo = campo_motivo.get()
        inicio_data = campo_data_inicial_bloqueio.get_date()
        fim_data = campo_data_final_bloqueio.get_date()
        nome = campo_nome.get()

        if inicio_data != fim_data:
            # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
            # data inicial
            inicio_data = inicio_data.strftime("%Y-%m-%d")
            
            # data final
            fim_data = fim_data.strftime("%Y-%m-%d")

            dados_catraca = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_catraca = pyodbc.connect(dados_catraca)
            cursor = conexao_catraca.cursor()
            try:
                cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 1, Funcionarios.InicioBloqueio = ?, Funcionarios.FimBloqueio = ? WHERE Funcionarios.Matricula = ?", inicio_data, fim_data, matricula )
                conexao_catraca.commit()
            except:
                pass
            
            conexao_catraca.close()

            dados_geral = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BASE>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()

            try:
                cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = ? WHERE matricula = ?",(motivo,matricula))
                conexao_geral.commit()
                #vamos registrar a alteração no log.
                registro = f"Bloqueou o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"3",registro))
                conexao_geral.commit()
            except:
                pass
            conexao_geral.close()

            busca()
            aviso.aviso("Funcionário bloqueado","AVISO",timeout=1500)
            
        
        else:
            aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)

    def desbloquear_funcionario():#função para desbloquear o funcionario.
        matricula = campo_matricula.get()
        nome = campo_nome.get()

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
            cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 0 WHERE Funcionarios.Matricula = ?", matricula )
            conexao.commit()
            
        except:
            pass

        conexao.close()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        try:
            cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = NULL WHERE matricula = ?",(matricula))
            conexao_geral.commit()
            #vamos registrar a alteração no log.
            registro = f"Desbloqueou o funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"4",registro))
            conexao_geral.commit()
        except:
            pass
        conexao_geral.close()
        busca()
        aviso.aviso("Funcionário desbloqueado","AVISO",timeout=1500)
         
    # CONFIGURAÇÕES DE TELA
    janela = TkinterDnD.Tk() #criando janela
    janela.title("Associa Provisório") #titulo
    janela.geometry("1070x550") #quadro
    janela.config(bg = "white")
    janela.iconbitmap(icon_path)
    janela.resizable(0,0)
    logo = PhotoImage(file=str(logo_path))
    label1 = Label( janela, image = logo, bg="white") 
    label1.place(x=0,y=0)

    comando_argumentos = partial(logoff,janela)

    # FUNCIONARIO
    usuario = str(usuario).split("'")
    usuario = usuario[1]
    Label(janela,bg="white",highlightbackground="black",highlightthickness=1,width=129,height=2).place(x=143,y=5)
    Label(janela,bg="white",text="LOGADO COMO: "+ usuario).place(x=150,y=10)
    Button(janela,fg="navy blue",text="Trocar senha",command=trocar_senha).place(x=870,y=11)
    Button(janela,fg="navy blue",text="Trocar senha",command=comando_argumentos).place(x=960,y=11)

    # DECORAÇÃO
    Label(janela,bg="white",width=65,height=6,highlightthickness=1,highlightbackground="black").place(x=15,y=55)
    Label(janela,text="FUNCIONÁRIO",bg="white").place(x=30,y=45)

    # CAMPOS
    campo_matricula = Entry(janela,)
    campo_matricula.place(x=100,y=70)
    texto_matricula = Label(janela,text="Matricula: ",bg="white")
    texto_matricula.place(x=30,y=70)

    campo_nome = Entry(janela, textvariable="nome",width=55)
    campo_nome.place(x=100,y=110)
    texto_nome = Label(janela,text="Nome: ",bg="white")
    texto_nome.place(x=30,y=110)

    # BOTOES
    botao_busca = Button(janela,text="Buscar",command=busca,fg="navy blue") #realiza a busca.
    botao_busca.place(x=240,y=68)

    botao_limpar = Button(janela,text="limpar",command=limpa,fg="navy blue")
    botao_limpar.place(x=300,y=68)

    # SITUACAO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=165)
    Label(janela,text="SITUAÇÃO",bg="white").place(x=30,y=155)

    # CAMPOS
    campo_bloqueio = Entry(janela, width=20)
    campo_bloqueio.place(x=100,y=185)
    texto_situacao = Label(janela,text="Situação: ",bg="white")
    texto_situacao.place(x=30,y=185)

    campo_motivo = Entry(janela,width=20)
    campo_motivo.place(x=300,y=185)
    texto_motivo = Label(janela,text="Motivo:",bg="white")
    texto_motivo.place(x=240,y=185)

    dica_cartao_provisorio = Label(janela,text="* Para bloquear, selecione o período de bloqueio",fg="red",bg="white")
    dica_cartao_provisorio.place(x=30,y=215)

    # CALENDARIO BLOQUEIO
    texto_data_inicial_bloqueio = Label(janela,text="Data inicial:",bg="white")
    texto_data_inicial_bloqueio.place(x=30,y=250)
    campo_data_inicial_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_inicial_bloqueio.place(x=100,y=250)

    texto_data_final_bloqueio = Label(janela,text="Data final:",bg="white")
    texto_data_final_bloqueio.place(x=207,y=250)
    campo_data_final_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_final_bloqueio.place(x=270,y=250)

    # BOTÕES
    botao_bloquear = Button(janela,text="Bloquear funcionário",command=bloquear_funcionario,fg="navy blue")
    botao_bloquear.place(x=30,y=285)

    botao_desbloquear = Button(janela,text="Desbloquear funcionário",command=desbloquear_funcionario,fg="navy blue")
    botao_desbloquear.place(x=207,y=285)

    # PROVISORIO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=340)
    Label(janela,text="PROVISÓRIO",bg="white").place(x=30,y=330)

    # CAMPOS
    campo_cartao_provisorio = Entry(janela,width=10)
    campo_cartao_provisorio.place(x=100,y=355)
    texto_cartao_provisorio = Label(janela,text="Cartão:",bg="white")
    texto_cartao_provisorio.place(x=30,y=355)
    dica_cartao_provisorio = Label(janela,text="* Para associar, insira os 5 dígitos do cartão e selecione o período",fg="red",bg="white")
    dica_cartao_provisorio.place(x=30,y=390)

    # CALENDARIO PROVISORIO
    texto_data_inicial = Label(janela,text="Data inicial:",bg="white")
    texto_data_inicial.place(x=30,y=420)
    data_inicial = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_inicial.place(x=100,y=420)

    texto_data_final = Label(janela,text="Data final:",bg="white")
    texto_data_final.place(x=207,y=420)
    data_final = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_final.place(x=270,y=420)

    # BOTOES
    botao_associar = Button(janela,text="Associar cartão",command=associar,fg="navy blue")
    botao_associar.place(x=30,y=455)

    botao_dessassociar = Button(janela,text="Desassociar cartão",command=desassociar,fg="navy blue")
    botao_dessassociar.place(x=207,y=455)

    # QUADRO DE USUARIOS.
    # DECORACAO
    Label(janela,bg="white",width=78,height=29,highlightthickness=1,highlightbackground="black").place(x=500,y=55)

    # CAMPOS
    texto_quadro = Label(janela,text="FUNCIONÁRIOS COM CARTÕES PROVISÓRIOS",bg="white")
    texto_quadro.place(x=650,y=45)

    campo_quadro = ttk.Treeview(janela,columns=("c1","c2","c3","c4","c5",),show='headings',height=18)
    campo_quadro.place(x=510,y=65)

    # COLUNAS
    # Cartão
    campo_quadro.column("#1",width=60)
    campo_quadro.heading("#1",text="CARTÃO")
    # Matricula
    campo_quadro.column("#2",width=80)
    campo_quadro.heading("#2",text="MATRÍCULA")
    # Nome
    campo_quadro.column("#3",width=250)
    campo_quadro.heading("#3",text="NOME")
    # Inicio
    campo_quadro.column("#4",width=70)
    campo_quadro.heading("#4",text="INICIO")
    # Fim
    campo_quadro.column("#5",width=70)
    campo_quadro.heading("#5",text="FIM")

    # BOTOES
    botao_quadro = Button(janela,text="Atualizar lista",fg="navy blue",command=associados)
    botao_quadro.place(x=510,y=465)

    # LOG
    botão_log = Button(janela,height=1,width=5, text="log", command=abrir_log,fg="navy blue") #botão para abrir o log.
    botão_log.place(x=997, y=515)

    #VERSÃO
    criador = Label(janela,bg="white",text="Criado por: Daniel Lopes Manfrini")
    criador.place(x=10, y=500)
    versao = Label(janela,bg="white",text="Versão: 2.0 Bloqueio e quadro de cartões.")
    versao.place(x=10, y=520)

    janela.mainloop()

def nivel2(janela_login,login,usuario):
    janela_login.destroy()

    #FUNÇÕES DE LOGS E USUARIOS
    def abrir_log():#Função para abrirmos o log para verificação.
        log.inicia_log()

    def trocar_senha(): # Realiza a troca de senha

        def trocar():

            senha = campo_senha.get()
            confirmacao = campo_confirmacao.get()

            dados = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            if senha == confirmacao:
                cursor.execute("UPDATE acesso SET senha = ? WHERE usuario = ?",senha)
                conexao.commit()
                #vamos registrar a alteração no log.
                registro = f"Realizou a troca de senha"
                cursor.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(usuario,"5",registro))
                conexao.commit()
                conexao.close
                tela_senha.destroy()
                
            else:
                aviso.aviso("Senhas não conferem","ERRO",timeout=1500)
            
            # TELA    
        tela_senha = TkinterDnD.Tk() #criando janela
        tela_senha.title("Associa Provisório") #titulo
        tela_senha.geometry("300x150") #quadro
        tela_senha.config(bg = "white")
        tela_senha.iconbitmap(icon_path)
        tela_senha.resizable(0,0)

        Label(tela_senha,bg="white",highlightthickness=1,highlightbackground="black",width=39,height=8).place(x=10,y=15)
        Label(tela_senha,text="TROCA DE SENHA",bg="white").place(x=30,y=5)

        Label(tela_senha,text="Senha:",bg="white").place(x=30,y=40)
        campo_senha = Entry(tela_senha,show="*")
        campo_senha.place(x=140,y=40)

        Label(tela_senha,text="Confirmar senha:",bg="white").place(x=30,y=70)
        campo_confirmacao = Entry(tela_senha,show="*")
        campo_confirmacao.place(x=140,y=70)

        Button(tela_senha,text="Trocar",command=trocar,fg="navy blue").place(x=30,y=105)

        tela_senha.mainloop()()

    def busca(): # Função para buscar o funcionário.

        # o nome atual no campo, mesmo se vazio, serve para em caso de uma nova consulta.
        campo_nome.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_motivo.delete(0,'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_head_atual.delete(0, 'end')
        campo_head_novo.delete(0, 'end')
        campo_chamado.delete(0, 'end')
        campo_uso.delete(0, 'end')
        caixa_motivo.set(value="")
        campo_bloqueio.delete(0, 'end')
        campo_motivo.delete(0,'end')

        # Coletamos a matrícula
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.

        if matricula != '':
            bloqueio()
            headsets()
            busca_cartao()
            # dados da conexão
            dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()


            # buscamos o nome do operador
            try:
                dados_usuario = cursor_geral.execute("SELECT Nome,Cartao,Pis,Departamento,Situacao,DataDemissao FROM Funcionarios WHERE Funcionarios.Matricula = ?", matricula ).fetchall()
                conexao_geral.commit()
                print(dados_usuario)
                nome = dados_usuario[0][0]
                cracha = dados_usuario[0][1]                
                pis = dados_usuario[0][2]
                setor = dados_usuario[0][3]
                situacao = dados_usuario[0][4]
                demissao = dados_usuario[0][5]
                
                if cracha == None:
                    cracha= ""

                if situacao == False:
                    situacao = "TRABALHANDO"
                else:
                    situacao = "DEMITIDO"

                if demissao == None:
                    demissao = ""

                campo_nome.insert(END,nome)
                
            except pyodbc.Error as erro:
                print (erro)
            conexao_geral.close()

        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
            
    def limpa(): # basicamente, limpa os campos.
    
        campo_nome.delete(0, 'end')
        campo_matricula.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_head_atual.delete(0, 'end')
        campo_head_novo.delete(0, 'end')
        campo_chamado.delete(0, 'end')
        campo_uso.delete(0, 'end')
        caixa_motivo.set(value="")

    # FUNÇÕES DE CARTÕES PROVISÓRIOS
    def busca_cartao():
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.
        if matricula != '':
            
            # dados da conexão
            dados = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            # buscamos o nome do operador
            try:
                linha = cursor.execute("SELECT Nome, COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
                conexao.commit()
                print(linha)
                linha = str(linha[0]).split(",")
                pessoa = str(linha[1]).strip(")").strip("\'")

                #vamos buscar o cartão
                cartao = cursor.execute("SELECT NumeroCartao FROM CartoesProvisorios WHERE COD_PESSOA = ?",pessoa).fetchall()
                conexao.commit()
                if str(cartao) != "[]":
                    cartao = str(cartao[0]).split(",")
                    cartao = str(cartao[0]).strip("(").strip("\'")
                    campo_cartao_provisorio.insert(END,cartao)
            except: 
                pass

    def associados(): # Busca todos os provisórios associados
        # limpar a tabela atual

        campo_quadro.delete(*campo_quadro.get_children())

        # Vamos buscar os cartões assoiados a funcionários.
        dados_catraca =(
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_catraca = pyodbc.connect(dados_catraca)
        cursor = conexao_catraca.cursor()

        # Verificar se o cartão existe.
        assoc = cursor.execute("SELECT	CartoesProvisorios.NumeroCartao,Funcionarios.Matricula,Funcionarios.Nome, CONVERT (VARCHAR(10),CartoesProvisorios.Inicio,103), CONVERT (VARCHAR(10),CartoesProvisorios.Fim,103) FROM dbo.Funcionarios, dbo.CartoesProvisorios WHERE Funcionarios.COD_PESSOA = CartoesProvisorios.COD_PESSOA").fetchall()
        conexao_catraca.commit()

        for linha in assoc:
            campo_quadro.insert("",END, values=(linha[0],linha[1],linha[2],linha[3],linha[4])) 

        conexao_catraca.close()

    def associar(): # Associa o cartão.

        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() #pegar cartao.
        nome = campo_nome.get() # Pegar o nome para log.
        inicio_data = data_inicial.get_date() # pegar data inicial.
        fim_data = data_final.get_date() #pegar data final.

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Adicionamos uma série de comparadores para verificar se os dados foream inseridos corretamente.
        if matricula != '':
            if cartao != '':
                if inicio_data != fim_data:
                    # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
                    # data inicial
                    inicio_data = inicio_data.strftime("%Y-%m-%d")
                    
                    # data final
                    fim_data = fim_data.strftime("%Y-%m-%d")
                    
                    # Dados.
                    dados_catraca =(
                        "driver={SQL Server Native Client 11.0};"
                        "server=<SERVIDOR>;"
                        "Database=<BANCO>;"
                        "UID=<USUARIO>;"
                        "PWD=<SENHA>;"
                    )
                    conexao_catraca = pyodbc.connect(dados_catraca)
                    cursor = conexao_catraca.cursor()

                    # Verificar se o cartão existe.
                    existe = cursor.execute("SELECT * FROM Cartoes WHERE NUM_CARTAO = ?",cartao).fetchall()
                    conexao_catraca.commit()
                    # se existir continua, se não, reporta erro.
                    if str(existe) != "[]":

                        # coletar o código da pessoa da base de funcionarios
                        pessoa = cursor.execute("SELECT COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula).fetchall()
                        conexao_catraca.commit()
                        pessoa = str(pessoa[0]).split(",")
                        pessoa = str(pessoa[0]).strip("(").strip(")")
                        
                        try:
                            cursor.execute("INSERT INTO CartoesProvisorios (NumeroCartao,COD_PESSOA,Inicio,Fim) VALUES (?,?,?,?)",(cartao,pessoa,inicio_data,fim_data))
                            conexao_catraca.commit()
                            # Limpando os dados para nova consulta.
                            busca()

                            #vamos registrar a alteração no log.
                            dados_geral = (
                            "driver={SQL Server Native Client 11.0};"
                            "server=<SERVIDOR>;"
                            "Database=<BASE>;"
                            "UID=<USUARIO>;"
                            "PWD=<SENHA>;"
                            )
                            conexao_geral = pyodbc.connect(dados_geral)
                            cursor_geral = conexao_geral.cursor()
                            registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"1",registro))
                            conexao_geral.close()
                            # tela pra avisar a conclusão.
                            
                            busca()
                            aviso.aviso("Associado com sucesso","AVISO",timeout=1500)
                        except pyodbc.Error as erro:
                            print(erro)

                        conexao_catraca.close()
                    
                    else:
                        aviso.aviso("Cartão não existe","ERRO",timeout=1500)                
                else:
                    aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)
            else:
                aviso.aviso("Insira um cartão","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    def desassociar(): # Dessassocia o cartão. 
        
        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() # Pegar cartao.
        nome = campo_nome.get() # Pegar nome para log

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Se o cartão existir vai excluir da tabela, se não, vai reportar erro.
        if matricula != "":
            if cartao != "":

                # Dados.
                dados_catraca =(
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor = conexao_catraca.cursor()
                
                # Tentar realizar a exclusão.
                try:
                    cursor.execute("DELETE FROM CartoesProvisorios WHERE NumeroCartao = ?",cartao)
                    conexao_catraca.commit()

                    #vamos registrar a alteração no log.
                    dados_geral = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BASE>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                    )
                    conexao_geral = pyodbc.connect(dados_geral)
                    cursor_geral = conexao_geral.cursor()
                    registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"2",registro))
                    conexao_geral.close()

                    # Tela Sucesso.
                    busca()
                    aviso.aviso("Dessassociado com sucesso","AVISO",timeout=1500)

                except pyodbc.Error as erro:
                    print(erro)

                conexao_catraca.close()

            else:
                aviso.aviso("Não há um cartão associado","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    # FUNÇÕES DE BLOQUEIO
    def bloqueio(): # Busca se o funcionário está bloqueado ou não.
        
        matricula = campo_matricula.get()
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
            situacao_bloqueio = cursor.execute("SELECT Bloqueado FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
            conexao.commit()

        except:
            pass

        if str(situacao_bloqueio) != "[]":
            situacao_bloqueio = str(situacao_bloqueio[0]).split(',')
            situacao_bloqueio = str(situacao_bloqueio[0]).strip('(').strip('\'')

            if situacao_bloqueio != "False":
                campo_bloqueio.insert(END,"BLOQUEADO")
            else:
                campo_bloqueio.insert(END,"LIBERADO")

            conexao.close()

    def bloquear_funcionario():#Função para bloquear o funcionario.

        matricula = campo_matricula.get()
        motivo = campo_motivo.get()
        inicio_data = campo_data_inicial_bloqueio.get_date()
        fim_data = campo_data_final_bloqueio.get_date()
        nome = campo_nome.get()

        if inicio_data != fim_data:
            # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
            # data inicial
            inicio_data = inicio_data.strftime("%Y-%m-%d")
            
            # data final
            fim_data = fim_data.strftime("%Y-%m-%d")

            dados_catraca = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_catraca = pyodbc.connect(dados_catraca)
            cursor = conexao_catraca.cursor()
            try:
                cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 1, Funcionarios.InicioBloqueio = ?, Funcionarios.FimBloqueio = ? WHERE Funcionarios.Matricula = ?", inicio_data, fim_data, matricula )
                conexao_catraca.commit()
            except:
                pass
            
            conexao_catraca.close()

            dados_geral = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BASE>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()

            try:
                cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = ? WHERE matricula = ?",(motivo,matricula))
                conexao_geral.commit()
                #vamos registrar a alteração no log.
                registro = f"Bloqueou o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"3",registro))
                conexao_geral.commit()
            except:
                pass
            conexao_geral.close()

            busca()
            aviso.aviso("Funcionário bloqueado","AVISO",timeout=1500)
            
        
        else:
            aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)

    def desbloquear_funcionario():#função para desbloquear o funcionario.
        matricula = campo_matricula.get()
        nome = campo_nome.get()

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
            cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 0 WHERE Funcionarios.Matricula = ?", matricula )
            conexao.commit()
            
        except:
            pass

        conexao.close()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        try:
            cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = NULL WHERE matricula = ?",(matricula))
            conexao_geral.commit()
            #vamos registrar a alteração no log.
            registro = f"Desbloqueou o funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"4",registro))
            conexao_geral.commit()
        except:
            pass
        conexao_geral.close()
        busca()
        aviso.aviso("Funcionário desbloqueado","AVISO",timeout=1500)
        

    # FUNÇÕES PARA MOVIMENTAÇÕES DE HEADSETS
    def headsets(): # realiza a busca dos headsets vinculados

        matricula = campo_matricula.get()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        query_head = cursor_geral.execute("SELECT Headsets.Lacre,Funcionarios.HeadDevolvido FROM Headsets,Funcionarios WHERE Headsets.Id = Funcionarios.Id_headset AND Funcionarios.Matricula = ?",matricula).fetchall()
        conexao_geral.commit()
        print (query_head)
        if str(query_head) != "[]":
            lacre = query_head[0][0]
            situacao = query_head[0][1]
            print (situacao)
            if situacao == False:
                situacao = "EM USO"
                campo_uso.insert(END,situacao)
            else:
                situacao = "DEVOLVIDO"
                campo_uso.insert(END,situacao)

            campo_head_atual.insert(END,lacre)

        conexao_geral.close()

    def associa_heads(): # Realiza a troca do headset
        
        head_antigo = campo_head_atual.get()
        head_novo = campo_head_novo.get()
        chamado = campo_chamado.get()
        motivo = caixa_motivo.get()
        posse = campo_matricula.get()
        nome = campo_nome.get()

        # definindo em bytes
        if str(motivo) == "INTEGRAÇÃO":
            motivo = 1
        elif str(motivo) == "TROCA":
            motivo = 2
        else:
            motivo = 0
        print(motivo)

        # Primeiro vamos buscar se o headset exite.
        dados_geral = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BASE>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()        
        existe = cursor_geral.execute("SELECT Lacre FROM Headsets WHERE Lacre = ?",head_novo).fetchall()
        conexao_geral.commit()
        print (existe)
        if str(existe) == "[]":

            if motivo == 1:
                print ("veio")
                try:
                    cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse) VALUES (?,?,)",(head_novo,posse))
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                except pyodbc.Error as e:
                    print(e)
                #vamos registrar a alteração no log.          
                registro = f"Entregou o HEAD: {head_novo} ao funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,6,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 2:
                # Primeiro verificamos se o chamado foi digitado.
                if chamado != '':
                    # ATUALIZAMOS O HEAD ANTIGO COM OS DADOS DO FUNCIONARIO
                    try:
                        cursor_geral.execute("UPDATE Headsets SET EmPosse=NULL,UltPosse = ? WHERE Lacre = ?",posse,head_antigo)
                        conexao_geral.commit()
                    except:
                        pass
                    # AGORA RALIZAMOS O INSERT DO HEAD NOVO
                    cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse) VALUES (?,?)",(head_novo,posse))
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_antigo,Id_headset_novo,Id_funcionario,id_tecnico,Chamado) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),(SELECT id FROM Headsets WHERE Lacre=?),?,?,?)",(motivo,head_antigo,head_novo,posse,login,chamado))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                    #vamos registrar a alteração no log.          
                    registro = f"Trocou o HEAD: {head_antigo} por {head_novo} para o funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,7,registro))
                    conexao_geral.commit()
                    conexao_geral.close()
                else:
                    aviso.aviso(texto="Insira o chamado",tipo="ERRO",timeout=1500)
        else:
            if motivo == 1:
                print ("errado")
                cursor_geral.execute("UPDATE Headsets SET EmPosse=?,Estoque=0 WHERE Lacre = ?",posse,head_novo)
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Entregou o HEAD: {head_novo} ao funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,6,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 2:
                if chamado != '':
                    try:
                        cursor_geral.execute("UPDATE Headsets SET EmPosse=NULL,UltPosse = ? WHERE Lacre = ?",posse,head_antigo)
                        conexao_geral.commit()
                    except:
                        pass
                    cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse=?,Estoque=0,Manutencao=0,Inativo=0 WHERE Lacre = ?",posse,head_novo)
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_antigo,Id_headset_novo,Id_funcionario,id_tecnico,Chamado) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),(SELECT id FROM Headsets WHERE Lacre=?),?,?,?)",(motivo,head_antigo,head_novo,posse,login,chamado))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                    #vamos registrar a alteração no log.          
                    registro = f"Trocou o HEAD: {head_antigo} por {head_novo} para o funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,7,registro))
                    conexao_geral.commit()
                    conexao_geral.close()
                else:
                    aviso.aviso(texto="Insira o chamado",tipo="ERRO",timeout=1500)
        busca()
        aviso.aviso("Headset associado com sucesso","AVISO",timeout=2000)
        
    
    # CONFIGURAÇÕES DE TELA
    janela = TkinterDnD.Tk() #criando janela
    janela.title("Associa Provisório") #titulo
    janela.geometry("1350x550") #quadro
    janela.config(bg = "white")
    janela.iconbitmap(icon_path)
    janela.resizable(0,0)
    logo = PhotoImage(file=str(logo_path))
    label1 = Label( janela, image = logo, bg="white") 
    label1.place(x=0,y=0)

    comando_argumentos = partial(logoff,janela)

    # FUNCIONARIO
    usuario = str(usuario).split("'")
    usuario = usuario[1]
    Label(janela,bg="white",highlightbackground="black",highlightthickness=1,width=169,height=2).place(x=143,y=5)
    Label(janela,bg="white",text="LOGADO COMO: "+ usuario).place(x=150,y=15)
    Button(janela,fg="navy blue",text="Trocar senha",command=trocar_senha).place(x=1150,y=11)
    Button(janela,fg="navy blue",text="Logoff",command=comando_argumentos).place(x=1240,y=13)

    # DECORAÇÃO
    Label(janela,bg="white",width=65,height=6,highlightthickness=1,highlightbackground="black").place(x=15,y=55)
    Label(janela,text="FUNCIONÁRIO",bg="white").place(x=30,y=45)

    # CAMPOS
    campo_matricula = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_matricula.place(x=100,y=70)
    texto_matricula = Label(janela,text="Matricula: ",bg="white")
    texto_matricula.place(x=30,y=70)

    campo_nome = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=55)
    campo_nome.place(x=100,y=110)
    texto_nome = Label(janela,text="Nome: ",bg="white")
    texto_nome.place(x=30,y=110)

    # BOTOES
    Button(janela,text="Buscar",command=busca,fg="navy blue").place(x=240,y=68) #realiza a busca.

    botao_limpar = Button(janela,text="limpar",command=limpa,fg="navy blue")
    botao_limpar.place(x=300,y=68)

    # SITUACAO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=168)
    Label(janela,text="SITUAÇÃO",bg="white").place(x=30,y=158)

    # CAMPOS
    campo_bloqueio = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=20)
    campo_bloqueio.place(x=100,y=188)
    texto_situacao = Label(janela,text="Situação: ",bg="white")
    texto_situacao.place(x=30,y=188)

    campo_motivo = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=20)
    campo_motivo.place(x=300,y=188)
    texto_motivo = Label(janela,text="Motivo:",bg="white")
    texto_motivo.place(x=240,y=188)

    dica_cartao_provisorio = Label(janela,text="* Para bloquear, selecione o período de bloqueio",fg="red",bg="white")
    dica_cartao_provisorio.place(x=30,y=218)

    # CALENDARIO BLOQUEIO
    texto_data_inicial_bloqueio = Label(janela,text="Data inicial:",bg="white")
    texto_data_inicial_bloqueio.place(x=30,y=248)
    campo_data_inicial_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_inicial_bloqueio.place(x=100,y=248)

    texto_data_final_bloqueio = Label(janela,text="Data final:",bg="white")
    texto_data_final_bloqueio.place(x=207,y=248)
    campo_data_final_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_final_bloqueio.place(x=270,y=248)

    # BOTÕES
    botao_bloquear = Button(janela,text="Bloquear funcionário",command=bloquear_funcionario,fg="navy blue")
    botao_bloquear.place(x=30,y=290)

    botao_desbloquear = Button(janela,text="Desbloquear funcionário",command=desbloquear_funcionario,fg="navy blue")
    botao_desbloquear.place(x=207,y=290)

    
    # PROVISORIO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=340)
    Label(janela,text="PROVISÓRIO",bg="white").place(x=30,y=330)

    # CAMPOS
    campo_cartao_provisorio = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=10)
    campo_cartao_provisorio.place(x=100,y=355)
    texto_cartao_provisorio = Label(janela,text="Cartão:",bg="white")
    texto_cartao_provisorio.place(x=30,y=355)
    dica_cartao_provisorio = Label(janela,text="* Para associar, insira os 5 dígitos do cartão e selecione o período",fg="red",bg="white")
    dica_cartao_provisorio.place(x=30,y=390)

    # CALENDARIO PROVISORIO
    texto_data_inicial = Label(janela,text="Data inicial:",bg="white")
    texto_data_inicial.place(x=30,y=420)
    data_inicial = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_inicial.place(x=100,y=420)

    texto_data_final = Label(janela,text="Data final:",bg="white")
    texto_data_final.place(x=207,y=420)
    data_final = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_final.place(x=270,y=420)

    # BOTOES
    botao_associar = Button(janela,text="Associar cartão",command=associar,fg="navy blue")
    botao_associar.place(x=30,y=455)

    botao_dessassociar = Button(janela,text="Desassociar cartão",command=desassociar,fg="navy blue")
    botao_dessassociar.place(x=207,y=455)

    # HEADSET
    Label(janela,bg="white",width=40,height=6,highlightbackground="black",highlightthickness=1).place(x=485,y=55)
    Label(janela,text="HEADSET",bg="white").place(x=500,y=45)

    # CAMPOS
    Label(janela,bg="white",text="Headset atual:").place(x=500,y=75)
    campo_head_atual = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_head_atual.place(x=600,y=75)

    Label(janela,bg="white",text="Situação:").place(x=500,y=110)
    campo_uso = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_uso.place(x=600,y=110)

    # TROCA/ENTREGA
    Label(janela,bg="white",width=40,height=10,highlightthickness=1,highlightbackground="black").place(x=485,y=168)
    Label(janela,text="ENTREGA/TROCA/BAIXA",bg="white").place(x=500,y=158)

    Label(janela,bg="white",text="Headset novo:").place(x=500,y=185)
    campo_head_novo = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_head_novo.place(x=600,y=185)

    Label(janela,bg="white",text="Motivo:").place(x=500,y=215)
    motivo = StringVar()
    caixa_motivo = Combobox(janela, textvariable=motivo, values=("INTEGRAÇÃO"))
    caixa_motivo.place(x=600,y=215,width=100)

    Label(janela,bg="white",fg="red",text="* Para troca informe o chamado.",).place(x=500,y=240)
    Label(janela,bg="white",text="Chamado:").place(x=500,y=260)
    campo_chamado = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_chamado.place(x=600,y=260)

    # BOTÕES
    Button(janela,fg="navy blue",text='Atualizar headset',command=associa_heads).place(x=500,y=294)

    # QUADRO DE USUARIOS
    # DECORACAO
    Label(janela,bg="white",width=78,height=29,highlightthickness=1,highlightbackground="black").place(x=780,y=55)

    # CAMPOS
    texto_quadro = Label(janela,text="FUNCIONÁRIOS COM CARTÕES PROVISÓRIOS",bg="white")
    texto_quadro.place(x=930,y=45)

    campo_quadro = ttk.Treeview(janela,columns=("c1","c2","c3","c4","c5",),show='headings',height=18)
    campo_quadro.place(x=790,y=65)

    # COLUNAS
    # Cartão
    campo_quadro.column("#1",width=60)
    campo_quadro.heading("#1",text="CARTÃO")
    # Matricula
    campo_quadro.column("#2",width=80)
    campo_quadro.heading("#2",text="MATRÍCULA")
    # Nome
    campo_quadro.column("#3",width=250)
    campo_quadro.heading("#3",text="NOME")
    # Inicio
    campo_quadro.column("#4",width=70)
    campo_quadro.heading("#4",text="INICIO")
    # Fim
    campo_quadro.column("#5",width=70)
    campo_quadro.heading("#5",text="FIM")

    # BOTOES
    botao_quadro = Button(janela,text="Atualizar lista",fg="navy blue",command=associados)
    botao_quadro.place(x=790,y=465)

    # LOG
    botão_log = Button(janela,height=1,width=5, text="log", command=abrir_log,fg="navy blue") #botão para abrir o log.
    botão_log.place(x=1297, y=515)

    #VERSÃO
    criador = Label(janela,bg="white",text="Criado por: Daniel Lopes Manfrini")
    criador.place(x=10, y=500)
    versao = Label(janela,bg="white",text="Versão: 2.0 Bloqueio e quadro de cartões.")
    versao.place(x=10, y=520)

    janela.mainloop()

def nivel3(janela_login,login,usuario):
    janela_login.destroy()

    # VAMOS PREENCHER AS OPÇÔES DA COMBOBOX de setor
    dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
    conexao_geral = pyodbc.connect(dados_geral)
    cursor_geral = conexao_geral.cursor()

    # SETORES
    setores = cursor_geral.execute("SELECT Departamento FROM Departamentos").fetchall()
    conexao_geral.commit()
    lista_setor = [setor[0] for setor in setores]

    #FUNÇÕES DE LOGS E USUARIOS
    def abrir_log():#Função para abrirmos o log para verificação.
        log.inicia_log()

    def trocar_senha(): # Realiza a troca de senha

        def trocar():

            senha = campo_senha.get()
            confirmacao = campo_confirmacao.get()

            dados = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            if senha == confirmacao:
                cursor.execute("UPDATE acesso SET senha = ? WHERE usuario = ?",senha)
                conexao.commit()
                #vamos registrar a alteração no log.
                registro = f"Realizou a troca de senha"
                cursor.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(usuario,"5",registro))
                conexao.commit()
                conexao.close
                tela_senha.destroy()
                
            else:
                aviso.aviso("Senhas não conferem","ERRO",timeout=1500)
            
            # TELA    
        tela_senha = TkinterDnD.Tk() #criando janela
        tela_senha.title("Associa Provisório") #titulo
        tela_senha.geometry("300x150") #quadro
        tela_senha.config(bg = "white")
        tela_senha.iconbitmap(icon_path)
        tela_senha.resizable(0,0)

        Label(tela_senha,bg="white",highlightthickness=1,highlightbackground="black",width=39,height=8).place(x=10,y=15)
        Label(tela_senha,text="TROCA DE SENHA",bg="white").place(x=30,y=5)

        Label(tela_senha,text="Senha:",bg="white").place(x=30,y=40)
        campo_senha = Entry(tela_senha,show="*")
        campo_senha.place(x=140,y=40)

        Label(tela_senha,text="Confirmar senha:",bg="white").place(x=30,y=70)
        campo_confirmacao = Entry(tela_senha,show="*")
        campo_confirmacao.place(x=140,y=70)

        Button(tela_senha,text="Trocar",command=trocar,fg="navy blue").place(x=30,y=105)

        tela_senha.mainloop()

    def busca(): # Função para buscar o funcionário.

        # o nome atual no campo, mesmo se vazio, serve para em caso de uma nova consulta.
        campo_nome.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_bloqueio.delete(0, 'end')
        campo_motivo.delete(0,'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_setor.delete(0, 'end')
        campo_cartao.delete(0, 'end')
        campo_pis.delete(0, 'end')
        campo_bloqueio.delete(0, 'end')
        campo_head_atual.delete(0, 'end')
        campo_head_novo.delete(0, 'end')
        campo_chamado.delete(0, 'end')
        campo_uso.delete(0, 'end')
        campo_situacao.delete(0, 'end')
        campo_demissao.delete(0, 'end')
        caixa_motivo.set(value="")
        check_desconto.deselect()

        # Coletamos a matrícula
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.

        if matricula != '':
            bloqueio()
            headsets()
            busca_cartao()
            # dados da conexão
            dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()


            # buscamos o nome do operador
            try:
                dados_usuario = cursor_geral.execute("SELECT Nome,Cartao,Pis,Departamento,Situacao,DataDemissao FROM Funcionarios WHERE Funcionarios.Matricula = ?", matricula ).fetchall()
                conexao_geral.commit()
                print(dados_usuario)
                nome = dados_usuario[0][0]
                cracha = dados_usuario[0][1]                
                pis = dados_usuario[0][2]
                setor = dados_usuario[0][3]
                situacao = dados_usuario[0][4]
                demissao = dados_usuario[0][5]
                
                if cracha == None:
                    cracha= ""

                if situacao == False:
                    situacao = "TRABALHANDO"
                else:
                    situacao = "DEMITIDO"

                if demissao == None:
                    demissao = ""

                campo_nome.insert(END,nome)
                campo_cartao.insert(END,cracha) 
                campo_pis.insert(END,pis)
                campo_setor.set(setor)
                campo_situacao.insert(END,situacao)
                campo_demissao.insert(END,demissao)
                
            except pyodbc.Error as erro:
                print (erro)
            conexao_geral.close()

        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
            
    def limpa(): # basicamente, limpa os campos.
    
        campo_nome.delete(0, 'end')
        campo_matricula.delete(0, 'end')
        campo_cartao_provisorio.delete(0, 'end')
        campo_setor.delete(0, 'end')
        campo_cartao.delete(0, 'end')
        campo_pis.delete(0, 'end')
        campo_bloqueio.delete(0, 'end')
        campo_head_atual.delete(0, 'end')
        campo_head_novo.delete(0, 'end')
        campo_chamado.delete(0, 'end')
        campo_uso.delete(0, 'end')
        campo_situacao.delete(0, 'end')
        campo_demissao.delete(0, 'end')
        caixa_motivo.set(value="")
        check_desconto.deselect()

    # FUNÇÕES DE CARTÕES PROVISÓRIOS
    def busca_cartao():
        matricula = campo_matricula.get()

        # Buscamos a situação chamando uma função.
        if matricula != '':
            
            # dados da conexão
            dados = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao = pyodbc.connect(dados)
            cursor = conexao.cursor()

            # buscamos o nome do operador
            try:
                linha = cursor.execute("SELECT Nome, COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
                conexao.commit()
                print(linha)
                linha = str(linha[0]).split(",")
                pessoa = str(linha[1]).strip(")").strip("\'")

                #vamos buscar o cartão
                cartao = cursor.execute("SELECT NumeroCartao FROM CartoesProvisorios WHERE COD_PESSOA = ?",pessoa).fetchall()
                conexao.commit()
                if str(cartao) != "[]":
                    cartao = str(cartao[0]).split(",")
                    cartao = str(cartao[0]).strip("(").strip("\'")
                    campo_cartao_provisorio.insert(END,cartao)
            except: 
                pass

    def associados(): # Busca todos os provisórios associados
        # limpar a tabela atual

        campo_quadro.delete(*campo_quadro.get_children())

        # Vamos buscar os cartões assoiados a funcionários.
        dados_catraca =(
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_catraca = pyodbc.connect(dados_catraca)
        cursor = conexao_catraca.cursor()

        # Verificar se o cartão existe.
        assoc = cursor.execute("SELECT	CartoesProvisorios.NumeroCartao,Funcionarios.Matricula,Funcionarios.Nome, CONVERT (VARCHAR(10),CartoesProvisorios.Inicio,103), CONVERT (VARCHAR(10),CartoesProvisorios.Fim,103) FROM dbo.Funcionarios, dbo.CartoesProvisorios WHERE Funcionarios.COD_PESSOA = CartoesProvisorios.COD_PESSOA").fetchall()
        conexao_catraca.commit()

        for linha in assoc:
            campo_quadro.insert("",END, values=(linha[0],linha[1],linha[2],linha[3],linha[4])) 

        conexao_catraca.close()

    def associar(): # Associa o cartão.

        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() #pegar cartao.
        nome = campo_nome.get() # Pegar o nome para log.
        inicio_data = data_inicial.get_date() # pegar data inicial.
        fim_data = data_final.get_date() #pegar data final.

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Adicionamos uma série de comparadores para verificar se os dados foream inseridos corretamente.
        if matricula != '':
            if cartao != '':
                if inicio_data != fim_data:
                    # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
                    # data inicial
                    inicio_data = inicio_data.strftime("%Y-%m-%d")
                    
                    # data final
                    fim_data = fim_data.strftime("%Y-%m-%d")
                    
                    # Dados.
                    dados_catraca =(
                        "driver={SQL Server Native Client 11.0};"
                        "server=<SERVIDOR>;"
                        "Database=<BANCO>;"
                        "UID=<USUARIO>;"
                        "PWD=<SENHA>;"
                    )
                    conexao_catraca = pyodbc.connect(dados_catraca)
                    cursor = conexao_catraca.cursor()

                    # Verificar se o cartão existe.
                    existe = cursor.execute("SELECT * FROM Cartoes WHERE NUM_CARTAO = ?",cartao).fetchall()
                    conexao_catraca.commit()
                    # se existir continua, se não, reporta erro.
                    if str(existe) != "[]":

                        # coletar o código da pessoa da base de funcionarios
                        pessoa = cursor.execute("SELECT COD_PESSOA FROM Funcionarios WHERE Matricula = ?", matricula).fetchall()
                        conexao_catraca.commit()
                        pessoa = str(pessoa[0]).split(",")
                        pessoa = str(pessoa[0]).strip("(").strip(")")
                        
                        try:
                            cursor.execute("INSERT INTO CartoesProvisorios (NumeroCartao,COD_PESSOA,Inicio,Fim) VALUES (?,?,?,?)",(cartao,pessoa,inicio_data,fim_data))
                            conexao_catraca.commit()
                            # Limpando os dados para nova consulta.
                            busca()

                            #vamos registrar a alteração no log.
                            dados_geral = (
                            "driver={SQL Server Native Client 11.0};"
                            "server=<SERVIDOR>;"
                            "Database=<BASE>;"
                            "UID=<USUARIO>;"
                            "PWD=<SENHA>;"
                            )
                            conexao_geral = pyodbc.connect(dados_geral)
                            cursor_geral = conexao_geral.cursor()
                            registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"1",registro))
                            conexao_geral.close()
                            # tela pra avisar a conclusão.
                            
                            busca()
                            aviso.aviso("Associado com sucesso","AVISO",timeout=1500)
                        except pyodbc.Error as erro:
                            print(erro)

                        conexao_catraca.close()
                    
                    else:
                        aviso.aviso("Cartão não existe","ERRO",timeout=1500)                
                else:
                    aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)
            else:
                aviso.aviso("Insira um cartão","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    def desassociar(): # Dessassocia o cartão. 
        
        # Coletar os dados
        matricula = campo_matricula.get() # Pegar matricula.
        cartao = campo_cartao_provisorio.get() # Pegar cartao.
        nome = campo_nome.get() # Pegar nome para log

        # Remover os zeros a esquerda.
        cartao_com_zeros = [cartao]
        cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
        cartao = cartao_sem_zeros[0]

        # Se o cartão existir vai excluir da tabela, se não, vai reportar erro.
        if matricula != "":
            if cartao != "":

                # Dados.
                dados_catraca =(
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor = conexao_catraca.cursor()
                
                # Tentar realizar a exclusão.
                try:
                    cursor.execute("DELETE FROM CartoesProvisorios WHERE NumeroCartao = ?",cartao)
                    conexao_catraca.commit()

                    #vamos registrar a alteração no log.
                    dados_geral = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BASE>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                    )
                    conexao_geral = pyodbc.connect(dados_geral)
                    cursor_geral = conexao_geral.cursor()
                    registro = f"Associou o cartão: {cartao} ao funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"2",registro))
                    conexao_geral.close()

                    # Tela Sucesso.
                    busca()
                    aviso.aviso("Dessassociado com sucesso","AVISO",timeout=1500)

                except pyodbc.Error as erro:
                    print(erro)

                conexao_catraca.close()

            else:
                aviso.aviso("Não há um cartão associado","ERRO",timeout=1500)
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    # FUNÇÕES DE BLOQUEIO
    def bloqueio(): # Busca se o funcionário está bloqueado ou não.
        
        matricula = campo_matricula.get()
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
            situacao_bloqueio = cursor.execute("SELECT Bloqueado FROM Funcionarios WHERE Matricula = ?", matricula ).fetchall()
            conexao.commit()

        except:
            pass

        if str(situacao_bloqueio) != "[]":
            situacao_bloqueio = str(situacao_bloqueio[0]).split(',')
            situacao_bloqueio = str(situacao_bloqueio[0]).strip('(').strip('\'')

            if situacao_bloqueio != "False":
                campo_bloqueio.insert(END,"BLOQUEADO")
            else:
                campo_bloqueio.insert(END,"LIBERADO")

            conexao.close()

    def bloquear_funcionario():#Função para bloquear o funcionario.

        matricula = campo_matricula.get()
        motivo = campo_motivo.get()
        inicio_data = campo_data_inicial_bloqueio.get_date()
        fim_data = campo_data_final_bloqueio.get_date()
        nome = campo_nome.get()

        if inicio_data != fim_data:
            # Realizar as conversões das datas para a aceita pela database que é dd-mm-aaa hh:mm:ss.mimimimi
            # data inicial
            inicio_data = inicio_data.strftime("%Y-%m-%d")
            
            # data final
            fim_data = fim_data.strftime("%Y-%m-%d")

            dados_catraca = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_catraca = pyodbc.connect(dados_catraca)
            cursor = conexao_catraca.cursor()
            try:
                cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 1, Funcionarios.InicioBloqueio = ?, Funcionarios.FimBloqueio = ? WHERE Funcionarios.Matricula = ?", inicio_data, fim_data, matricula )
                conexao_catraca.commit()
            except:
                pass
            
            conexao_catraca.close()

            dados_geral = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BASE>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
            conexao_geral = pyodbc.connect(dados_geral)
            cursor_geral = conexao_geral.cursor()

            try:
                cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = ? WHERE matricula = ?",(motivo,matricula))
                conexao_geral.commit()
                #vamos registrar a alteração no log.
                registro = f"Bloqueou o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"3",registro))
                conexao_geral.commit()
            except:
                pass
            conexao_geral.close()

            busca()
            aviso.aviso("Funcionário bloqueado","AVISO",timeout=1500)
            
        
        else:
            aviso.aviso("Os períodos não podem ser iguais","ERRO",timeout=1500)

    def desbloquear_funcionario():#função para desbloquear o funcionario.
        matricula = campo_matricula.get()
        nome = campo_nome.get()

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
            cursor.execute("UPDATE Funcionarios SET Funcionarios.Bloqueado = 0 WHERE Funcionarios.Matricula = ?", matricula )
            conexao.commit()
            
        except:
            pass

        conexao.close()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        try:
            cursor_geral.execute("UPDATE Funcionarios SET MotivoBloqueio = NULL WHERE matricula = ?",(matricula))
            conexao_geral.commit()
            #vamos registrar a alteração no log.
            registro = f"Desbloqueou o funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,"4",registro))
            conexao_geral.commit()
        except:
            pass
        conexao_geral.close()
        busca()
        aviso.aviso("Funcionário desbloqueado","AVISO",timeout=1500)
        
    # FUNÇÕES PARA MOVIMENTAÇÕES DE HEADSETS
    def headsets(): # realiza a busca dos headsets vinculados

        matricula = campo_matricula.get()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        query_head = cursor_geral.execute("SELECT T1.Lacre AS lacre_head,Funcionarios.HeadDevolvido,T2.Lacre AS Lacre_emprestimo FROM Funcionarios LEFT JOIN Headsets AS T1 ON T1.Id = Funcionarios.Id_headset LEFT JOIN Headsets AS T2 On T2.Id = Funcionarios.Id_head_emprestimo WHERE Funcionarios.Matricula = ?",matricula).fetchall()
        conexao_geral.commit()
        print (query_head) 
        lacre = query_head[0][0]
        situacao = query_head[0][1]
        emprestimo = query_head[0][2]

        if lacre == None and emprestimo == None:

            botao_associa_head['state'] = NORMAL
            botao_baixa_head['state'] = DISABLED
            botao_baixa_emprestimo['state'] = DISABLED
                    
        else:

            if emprestimo == None:
                if situacao == False:
                    situacao = "EM USO"
                    campo_uso.insert(END,situacao)
                else:
                    situacao = "DEVOLVIDO"
                    campo_uso.insert(END,situacao)
                botao_baixa_emprestimo['state'] = DISABLED
                botao_associa_head['state'] = NORMAL
                botao_baixa_head['state'] = NORMAL
                campo_head_atual.insert(END,lacre)
            else:
                situacao = "EMPRESTADO"
                campo_uso.insert(END,situacao)
                campo_head_atual.insert(END,emprestimo)
                botao_associa_head['state'] = DISABLED
                botao_baixa_head['state'] = DISABLED
                botao_baixa_emprestimo['state'] = NORMAL
            

        conexao_geral.close()

    def abrir_heads():
        heads.tela_head()

    def associa_heads(): # Realiza a troca do headset
        
        head_antigo = campo_head_atual.get()
        head_novo = campo_head_novo.get()
        chamado = campo_chamado.get()
        motivo = caixa_motivo.get()
        posse = campo_matricula.get()
        nome = campo_nome.get()
        realizar_desconto = variavel_desconto.get()

        # definindo em bytes
        if str(motivo) == "INTEGRAÇÃO":
            motivo = 1
        elif str(motivo) == "TROCA":
            motivo = 2
        elif str(motivo) == "CORREÇÃO":
            motivo = 3
        elif str(motivo) == "EMPRÉSTIMO":
            motivo = 4
        else:
            motivo = 0
        print(motivo)

        # Primeiro vamos buscar se o headset exite.
        dados_geral = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BASE>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()        
        existe = cursor_geral.execute("SELECT Lacre FROM Headsets WHERE Lacre = ?",head_novo).fetchall()
        conexao_geral.commit()
        print (existe)
        if str(existe) == "[]":

            if motivo == 1:
                print ("veio")
                try:
                    cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse) VALUES (?,?)",(head_novo,posse))
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                except pyodbc.Error as e:
                    print(e)
                #vamos registrar a alteração no log.          
                registro = f"Entregou o HEAD: {head_novo} ao funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,6,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 2:
                # Primeiro verificamos se o chamado foi digitado.
                if chamado != '':
                    # ATUALIZAMOS O HEAD ANTIGO COM OS DADOS DO FUNCIONARIO
                    try:
                        cursor_geral.execute("UPDATE Headsets SET EmPosse=NULL,UltPosse = ? WHERE Lacre = ?",posse,head_antigo)
                        conexao_geral.commit()
                    except:
                        pass
                    # AGORA RALIZAMOS O INSERT DO HEAD NOVO
                    cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse) VALUES (?,?)",(head_novo,posse))
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_antigo,Id_headset_novo,Id_funcionario,id_tecnico,Chamado) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),(SELECT id FROM Headsets WHERE Lacre=?),?,?,?)",(motivo,head_antigo,head_novo,posse,login,chamado))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                    #vamos registrar a alteração no log.          
                    registro = f"Trocou o HEAD: {head_antigo} por {head_novo} para o funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,7,registro))
                    conexao_geral.commit()
                    conexao_geral.close()
                else:
                    aviso.aviso(texto="Insira o chamado",tipo="ERRO",timeout=1500)

            if motivo == 3:
                   # AGORA RALIZAMOS O INSERT DO HEAD NOVO
                cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse) VALUES (?,?)",(head_novo,posse))
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = (SELECT Id FROM Headsets WHERE Lacre = ?) WHERE matricula = ?",head_novo,posse)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Realizou a correção do HEAD: {head_novo} para o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,12,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 4:
                   # AGORA RALIZAMOS O INSERT DO HEAD NOVO
                cursor_geral.execute("INSERT INTO Headsets (Lacre,EmPosse,Emprestado) VALUES (?,?,?)",(head_novo,posse,1))
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.Id_head_emprestimo = (SELECT Id FROM Headsets WHERE Lacre = ?) WHERE matricula = ?",head_novo,posse)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Realizou O empréstimo do HEAD: {head_novo} para o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,13,registro))
                conexao_geral.commit()
                conexao_geral.close()

        else:
            if motivo == 1:
                print ("errado")
                cursor_geral.execute("UPDATE Headsets SET EmPosse=?,Estoque=0,Manutencao=0,Inativo=0 WHERE Lacre = ?",posse,head_novo)
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Entregou o HEAD: {head_novo} ao funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,6,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 2:
                if chamado != '':
                    try:
                        cursor_geral.execute("UPDATE Headsets SET EmPosse=NULL,UltPosse = ? WHERE Lacre = ?",posse,head_antigo)
                        conexao_geral.commit()
                    except:
                        pass
                    cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse=?,Estoque=0,Manutencao=0,Inativo=0 WHERE Lacre = ?",posse,head_novo)
                    conexao_geral.commit()
                    cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_antigo,Id_headset_novo,Id_funcionario,id_tecnico,Chamado) VALUES (?,(SELECT Id FROM Headsets WHERE lacre=?),(SELECT id FROM Headsets WHERE Lacre=?),?,?,?)",(motivo,head_antigo,head_novo,posse,login,chamado))
                    conexao_geral.commit()
                    cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = Headsets.Id FROM Funcionarios INNER JOIN Headsets ON Funcionarios.matricula = Headsets.EmPosse WHERE Headsets.Lacre = ? ",head_novo)
                    conexao_geral.commit()
        
                    #vamos registrar a alteração no log.          
                    registro = f"Trocou o HEAD: {head_antigo} por {head_novo} para o funcionário: {nome}"
                    cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,7,registro))
                    conexao_geral.commit()
                    conexao_geral.close()
                else:
                    aviso.aviso(texto="Insira o chamado",tipo="ERRO",timeout=1500)

            if motivo == 3:
                cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse=?,Estoque=0,Manutencao=0,Inativo=0 WHERE Lacre = ?",posse,head_novo)
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT id FROM Headsets WHERE Lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.RecebeuHead = 1, Funcionarios.Id_headset = (SELECT Id FROM Headsets WHERE Lacre = ?) WHERE Matricula = ?",head_novo,posse)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Realizou a correção do HEAD: {head_novo} para o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,12,registro))
                conexao_geral.commit()
                conexao_geral.close()

            if motivo == 4:
                cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse=?,Estoque=0,Manutencao=0,Inativo=0,Emprestado=1 WHERE Lacre = ?",posse,head_novo)
                conexao_geral.commit()
                cursor_geral.execute("INSERT INTO TrocasHeadsets (Id_motivo,Id_headset_novo,Id_funcionario,id_tecnico) VALUES (?,(SELECT id FROM Headsets WHERE Lacre=?),?,?)",(motivo,head_novo,posse,login))
                conexao_geral.commit()
                cursor_geral.execute("UPDATE Funcionarios SET Funcionarios.Id_head_emprestimo = (SELECT Id FROM Headsets WHERE Lacre = ?) WHERE matricula = ?",head_novo,posse)
                conexao_geral.commit()
    
                #vamos registrar a alteração no log.          
                registro = f"Realizou O empréstimo do HEAD: {head_novo} para o funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,13,registro))
                conexao_geral.commit()
                conexao_geral.close()

        if realizar_desconto == 1:
            desconto()
        busca()    
        aviso.aviso("Headset associado com sucesso","AVISO",timeout=2000)
                
    def dar_baixa(): # Realiza baixa do head
        matricula = campo_matricula.get()
        lacre = campo_head_atual.get()
        nome = campo_nome.get()

        dados_geral = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BASE>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        
        try:
            cursor_geral.execute("UPDATE Funcionarios SET HeadDevolvido = 1 WHERE Matricula = ?",matricula)
            conexao_geral.commit()
            cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse = NULL WHERE EmPosse = ?",matricula)
            conexao_geral.commit()
            registro = f"Realizou a  baixa do HEAD: {lacre} do funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,8,registro))
            conexao_geral.commit()
            conexao_geral.close()
        except pyodbc.Error as erro:
            print (erro)
        
        busca()    
        aviso.aviso("Baixa de Head concluída","AVISO",timeout=1500)

    def dar_baixa_emprestimo():
        matricula = campo_matricula.get()
        lacre = campo_head_atual.get()
        nome = campo_nome.get()

        dados_geral = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BASE>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        
        try:
            cursor_geral.execute("UPDATE Funcionarios SET Id_head_emprestimo = NULL WHERE Matricula = ?",matricula)
            conexao_geral.commit()
            cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse = NULL WHERE EmPosse = ?",matricula)
            conexao_geral.commit()
            registro = f"Realizou a  baixa do HEAD: {lacre} emprestado ao funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,13,registro))
            conexao_geral.commit()
            conexao_geral.close()
        except pyodbc.Error as erro:
            print (erro)
        
        busca()    
        aviso.aviso("Baixa de Head concluída","AVISO",timeout=1500)

    def desconto(): # Vai registrar o desconto
        # Vamos pegar os dados para continuar
        matricula = campo_matricula.get()
        lacre_atual = campo_head_atual.get()
        lacre_novo = campo_head_novo.get()

        # Conexão ao banco de dados
        dados_geral = (
        "driver={SQL Server Native Client 11.0};"
        "server=<SERVIDOR>;"
        "Database=<BASE>;"
        "UID=<USUARIO>;"
        "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        
        # vamos primeiro registrar o desconto na tabela funcionarios
        try:
            cursor_geral.execute("UPDATE Funcionarios SET QuantidadeDescontos=(SELECT SUM(QuantidadeDescontos + 1) FROM Funcionarios WHERE Matricula=?) WHERE Matricula=?",matricula,matricula)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print (erro)

        # agora vamos registrar o equipamento na tabela de desconto.
        try:
            cursor_geral.execute("INSERT INTO Descontos (Id_funcionario,Id_EquipamentoAntigo,Id_equipamentoNovo) VALUES (?,(SELECT Id FROM Headsets WHERE Lacre = ?),(SELECT Id FROM Headsets WHERE Lacre = ?))",matricula,lacre_atual,lacre_novo)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print (erro)
        
        conexao_geral.close()
        busca()

# FUNÇÕES DE CADASTRO OU ATUALIZAÇÃO DE ACESSOS
    def cadastro(): # Cadastro ok

        # Começamos pegando as informações dos campos.
        matricula = campo_matricula.get()
        nome = campo_nome.get()
        setor = campo_setor.get()
        pis = campo_pis.get()
        cartao = campo_cartao.get()

        if str(matricula) != '':
            # vamos verificar se o pis está vazio.
            if pis == '':
                aviso.aviso("Insira o PIS","ERRO",timeout=1500)

            else:
            # Vamos verificar se o campo está vazio.
                if cartao == '':
                    cartao = None
                else:
                    # Vamos pegar o cartão mas com uma pequena função para
                    # retirar os zerros a esquerda e tranformar em um int
                    cartao_campo = campo_cartao.get()
                    cartao_com_zeros = [cartao_campo]
                    cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
                    cartao = cartao_sem_zeros[0]
                    cartao = int(str(cartao))

                # E depois apagamos para evitar sobreescrever
                campo_matricula.delete(0, 'end')
                campo_nome.delete(0, 'end') 
                campo_cartao.delete(0, 'end')
                campo_pis.delete(0, 'end')
                campo_setor.set('')

                # Iniciamos cadastrando nas catracas
                dados_catraca = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor = conexao_catraca.cursor()
                
                try: # vamos verificar se a chave já existe.
                    chave_catraca = cursor.execute("SELECT Funcionarios.Matricula FROM Funcionarios WHERE Matricula = ?", matricula).fetchall()
                    conexao_catraca.commit()
                    chave_catraca =str(chave_catraca)
                    if str(chave_catraca) == "[]":
                        try: 
                            # vamos buscar o código do departamento!
                            departamento = cursor.execute("SELECT Departamentos.COD_DEPARTAMENTO FROM Departamentos WHERE Departamentos.Descricao LIKE ?", setor ).fetchall()
                            conexao_catraca.commit()
                            departamento = str(departamento)
                            # Se o nome do departamento for vazio ele vai criar esse departamento
                            if str(departamento) == "[]":
                                try:
                                    cursor.execute("INSERT INTO Departamentos (Descricao,COD_EMPRESA) VALUES (?,?)",(setor,2))
                                    conexao_catraca.commit()
                                    departamento = cursor.execute("SELECT Departamentos.COD_DEPARTAMENTO FROM Departamentos WHERE Departamentos.Descricao LIKE ?", setor ).fetchall()
                                    conexao_catraca.commit()
                                    cod_dep = str(departamento).split(",")[0].strip("[").strip("(")
                                    print(f"Setor {departamento}: cadastrado com sucesso")

                                except pyodbc.Error as erro:
                                    print(f"Erro ao cadastrar o setor: {departamento}: {erro}")

                            # Se o setor exixtir vai extrair o código do setor.    
                            else:
                                cod_dep = str(departamento).split(",")[0].strip("[").strip("(")
                                print(cod_dep)

                            try: #vamos criar o código de pessoa que serve de chave.
                                cursor.execute("INSERT INTO Pessoas (Tipo) VALUES (?)",(0))
                                conexao_catraca.commit()
                                pessoa = cursor.execute("SELECT TOP 1 COD_PESSOA FROM Pessoas ORDER BY COD_PESSOA DESC").fetchall()
                                conexao_catraca.commit()
                                cod_pes = str(pessoa[0]).split(",")[0].strip("(")
                                print(cod_pes)
                                print(f"cod de pessoa gerado com sucesso")

                            except pyodbc.Error as erro:
                                print(f"erro ao gerar cod de pessoa: {erro}")
                            # aqui criamos o cadastro do funcionário.
                            try: 
                                cursor.execute("INSERT INTO Funcionarios (COD_PESSOA,Nome,COD_DEPARTAMENTO,COD_ZT,COD_PERFIL,IgnorarRota,IgnorarAntiPassback,IgnorarEntradas,Bloqueado,Matricula) VALUES (?,?,?,?,?,?,?,?,?,?)",(cod_pes,nome,cod_dep,1,1,1,0,1,0,matricula))
                                conexao_catraca.commit()

                            except pyodbc as erro:
                                print (f"{hora}: cadastro do funcionario {matricula} não realizado: {erro}")
                            # agora tentamos o cadastro do cartão        
                            
                            if str(cartao) != None:
                                try:
                                    cursor.execute('INSERT  INTO Cartoes (NUM_CARTAO,COD_PESSOA,CodigoDeBarras,DarBaixa,Offline,SemDigital,EnviadoListaSemDigital) VALUES (?,?,?,?,?,?,?)', (cartao,cod_pes,cartao,0,0,0,0)) 
                                    conexao_catraca.commit()
                                except pyodbc.Error as erro:
                                    print (erro)
                        except pyodbc as erro:
                            print (f"{hora}: falha na busca ou criação do departamento: {setor}: {erro}")

                    else:
                        print (f"{hora}: A chave primária {matricula} já existe no banco das catracas.\n")
                except:
                    pass

                conexao_catraca.close()             
                dados_central = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BASE>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
                )
                conexao_central = pyodbc.connect(dados_central)
                cursor = conexao_central.cursor()
                try: # vamos verificar se a chave já existe.
                    chave_central = cursor.execute("SELECT Funcionarios.Matricula FROM Funcionarios WHERE Matricula = ?", matricula).fetchall()
                    conexao_central.commit()
                    chave_central = str(chave_central)
                    print (chave_central,"Central")
                    if chave_central == '[]':
                        try: 
                            cursor.execute("INSERT INTO Funcionarios (Matricula,Nome,Pis,Id_departamento,Cartao) VALUES (?,?,?,(SELECT Id FROM Departamentos WHERE Departamentos.Departamento LIKE ?),?)",(matricula,nome,pis,setor,cartao))
                            conexao_central.commit()
                            print("Cadastro?")

                        except pyodbc as erro:
                            print (f"{hora}: cadastro do funcionario no banco geral: {matricula} não realizado: {erro}")


                    else:
                        print (f"{hora}: A chave primária {matricula} já existe no banco central.\n")                 
                except:
                    pass

                conexao_central.close()

                #agora vamos realizar  a importação no access
                dados_access = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\172.10.20.60\SoapAdmin3.5\Access.mdb;'
                conexao_access = pyodbc.connect(dados_access)
                cursor = conexao_access.cursor()

                nome = nome.split()
                if len(nome) < 1:
                    # separamos o nome e sobrenome        
                    nome_separado = nome.split()[0]
                    caraceteres_nome = len(nome_separado) + 1
                    sobrenome = nome[caraceteres_nome:]
                else:
                    sobrenome = None

                if cartao != None:
                    #convertemos o cartão para exadecimal
                    cartao_exa = hex(cartao)
                    caracetres = len(cartao_exa)

                    #contamos os caracteres epegamos apenas os 6 ultimos digitos
                    if caracetres > 10:
                        cartao_exa_cortado = cartao_exa[6:]
                    else:
                        cartao_exa_cortado = cartao_exa[5:]

                    #convertemos novamente para decimal na base 16
                    cartao_decimal = int(cartao_exa_cortado,16)

                # realizamos o update
                # vamos verificar se a chave já existe.    
                try: 
                    chave_access = cursor.execute("SELECT USERINFO.Badgenumber FROM USERINFO WHERE Badgenumber = ?", matricula).fetchall()
                    conexao_access.commit()
                    chave_access = str(chave_access)
                    if str(chave_access) == "[]":
                        try:
                            cursor.execute("""INSERT INTO USERINFO (Badgenumber, SSN, name, lastname, Gender, CardNo) 
                            VALUES (?,?,?,?,?,?)""",
                            (matricula, 0, nome_separado, sobrenome, "M", cartao_decimal))
                            conexao_access.commit()
                            print (f"{hora}: cadastro do funcionario no access {matricula} com sucesso.")

                        except pyodbc as erro:
                            print (f"{hora}: cadastro do funcionario no banco geral: {matricula} não realizado: {erro}")

                    else:
                        print (f"{hora}: A chave primária {matricula} já existe no banco access.\n")

                except pyodbc.Error as erro:
                    print (erro)
                conexao_access.close()
                
                busca()
                aviso.aviso("Cadastrado com sucesso","AVISO",timeout=1500)

                # AGORA VAMOS REGISTRAR O LOG DE CRIAÇÂO
                conexao_geral = pyodbc.connect(dados_central)
                cursor_geral = conexao_geral.cursor()
                registro = f"Realizou o cadastro do funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES",(usuario,10,registro))
                conexao_geral.close()
        else:
            aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
        
    def atualizar(): # Atualiza os dados dos funcionários. testar
        
        with open(log_path, "a") as log:
            # Aqui vamos fegar todos os dados nos campos do programa
            matricula = campo_matricula.get()
            nome = campo_nome.get()
            setor = campo_setor.get()
            pis = campo_pis.get()

            if str(matricula) != "":
                # Vamos pegar o cartão mas com uma pequena função para
                # retirar os zerros a esquerda e tranformar em um int
                cartao_campo = campo_cartao.get()
                cartao_com_zeros = [cartao_campo]
                cartao_sem_zeros = [ele.lstrip('0') for ele in cartao_com_zeros]
                cartao = cartao_sem_zeros[0]
                cartao = int(str(cartao))

                # Dados de acesso ao banco Geral
                dados_geral = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BASE>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_geral = pyodbc.connect(dados_geral)
                cursor_geral = conexao_geral.cursor()
                # Realizamos o Update das informações.
                try:
                    cursor_geral.execute("UPDATE Funcionarios SET Matricula=? , Nome=?, Id_departamento=(SELECT Id FROM Departamentos WHERE Departamentos.Departamento LIKE ?), Cartao=?, Pis=?  WHERE Matricula = ?",(matricula,nome,setor,cartao,pis,matricula))
                    conexao_geral.commit()
                    log.write(f"{hora}: os dados do usuario {nome} foram atualizados no banco central.\n")
                except pyodbc.Error() as erro:
                    log.write(f"{hora}: os dados do usuario {nome} não foram atualizados no banco central erro: {erro}\n")

                conexao_geral.close()
                # Agora no banco das catracas.
                dados_catraca = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor_catraca = conexao_catraca.cursor()
            
                try:
                    #primeiro vamos buscar o código de pessoa do usuario!
                    nome_usuario = cursor_catraca.execute("SELECT Funcionarios.COD_PESSOA FROM Funcionarios WHERE Funcionarios.Matricula = ?", matricula ).fetchall()
                    conexao_catraca.commit()
                    nome_usuario = str(nome_usuario)
                    copess = nome_usuario.split(",")[0]
                    copess = str(copess)
                    copess = copess.split("(")[1]
                    #com os dados em mãos vamos buscar se já existe um cartão para o funcionário.
                    try:
                        query = cursor_catraca.execute('SELECT NUM_CARTAO FROM Cartoes WHERE COD_PESSOA = ?',copess).fetchall()
                        conexao_catraca.commit()
                        
                        if str(query) != "[]":
                            cursor_catraca.execute('UPDATE Cartoes SET NUM_CARTAO = ?,CodigoDeBarras = ? WHERE COD_PESSOA = ?',cartao,cartao,copess) 
                            conexao_catraca.commit()
                        else:
                            cursor_catraca.execute('INSERT  INTO Cartoes (NUM_CARTAO,COD_PESSOA,CodigoDeBarras,DarBaixa,Offline,SemDigital,EnviadoListaSemDigital) VALUES (?,?,?,?,?,?,?)', (cartao,copess,cartao,0,0,0,0)) 
                            conexao_catraca.commit()

                    except pyodbc.Error as erro:
                        print (erro)
                
                except pyodbc.Error as erro:
                    print (erro)

                conexao_catraca.close()

                #agora vamos fazer a atualização dos dados no soap admin.
                dados = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\172.10.20.60\SoapAdmin3.5\Access.mdb;'
                conexao = pyodbc.connect(dados)
                cursor = conexao.cursor()
                cartao = int(cartao)

                #separando o nome do sobrenome
                nome_separado = nome.split()[0]
                caraceteres_nome = len(nome_separado) + 1
                sobrenome = nome[caraceteres_nome:]

                #convertemos o cartão para exadecimal
                cartao_exa = hex(cartao)
                caracetres = len(cartao_exa)

                #contamos os caracteres epegamos apenas os 6 ultimos digitos
                if caracetres > 10:
                    cartao_exa_cortado = cartao_exa[6:]
                else:
                    cartao_exa_cortado = cartao_exa[5:]

                #convertemos novamente para decimal na base 16
                cartao_decimal = int(cartao_exa_cortado,16)
                # realizamos o update
                try:
                    cursor.execute("UPDATE USERINFO SET Badgenumber = ?,name = ?, lastname = ?,CardNo = ? WHERE Badgenumber = ?",(matricula, nome_separado, sobrenome, cartao_decimal, matricula))
                    conexao.commit()
                    log.write(f"{hora}: os dados do usuario {nome} foram atualizados no banco Access.\n")
                except pyodbc.Error as erro:
                    log.write(f"{hora}: os dados do usuario {nome} foram atualizados no banco central erro: {erro}\n")
                conexao.close()

                conexao_geral = pyodbc.connect(dados_geral)
                cursor_geral = conexao_geral.cursor()
                registro = f"Atualizou o cadastro do funcionário: {nome}"
                cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,11,registro))
                conexao_geral.close()

                busca()
                aviso.aviso("Atualizado com sucesso","AVISO",timeout=1500)

            else:
                aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
        
    def baixa_funcionario():
        matricula = campo_matricula.get()
        nome = campo_nome.get()
        demissao = dia.strftime('%Y-%m-%d')

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BASE>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        # Realizamos o Update das informações.
        try:
            cursor_geral.execute("UPDATE Funcionarios SET Situacao = 1,DataDemissao=?  WHERE Matricula = ?",demissao,matricula)
            conexao_geral.commit()
            registro = f"Realizou a baixa manual do funcionário: {nome}"
            cursor_geral.execute("INSERT INTO RegistroAcoes (Usuario,tipo,Registro) VALUES (?,?,?)",(login,9,registro))
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()

        busca()
        aviso.aviso("Baixa de funcionário concluída","AVISO",timeout=1500)
        
    # FUNÇÕES DE EXCLUSÃO
    def excluir_catraca(): # Excluindo OK
        
        with open(log_path, "a") as log:
            matricula = campo_matricula.get()

            if str(matricula) != '':

                nome = campo_nome.get()
                dados_catraca = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BANCO>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_catraca = pyodbc.connect(dados_catraca)
                cursor_catraca = conexao_catraca.cursor()
            
                try:
                    cursor_catraca.execute(
                        "DELETE Cartoes FROM Cartoes LEFT OUTER JOIN Funcionarios ON Funcionarios.COD_PESSOA = Cartoes.COD_PESSOA WHERE Funcionarios.Matricula = ?", matricula
                    )
                    conexao_catraca.commit()
                    cursor_catraca.execute("DELETE FROM Funcionarios WHERE Matricula = ?", matricula)
                    conexao_catraca.commit()
                    log.write(f"{hora}: O usuario {nome} foi excluido das catracas.\n")
                except pyodbc.Error() as erro:
                    print(erro)

                conexao_catraca.close()

                dados_geral = (
                    "driver={SQL Server Native Client 11.0};"
                    "server=<SERVIDOR>;"
                    "Database=<BASE>;"
                    "UID=<USUARIO>;"
                    "PWD=<SENHA>;"
                )
                conexao_geral = pyodbc.connect(dados_geral)
                cursor_geral  = conexao_geral.cursor()
            
                try:
                    cursor_geral.execute("DELETE FROM Funcionarios WHERE Matricula = ?", matricula)
                    conexao_geral.commit()
                    log.write(f"{hora}: O usuario {nome} foi excluido das banco geral.\n")
                except pyodbc.Error() as erro:
                    print(erro)
                    
                conexao_geral.close()
                aviso.aviso("Excluído das catracas","AVISO",timeout=1500)
            else:
                aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    def excluir_controlid(): # Excluindo OK
        with open(log_path, "a") as log:

            matricula = campo_matricula.get()

            if str(matricula) != '':
                nome = campo_nome.get()
                pis = campo_pis.get()
                host = ["IPS DE CADA EQUIPAMENTO"]
                for link in host:        
                    url_login = link + '/login.fcgi'
                    dados = {
                    'login': 'admin', 
                    'password': 'admin'
                    }
                    response = requests.post(url=url_login, json=dados, verify=False)
                    chave = response.json()
                    chave = str(chave).split("'")[3]             
                    url_remove = link + '/remove_users.fcgi?session=' + chave 
                    funcionarios ={
                    'users':[
                        int(pis)
                    ]
                    }
                    response = requests.post(url=url_remove, json=funcionarios, verify=False)
                    resposta = str(response)
                    if resposta == "<Response [200]>":
                        log.write(f"{hora}: o usuario {nome} foi excluido do ponto\n")       

                url_logoff = link + '/logout.fcgi'
                dados = json.dumps({
                "  session": chave
                })
                cabecalho = {
                'Content-Type': 'application/json'
                }

                response = requests.request("POST", url=url_logoff, headers=cabecalho, data=dados, verify=False)

                aviso.aviso("Excluído do ponto","AVISO",timeout=1500)
            else:
                aviso.aviso("Insira a matrícula","ERRO",timeout=1500)

    def excluir_access(): # Excluindo ok

        with open(log_path, "a") as log:

            matricula = campo_matricula.get()

            if str(matricula) != '':
                nome = campo_nome.get()
                dados_acess = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=<CAMINHO>;'
                conexao_access = pyodbc.connect(dados_acess)            
                cursor_access = conexao_access.cursor()
                
                try:
                    conexao_access.execute('DELETE FROM USERINFO WHERE USERID = ?', (matricula))
                    cursor_access.commit()
                    log.write(f"{hora}: o usuario {nome} foi removido da porta.\n")
                except pyodbc.Error as erro:
                    print (erro)
            
                conexao_access.close()

                aviso.aviso("Excluido da porta","AVISO",timeout=1500)
            else:
                aviso.aviso("Insira a matrícula","ERRO",timeout=1500)
    
    # CONFIGURAÇÕES DE TELA
    janela = TkinterDnD.Tk() #criando janela
    janela.title("Associa Provisório") #titulo
    janela.geometry("1350x625") #quadro
    janela.config(bg = "white")
    janela.iconbitmap(icon_path)
    janela.resizable(0,0)
    logo = PhotoImage(file=str(logo_path))
    label1 = Label( janela, image = logo, bg="white") 
    label1.place(x=0,y=0)

    comando_argumentos = partial(logoff,janela)

    # FUNCIONARIO
    usuario = str(usuario).split("'")
    usuario = usuario[1]
    Label(janela,bg="white",highlightbackground="black",highlightthickness=1,width=169,height=2).place(x=143,y=5)
    Label(janela,bg="white",text="LOGADO COMO: "+ usuario).place(x=150,y=15)
    Button(janela,fg="navy blue",text="Headsets",command=abrir_heads,width=10).place(x=1060,y=11)
    Button(janela,fg="navy blue",text="Trocar senha",command=trocar_senha,width=10).place(x=1150,y=11)
    Button(janela,fg="navy blue",text="Logoff",command=comando_argumentos,width=10).place(x=1240,y=11)

    # DECORAÇÃO
    Label(janela,bg="white",width=65,height=14,highlightthickness=1,highlightbackground="black").place(x=15,y=55)
    Label(janela,text="FUNCIONÁRIO",bg="white").place(x=30,y=45)

    # CAMPOS
    campo_matricula = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_matricula.place(x=100,y=70)
    texto_matricula = Label(janela,text="Matricula: ",bg="white")
    texto_matricula.place(x=30,y=70)

    campo_nome = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1, textvariable="nome",width=55)
    campo_nome.place(x=100,y=110)
    texto_nome = Label(janela,text="Nome: ",bg="white")
    texto_nome.place(x=30,y=110)

    # SETOR
    Label(janela,bg="white",text="Setor: ").place(x=30,y=150)
    setor = StringVar()
    campo_setor = Combobox(janela, textvariable=setor, values=lista_setor)
    campo_setor.place(x=100,y=150,width=185)

    # CARTÃO
    Label(janela,bg="white",text="Cartão: ").place(x=30,y=190)
    campo_cartao = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_cartao.place(x=100,y=190)

    # SITUAÇÃO
    Label(janela,bg="white",text="Situação: ").place(x=240,y=190)
    campo_situacao = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_situacao.place(x=300,y=190)

    # DEMISSÃO
    Label(janela,bg="white",text="Demissão: ").place(x=240,y=230)
    campo_demissao = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_demissao.place(x=300,y=230)

    # PIS
    Label(janela,bg="white",text="Pis: ").place(x=30,y=230)
    campo_pis = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1, textvariable="pis")
    campo_pis.place(x=100,y=230)


    # BOTOES
    botao_busca = Button(janela,text="Buscar",command=busca,fg="navy blue") #realiza a busca.
    botao_busca.place(x=240,y=68)

    botao_limpar = Button(janela,text="limpar",command=limpa,fg="navy blue")
    botao_limpar.place(x=300,y=68)

    # SITUACAO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=285)
    Label(janela,text="ENTRADA NA EMPRESA",bg="white").place(x=30,y=275)

    # CAMPOS
    Label(janela,text="Situação: ",bg="white").place(x=30,y=305)
    campo_bloqueio = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1, width=20)
    campo_bloqueio.place(x=100,y=305)

    Label(janela,text="Motivo:",bg="white").place(x=240,y=305)
    campo_motivo = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=20)
    campo_motivo.place(x=300,y=305)

    Label(janela,text="* Para bloquear, selecione o período de bloqueio",fg="red",bg="white").place(x=30,y=335)

    # CALENDARIO BLOQUEIO
    Label(janela,text="Data inicial:",bg="white").place(x=30,y=370)
    campo_data_inicial_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_inicial_bloqueio.place(x=100,y=370)

    Label(janela,text="Data final:",bg="white").place(x=207,y=370)
    campo_data_final_bloqueio = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    campo_data_final_bloqueio.place(x=270,y=370)

    # BOTÕES
    Button(janela,text="Bloquear funcionário",command=bloquear_funcionario,fg="navy blue").place(x=30,y=405)
    Button(janela,text="Desbloquear funcionário",command=desbloquear_funcionario,fg="navy blue").place(x=207,y=405)

    # PROVISORIO
    # DECORACAO
    Label(janela,bg="white",width=65,height=10,highlightthickness=1,highlightbackground="black").place(x=15,y=455)
    Label(janela,text="PROVISÓRIO",bg="white").place(x=30,y=445)

    # CAMPOS
    campo_cartao_provisorio = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=10)
    campo_cartao_provisorio.place(x=100,y=475)
    Label(janela,text="Cartão:",bg="white").place(x=30,y=475)

    Label(janela,text="* Para associar, insira os 5 dígitos do cartão e selecione o período",fg="red",bg="white").place(x=30,y=505)


    # CALENDARIO PROVISORIO
    Label(janela,text="Data inicial:",bg="white").place(x=30,y=540)
    data_inicial = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_inicial.place(x=100,y=540)

    Label(janela,text="Data final:",bg="white").place(x=207,y=540)
    data_final = tkcalendar.DateEntry(janela,date_pattern='dd/mm/yyyy')
    data_final.place(x=270,y=540)

    # BOTOES
    botao_associar = Button(janela,text="Associar cartão",command=associar,fg="navy blue")
    botao_associar.place(x=30,y=575)

    botao_dessassociar = Button(janela,text="Desassociar cartão",command=desassociar,fg="navy blue")
    botao_dessassociar.place(x=207,y=575)

    # OPÇÕES
    # DECORAÇÃO OPCOES
    Label(janela,bg="white",highlightbackground="black",highlightthickness=1,width=40,height=4).place(x=485,y=55)
    Label(janela,text="CADASTRO",bg="white").place(x=500,y=45)

    # BOTOES OPCOES
    Button(janela,height=1,width=10,fg="navy blue",text="Atualizar",command=atualizar).place(x=500,y=75)
    Button(janela,height=1,width=10,text="Cadastrar",fg="navy blue", command=cadastro).place(x=590,y=75)
    Button(janela,height=1,width=10,text="Baixa",fg="navy blue", command=baixa_funcionario).place(x=680,y=75)

    # DECORACAO EXCLUSAO
    Label(janela,bg="white",highlightbackground="black",highlightthickness=1,width=40,height=4).place(x=485,y=140)
    Label(janela,text="EXCLUIR CADASTRO",bg="white").place(x=500,y=130)

    # BOTOES EXCLUSAO
    Button(janela,height=1,width=10,fg="navy blue",text="Catracas",command=excluir_catraca).place(x=500,y=160)
    Button(janela,height=1,width=10,fg="navy blue",text="Ponto",command=excluir_controlid).place(x=590,y=160)
    Button(janela,height=1,width=10,fg="navy blue",text="Porta",command=excluir_access).place(x=680,y=160)

    ############################################# HEADSET ################################################
    Label(janela,bg="white",width=40,height=6,highlightbackground="black",highlightthickness=1).place(x=485,y=225)
    Label(janela,text="HEADSET",bg="white").place(x=500,y=215)

    # CAMPOS
    Label(janela,bg="white",text="Headset atual:").place(x=500,y=245)
    campo_head_atual = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_head_atual.place(x=600,y=245)

    Label(janela,bg="white",text="Situação:").place(x=500,y=280)
    campo_uso = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_uso.place(x=600,y=280)

    # TROCA/ENTREGA
    Label(janela,bg="white",width=40,height=18,highlightthickness=1,highlightbackground="black").place(x=485,y=335)
    Label(janela,text="ENTREGA/TROCA/BAIXA",bg="white").place(x=500,y=325)

    # LACRE DO HEADSET NOVO
    Label(janela,bg="white",text="Headset novo:").place(x=500,y=355)
    campo_head_novo = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_head_novo.place(x=600,y=355)

    # MOTIVO DA TROCA
    def teste(event): # Essa funão testa a opção para habilitar o campo do chamado.
            mot = motivo.get()
            if mot == "TROCA":
                campo_chamado['state'] = NORMAL
    Label(janela,bg="white",text="Motivo:").place(x=500,y=390)
    motivo = StringVar()
    caixa_motivo = Combobox(janela, textvariable=motivo, values=("TROCA","INTEGRAÇÃO","CORREÇÃO","EMPRÉSTIMO"),)
    caixa_motivo.place(x=600,y=390,width=100)
    caixa_motivo.bind("<<ComboboxSelected>>", teste)

    # CHAMADO
    Label(janela,bg="white",fg="red",text="* Para troca informe o chamado.").place(x=500,y=420)
    Label(janela,bg="white",text="Chamado:").place(x=500,y=450)
    campo_chamado = Entry(janela,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,state=DISABLED)
    campo_chamado.place(x=600,y=450)

    # CHECKBOX PARA DESCONTO
    variavel_desconto = IntVar()
    check_desconto = Checkbutton(janela,bg="white",fg="red",text="* Marque a caixa para registrar desconto\n por danos no equipamento ou perca.",variable=variavel_desconto, offvalue=0, onvalue=1)
    check_desconto.place(x=500,y=475)

    # BOTÕES
    botao_associa_head = Button(janela,fg="navy blue",width=30,text='Atualizar headset',command=associa_heads)
    botao_associa_head.place(x=520,y=520)

    botao_baixa_head = Button(janela,fg="navy blue",width=30,text='Dar baixa no headset',command=dar_baixa)
    botao_baixa_head.place(x=520,y=550)

    botao_baixa_emprestimo = Button(janela,fg="navy blue",width=30,text='Dar baixa no empréstimo',command=dar_baixa_emprestimo)
    botao_baixa_emprestimo.place(x=520,y=580)

    # QUADRO DE USUARIOS.
    # DECORACAO
    Label(janela,bg="white",width=78,height=31,highlightthickness=1,highlightbackground="black").place(x=780,y=55)

    # CAMPOS
    texto_quadro = Label(janela,text="FUNCIONÁRIOS COM CARTÕES PROVISÓRIOS",bg="white")
    texto_quadro.place(x=930,y=45)

    campo_quadro = ttk.Treeview(janela,columns=("c1","c2","c3","c4","c5",),show='headings',height=19)
    campo_quadro.place(x=790,y=65)

    # COLUNAS
    # Cartão
    campo_quadro.column("#1",width=60)
    campo_quadro.heading("#1",text="CARTÃO")
    # Matricula
    campo_quadro.column("#2",width=80)
    campo_quadro.heading("#2",text="MATRÍCULA")
    # Nome
    campo_quadro.column("#3",width=250)
    campo_quadro.heading("#3",text="NOME")
    # Inicio
    campo_quadro.column("#4",width=70)
    campo_quadro.heading("#4",text="INICIO")
    # Fim
    campo_quadro.column("#5",width=70)
    campo_quadro.heading("#5",text="FIM")

    # BOTOES
    Button(janela,text="Atualizar lista",fg="navy blue",command=associados).place(x=790,y=490)

    # LOG
    Button(janela,height=1,width=5, text="log", command=abrir_log,fg="navy blue").place(x=1289,y=588) #botão para abrir o log.

    #VERSÃO
    Label(janela,bg="white",text="Criado por: Daniel Lopes Manfrini.").place(x=790, y=540)
    Label(janela,bg="white",text="Versão: 2.5").place(x=790, y=560)
    Label(janela,bg="white",text="Modelo: Admin").place(x=790, y=580)
    janela.mainloop()