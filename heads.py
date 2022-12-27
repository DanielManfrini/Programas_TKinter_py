#Importadores GLOBAL
from datetime import datetime
import re

#importadores SQL
import pyodbc

#variaveis Tkinter
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox
import tkcalendar

#importadores caminhos
import os
import sys

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
icon_path = os.path.join(application_path, icon_name)


# FUNÇÕES
# BUSCAR DADOS
def tela_head():
    
    def busca(): # Busca OK
        # Limpamos os dados em caso de nova pesquisa.
        campo_funcionario.delete(0,'end')
        campo_chamado.delete(0,'end')
        campo_situacao.delete(0,'end')
        caixa_marca.set(value='')
        caixa_defeito.set(value='')
        caixa_vendedor.set(value='')
        campo_garantia.delete(0,'end')
        check_desconto.deselect()

        # Pegamos o lacre para realizar a busca
        lacre = campo_lacre.get()
        print(lacre)
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            head = cursor_geral.execute("SELECT Headsets.EmPosse,Headsets.Chamado,Headsets.Manutencao,Headsets.Inativo,Headsets.Estoque,MarcasHeadsets.Marca,Headsets.Garantia FROM Headsets LEFT JOIN MarcasHeadsets ON MarcasHeadsets.id = Headsets.Id_marca WHERE lacre = ?",lacre).fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:

            print(erro)
        print (head)
        matricula = head[0][0]
        chamado = head[0][1]
        manutencao = head[0][2]
        inativo = head[0][3]
        estoque = head[0][4]
        marca = head[0][5]
        garantia = head[0][6]

        try:
            defeito = cursor_geral.execute("SELECT Defeito FROM DefeitosHeadsets INNER JOIN Headsets ON DefeitosHeadsets.Id = Headsets.Id_defeito WHERE Lacre = ?",lacre).fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        
        try:
            vendedor = cursor_geral.execute("SELECT Vendedor FROM Vendedores INNER JOIN Headsets ON Vendedores.Id = Headsets.Id_vendedor WHERE Lacre = ?",lacre).fetchall()
            conexao_geral.commit()               
        except pyodbc.Error as erro:
            print(erro)


        # Verificamos se os dados estão vazios
        if matricula == None:
            matricula = "NÃO ASSOCIADO"
        if chamado == None:
            chamado = "NÃO REGISTRADO"
        if marca == None:
            marca = "NÃO REGISTRADO"
        if garantia == None:
            garantia = "NÃO"
        else:
            garantia = "SIM"

        # VAMOS VERIFICAR SE ESTÀ INATIVO
        if estoque == False:
            if inativo == False:
                # VAMOS VERIFICAR SE ESTÀ EM MANUTENÇÂO
                if manutencao != False:            
                    estado = "MANUTENÇÃO"       
                else:
                    estado = "EM USO" 
            else:
                estado = "INATIVO"
        else:
            estado = "ESTOQUE"

        if str(defeito) != "[]":
                    defeito = defeito[0][0]
                    caixa_defeito.set(defeito) 

        if str(vendedor) != "[]":
                    vendedor = vendedor[0][0]
                    caixa_vendedor.set(vendedor) 

        if vendedor == None:
            vendedor = "NÃO REGISTRADO"            

        # Com os dados em mãos vamos injetar
        campo_funcionario.insert(END,matricula)
        campo_chamado.insert(END,chamado)
        campo_situacao.insert(END,estado)
        caixa_marca.set(marca)
        variavel_lacre_atual.set(lacre)
        caixa_vendedor.set(vendedor)
        campo_garantia.insert(END,garantia) 

        conexao_geral.close()
    # TELA DE HEADS COM DEFEITO.
    def busca_defeitos():
        # limpar a tabela atual
        count_assoc = 0
        campo_quadro.delete(*campo_quadro.get_children())

        # Vamos buscar os cartões assoiados a funcionários.
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        # Verificar se o cartão existe.
        assoc = cursor_geral.execute("SELECT Headsets.Lacre,Headsets.Chamado,DefeitosHeadsets.Defeito FROM Headsets FULL JOIN DefeitosHeadsets ON DefeitosHeadsets.Id = Headsets.Id_defeito WHERE Manutencao = 1 AND Inativo != 1").fetchall()
        conexao_geral.commit()
        for linha in assoc:
            lacre = linha[0]
            defeito = linha[2]
            chamado = linha[1]
            count_assoc += 1

            if defeito == None:
                defeito = "NÃO REGISTRADO"
            if chamado == None:
                chamado = "NÃO REGISTRADO"

            campo_quadro.insert("",END, values=(lacre,defeito,chamado)) 
            valor_quantidade['text'] = count_assoc

        conexao_geral.close()
    # ATUALIZA O USO
    def uso():
        # PEGAMOS A MATRICULA
        matricula = campo_funcionario.get()

        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        if matricula.count() == 6 :
            cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse = NULL WHERE EmPosse = ?",matricula)
            conexao_geral.commit()        
    # CADASTRAR LACRE.
    def cadastro_lacre():
        # Lacre
        lacre = campo_lacre.get()
        marca = variavel_marca.get()
        vendedor = variavel_vendedor.get()

        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("INSERT INTO Headsets (lacre,Id_marca,id_vendedor) VALUES (?,(SELECT Id FROM MarcasHeadsets WHERE Marca LIKE ?),(SELECT id FROM Vendedores WHERE Vendedor LIKE ?))",(lacre,marca,vendedor))
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # ATUALIZAR LACRE.
    def atualizar_lacre():
        # Lacre
        lacre_atual = variavel_lacre_atual.get()
        lacre_novo = campo_lacre.get()
        marca = caixa_marca.get()
        vendedor = caixa_vendedor.get()

        print("VARIAVEL HEAD ATUAL: ",lacre_atual)
        print("marca: ",marca)
        print("vendedor: ",vendedor)
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("UPDATE Headsets SET lacre = ?, Id_marca = (SELECT Id FROM MarcasHeadsets WHERE Marca LIKE ?),id_vendedor = (SELECT id FROM Vendedores WHERE Vendedor LIKE ? ) WHERE lacre = ?",(lacre_novo,marca,vendedor,lacre_atual))
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # CADASTRAR DEFEITOS.
    def cadastra_defeito():
        # Lacre
        lacre = campo_lacre.get()
        defeito = caixa_defeito.get()
        garantia = variavel_garantia.get()
        print(garantia)
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("UPDATE Headsets SET Manutencao=1,Estoque=0, Id_defeito=(SELECT Id FROM DefeitosHeadsets WHERE Defeito LIKE ?),Inativo = 0 WHERE Lacre = ?",defeito,lacre)
            conexao_geral.commit()
            if garantia == 1:
                cursor_geral.execute("UPDATE Headsets SET Garantia = 1 WHERE Lacre = ?",lacre)
                conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # REMOVER DEFEITOS.
    def remover_defeito():
        # Lacre
        lacre = campo_lacre.get()
        defeito = caixa_defeito.get()
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("UPDATE Headsets SET Manutencao=0,Estoque=1, Id_defeito=NULL,Inativo = 0 WHERE Lacre = ?",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # ATIVAR LACRE.
    def ativar_lacre():
        lacre = campo_lacre.get()
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("UPDATE Headsets SET Inativo=0,Estoque=1 WHERE lacre = ?",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # INATIVAR LACRE.
    def inativar_lacre():
        lacre = campo_lacre.get()
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()
        try:
            cursor_geral.execute("UPDATE Headsets SET Inativo = 1,Estoque=0 WHERE lacre = ?",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
        conexao_geral.close()
        busca()
    # ENTREGUE AO TREINAMENTO.
    def entrega_treinamento():
        # coletamos dados
        lacre = campo_lacre.get()
        admissao = calendario_admissao.get()
        admissao = arrow.get(str(admissao),'DD/MM/YYYY')
        admissao = admissao.format('YYYY-MM-DD')
        
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor() 

        try:
            cursor_geral.execute("INSERT INTO Treinamento (Id_headset,Admissao)VALUES((SELECT Id FROM Headsets WHERE Lacre = ?),?)",lacre,admissao)
            conexao_geral.commit()
            cursor_geral.execute("UPDATE Headsets SET Estoque=0 WHERE Lacre=?",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)    
        conexao_geral.close()    
    # DEVOLVIDO PELO TREINAMENTO.
    def devolvido_treinamento():
        # coletamos dados
        lacre = campo_lacre.get()
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()       
        try:
            cursor_geral.execute("UPDATE Treinamento SET Devolvido = 1 WHERE (SELECT Id FROM Headsets WHERE Lacre = ?) = Id_headset",lacre)
            conexao_geral.commit()
            cursor_geral.execute("UPDATE Headsets SET Estoque=1 WHERE Lacre=?",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)    
        conexao_geral.close()
    # buscar headsets associados ao treinamento.
    def busca_treinamento():
        # limpar a tabela atual
        count_assoc = 0
        campo_quadro_treinamento.delete(*campo_quadro_treinamento.get_children())

        # Vamos buscar os cartões assoiados a funcionários.
        # Dados do banco geral
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        # Verificar se o cartão existe.
        assoc = cursor_geral.execute("SELECT Treinamento.Admissao,Headsets.Lacre FROM Treinamento INNER JOIN Headsets ON Headsets.Id = Treinamento.Id_headset WHERE Treinamento.EmUso=0 AND Treinamento.Devolvido=0 ORDER BY Treinamento.Admissao").fetchall()
        conexao_geral.commit()
        for linha in assoc:
            admissao = linha[0]
            lacre = linha[1]
            count_assoc += 1

            if admissao == None:
                admissao = "NÃO REGISTRADO"
            if lacre == None:
                lacre = "NÃO REGISTRADO"

            campo_quadro_treinamento.insert("",END, values=(lacre,admissao))
    # Volocar o hgead em estoque.
    def estoque():
        # coletamos dados
        lacre = campo_lacre.get()
        dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()       
        try:
            cursor_geral.execute("UPDATE Headsets SET Estoque=1,Manutencao=0,Inativo=0 WHERE Lacre=?",lacre)
            conexao_geral.commit()
            cursor_geral.execute("UPDATE Headsets SET UltPosse=EmPosse,EmPosse=NULL WHERE Lacre = ? AND EmPosse IS NOT NULL",lacre)
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)    
        conexao_geral.close()
        busca()

    # VAMOS PREENCHER AS OPÇÔES DAS COMBOBOX
    dados_geral = (
            "driver={SQL Server Native Client 11.0};"
            "server=<SERVIDOR>;"
            "Database=<BANCO>;"
            "UID=<USUARIO>;"
            "PWD=<SENHA>;"
        )
    conexao_geral = pyodbc.connect(dados_geral)
    cursor_geral = conexao_geral.cursor()

    # MARCAS DO HEADSET
    marcas = cursor_geral.execute("SELECT Marca FROM MarcasHeadsets").fetchall()
    conexao_geral.commit()
    lista_marcas = [marca[0] for marca in marcas]

    # DEFEITOS DO HEADSET
    defeitos = cursor_geral.execute("SELECT Defeito FROM DefeitosHeadsets").fetchall()
    conexao_geral.commit()
    lista_defeitos = [defeito[0] for defeito in defeitos]

    # VENDEDORES DE HEADSET
    vendedor = cursor_geral.execute("SELECT Vendedor FROM Vendedores").fetchall()
    conexao_geral.commit()
    lista_vendedor = [vendedor[0] for vendedor in vendedor]

    # SAINDO
    conexao_geral.close()

    janela_headsets = Tk() #criando janela_headsets
    janela_headsets.title("HEADSETS") #titulo
    janela_headsets.geometry("875x505") #quadro
    janela_headsets.config(bg = "white")
    janela_headsets.iconbitmap(icon_path)
    janela_headsets.resizable(0,0)

    # VARiAVEIS DE MOVIMENTACAO
    variavel_lacre_atual = StringVar()

    ########################################## HEADSET #########################################################################################

    # DECORAÇÂO
    Label(janela_headsets,bg="white",width=40,height=13,highlightbackground="navyblue",highlightthickness=1).place(x=5,y=10)
    Label(janela_headsets,text="HEADSET",bg="white",fg="navy blue").place(x=10,y=0)


    # HEADSET
    Label(janela_headsets,bg="white",text="Lacre:").place(x=10,y=30)
    campo_lacre = Entry(janela_headsets,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_lacre.place(x=90,y=30)

    botao_busca = Button(janela_headsets,text="Buscar",command=busca,fg="navy blue") #realiza a busca.
    botao_busca.place(x=230,y=29)

    # FUNCIONARIO QUE ESTÁ USANDO
    Label(janela_headsets,bg="white",text="Funcionário:").place(x=10,y=65)
    campo_funcionario = Entry(janela_headsets,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_funcionario.place(x=90,y=65)

    # MARCA DO HEAD
    Label(janela_headsets,bg="white",text="Marca:").place(x=10,y=100)
    variavel_marca = StringVar()
    caixa_marca = Combobox(janela_headsets, textvariable=variavel_marca, values=lista_marcas)
    caixa_marca.place(x=90,y=100,width=125)

    # VENDEDOR DO HEAD
    Label(janela_headsets,bg="white",text="Vendedor:").place(x=10,y=135)
    variavel_vendedor = StringVar()
    caixa_vendedor = Combobox(janela_headsets, textvariable=variavel_vendedor, values=lista_vendedor)
    caixa_vendedor.place(x=90,y=135,width=125)

    # BOTOES
    # CADASTRAR
    Button(janela_headsets,width=15,text="Cadastrar lacre",fg="navy blue",command=cadastro_lacre).place(x=25,y=175)
    # ATUALIZAR
    Button(janela_headsets,width=15,text="Atualizar lacre",fg="navy blue",command=atualizar_lacre).place(x=155,y=175)

    ######################################## STATUS ########################################################################################################

    # DECORAÇÃO STATUS
    Label(janela_headsets,bg="white",width=40,height=17,highlightthickness=1,highlightbackground="navyblue").place(x=5,y=235)
    Label(janela_headsets,text="SITUAÇÃO",bg="white",fg="navy blue").place(x=10,y=225)

    # SITUAÇÂO DO HEAD
    Label(janela_headsets,bg="white",text="Situação:").place(x=10,y=255)
    campo_situacao = Entry(janela_headsets,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_situacao.place(x=90,y=255)

    # DEFEITO QUE APRESENTA
    variavel_defeito = StringVar()
    Label(janela_headsets,bg="white",text="Defeito:").place(x=10,y=290)
    caixa_defeito = Combobox(janela_headsets, textvariable=variavel_defeito, values=lista_defeitos)
    caixa_defeito.place(x=90,y=290,width=125)

    # GARANTIA
    Label(janela_headsets,bg="white",text="Garantia:").place(x=10,y=325)
    campo_garantia = Entry(janela_headsets,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1,width=5)
    campo_garantia.place(x=90,y=325)

    # ULTIMO CHAMADO
    Label(janela_headsets,bg="white",text="Chamado:").place(x=10,y=360)
    campo_chamado = Entry(janela_headsets,highlightcolor="#828790",highlightbackground="#828790",highlightthickness=1)
    campo_chamado.place(x=90,y=360)

    # CADASTRAR DEFEITO
    Button(janela_headsets,width=15,text="Registrar defeito",fg="navy blue",command=cadastra_defeito).place(x=25,y=400)

    # CHECKBOX PARA GARANTIA
    variavel_garantia = IntVar()
    check_desconto = Checkbutton(janela_headsets,bg="white",fg="navy blue",text="Está na garantia?",variable=variavel_garantia, offvalue=0, onvalue=1)
    check_desconto.place(x=155,y=400)

    # REMOVER DEFEITO
    Button(janela_headsets,width=15,text="Remover defeito",fg="navy blue",command=remover_defeito).place(x=25,y=432)

    # ESTOQUE
    Button(janela_headsets,width=15,text="Colocar em estoque",fg="navy blue",command=estoque).place(x=155,y=432)

    # ATIVAR
    Button(janela_headsets,width=15,text="Ativar lacre",fg="navy blue",command=ativar_lacre).place(x=25,y=464)
    # INATIVAR
    Button(janela_headsets,width=15,text="Inativar lacre",fg="navy blue",command=inativar_lacre).place(x=155,y=464)


    ############################################# QUADRO DE MANUTENÇÃO ##########################################################################

    # QUADRO DE HEADSETS.
    # DECORACAO
    Label(janela_headsets,bg="white",width=48,height=32,highlightthickness=1,highlightbackground="navy blue").place(x=300,y=10)
    Label(janela_headsets,text="HEADSETS EM MANUTENÇÃO",bg="white",fg="navy blue").place(x=305,y=0)

    # CAMPOS
    campo_quadro = ttk.Treeview(janela_headsets,columns=("c1","c2","c3"),show='headings',height=20)
    campo_quadro.place(x=305,y=20)

    # COLUNAS
    # HEADSET
    campo_quadro.column("#1",width=70)
    campo_quadro.heading("#1",text="HEADSET")
    # DEFEITO
    campo_quadro.column("#2",width=130)
    campo_quadro.heading("#2",text="DEFEITO")
    # CHAMADO
    campo_quadro.column("#3",width=130)
    campo_quadro.heading("#3",text="CHAMADO")


    # BOTOES
    botao_quadro = Button(janela_headsets,text="Atualizar lista",fg="navy blue",command=busca_defeitos)
    botao_quadro.place(x=310,y=465)

    # DASH
    Label(janela_headsets,text='Quantidade:          ',bg="white",highlightbackground="navy blue",highlightthickness=1,height=2).place(x=534,y=455)
    valor_quantidade = Label(janela_headsets,text='',bg="white")
    valor_quantidade.place(x=605,y=463)

    ########################################### TREINAMENTO ##############################################################################################

    # DECORACAO
    Label(janela_headsets,bg="white",width=30,height=32,highlightthickness=1,highlightbackground="navy blue").place(x=652,y=10)
    Label(janela_headsets,text="HEADSETS COM O TREINAMENTO",bg="white",fg="navy blue").place(x=657,y=0)

    # DATA DA ADMISSAO
    Label(janela_headsets,text="Admissão:",bg="white").place(x=657,y=25)
    calendario_admissao = tkcalendar.DateEntry(janela_headsets,date_pattern='dd/mm/yyyy',width=10)
    calendario_admissao.place(x=730,y=25)

    # BOTOES
    # ENTREGA
    Button(janela_headsets,width=15,text="Entregue",fg="navy blue",command=entrega_treinamento).place(x=657,y=50)
    # DEVOLUCAO
    Button(janela_headsets,width=15,text="Devolvido",fg="navy blue",command=devolvido_treinamento).place(x=657,y=85)

    # QUADRO DE HEADSETS COM O TREINAMENTO.
    # CAMPOS
    campo_quadro_treinamento = ttk.Treeview(janela_headsets,columns=("c1","c2"),show='headings',height=15)
    campo_quadro_treinamento.place(x=657,y=120)

    # COLUNAS
    # HEADSET
    campo_quadro_treinamento.column("#1",width=75)
    campo_quadro_treinamento.heading("#1",text="HEADSET")
    # DEFEITO
    campo_quadro_treinamento.column("#2",width=130)
    campo_quadro_treinamento.heading("#2",text="ADMISSÃO")

    # BOTOES
    botao_quadro_treinamento = Button(janela_headsets,text="Atualizar lista",fg="navy blue",command=busca_treinamento)
    botao_quadro_treinamento.place(x=657,y=465)

    #Label(janela_headsets,text="Criado por: Daniel Lopes manfrini").place(x=8,y=475)

    janela_headsets.mainloop()

if __name__ == '__main__':
    tela_head()