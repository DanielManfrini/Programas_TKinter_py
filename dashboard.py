#Importadores GLOBAL
from datetime import datetime

#importadores SQL
import pyodbc

#variaveis Tkinter
from tkinter import *
from tkinter.ttk import Combobox

#importadores caminhos
import os
import sys

#CAMINHOS
#Caminho da logo.
logo_name = 'logo.png'
icon_name = 'icon.ico'
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
logo_path = os.path.join(application_path, logo_name)
icon_path = os.path.join(application_path, icon_name)

def dash():
    def busca():
        data = datetime.today()
        dia = data.date()
        dia_atual = data.strftime('%Y-%m-%d')
        primeiro_dia = dia.strftime('%Y-%m-01')

        mes_busca = mes.get()

        if str(mes_busca) == "JANEIRO":
            mes_busca = '01'
        if str(mes_busca) == "FEVEREIRO":
            mes_busca = '02'
        if str(mes_busca) == "MARÇO":
            mes_busca = '03'
        if str(mes_busca) == "ABRIL":
            mes_busca = '04'
        if str(mes_busca) == "MAIO":
            mes_busca = '05'
        if str(mes_busca) == "JUNHO":
            mes_busca = '06'
        if str(mes_busca) == "JULHO":
            mes_busca = '07'
        if str(mes_busca) == "AGOSTO":
            mes_busca = '08'
        if str(mes_busca) == "SETEMBRO":
            mes_busca = '09'
        if str(mes_busca) == "OUTUBRO":
            mes_busca = '10'
        if str(mes_busca) == "NOVEMBRO":
            mes_busca = '11'
        if str(mes_busca) == "DEZEMBRO":
            mes_busca = '12'

        inicio_mes = ('2022-'+mes_busca+'-01')
        fim_mes = ('2022-'+mes_busca+'-31')

        print (primeiro_dia,dia_atual,inicio_mes,fim_mes)
        # Conectando ao banco

        dados_geral = (
                "driver={SQL Server Native Client 11.0};"
                "server=<SERVIDOR>;"
                "Database=<BANCO>;"
                "UID=<USUARIO>;"
                "PWD=<SENHA>;"
            )
        conexao_geral = pyodbc.connect(dados_geral)
        cursor_geral = conexao_geral.cursor()

        # ESTOQUE
        try:
            estoque = cursor_geral.execute("SELECT lacre FROM HEADSETS WHERE Estoque=1 AND Inativo = 0 AND Manutencao = 0 AND Emprestado = 0 AND Id_marca != 4").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            manutencao = cursor_geral.execute("SELECT lacre FROM HEADSETS WHERE Manutencao=1 AND Inativo=0 ").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            trocas = cursor_geral.execute(f"SELECT * FROM Trocasheadsets WHERE Id_motivo = 2 AND Data BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()
        except:
            fim_mes = ('2022-'+mes_busca+'-30')
            try:
                trocas = cursor_geral.execute(f"SELECT * FROM Trocasheadsets WHERE Id_motivo = 2 AND Data BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
                conexao_geral.commit()
            except pyodbc.Error as erro:
                print(erro)
                conexao_geral.close()

        try:
            entregas = cursor_geral.execute(f"SELECT * FROM Trocasheadsets WHERE Id_motivo = 1 AND Id_headset_novo NOT IN (SELECT Id_headset FROM Treinamento) AND Data BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            treinamento = cursor_geral.execute(f"SELECT * FROM Treinamento WHERE Data BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()

        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            emprestimo = cursor_geral.execute("SELECT lacre FROM HEADSETS WHERE Emprestado = 1").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            devolvidos = cursor_geral.execute(f"SELECT * FROM Funcionarios WHERE HeadDevolvido = 1 AND DataDemissao BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            demitidos = cursor_geral.execute(f"SELECT * FROM Funcionarios WHERE Situacao = 1 AND RecebeuHead = 1 AND HeadDevolvido = 0 AND DataDemissao BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        try:
            descontos = cursor_geral.execute(f"SELECT * FROM Descontos WHERE Data BETWEEN '{inicio_mes}' AND '{fim_mes}'").fetchall()
            conexao_geral.commit()
        except pyodbc.Error as erro:
            print(erro)
            conexao_geral.close()

        conexao_geral.close()

        # Filtrando os dados
        #estoque
        count = 0
        for lacre in estoque:
            count += 1
        texto_estoque['text'] = count 

        # trocas
        count = 0
        for troca in trocas:
            count += 1
        texto_trocas['text'] = count 

        # entregas
        count = 0
        for entrega in entregas:
            count += 1
        texto_entrega['text'] = count

        # Manutencao
        count = 0
        for lacre in manutencao:
            count += 1
        texto_manutencao['text'] = count 

        # treinamento
        count = 0
        for lacre in treinamento:
            count += 1
        texto_integracao['text'] = count 

        count = 0
        for lacre in emprestimo:
            count += 1
        texto_emprestimo['text'] = count

        # devolvidos treinamento
        count = 0
        for lacre in devolvidos:
            count += 1
        texto_recolhimentos['text'] = count

        # devolvidos treinamento
        count = 0
        for lacre in demitidos:
            count += 1
        texto_demissoes['text'] = count

        # devolvidos treinamento
        count = 0
        valor = 0
        for desconto in descontos:
            count += 1
            valor += 60
        texto_descontos['text'] = count
        texto_retorno['text'] = str(valor)+"$"

    #inicio da tela
    janela_dash = Tk()
    janela_dash.title("Dashboard") #titulo
    janela_dash.geometry("390x505") #quadro
    janela_dash.config(bg = "white")
    janela_dash.resizable(0,0)
    janela_dash.iconbitmap(icon_path)
    logo = PhotoImage(file=str(logo_path))
    Label(janela_dash, image = logo, bg="white").place(x=0,y=0)

    # DECORAÇÃO
    Button(janela_dash,fg="navy blue",text="ATUALIZAR",command=busca).place(x=300,y=15)

    Label(janela_dash,bg="white",text="MÊS:").place(x=130,y=17)

    lista_mes = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]

    data = datetime.today()
    dia = data.date()
    qual_mes = dia.strftime('%m')
    qual_mes = int(qual_mes) - 1
    qual_mes = lista_mes[qual_mes]
    print(qual_mes)

    mes = StringVar()
    caixa_mes = Combobox(janela_dash,textvariable=mes,width=15,values=("JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"))
    caixa_mes.place(x=170,y=17)

    caixa_mes.insert('0',qual_mes)

    Label(janela_dash,bg="white",width=53,height=29,highlightbackground="navy blue",highlightthickness=1).place(x=5,y=55)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=4,width=50).place(x=15,y=70)
    Label(janela_dash,bg="white",text="ESTOQUE",fg="navy blue").place(x=30,y=60)
    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=107)

    Label(janela_dash,text="EM ESTOQUE: ",bg="white").place(x=30,y=80)
    texto_estoque = Label(janela_dash,text="",bg="white")
    texto_estoque.place(x=320,y=80)
    Label(janela_dash,bg="white",text="MANUTENÇÃO: ").place(x=30,y=115)
    texto_manutencao = Label(janela_dash,text="",bg="white")
    texto_manutencao.place(x=320,y=115)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=155)
    Label(janela_dash,bg="white",text="MOVIMENTAÇÕES",fg="navy blue").place(x=30,y=145)

    Label(janela_dash,bg="white",text="ENTREGAS: ").place(x=30,y=165)
    texto_entrega = Label(janela_dash,text="",bg="white")
    texto_entrega.place(x=320,y=165)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=192)
    Label(janela_dash,bg="white",text="TROCAS: ").place(x=30,y=200)
    texto_trocas = Label(janela_dash,text="",bg="white")
    texto_trocas.place(x=320,y=200)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=228)
    Label(janela_dash,bg="white",text="INTEGRAÇÃO: ").place(x=30,y=236)
    texto_integracao = Label(janela_dash,text="",bg="white")
    texto_integracao.place(x=320,y=236)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=264)
    Label(janela_dash,bg="white",text="EMPRÉSTIMO: ").place(x=30,y=273)
    texto_emprestimo = Label(janela_dash,text="",bg="white")
    texto_emprestimo.place(x=320,y=273)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=315)
    Label(janela_dash,bg="white",text="TURNOVER",fg="navy blue").place(x=30,y=305)

    Label(janela_dash,bg="white",text="HEADSETS RECOLHIDOS: ").place(x=30,y=325)
    texto_recolhimentos = Label(janela_dash,text="",bg="white")
    texto_recolhimentos.place(x=320,y=325)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=352)
    Label(janela_dash,bg="white",text="HEADSETS NÃO RECOLHIDOS: ").place(x=30,y=362)
    texto_demissoes = Label(janela_dash,text="",bg="white")
    texto_demissoes.place(x=320,y=362)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=405)
    Label(janela_dash,bg="white",text="COBRANÇA",fg="navy blue").place(x=30,y=395)

    Label(janela_dash,bg="white",text="DESCONTOS EM FOLHA: ").place(x=30,y=415)
    texto_descontos = Label(janela_dash,text="",bg="white")
    texto_descontos.place(x=320,y=415)

    Label(janela_dash,bg="white",highlightbackground='navy blue',highlightthickness=1,height=2,width=50).place(x=15,y=440)
    Label(janela_dash,bg="white",text="RETORNO PARA MANUTENÇÃO: ").place(x=30,y=450)
    texto_retorno = Label(janela_dash,text="",bg="white")
    texto_retorno.place(x=320,y=450)

    janela_dash.mainloop()

if __name__ == '__main__':
    dash()