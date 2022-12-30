import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import date, datetime
import xml.etree.ElementTree as ET
import requests
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os

dia_atual = date.today()
ano_atual = dia_atual.year

colorfont1 = "#dad6ca"
corfundo2 = "#1bb0ce"
corfundo = "#4f8699"
corfundo1 = "#6a5e72"
corfundo3 = "#563444"

moedas = requests.get('https://economia.awesomeapi.com.br/xml/available/uniq')
requisicao = requests.get("https://economia.awesomeapi.com.br/json/all")
dicionario_moedas = requisicao.json()

moedas_xml = ET.fromstring(moedas.content)

moedas_final = [moeda.tag for moeda in moedas_xml if moeda.tag != 'xml']


def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requiMoeda = requests.get(link)
    cotacao = requiMoeda.json()
    try:
        valor_moeda = cotacao[0]['bid']
        label_resultadocotacao['text'] = f'A cotacao da moeda {moeda} no dia {data_cotacao} foi de: R$ {valor_moeda}'
    except Exception:
        label_resultadocotacao['text'] = 'Não foi possivel pegar a cotação dessa moeda!'


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(
        title='Selecione o arquivo de moeda')
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo selecionado: {caminho_arquivo}'
    else:
        label_arquivoselecionado['text'] = 'Nenhum arquivo selecionado'


def atualizar_cotacoes():
    try:
        # ler o dataframe de moedas
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/10000?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'
            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')

                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
                
        if os.path.isfile("ArquivoCotacoes.xlsx"):
            os.remove("ArquivoCotacoes.xlsx")

        df.to_excel("ArquivoCotacoes.xlsx")
        label_arquivoatualizado['text'] = "Arquivo atualizado com sucesso"
    except Exception:
        label_arquivoatualizado['text'] = "Selecione um arquivo no xlsx no formato correto"


janela = tk.Tk()

janela.title('Ferramenta de cotação de moedas')
janela.configure(bg=corfundo)

label_cotacaomoeda = tk.Label(
    text='Cotação de uma moeda especifica', borderwidth=2, relief='solid', bg=corfundo1, fg=colorfont1)
label_cotacaomoeda.grid(row=0, column=0, columnspan=3,
                        sticky="NSEW", padx=10, pady=10)

label_selecionemoeda = tk.Label(
    text='Selecione moeda.:', bg=corfundo, fg=colorfont1, anchor='e')
label_selecionemoeda.grid(row=1, column=0, columnspan=2,
                          sticky='nsew', padx=10, pady=10)

combobox_selecionarmoeda = ttk.Combobox(
    values=moedas_final)
combobox_selecionarmoeda.grid(
    row=1, column=2, sticky='nsew', padx=10,  pady=10)

label_selecioneodia = tk.Label(
    text='Selecione o dia que deseja pegar a cotação.:', anchor='e', bg=corfundo, fg=colorfont1)
label_selecioneodia.grid(row=2, column=0, columnspan=2,
                         sticky='nsew', padx=10, pady=10)

calendario_moeda = DateEntry(year=ano_atual, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

label_resultadocotacao = tk.Label(
    text="", bg=corfundo, fg=colorfont1, anchor='e')
label_resultadocotacao.grid(
    row=3, column=0, columnspan=2, sticky='nsew', padx=10, pady=10)

button_pegarcotacao = tk.Button(
    text='Pegar Cotacao', command=pegar_cotacao, background=corfundo2)
button_pegarcotacao.grid(
    row=3, column=2, sticky='nsew', padx=10, pady=10)


# Cotação de varias moedas

label_cotacoesmoedas = tk.Label(
    text='Cotações de multiplas moedas especificas', borderwidth=2, relief='solid', bg=corfundo1, fg=colorfont1)
label_cotacoesmoedas.grid(row=4, column=0, columnspan=3,
                          sticky="NSEW", padx=10, pady=10)


label_cotacaovariasmoedas = tk.Label(
    text='Selecione um arquivo em Excel com as Moedas na Coluna A.:', bg=corfundo, fg=colorfont1)
label_cotacaovariasmoedas.grid(row=5, column=0, columnspan=2,
                               sticky='nsew', padx=10, pady=10)

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(
    text='Clique para selecionar', command=selecionar_arquivo, background=corfundo2)
botao_selecionararquivo.grid(
    row=5, column=2, sticky='nsew', padx=10,  pady=10)

label_arquivoselecionado = tk.Label(
    text='Nenhum arquivo selecionado', anchor='e', bg=corfundo, fg=colorfont1)
label_arquivoselecionado.grid(
    row=6, column=0, columnspan=3, sticky='nsew', padx=10, pady=10)


label_datainicial = tk.Label(
    text='Data Inicial.:', bg=corfundo, fg=colorfont1, anchor='e')
label_datafinal = tk.Label(
    text='DataFinal.:', bg=corfundo, fg=colorfont1, anchor='e')
label_datainicial.grid(row=7, column=0, columnspan=2,
                       sticky='nsew', padx=10, pady=10)
label_datafinal.grid(row=8, column=0, columnspan=2,
                     sticky='nsew', padx=10, pady=10)

calendario_datainicial = DateEntry(year=ano_atual, locale='pt_br')
calendario_datafinal = DateEntry(year=ano_atual, locale='pt_br')
calendario_datafinal.grid(row=8, column=2, padx=10, pady=10, sticky='nsew')
calendario_datainicial.grid(row=7, column=2, padx=10, pady=10, sticky='nsew')

botao_atualizarcotacoes = tk.Button(
    text='Atualizar Cotações', command=atualizar_cotacoes, background=corfundo2)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')
label_arquivoatualizado = tk.Label(
    text='', bg=corfundo, fg=colorfont1, anchor='e')
label_arquivoatualizado.grid(
    row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')


botao_fechar = tk.Button(
    text='Fechar', command=janela.quit, background=corfundo2)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()
