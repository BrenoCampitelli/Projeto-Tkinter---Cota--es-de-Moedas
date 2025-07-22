import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import pandas as pd
import requests
from datetime import datetime
import numpy as np


requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_textocotacao['text'] = f'A cotação de {moeda} no dia {data_cotacao} foi de: R${valor_moeda}.'


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'


# def atualizar_cotacoes():
#     # ler o dataframe de moedas
#     df = pd.read_excel()
#     moedas = df.iloc[:, 0]
#
#     # pegar a data de inicio e data de fim das cotações
#     data_inicial = calendario_datainicial.get()
#     data_final = calendario_datafinal.get()
#
#     ano_inicial = data_inicial[-4:]
#     mes_inicial = data_inicial[3:5]
#     dia_inicial = data_inicial[:2]
#
#     ano_final = data_final[-4:]
#     mes_final = data_final[3:5]
#     dia_final = data_final[:2]
#
#     # para cada moeda
#         # pegar todas das cotações daquela moeda
#         # criar uma coluna em um novo dataframe com todas as cotações daquela moeda
#     for moeda in moedas:
#         link = (f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?'
#                 f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&'
#                 f'end_date={ano_final}{mes_final}{dia_final}')
#
#         requisicao_moeda = requests.get(link)
#         cotacoes = requisicao_moeda.json()
#         for cotacao in cotacaos:
#             timestamp = int(cotacao['timestamp'])
#             bid = float(cotacao['bid'])
#             data = datetime.fromtimestamp(timestamp)
#             data = datetime.strftime('%d/%m/%Y')
#             if data not in df:
#                 print(data)
#                 df[data] = np.nan
#
#             df.loc[df.iloc[:, 0] == moeda , data] = bid
#
#     # criar um arquivo com todas as cotações
#     df.to_excel('Teste.xlsx')
#     label_atualizarcotacao['Text'] = 'Arquivo Atualizado com Sucesso!'

def atualizar_cotacoes():
    caminho = var_caminhoarquivo.get()  # Pega o caminho do arquivo Excel selecionado

    if not caminho:
        messagebox.showerror("Erro", "Selecione um arquivo antes de atualizar.")
        return

    try:
        # Ler o dataframe de moedas a partir do arquivo Excel
        df = pd.read_excel(caminho)
        moedas = df.iloc[:, 0]

        # Pegar as datas selecionadas
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        # Loop por cada moeda no Excel
        for moeda in moedas:
            link = (
                f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?'
                f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&'
                f'end_date={ano_final}{mes_final}{dia_final}'
            )

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()

            for cotacao in cotacoes:  # Corrigido: antes estava `cotacaos`
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp).strftime('%d/%m/%Y')

                if data not in df.columns:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid

        # Salva o novo Excel com as cotações atualizadas
        df.to_excel('Teste.xlsx', index=False)
        label_atualizarcotacao['text'] = 'Arquivo Atualizado com Sucesso!'  # Corrigido: 'Text' → 'text'

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


janela = tk.Tk()

janela.title('Ferramenta de Cotação de Moedas')

label_cotacaomoeda = tk.Label(text='Cotação de uma Moeda Específica', borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionarmoeda = tk.Label(text='Selecionar Moeda:', anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

label_selecionardia = tk.Label(text='Selecione o dia que deseja pegar a cotação:', anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2025, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

label_textocotacao = tk.Label(text = '')
label_textocotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

botao_pegarcotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')


# Cotação de várias moedas

label_cotacaovariasmoedas = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacaovariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionararquivo = tk.Label(text='Selecione um arquivo em Excel com as Moedas na Coluna A')
label_selecionararquivo.grid(row=5, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

label_arquivoselecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_datainicial = tk.Label(text='Data Inicial:', anchor='e')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')

label_datafinal = tk.Label(text='Data Final:', anchor='e')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

calendario_datainicial = DateEntry(year=2025, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')

calendario_datafinal = DateEntry(year=2025, locale='pt_br')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')

botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')

label_atualizarcotacao = tk.Label(text='')
label_atualizarcotacao.grid(row=9, column=1, padx=10, pady=10, sticky='nswe', columnspan=2)

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

janela.mainloop()
