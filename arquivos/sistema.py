import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
from datetime import datetime
import openpyxl
import requests
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()


lista_moedas =  list(dicionario_moedas.keys())

def pegar_cotacao():
    moeda = combobox_selecionar_moedas.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]

    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"

    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda  = cotacao[0]['bid']
    print(valor_moeda)
    label_texto_cotacao['text'] = f'A cotação da {moeda} no dia {data_cotacao} foi de R${valor_moeda}'


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='selecione o caminho do arquivo de moeda')
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivo_selecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"
        

def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario_data_inicial.get()
        data_final = calendario_data_final.get()

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}"

        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao['bid'])
            data = datetime.timestamp(timestamp)
            data = datetime.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan

            df.loc[df.iloc[:,0] == moeda, data] = bid

        df.to_excel('texte.xlsx')
        label_atualizar_cotacoes['text'] = "Arquivo Atualizado com Sucesso"

    except:
        label_atualizar_cotacoes['text'] ="Selecione um Arquivo Excel no Formato Correto"


janela = tk.Tk()

janela.title('Ferramenta de Cotação de Moeda')


label_cotacao_moeda = tk.Label(text='Cotação de uma Moeda Específica', borderwidth=2, relief='solid')
label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)


label_selecionar_moeda = tk.Label(text='Selecionar Moeda', anchor='e')
label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

# Combobox para seleção da moeda
combobox_selecionar_moedas = ttk.Combobox(value=lista_moedas)
combobox_selecionar_moedas.grid(row=1, column=2, padx=10, pady=10, sticky='nswe', columnspan=1)


label_selecionar_dia_cotacao = tk.Label(text='Selecione o Dia que Deseja Pegar a Cotação', anchor='e')
label_selecionar_dia_cotacao.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)


calendario_moeda = DateEntry(year=2023, locale='pt_br', )
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')


label_texto_cotacao = tk.Label(text='')
label_texto_cotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe')

botao_pegar_cotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')


# Configuração e posicionamento do rótulo para cotação de várias moedas
label_cotacao_varias_moeda = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacao_varias_moeda.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)


label_secionar_arquivo = tk.Label(text='Selecione um Arquivo em Excel com as Moedas na Coluna A')
label_secionar_arquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')


var_caminhoarquivo = tk.StringVar()

botao_secionar_arquivo = tk.Button(text='Selecione o Arquivo', command=selecionar_arquivo)
botao_secionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

label_arquivo_selecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nswe')

label_data_inicial = tk.Label(text='Data Inicial')
label_data_inicial.grid(row=7, column=0, padx=10,pady=10,sticky='nswe')

calendario_data_inicial =DateEntry(year=2023, locate='pt_br')
calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10,sticky='nswe')

label_data_final = tk.Label(text='Data Final')
label_data_final.grid(row=8, column=0, padx=10,pady=10,sticky='nswe')

calendario_data_final =DateEntry(year=2023, locate='pt_br')
calendario_data_final.grid(row=8, column=1, padx=10, pady=10,sticky='nswe')



label_atualizar_cotacoes = tk.Label(text='')
label_atualizar_cotacoes.grid(row=9, column=2, columnspan=2, padx=10,pady=10, sticky='nswe')

botao_atualizar_cotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0,padx=10,pady=10, sticky='nswe')

botao_fechar = tk.Button(text='fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2,padx=10,pady=10, sticky='nswe')

janela.mainloop()