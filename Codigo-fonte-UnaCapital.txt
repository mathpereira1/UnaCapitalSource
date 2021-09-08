import time
import os
import pandas as pd
import plotly.graph_objects as go
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def acesso_url():
    url = "https://candidato:DEVUNA2021@dev2021lab.unacapital.com.br/#"
    option = Options()
    option.headless = True
    driver = webdriver.Chrome(options=option)
    driver.get(url)
    time.sleep(5)
    return driver


def encontra_elemento():
    element = driver.find_element_by_xpath(
        "//div[@class='table-responsive']//table")
    html_content = element.get_attribute('outerHTML')
    return html_content


def gera_diretorio():
    if not os.path.exists("images"):
        os.mkdir("images")


def gera_grafico():
    fig = go.Figure(data=[go.Candlestick(x=data_frame['Data'],
                                         open=data_frame['Abertura'],
                                         high=data_frame['Máxima'],
                                         low=data_frame['Mínima'],
                                         close=data_frame['Fechamento'])])
    # Inversão do Eixo X para que o ponto inicial do gráficos seja a primeira data, ou seja, partindo da mais anterior à atual
    fig['layout']['xaxis']['autorange'] = "reversed"

    fig.write_image("images/fig1.png")
    return fig


def mostra_grafico():
    graph = input("Deseja visualizar o gráfico interativo ? (y/n)")
    while graph not in ['Y', 'y', 'n', 'N']:
        print("Erro, favor digitar uma das opções entre 'y' (Sim) ou 'n' (Não).")
        graph = input("Deseja visualizar o gráfico interativo ? (y/n)")
    if graph in ['Y', 'y']:
        fig.show()


def gera_excel():
    datatoexcel = pd.ExcelWriter('Cotacoes.xlsx')

    data_frame.to_excel(datatoexcel)

    datatoexcel.save()

    print('Os dados foram escritos em excel com sucesso.')


def mostra_excel():
    excel = input("Deseja abrir a planilha com os dados ? (y/n)")
    while excel not in ['Y', 'y', 'n', 'N']:
        print("Erro, favor digitar uma das opções entre 'y' (Sim) ou 'n' (Não).")
        excel = input("Deseja abrir a planilha com os dados ? (y/n)")
    if excel in ['Y', 'y']:
        os.system('start excel.exe Cotacoes.xlsx')


driver = acesso_url()

html_content = encontra_elemento()

# BeautifulSoup - Parsear o conteúdo HTML
soup = BeautifulSoup(html_content, 'html.parser')
table = soup.find(name='table')

# Panda - Estruturar conteúdo em um Data Frame
data_frame = pd.read_html( str(table) )[0]
data_frame.columns = ['Data', 'Abertura', 'Máxima', 'Mínima', 'Fechamento', 'Média', 'Variação', 'Volume', 'Ações', 'Negociações']

gera_diretorio()

gera_excel()

fig = gera_grafico()

mostra_excel()

mostra_grafico()
