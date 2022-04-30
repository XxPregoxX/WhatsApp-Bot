from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import urllib.parse
import PySimpleGUI as Sg
# Lê a planilia do excel
Phones = pd.read_excel("3.xlsx")
# i == ao total de vezes que o laço se repete, ou seja é a quantidade de mensagens enviadas
i = 0
# Inteface inicial
Sg.theme("Black")
layout = [
 [Sg.Text("Mensagem"), Sg.Input(size=(20, 1))],
 [Sg.Text("Xpath", size=(8, 1)), Sg.Input(size=(20, 1))],
 [Sg.Button("Rodar")]
]
t = 0
janela = Sg.Window("Insira a mensagem que deseja enviar", layout)

while True:
    event, message = janela.read()
# Lê a mensagem escrita na interface
    parse = (message[0])
# Lê o Xpath escrito na interface
    Xpath = (message[1])
    if message == Sg.WIN_CLOSED:
        break
    elif event == "Rodar":
        t += 1
        break
# Inicia o procedimento
if t == 1:
    # Prepara o texto
    text = urllib.parse.quote(parse)
    driver = webdriver.Chrome()

    # Abre o Whatsapp
    driver.get("https://web.whatsapp.com/")
    # Caso a tabela de contatos lateral não exista ele fica esperando até aparecer
    while len(driver.find_elements_by_id("side")) < 1:
        time.sleep(10)
    # Ele lê a primeira linha que deve ser preenchida com a palavra "título" e vai lendo os numeros um por um à medida
    # que o laço vai se repetindo
    for i, nada in enumerate(Phones["titulo"]):

        numero = Phones.loc[i, "titulo"]
        # Ele abre o navegador "Google Chrome" e abre o Whatsapp
        link = f"https://web.whatsapp.com/send?phone={numero}&text={text}"

        driver.get(link)
        # Caso a tabela de contatos lateral não exista ele fica esperando até aparecer
        while len(driver.find_elements_by_id("side")) < 1:
            time.sleep(10)
        # Lê o xpath inserido e verifica se existe, caso exista ele aperta ENTER e envia a mensagem,
        # A verificação é feita para evitar erros
        if len(driver.find_elements_by_xpath(Xpath)) >= 1:
            driver.find_element_by_xpath(Xpath).send_keys(Keys.ENTER)
        # Caso o xpath seja invalido, uma interface é aberta para mostrar quantas mensagens foram enviadas
        # isso serve para quando o programa trminar de ler ou caso o numero seja invalido, serve para varios outros
        # erros, mas o objetivo principal é para informar o usuário aonde parou para que não precise ser feito
        # manualmente
        else:
            layout2 = [
                [Sg.Text(i)]
            ]

            colher = Sg.Window("Mensagens enviadas", layout2)
            while True:
                d, f = colher.read()
                # Se a tabela de erros for fechada, o programa encerra
                if f == Sg.WIN_CLOSED:
                    raise ValueError
        # Espera 1 minuto para evitar o banimento do aplicativo
        time.sleep(60)
