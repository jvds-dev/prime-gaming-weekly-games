from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

import re

import time
import openpyxl

def colored_print(text, color='reset'):
    colors = {
        "reset": "\033[0m",
        "red": "\033[91m",
        "green": "\033[92m",
        "yellow": "\033[93m",
        "blue": "\033[94m",
        "magenta": "\033[95m",
        "cyan": "\033[96m",
        "white": "\033[97m",
    }

    if color in colors:
        color_code = colors[color]
        print(f"{color_code}{text}{colors['reset']}")
    else:
        print(text)

def webdriver_config():
    global driver
    chrome_driver_path = r'./chromedriver.exe'
    service = Service(chrome_driver_path)
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get('https://gaming.amazon.com/home')
    time.sleep(2)

def locate_weekly_games():
    button = driver.find_element(By.CLASS_NAME, 'offer-filters__button__title--Game')
    button.click()
    time.sleep(1)

def scroll_to_element(element):
    driver.execute_script("arguments[0].scrollIntoView();", element)

def get_all_games():
    game_divs = driver.find_elements(By.CLASS_NAME, 'tw-block')
    global games
    games = []

    for game_div in game_divs:
        # Rolar até elementos antes de extrair texto
        scroll_to_element(game_div) #garante que elementos serão carregados
        all_text = game_div.text
        game_name = re.findall(r'dias\n(.+?)\n', all_text)
        
        if game_name:
            games.append(game_name)


    for i, game in enumerate(games): #Transforma os elementos de lista para string ( [['um'], ['dois']] --> ['um', 'dois])
        if isinstance(game, list):
            games[i] = game[0]

    games = list(set(games)) #Remove itens repitidos

def create_and_save_sheet():
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet["A1"] = 'Nome do jogo'

    index = 2
    for game in games:
        sheet[f"A{index}"] = game
        index+=1
    
    workbook.save('Free-Weekly-Games.xlsx')
    workbook.close()


def start():
    colored_print('1. Iniciando selenium...', 'magenta')
    webdriver_config()
    colored_print('     > Selenium iniciado', 'green')

    colored_print('2. Localizando jogos...', 'magenta')
    locate_weekly_games()
    get_all_games()
    colored_print(f'     > {len(games)} Jogos identificados', 'green') 

    colored_print('3. Criando e salvando informações na planilha...', 'magenta')
    create_and_save_sheet()
    colored_print('     > Planilha salva', 'green')

    colored_print('< EXECUÇÃO FINALIZADA >', 'yellow')

start()