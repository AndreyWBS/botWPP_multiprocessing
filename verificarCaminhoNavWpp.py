import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import pandas as pd

novoCaminho = ""


def verificar_WPP_logado(valor):
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless=new")
    caminho_nav = f"user-data-dir={valor}"
    options.add_argument(caminho_nav)
    driver = webdriver.Chrome(options=options)
    driver.get('https://web.whatsapp.com/')
    wait = WebDriverWait(driver, 120)
    elemento = wait.until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, '#app > div > div > div._2Ts6i._3RGKj > header > div._3WByx')))
    return [True, driver]


# ._3iLTh


def pegarCaminho():
    seletor = "#profile_path"
    driver = webdriver.Chrome()
    driver.get('chrome://version/')
    wait = WebDriverWait(driver, 120)
    elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, seletor)))
    caminho = driver.find_element(By.CSS_SELECTOR, seletor)
    return caminho


def salvarPlanilha(caminho, planilha, i):
    planilha.at[i, 'navegadorWPP'] = str(caminho.text)
    planilha.to_excel(r'C:\Users\Micro\Desktop\gerenciar_PLanilhas\informacoesWPP.xlsx', index=False)


def colocarWPP(caminho):
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless=new")
    caminho_nav = f"user-data-dir={caminho}"
    options.add_argument(caminho_nav)
    driver = webdriver.Chrome(options=options)
    driver.get('https://web.whatsapp.com/')
    wait = WebDriverWait(driver, 10)
    elemento = wait.until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, '#app > div > div > div._2Ts6i._3RGKj > header > div._3WByx')))



def verificarNav(i, planilha):
    print(planilha.loc[i, 'navegadorWPP'])
    caminhoNAV = planilha.loc[i, 'navegadorWPP']
    if "nan" != str(caminhoNAV):
        print("o caminho não ta vazio")
        logado = verificar_WPP_logado(str(caminhoNAV).replace("\Default", ""))
        if logado[0]:
            return logado[1]

    else:
        caminho = pegarCaminho()
        salvarPlanilha(caminho, planilha, i)
        print(caminho.text)
        colocarWPP(str(caminho.text))
        print("o caminho esta vazio")
        return caminho
