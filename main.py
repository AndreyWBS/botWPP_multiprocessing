import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import pandas as pd

nomePLA = ""
nomeCVS = ""
caminhoPlanilha = r"C:\Users\Micro\Desktop\planilhasBot\planilhaCCKbot7000-8000.xlsx"

planilha_de_numeros_env = {
    "numero": [],
    "codigo": []
}
planilha_de_numeros_Erro = {
    "numero": [],
    "codigo": []
}


def procurarElementocomSeletor(seletor, driver):
    wait = WebDriverWait(driver, 120)
    elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, seletor)))
    return driver.find_element(By.CSS_SELECTOR, seletor)


def procurarElementscomseletor(seletor, driver):
    wait = WebDriverWait(driver, 25)
    elemento = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, seletor)))
    return driver.find_elements(By.CSS_SELECTOR, seletor)


def clicar_seletor_elemento(seletor, elementoP, driver):
    wait = WebDriverWait(driver, 25)
    elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, seletor)))
    elementoP.find_element(By.CSS_SELECTOR, seletor).click()


def escrever(seletor, txt, driver):
    wait = WebDriverWait(driver, 25)
    elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, seletor)))
    for letra in txt:
        if letra == "#":
            driver.find_element(By.CSS_SELECTOR, seletor).send_keys(Keys.CONTROL, Keys.ENTER)
        else:
            driver.find_element(By.CSS_SELECTOR, seletor).send_keys(letra)
    driver.find_element(By.CSS_SELECTOR, seletor).send_keys(Keys.ENTER)


def gerarPLanilahas(nomeWPP):
    df_Planilha_com_ERRO = pd.DataFrame(planilha_de_numeros_Erro)
    df_Planilha_com_ENV = pd.DataFrame(planilha_de_numeros_env)
    nome_env = f"{nomeWPP}numeros_env.xlsx"
    nome_erro = f"{nomeWPP}numeros_erros.xlsx"
    caminho_planilha = r"C:\Users\Micro\Desktop\planilhasGeradasBOT\acionamentosWPP\pla"

    df_Planilha_com_ERRO.to_excel(caminho_planilha + nome_erro, index=False)
    df_Planilha_com_ENV.to_excel(caminho_planilha + nome_env, index=False)


def obterPrimeiraLinha(texto):
    linhas = texto.split("\n")
    return linhas[0]


def comecar(ondeMandarNumeros, TextoMandar, nomeWPP, caminhoPlanilha, driver):
    # TODO : fazer um if que verificar se tem nome e carteira na planilha e se tiver colocar dentro do texto no lugar
    #  das {} (chaves)

    contatos_df = pd.read_excel(caminhoPlanilha)

    try:

        nomeCVS = procurarElementscomseletor("._21S-L .l7jjieqr", driver)

        for i, mensagem in enumerate(contatos_df['numero']):
            try:
                len(contatos_df["nome"])
                TextoMandar = str(TextoMandar).replace("{nome}", str(contatos_df["nome"]))
            except:
                print("não é do diversos")

            numeroTELL = contatos_df.loc[i, "numero"]
            codigo = contatos_df.loc[i, "codigo"]
            print(str(numeroTELL))

            for elemt in nomeCVS:
                print(ondeMandarNumeros)
                if elemt.text == ondeMandarNumeros:
                    print("oi")
                    elemt.click()
                    escrever("#main .iq0m558w", str(numeroTELL), driver)
                    conversa = procurarElementscomseletor("._1jHIY , .ooty25bp", driver)

                    for linha in conversa:

                        linhatxt = obterPrimeiraLinha(linha.text)

                        if linhatxt == str(numeroTELL):
                            print("esse numero é igual")

                            time.sleep(3)
                            clicar_seletor_elemento("span._11JPr.selectable-text.copyable-text > span > a", linha,
                                                    driver)

                            try:
                                conversarCOM = procurarElementocomSeletor(".iWqod._1MZM5._2BNs3.nqtxkp62.btzf6ewn",
                                                                          driver)
                                if conversarCOM.get_attribute("aria-label") == "Conversar com ":
                                    conversarCOM.click()
                                    time.sleep(5)

                                    texto = str(TextoMandar)
                                    escrever("._3Uu1_  .g0rxnol2  .iq0m558w", "boa tarde", driver)
                                    escrever("._3Uu1_  .g0rxnol2  .iq0m558w", texto, driver)
                                    planilha_de_numeros_env["numero"].append(str(numeroTELL))
                                    planilha_de_numeros_env["codigo"].append(str(codigo))
                                    gerarPLanilahas(nomeWPP)
                                    break
                            except:

                                planilha_de_numeros_Erro["numero"].append(str(numeroTELL))
                                planilha_de_numeros_Erro["codigo"].append(str(codigo))
                                gerarPLanilahas(nomeWPP)
                                break
                    nomeCVS = procurarElementscomseletor("._21S-L .l7jjieqr", driver)

                    break

        gerarPLanilahas(nomeWPP)
        driver.quit()
    except:
        gerarPLanilahas(nomePLA)
