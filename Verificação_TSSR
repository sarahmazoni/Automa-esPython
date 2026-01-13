from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import os
from datetime import datetime

# ===============================
# CARREGAR VARIÁVEIS DE AMBIENTE
# ===============================
load_dotenv()

CAMINHO_OCS = os.getenv("CAMINHO_OCS")
URL = os.getenv("MRIC_URL")
USUARIO = os.getenv("MRIC_USUARIO")
SENHA = os.getenv("MRIC_SENHA")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR")
LOG_PATH = "log_execucao_TSSR.txt"


def registrar_log(mensagem):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(mensagem + "\n")


def aguardar_carregamento(driver):
    try:
        loading_element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'load-spinner__content')]")
            )
        )
        WebDriverWait(driver, 60).until_not(
            EC.visibility_of(loading_element)
        )
    except:
        pass


def tratar_popup_sem_dados(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(text(),'Nenhum dado encontrado')]")
            )
        )

        botao_ok = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/div/div[2]/div/div[2]/div/ul/li/div/span")
            )
        )
        botao_ok.click()
        time.sleep(2)
        return True
    except:
        return False


def iniciar_bot():
    options = webdriver.ChromeOptions()
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get(URL)

    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "details-button"))
        ).click()
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "proceed-link"))
        ).click()
    except:
        pass

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div/div/div/div/div[2]/div[2]/input")
        )
    ).send_keys(USUARIO + Keys.RETURN)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div/div/div/div/div[3]/div[2]/div/input")
        )
    ).send_keys(SENHA + Keys.RETURN)

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div/div/div/div/div[6]/div[2]/span")
        )
    ).click()

    aguardar_carregamento(driver)
    return driver


def acessar_ordens(driver):
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[1]/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[5]/div[2]/ul/li[1]/div",
            )
        )
    ).click()


def tratar_ordem(driver, oc):
    actions = ActionChains(driver)
    baixou_algo = False

    campo_oc = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/input",
            )
        )
    )
    campo_oc.send_keys(Keys.CONTROL + "a")
    campo_oc.send_keys(oc)
    campo_oc.send_keys(Keys.RETURN)
    aguardar_carregamento(driver)

    primeiro_nome_atividade = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[6]/div/div[3",
            )
        )
    )
    actions.double_click(primeiro_nome_atividade).perform()
    aguardar_carregamento(driver)
    time.sleep(5)

    aba_anexo = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/div[1]/ul/li[2]/span",
            )
        )
    )
    aba_anexo.click()
    time.sleep(5)

    aba_setinha = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div[2]/ul/li/div/i",
            )
        )
    )
    aba_setinha.click()

    campo_busca = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/div[1]/ul/li[2]/div/div/input",
            )
        )
    )
    campo_busca.clear()
    campo_busca.send_keys("OS")
    campo_busca.send_keys(Keys.RETURN)
    aguardar_carregamento(driver)

    if tratar_popup_sem_dados(driver):
        registrar_log(f"{datetime.now()} | OC {oc} | SEM TSSR")
        return

    xpath_new = "//div[normalize-space(text())='New']"
    elementos_new = WebDriverWait(driver, 15).until(
        EC.visibility_of_all_elements_located((By.XPATH, xpath_new))
    )

    for _ in elementos_new:
        try:
            btn_baixar = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div[11]/div/div/div/div/div[2]/ul/li/div[1]/div/span",
                    )
                )
            )
            btn_baixar.click()
            time.sleep(10)
            baixou_algo = True
        except:
            pass

    if baixou_algo:
        registrar_log(f"{datetime.now()} | OC {oc} | BAIXADO")
    else:
        registrar_log(f"{datetime.now()} | OC {oc} | NAO BAIXOU")


def main():
    if not CAMINHO_OCS or not os.path.exists(CAMINHO_OCS):
        print("Arquivo de OCs não encontrado.")
        return

    with open(CAMINHO_OCS, "r") as f:
        ocs = [linha.strip() for linha in f if linha.strip()]

    driver = iniciar_bot()
    acessar_ordens(driver)

    try:
        for oc in ocs:
            tratar_ordem(driver, oc)
    finally:
        try:
            driver.quit()
        except:
            pass


if __name__ == "__main__":
    main()
