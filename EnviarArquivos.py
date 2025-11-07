import os
import time
import csv
import shutil
import pyautogui
import pyperclip
import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# CONFIGURAÇÕES
# ============================================================

base_path = rf"C:\Users\{os.getlogin()}\Nokia\MN LAT ENG - AUTO\200 UPLOAD NETFLOW"

todas_ocs = os.path.join(base_path, "201 INPUT TSSR REMOPT", "OCs - TSSR REMOPT.txt")
pasta_zip = os.path.join(base_path, "203 INPUT TSSR INATEL")
pasta_carregados = os.path.join(base_path, "290 CARREGADOS")
log_arquivo_txt = os.path.join(base_path, "upload_log.txt")
log_arquivo_csv = os.path.join(base_path, "upload_log.csv")

url = ""
usuario = ""

# senha é lida de variável de ambiente NETFLOW_PASSWORD ou via prompt seguro
senha = os.environ.get("NETFLOW_PASSWORD")
if not senha:
    senha = getpass.getpass("Senha Netflow: ")

# ============================================================
# ETAPA 1 — GERAR TXT AUTOMATICAMENTE COM AS OCs
# ============================================================

arquivos = [f for f in os.listdir(pasta_zip) if f.endswith('.zip')]
ocs_extraidas = []

with open(todas_ocs, 'w') as txt:
    for arquivo in arquivos:
        numero = arquivo.split('_')[0].strip()
        if numero.isdigit():
            txt.write(numero + '\n')
            ocs_extraidas.append(numero)

if ocs_extraidas:
    primeira_oc = ocs_extraidas[0]
    with open(todas_ocs, 'a') as txt:
        txt.write(primeira_oc + '\n')
    print(f"Primeira OC ({primeira_oc}) duplicada para processar duas vezes.")
else:
    print("Nenhum arquivo ZIP válido encontrado para gerar OCs.")

print(f"{len(ocs_extraidas)} OCs extraídas e salvas em: {todas_ocs}")

# ============================================================
# CRIAR LOG
# ============================================================

if not os.path.exists(log_arquivo_csv):
    with open(log_arquivo_csv, mode='w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['OC', 'Status', 'Detalhe'])

def registrar_log(oc, status, detalhe=""):
    with open(log_arquivo_txt, "a") as log_txt:
        log_txt.write(f"{oc} - {status} - {detalhe}\n")
    with open(log_arquivo_csv, "a", newline='') as log_csv:
        writer = csv.writer(log_csv)
        writer.writerow([oc, status, detalhe])

# ============================================================
# FUNÇÃO DE ESPERA
# ============================================================

def aguardar_loading_generico(driver, timeout=120):
    """Espera até que indicadores de loading desapareçam."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((
                By.XPATH,
                "//*[contains(text(), 'Loading') or contains(@class, 'spinner') or contains(@class, 'loading')]"
            ))
        )
    except:
        pass

# ============================================================
# INICIAR NAVEGADOR
# ============================================================

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get(url)
print("Abrindo o Netflow no Chrome...")
time.sleep(3)

# Ignorar aviso de segurança se existir
try:
    driver.find_element(By.ID, "details-button").click()
    time.sleep(1)
    driver.find_element(By.ID, "proceed-link").click()
    print("Ignorando aviso de segurança...")
except:
    pass

# ============================================================
# LOGIN
# ============================================================

try:
    WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, "//input[contains(@placeholder, 'User') or contains(@type, 'text')]"))
    ).send_keys(usuario, Keys.RETURN)
    print("Usuário inserido.")
    time.sleep(2)

    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Pass') or contains(@type, 'password')]"))
    ).send_keys(senha, Keys.RETURN)
    print("Senha inserida.")
    time.sleep(2)

    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'login-button') or text()='Login']"))
    ).click()
    print("Login realizado com sucesso.")
    aguardar_loading_generico(driver)
except Exception as e:
    print(f"Erro no login: {e}")
    driver.quit()
    exit()

# ============================================================
# NAVEGAR ATÉ GERENCIAR ORDEM COMPLEXA
# ============================================================

try:
    menu_oc = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'menu-item') and contains(text(), 'Gerenciar Ordem Complexa')]"))
    )
    menu_oc.click()
    print("Navegando para 'Gerenciar Ordem Complexa'...")
    aguardar_loading_generico(driver)
except:
    print("Não foi possível localizar o menu de 'Gerenciar Ordem Complexa'.")

# ============================================================
# PROCESSAR TODAS AS OCs
# ============================================================

with open(todas_ocs, 'r') as f:
    ocs = [oc.strip() for oc in f if oc.strip()]

if not ocs:
    print("Nenhuma OC encontrada no TXT.")
else:
    for oc in ocs:
        try:
            print(f"\nProcessando OC: {oc}")

            campo_oc = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='ID Ordem Complexa' and @type='text']"))
            )
            campo_oc.clear()
            campo_oc.send_keys(oc, Keys.RETURN)
            aguardar_loading_generico(driver)

            # Acessar OC
            botao_oc_anexado = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'oc-link') or contains(text(), 'OC')]"))
            )
            botao_oc_anexado.click()
            print("OC aberta.")
            time.sleep(2)

            # Aba documentos
            aba_documentos = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Documentos')]"))
            )
            aba_documentos.click()
            print("Aba de documentos acessada.")
            time.sleep(2)

            # Duplo clique no TSSR
            tssr_box = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'TSSR')]"))
            )
            ActionChains(driver).move_to_element(tssr_box).double_click(tssr_box).perform()
            print("Documento TSSR selecionado.")
            time.sleep(2)

            # Aba adicionar
            aba_adicionar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Adicionar')]"))
            )
            aba_adicionar.click()
            print("Aba 'Adicionar' aberta.")
            time.sleep(2)

            # Botão carregar
            botao_carregar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Carregar') or contains(@class, 'upload-button')]"))
            )
            botao_carregar.click()
            print("Botão 'Carregar' clicado.")
            time.sleep(2)

            # Procurar arquivo ZIP
            arquivo_oc = None
            for arquivo in os.listdir(pasta_zip):
                if arquivo.startswith(oc) and arquivo.endswith('.zip'):
                    arquivo_oc = os.path.join(pasta_zip, arquivo)
                    break

            if not arquivo_oc:
                print("Nenhum arquivo ZIP correspondente encontrado.")
                registrar_log(oc, "FALHA", "ZIP não encontrado")
                continue

            print(f"ZIP encontrado: {arquivo_oc}")

            # Selecionar arquivo via clipboard
            pyperclip.copy(arquivo_oc)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            print("Upload iniciado...")

            aguardar_loading_generico(driver, timeout=120)
            print("Upload finalizado.")
            registrar_log(oc, "SUCESSO", "Upload finalizado")

            # Mover arquivo
            try:
                nome_arquivo = os.path.basename(arquivo_oc)
                destino = os.path.join(pasta_carregados, nome_arquivo)
                shutil.move(arquivo_oc, destino)
                print(f"Arquivo movido para: {destino}")
            except Exception as e:
                print(f"Erro ao mover arquivo: {e}")

            # Fechar aba da OC
            btn_x = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'tab-close-button') or contains(@title, 'Fechar')]"))
            )
            btn_x.click()
            print("Aba da OC fechada.")

        except Exception as e:
            print(f"Erro ao processar OC {oc}: {e}")
            registrar_log(oc, "FALHA", str(e))

# ============================================================
# LOGOUT
# ============================================================

try:
    driver.find_element(By.XPATH, "//button[contains(@class, 'user-menu-button')]").click()
    time.sleep(1)
    driver.find_element(By.XPATH, "//button[contains(@class, 'logout-button') or text()='Logout']").click()
    print("Logout realizado.")
except:
    print("Erro ao tentar deslogar.")

input("Aperte Enter para encerrar...")
driver.quit()
print("Automação finalizada com sucesso!")
