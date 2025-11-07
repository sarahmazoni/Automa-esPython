import os
import smtplib
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import datetime

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

# Pasta base para salvar exports
usuario_local = os.getlogin()
pasta_base = os.path.join("C:", "Users", usuario_local, "Documents", "Netflow_Automation", "exports")
os.makedirs(pasta_base, exist_ok=True)

# Caminhos completos dos arquivos de export
arquivo_export_1 = os.path.join(pasta_base, "exportCSV_OC_Activated.zip")
arquivo_export_2 = os.path.join(pasta_base, "exportCSV_Others.zip")

# Credenciais e configurações
url = "https://.com"
usuario = "seu_usuario"
senha = "sua_senha"

usuario_email = "seu.email.com"
senha_email = "sua_senha_email"
destinatario = "destinatario.com"

# ============================================================
# LIMPAR ARQUIVOS ANTIGOS
# ============================================================

for arquivo in [arquivo_export_1, arquivo_export_2]:
    if os.path.exists(arquivo):
        try:
            os.remove(arquivo)
            print(f"[INFO] Arquivo antigo '{os.path.basename(arquivo)}' removido.")
        except Exception as e:
            print(f"[AVISO] Não foi possível excluir '{arquivo}': {e}")

# ============================================================
# INICIALIZAR O CHROME
# ============================================================

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

driver.get(url)
print("[INFO] Abrindo o Netflow no Chrome...")

# Ignorar aviso de segurança, se houver
try:
    driver.find_element(By.ID, "details-button").click()
    time.sleep(2)
    driver.find_element(By.ID, "proceed-link").click()
    print("[INFO] Aviso de segurança ignorado.")
except:
    pass

# ============================================================
# LOGIN
# ============================================================

try:
    campo_login = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'Username') or contains(@name, 'user')]")
    campo_login.send_keys(usuario)
    campo_login.send_keys(Keys.RETURN)
    print("[INFO] Usuário inserido.")
    time.sleep(2)

    campo_senha = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'Password') or contains(@type, 'password')]")
    campo_senha.send_keys(senha)
    campo_senha.send_keys(Keys.RETURN)
    print("[INFO] Senha inserida.")
    time.sleep(2)

    botao_login = driver.find_element(By.XPATH, "//button[contains(@class, 'login-button') or text()='Login']")
    botao_login.click()
    print("[INFO] Login realizado com sucesso.")
    time.sleep(10)

except Exception as e:
    print(f"[ERRO] Falha no login: {e}")
    driver.quit()
    exit()

# ============================================================
# FUNÇÃO DE EXPORTAÇÃO
# ============================================================

def realizar_export(xpath_favorito, tempo_espera, nome_export):
    """Executa uma exportação completa dentro do Netflow."""
    try:
        print(f"[INFO] Iniciando exportação: {nome_export}")
        driver.find_element(By.XPATH, xpath_favorito).click()
        time.sleep(8)

        # Esperar carregamento
        while True:
            try:
                carregando = driver.find_element(By.XPATH, "//div[contains(@class, 'load-spinner__content')]")
                if carregando.is_displayed():
                    print("[INFO] Aguardando carregamento...")
                    time.sleep(5)
                else:
                    break
            except:
                break

        # Exportar todos os registros
        botao_exportar = driver.find_element(By.XPATH, "//button[contains(@class, 'export-button')]")
        botao_exportar.click()
        time.sleep(2)

        opcao_todos = driver.find_element(By.XPATH, "//span[contains(text(), 'Todos os registros') or contains(@class, 'export-all')]")
        opcao_todos.click()
        time.sleep(2)

        botao_ok = driver.find_element(By.XPATH, "//button[contains(@class, 'confirm-button') or text()='OK']")
        botao_ok.click()

        print(f"[INFO] Exportação '{nome_export}' iniciada. Aguardando {tempo_espera / 60:.0f} min...")
        time.sleep(tempo_espera)

    except Exception as e:
        print(f"[ERRO] Falha na exportação '{nome_export}': {e}")

# ============================================================
# EXECUTAR EXPORTAÇÕES
# ============================================================

# Primeira exportação - OC ACTIVATED
realizar_export(
    xpath_favorito="//div[contains(text(), 'EXPORT COMPLETO OC ACTIVATED') or contains(@class, 'favorite-item-oc-activated')]",
    tempo_espera=240,
    nome_export="EXPORT COMPLETO OC ACTIVATED"
)

# Fechar aba atual
try:
    driver.find_element(By.XPATH, "//button[contains(@class, 'tab-close-button')]").click()
    print("[INFO] Aba anterior fechada.")
except:
    print("[AVISO] Não foi possível fechar a aba anterior.")

# Segunda exportação - DEMAIS OCS
realizar_export(
    xpath_favorito="//div[contains(text(), 'EXPORT COMPLETO DEMAIS OCS') or contains(@class, 'favorite-item-others')]",
    tempo_espera=180,
    nome_export="EXPORT COMPLETO DEMAIS OCS"
)

# ============================================================
# LOGOUT
# ============================================================

try:
    print("[INFO] Finalizando sessão e realizando logout...")
    menu_usuario = driver.find_element(By.XPATH, "//button[contains(@class, 'user-menu-button')]")
    menu_usuario.click()
    time.sleep(2)

    botao_logout = driver.find_element(By.XPATH, "//button[contains(@class, 'logout-button') or text()='Logout']")
    botao_logout.click()
    print("[INFO] Logout realizado com sucesso.")
    time.sleep(5)
except Exception as e:
    print(f"[AVISO] Erro ao deslogar: {e}")

# ============================================================
# ENVIO DE EMAIL
# ============================================================

data_hora_atual = datetime.datetime.now()
data_hora_str = data_hora_atual.strftime("%d/%m/%Y %H:%M:%S")

try:
    msg = EmailMessage()
    msg["Subject"] = f"Exportações Netflow - {data_hora_str}"
    msg["From"] = usuario_email
    msg["To"] = destinatario
    msg.set_content(f"Exportações concluídas e enviadas automaticamente via Python às {data_hora_str}.")

    # Anexos
    for arquivo in [arquivo_export_1, arquivo_export_2]:
        if os.path.exists(arquivo):
            with open(arquivo, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=os.path.basename(arquivo))

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(usuario_email, senha_email)
        smtp.send_message(msg)
        print("[INFO] E-mail enviado com sucesso!")

except Exception as e:
    print(f"[ERRO] Falha ao enviar o e-mail: {e}")

# ============================================================
# FINALIZAÇÃO
# ============================================================

print("[INFO] Encerrando automação...")
driver.quit()
