import shutil
import os
import sys
import pandas as pd
import time
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def criar_diretorio(diretorio):
    try:
        os.makedirs(diretorio, exist_ok=True)
        print(f"Diretório '{diretorio}' criado com sucesso!")
    except PermissionError:
        print("Permissão negada para criar o diretório.")
    except Exception as e:
        print(f"Ocorreu um erro ao criar o diretório: {e}")

def copiar_diretorio(origem, destino):
    try:
        if os.path.exists(destino):
            shutil.rmtree(destino)  # Remove o diretório de destino se ele já existir
        shutil.copytree(origem, destino)
        print(f"Diretório '{origem}' copiado para '{destino}' com sucesso!")
    except FileNotFoundError:
        print("Diretório de origem não encontrado.")
    except PermissionError:
        print("Permissão negada para acessar o diretório de origem ou destino.")
    except Exception as e:
        print(f"Ocorreu um erro ao copiar o diretório: {e}")

def excluir_diretorios(diretorio, excecao=None):
    for pasta in os.listdir(diretorio):
        caminho_pasta = os.path.join(diretorio, pasta)
        if os.path.isdir(caminho_pasta) and pasta != excecao:
            shutil.rmtree(caminho_pasta)
            print(f"Pasta {pasta} excluída.")
    print("Processo de exclusão concluído.")

def inicializar_driver(user_data_dir):
    opcoes = Options()
    opcoes.add_argument(f"--user-data-dir={user_data_dir}")
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=opcoes)
    driver.get("https://messages.google.com/web/conversations?hl=pt-BR")
    driver.maximize_window()
    time.sleep(15)
    return driver

def enviar_mensagem(driver, numero_telefone):
    try:
        botao_escrever_mensagem = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/div/mw-fab-link/a/span[2]'))
        )
        botao_escrever_mensagem.click()
        barra_pesquisa = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/div/div/input'))
        )
        barra_pesquisa.send_keys(numero_telefone)
        botao_clicar_numero = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contact-selector-button/button/span[4]'))
        )
        botao_clicar_numero.click()
        mensagem = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/div[2]/div/div/mws-autosize-textarea/textarea'))
        )
        mensagem.send_keys("Olá! Agora você pode abrir sua conta no banco digital Unixbank! Aproveite a conveniência dos serviços bancários online. Visite nosso site ou baixe nosso aplicativo para começar https://unixbank.com.br/unixconta. Qualquer dúvida, estamos à disposição. Equipe Unixbank.")
        botao_enviar = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button'))
        )
        botao_enviar.click()
        return True
    except Exception as e:
        print(f"Erro ao enviar mensagem para o número {numero_telefone}: {e}")
        return False

def main():
    script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
    data_dir = os.path.join(script_dir, 'data')
    backup_dir = os.path.join(script_dir, 'backup')
    excel_files = glob.glob(os.path.join(script_dir, '*.xlsx'))

    if not excel_files:
        print("Erro: Nenhum arquivo Excel encontrado na pasta.")
        sys.exit(1)  # Encerra o programa se nenhum arquivo Excel for encontrado

    excel_path = excel_files[0]
    erros_path = os.path.join(script_dir, 'Erros.xlsx')

    criar_diretorio(data_dir)
    criar_diretorio(backup_dir)

    df = pd.read_excel(excel_path)
    df['Telefone'] = df['Telefone'].astype(str)
    numeros_invalidos = [telefone for telefone in df['Telefone'] if len(telefone) < 3 or telefone[2] not in ("9", "2", "3") or len(telefone) != 11]

    df_erros = pd.DataFrame(numeros_invalidos, columns=['TELEFONE'])
    df_erros.to_excel(erros_path, index=False)

    origem = os.path.join(data_dir, 'Default', 'IndexedDB', 'https_messages.google.com_0.indexeddb.leveldb')
    destino = os.path.join(backup_dir, 'https_messages.google.com_0.indexeddb.leveldb')
    driver = inicializar_driver(data_dir)
    copiar_diretorio(origem, destino)
    enviados = 0
    qtd = 0
    try:
        for numero_telefone in df['Telefone']:
            if numero_telefone not in numeros_invalidos:
                sucesso = enviar_mensagem(driver, numero_telefone)
                if sucesso:
                    qtd += 1
                    enviados +=1
                    print(f"Mensagens enviadas: {enviados}")
                    if qtd > 100:
                        driver.quit()
                        excluir_diretorios(data_dir, 'Default')
                        excluir_diretorios(os.path.join(data_dir, 'Default'), 'IndexedDB')
                        copiar_diretorio(destino, origem)
                        driver = inicializar_driver(data_dir)
                        qtd = 0
                    time.sleep(1)
                else:
                    driver.quit()
                    excluir_diretorios(data_dir, 'Default')
                    excluir_diretorios(os.path.join(data_dir, 'Default'), 'IndexedDB')
                    copiar_diretorio(destino, origem)
                    driver = inicializar_driver(data_dir)
    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
