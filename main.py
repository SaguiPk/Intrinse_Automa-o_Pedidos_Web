import streamlit as st
import pandas as pd
import datetime as dt
import os
import time
from fuzzywuzzy import fuzz
import io
import zipfile

import threading
from queue import Queue

from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains


log_lock = threading.Lock()

# Pasta com os resultados da automação
UPLOAD_CODE = 'uploads/resultados'
if not os.path.exists(UPLOAD_CODE):
    os.makedirs(UPLOAD_CODE)

# Pasta com os Encaminhamentos
UPLOAD_FOLDER = "uploads/encaminhamentos"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Variáveis de Sessões
if 'arquivo_pacientes' not in st.session_state:
    st.session_state.arquivo_pacientes = None
if 'psico_select' not in st.session_state:
    st.session_state.psico_select = None
if 'pacientes_select' not in st.session_state:
    st.session_state.pacientes_select = None
if 'tipo_select' not in st.session_state:
    st.session_state.tipo_select = "Todos"
if 'files_encamin' not in st.session_state:
    st.session_state.files_encamin = None
if 'saved_files_paths' not in st.session_state:
    st.session_state.saved_files_paths = []
if 'pedidos' not in st.session_state:
    st.session_state.pedidos = None
if 'auto_incoberta' not in st.session_state:
    st.session_state.auto_incoberta = False
if 'process_running' not in st.session_state:
    st.session_state.process_running = False
if 'stop_event' not in st.session_state:
    st.session_state.stop_event = threading.Event()
if 'run_thread' not in st.session_state:
    st.session_state.run_thread = None
if 'status_log' not in st.session_state:
    st.session_state.status_log = {}
if 'download' not in st.session_state:
    st.session_state.download = False


# Função para salvar no banco de dados
def salvar_no_banco_de_dados(result, cod_guia, senha_guia, paciente, psico_selec):
    name_file = f'{psico_selec}.txt'
    file_path = os.path.join(UPLOAD_CODE, name_file)
    with open(file_path, "a", encoding='utf-8') as file:
        file.write(f'{result}\n')
        file.write(50*'_')
        file.write('\n')

    # # 1. Configurar a conexão com o banco de dados
    # engine = create_engine("sqlite:///database/pacientes_bd.db///:memory:", echo=True)
    #
    #
    # try:
    #     # Exemplo com SQL puro
    #     session.execute(text("UPDATE pacientes SET cod_guia = :cod, senha_guia = :senha WHERE nome = :p_id"),
    #         {"cod": cod_guia, "senha": senha_guia, "p_id": paciente})
    #
    #     # 4. Confirmar as alterações
    #     session.commit()
    #     print(f"Paciente {paciente} atualizado com sucesso!")
    #
    # except Exception as e:
    #     # Se algo der errado, desfaça as alterações
    #     session.rollback()
    #     print(f"Ocorreu um erro:\n{e}")
    #
    # finally:
    #     # 5. Fechar a sessão
    #     session.close()

    time.sleep(1)
    return True


def ajustar_zoom(driver, zoom_level=0.5):
    """
    Ajusta o zoom da página para garantir que elementos fiquem visíveis
    :param driver: Instância do WebDriver
    :param zoom_level: Nível de zoom desejado (0.5 = 50%, 1.0 = 100%, etc)
    """
    try:
        # Método 1: Usando ActionChains para simular Ctrl + -
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).send_keys(Keys.SUBTRACT).key_up(Keys.CONTROL)
        actions.perform()
        time.sleep(0.5)

        # Método 2: Executando JavaScript para definir o zoom diretamente
        driver.execute_script(f"document.body.style.zoom='{zoom_level}'")
        time.sleep(1)
        return True
    except Exception as e:
        print(f"Erro ao ajustar zoom: {e}")
        return False

# Função da Automação
def run_automation(message_queue, pedidos, psico_selec, files_encamin, stop_event, modo='especifico', auto_incoberta=False):
    try:
        # --- Configuração do Selenium ---
        options = Options()
        if auto_incoberta:
            options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        # Silenciar logs desnecessários
        options.add_argument("--log-level=3")  # Apenas erros críticos
        options.add_argument("--disable-logging")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        try:
            #chrome_options = webdriver.ChromeOptions()
            chrome_install = ChromeDriverManager().install()
            folder = os.path.dirname(chrome_install)
            chromedriver_path = os.path.join(folder, "chromedriver.exe")
            service = Service(chromedriver_path)
            driver = webdriver.Chrome(service=service, options=options)
            message_queue.put({'paciente': 'Abertura', 'log': 'Site Conectado'})

            # --- Início da Automação ---
            driver.get("https://saude.sulamericaseguros.com.br/prestador/login/")  # Usando um site de demo seguro
            time.sleep(3)  # Espera para visualização

        except Exception as e:
            message_queue.put({'paciente': 'Abertura', 'log': 'ERRO NA CONEXÃO'})
            message_queue.put({'paciente': 'Abertura', 'log': e})
            driver.quit()
            return

        try:
            if stop_event.is_set():
                message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                driver.quit()
                return
            driver.find_element(By.ID, 'code').send_keys('100000015245')
            driver.find_element(By.ID, 'user').send_keys('MASTER')
            driver.find_element(By.ID, 'senha').send_keys('@cre5cer')
            driver.find_element(By.ID, 'entrarLogin').click()
            time.sleep(2)
            message_queue.put({'paciente': 'Abertura', 'log': 'Login com sucesso!'})
        except Exception as e:
            message_queue.put({'paciente': 'Abertura', 'log': 'ERRO NO LOGIN'})
            message_queue.put({'paciente': 'Abertura', 'log': e})
            driver.quit()
            return

        if stop_event.is_set():
            message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
            driver.quit()
            return

        message_queue.put({'paciente': 'Abertura', 'log': 'Iniciar Pedidos'})
        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
        driver.find_element(By.XPATH, '//*[@id="sas-box-lgpd-info"]/div/div[2]/button').click()
        time.sleep(2)

        contagem = 0
        for nu_lin, cliente in pedidos.iterrows():
            # Verificar se o processo deve ser interrompido
            if stop_event.is_set():
                message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                driver.quit()
                break
            cliente_cols = cliente[['paciente', 'convenio', 'codigo', 'dr', 'crm', 'cbo', 'n_sessoes']]
            if modo == 'geral':
                psico_selec = cliente['psico']
            contagem += 1
            if contagem == 5:
                break
            contagem_msg = f'{contagem}/{len(pedidos)} Paciente: {cliente.iloc[0].upper()}'
            message_queue.put({'paciente': cliente.iloc[0], 'log': contagem_msg})

            indices_vazios = cliente_cols[cliente_cols.isnull() | (cliente_cols == 'vazio')].index.tolist()
            if indices_vazios:
                message_queue.put({'paciente': cliente.iloc[0], 'log': f'Paciente {cliente.iloc[0]} possui dados vazios: {indices_vazios}'})
                result = f'{paciente} possui dados vazios - {indices_vazios}'
                name_file = f'{psico_selec}.txt'
                file_path = os.path.join(UPLOAD_CODE, name_file)
                with open(file_path, "a", encoding='utf-8') as file:
                    file.write(f'{result}\n')
                    file.write(50 * '_')
                    file.write('\n')
                continue
            else:
                try:
                    # Verificar se o processo deve ser interrompido
                    if stop_event.is_set():
                        message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                        driver.quit()
                        break
                    paciente = cliente['paciente'] #.iloc[0]
                    convenio = cliente['convenio'] #.iloc[1]
                    atentimento = 'Convencional'
                    cod_pedido = 50000470
                    if 'ON' in str(convenio):
                        atentimento = 'Video'
                        cod_pedido = 66000912
                    codigo_paciente = cliente['codigo']
                    nome_dr = cliente['dr']
                    crm = cliente['crm']
                    crm = str(crm).strip()
                    estado = 'SP'
                    print(paciente)
                    if crm[-2:].isalpha():
                        crm = int(crm[:-3])
                        estado = crm[-2:]
                    cbo = int(cliente['cbo'])
                    data_pedido = dt.datetime.now().strftime("%d/%m/%Y").encode('latin1').decode('utf-8')
                    num_ses = int(cliente['n_sessoes'])

                    message_queue.put({'paciente': paciente, 'log': 'Dados do paciente'})
                    message_queue.put({'paciente': paciente, 'log': f'Atend: {atentimento} / {cod_pedido}   |   N sessões: {num_ses}   |   Data: {data_pedido}'})
                    print(f'Atend: {atentimento} / {cod_pedido}   |   N sessões: {num_ses}   |   Data: {data_pedido}')
                    time.sleep(2)

                except Exception as e:
                    message_queue.put({'paciente': paciente, 'log': 'ERRO na extração dos dados'})
                    message_queue.put({'paciente': paciente, 'log': e})
                    driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                    time.sleep(1.5)

                    result = f'{paciente} ERRO na extração dos dados'
                    name_file = f'{psico_selec}.txt'
                    file_path = os.path.join(UPLOAD_CODE, name_file)
                    with open(file_path, "a", encoding='utf-8') as file:
                        file.write(f'{result}\n')
                        file.write(50 * '_')
                        file.write('\n')
                    continue

                try:
                    try:
                        ajustar_zoom(driver, 0.7)
                        popup = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="recadastramento-div"]/p[1]')))
                        popup.click()
                    except:
                        pass
                    # Verificar se o processo deve ser interrompido
                    if stop_event.is_set():
                        message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                        driver.quit()
                        break
                    # CODIGO DE SOLICITAÇÃO - cod beneficiario
                    driver.find_element(By.ID, 'codigo-beneficiario-1').send_keys(codigo_paciente)  # Codigo do Beneficiário
                    time.sleep(1)
                    driver.find_element(By.CLASS_NAME, 'sas-form-submit').click()
                except Exception as e:
                    message_queue.put({'paciente': paciente, 'log': 'ERRO - codigo de beneficiario'})
                    message_queue.put({'paciente': paciente, 'log': e})
                    driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                    time.sleep(1.5)

                    result = f'{paciente} ERRO - código de beneficiario'
                    name_file = f'{psico_selec}.txt'
                    file_path = os.path.join(UPLOAD_CODE, name_file)
                    with open(file_path, "a", encoding='utf-8') as file:
                        file.write(f'{result}\n')
                        file.write(50 * '_')
                        file.write('\n')
                    continue

                try:
                    # Verificar se o processo deve ser interrompido
                    if stop_event.is_set():
                        message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                        driver.quit()
                        break
                    # SELECIONAR TIPO DE SOLIC.
                    time.sleep(1)
                    driver.find_element(By.ID, 'btn-eletivo').click()
                    time.sleep(1)
                    driver.find_element(By.ID, 'btn-sp-sadt').click()

                except Exception as e:
                    message_queue.put({'paciente': paciente, 'log': 'ERRO ao selecionar o tipo de solicitação'})
                    message_queue.put({'paciente': paciente, 'log': e})
                    driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                    time.sleep(1.5)

                    result = f'{paciente} ERRO ao selecionar o tipo de solicitação'
                    name_file = f'{psico_selec}.txt'
                    file_path = os.path.join(UPLOAD_CODE, name_file)
                    with open(file_path, "a", encoding='utf-8') as file:
                        file.write(f'{result}\n')
                        file.write(50 * '_')
                        file.write('\n')
                    continue

                message_queue.put({'paciente': paciente, 'log': "Preencher solicitação"})
                try:
                    # Verificar se o processo deve ser interrompido
                    if stop_event.is_set():
                        message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                        driver.quit()
                        break
                    # INICIAR SOLICITAÇÃO SP_SADT...
                    time.sleep(1)
                    # KEY CLINICA
                    driver.find_element(By.ID, 'solicitacao-sp-sadt.numero-guia-prestador').send_keys('100000015245')
                    # NOME DO MEDICO
                    driver.find_element(By.NAME,'solicitacao-sp-sadt.executante-solicitante.nome-profissional-solicitante').send_keys(nome_dr)  # NOME DR
                    conselho = Select(driver.find_element(By.ID, 'conselho-profissional'))  # CAIXA DE VALORES CONCELHOS
                    conselho.select_by_visible_text('CRM')  # CRM
                    time.sleep(1)
                    # ESTADO
                    but_estado = Select(driver.find_element(By.ID, 'uf-conselho-profissional'))  # CAIXA DE VALORES ESTADOS
                    but_estado.select_by_visible_text(estado)  # SP
                    time.sleep(0.5)
                    # CRM
                    driver.find_element(By.CLASS_NAME, 'numero-conselho').send_keys(crm)  # NUMERO DO CRM
                    time.sleep(0.5)
                    # CBO
                    but_cbo = driver.find_element(By.ID, 'busca-codigo-cbo')
                    but_cbo.send_keys(int(cbo))
                    web = WebDriverWait(driver, 2.5)
                    drop = web.until(EC.element_to_be_clickable((By.ID, 'ui-id-1')))  # SELECIONAR O ITEM DA LISTA
                    time.sleep(1)
                    drop.click()  # CLICAR cbo
                    # DATA
                    time.sleep(1)
                    driver.find_element(By.ID, 'data-atendimento').send_keys(data_pedido)  # DATA DO PEDIDO

                    time.sleep(0.5)
                    recem = Select(driver.find_element(By.ID, 'recem-nato'))  # CAIDE DE VALORES RECEM NATO
                    recem.select_by_visible_text('Não')  # NÃO
                    time.sleep(0.5)
                    # ATENDIMENTO
                    tipo_proced = Select(driver.find_element(By.XPATH,'//*[@id="formSolicitaProcedimento"]/div[4]/div[3]/div[2]/select'))  # CAIDA DE VALORES TIPO DO ATENDIEMNTO
                    tipo_proced.select_by_visible_text(atentimento)  # CONVENCIONAL OU ONLINE
                    time.sleep(0.5)
                    # CODIDO DO TIPO DE ATENTIMENTO
                    driver.find_element(By.NAME, 'codigo-procedimento').send_keys(cod_pedido)  # COD DO PEDIDO 50001221, 50000470
                    time.sleep(1)
                    # CLICAR BUT INLCUIR PEDIDO
                    button = driver.find_element(By.ID, "btn-incluir-procedimento")  # CLICAR NO BUT INCLUIR O PEDIDO
                    driver.execute_script("arguments[0].scrollIntoView(true);", button)
                    button.click()  # CLICAR INCLUIR PEDIDO
                    time.sleep(1.5)  # esperar atualizar...
                    # QUANT DE PEDIDOS
                    qnt = driver.find_element(By.NAME, 'quantidade-solicitada')
                    qnt.send_keys(Keys.CONTROL + 'a')
                    qnt.send_keys(str(num_ses))
                    time.sleep(2)

                    # VERIFICAR A VALIDAÇÃO
                    driver.find_element(By.ID, 'btn-validar-procedimento').click()
                    time.sleep(3)
                    text_validacao = driver.find_element(By.XPATH,'//*[@id="tabelaSolicitaProcedimento"]/tbody/tr/td[7]').text  # '//*[@id="tabelaSolicitaProcedimento"]/tbody/tr/td[6]/text()').text
                    time.sleep(1)

                    message_queue.put({'paciente': paciente, 'log': f'Status de Validação do pedido: {text_validacao}'})

                except Exception as e:
                    message_queue.put({'paciente': paciente, 'log': 'ERRO ao preencher a solicitação'})
                    message_queue.put({'paciente': paciente, 'log': e})
                    driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                    time.sleep(1.5)

                    result = f'{paciente} ERRO ao preencher a solicitação'
                    name_file = f'{psico_selec}.txt'
                    file_path = os.path.join(UPLOAD_CODE, name_file)
                    with open(file_path, "a", encoding='utf-8') as file:
                        file.write(f'{result}\n')
                        file.write(50 * '_')
                        file.write('\n')
                    continue

                try:
                    alert = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'btnLayerFecharuserDialog')))
                    alert.click()
                except:
                    pass
                msg_motivo = None
                codigo_G = 10
                codigo_S = 10

                # Verificar se o processo deve ser interrompido
                if stop_event.is_set():
                    message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                    driver.quit()
                    break

                if text_validacao.lower() == 'validado' and 'não' not in text_validacao.lower():
                    print('Pedido Valido')
                    message_queue.put({'paciente': paciente, 'log': 'Pedido Valido'})
                    # ANEXANDO ARQ PDF ----------------------------------------------------------------------
                    #pasta = files_encamin # r"C:\Users\gui_b\PycharmProjects\IntrinseAutoPedidos\Encaminhamentos"
                    arquivo = f'{paciente.replace("  ", " ")}.'
                    anexado = False
                    caminho_encaminhamento = None

                    def buscar_encaminhamento(arquivo, anexado, caminho_completo):
                        pasta = UPLOAD_FOLDER
                        for item in os.listdir(pasta):
                            try:
                                nome = item[:-4]
                                print(nome)
                                simil = fuzz.ratio(arquivo.strip().lower(), nome.strip().lower())
                                if int(simil) > 85:
                                    caminho_completo = os.path.join(pasta, item)
                                    message_queue.put({'paciente': paciente, 'log': 'Encaminhamento encontrado'})
                                    time.sleep(1)
                                    anexado = True
                                    break
                                else:
                                    print('Nao achou')
                                    continue
                            except:
                                message_queue.put({'paciente': paciente, 'log': 'ERRO na busca do encaminhamento'})
                        return anexado, caminho_completo

                    anexado, caminho_encaminhamento = buscar_encaminhamento(arquivo, anexado, caminho_encaminhamento)

                    # if anexado is False:
                    #     message_queue.put({'paciente': paciente, 'log': 'ERRO Encaminhamento Não Encontrado'})
                    #     result = f'{paciente} ERRO - Encaminhamento não encontrado'
                    #     name_file = f'{psico_selec}.txt'
                    #     file_path = os.path.join(UPLOAD_CODE, name_file)
                    #     with open(file_path, "a", encoding='utf-8') as file:
                    #         file.write(f'{result}\n')
                    #         file.write(50 * '_')
                    #         file.write('\n')
                    #     driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                    #     time.sleep(1.5)
                    #     continue

                    print(caminho_encaminhamento)

                    anexar = driver.find_element(By.ID, 'upload-item')  # CLICAR NO BUT INCLUIR O PEDIDO
                    driver.execute_script("arguments[0].scrollIntoView(true);", anexar)  # SCROLL
                    time.sleep(1)
                    #anexar.send_keys(caminho_encaminhamento)

                    tipo_doc = Select(driver.find_element(By.ID, 'anexos-tipo-documento'))
                    tipo_doc.select_by_visible_text('Pedido do profissional de saúde')
                    time.sleep(1.5)

                    # navegador.find_element(By.ID, 'btn-anexar')
                    web = WebDriverWait(driver, 2)
                    drop = web.until(EC.element_to_be_clickable((By.ID, 'btn-anexar')))
                    #drop.click()

                    try:
                        web = WebDriverWait(driver, 1.5)
                        drop = web.until(EC.presence_of_element_located((By.XPATH,'/html/body/div[8]/div[2]/div/table/tbody/tr/td/div[1]/div/div/div[2]/table/tbody/tr/td/div')))
                        texto = drop.text
                        alerta_anexar = 'Selecione um arquivo para ser anexado!'

                        if texto in alerta_anexar:
                            driver.find_element(By.ID, 'btnLayerFecharuserDialog').click()
                            time.sleep(1)
                            result = f'{paciente} ERRO - Arquivo não anexado'
                            name_file = f'{psico_selec}.txt'
                            file_path = os.path.join(UPLOAD_CODE, name_file)
                            with open(file_path, "a", encoding='utf-8') as file:
                                file.write(f'{result}\n')
                                file.write(50 * '_')
                                file.write('\n')
                            driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                            time.sleep(1.5)
                            continue
                    except:
                        pass

                    try:
                        # Verificar se o processo deve ser interrompido
                        if stop_event.is_set():
                            message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                            driver.quit()
                            break
                        message_queue.put({'paciente': paciente, 'log': 'Encaminhamento anexado com sucesso!'})
                        wait = WebDriverWait(driver, 10)
                        confirmar = wait.until(EC.element_to_be_clickable((By.ID, 'btn-confirmar-solicitacao')))
                        driver.execute_script("arguments[0].scrollIntoView(true);", confirmar)
                        #confirmar.click()
                        message_queue.put({'paciente': paciente, 'log': 'Confirmar pedido'})
                        time.sleep(1.5)
                    except:
                        result = f'{paciente} ERRO na confiormação do pedido'
                        name_file = f'{psico_selec}.txt'
                        file_path = os.path.join(UPLOAD_CODE, name_file)
                        with open(file_path, "a", encoding='utf-8') as file:
                            file.write(f'{result}\n')
                            file.write(50 * '_')
                            file.write('\n')
                        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                        time.sleep(1.5)
                        continue

                    try:
                        # Verificar se o processo deve ser interrompido
                        if stop_event.is_set():
                            message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                            driver.quit()
                            break
                        #codigo_G = driver.find_element(By.XPATH,'//*[@id="Form_8A61F4B14103A5DA014103CDBDE40BE4"]/div/div[2]/div[1]/div[2]/div[1]/span[2]').text
                        #codigo_S = driver.find_element(By.XPATH,'//*[@id="Form_8A61F4B14103A5DA014103CDBDE40BE4"]/div/div[2]/div[1]/div[2]/div[2]/span[2]').text
                        message_queue.put({'paciente': paciente, 'log': 'Pedido realizado com sucesso!\n\n'
                                 f'Salvar os Códigos\n\n'
                                 f'Código G -> {codigo_G}\n\n'
                                 f'Código_S -> {codigo_S}'})

                        time.sleep(1)
                        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                        time.sleep(1.5)

                    except:
                        result = f'{paciente} ERRO na Extração dos Códigos G e S'
                        name_file = f'{psico_selec}.txt'
                        file_path = os.path.join(UPLOAD_CODE, name_file)
                        with open(file_path, "a", encoding='utf-8') as file:
                            file.write(f'{result}\n')
                            file.write(50 * '_')
                            file.write('\n')
                        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                        time.sleep(1.5)
                        continue

                else:
                    try:
                        # Verificar se o processo deve ser interrompido
                        if stop_event.is_set():
                            message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                            driver.quit()
                            break
                        driver.find_element(By.XPATH,'//*[@id="tabelaSolicitaProcedimento"]/tbody/tr/td[7]/a').click()
                        time.sleep(1)
                        msg_motivo = driver.find_element(By.XPATH,'/html/body/div[7]/div[2]/div/table/tbody/tr/td/div[1]/div/div/div[2]/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]').text

                        message_queue.put({'paciente': paciente, 'log': 'ERRO Pedido NÃO VALIDADO!\n\n'
                                                                        f'MOTIVO: {msg_motivo}'})  # /html/body/div[7]/div[2]/div/table/tbody/tr/td/div[1]/div/div/div[2]/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]

                        driver.find_element(By.ID, 'btnLayerFecharuserDialog').click()
                        time.sleep(1)
                        result = f'{paciente} ERRO - pedido não valido\n Motivo: {msg_motivo}'
                        name_file = f'{psico_selec}.txt'
                        file_path = os.path.join(UPLOAD_CODE, name_file)
                        with open(file_path, "a", encoding='utf-8') as file:
                            file.write(f'{result}\n')
                            file.write(50 * '_')
                            file.write('\n')
                        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                        time.sleep(1.5)
                    except:
                        driver.get(r'https://saude.sulamericaseguros.com.br/prestador/segurado/validacao-de-procedimentos-tiss-3/validacao-de-procedimentos/solicitacao/')
                        time.sleep(1.5)
                        continue

                codigo_gerado = (f'{paciente}\n\n'
                                 f'Pedido: {text_validacao}\n\n'
                                 f'Motivo {msg_motivo}\n\n'
                                 f'Código G: {codigo_G}\n\n'
                                 f'Código S: {codigo_S}')

                print(codigo_gerado)
                # Verificar se o processo deve ser interrompido
                if stop_event.is_set():
                    message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                    driver.quit()
                    break
                if salvar_no_banco_de_dados(codigo_gerado, codigo_G, codigo_S, paciente, psico_selec):
                    message_queue.put({'paciente': paciente, 'log': "✔️ Dados salvos com sucesso!"})
                else:
                    message_queue.put({'paciente': paciente, 'log': "❌ Erro ao salvar no banco de dados."})

                message_queue.put({'paciente': paciente, 'log': 'Realizar próximo pedido...\n\n'})
                message_queue.put({'paciente': 'Abertura', 'log': f'{codigo_gerado}\n\n{50*'_'}'})

            # Verificar se o processo deve ser interrompido
            if stop_event.is_set():
                message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
                driver.quit()
                break

        message_queue.put({'paciente': 'Abertura', 'log': f'Processo finalizado com sucesso!.\n\n'})

    except Exception as e:
        # Captura qualquer erro durante a automação
        message_queue.put({'paciente': paciente, 'log': f"🚨 Ocorreu um erro:\n\n {e}"})
        # Verificar se o processo deve ser interrompido
        if stop_event.is_set():
            message_queue.put({'paciente': 'Abertura', 'log': '⚠️ Processo interrompido pelo usuário'})
            driver.quit()
            return
    finally:
        # Garante que o navegador seja fechado, mesmo se ocorrer um erro
        if driver:
            driver.quit()
            message_queue.put({'paciente': 'Abertura', 'log': "🚪 Navegador fechado."})
        message_queue.put({'paciente': 'Abertura', 'log': "--- FIM_DO_PROCESSO ---"})
        st.session_state.process_running = False

# Função Encaminhamento
@st.dialog('uploads/encaminhamentos')
def emcaminhamentos():
    files_encam = st.file_uploader("Encaminhamentos", type=["pdf", "png", "jpeg"],accept_multiple_files=True)
    cool1, cool3, cool2 = st.columns([1.3,3,1.1])
    with cool1:
        if st.button('Cancelar', key='cancelar'):
            for name_file in os.listdir(UPLOAD_FOLDER):
                caminho_file = os.path.join(UPLOAD_FOLDER,name_file)
                os.remove(caminho_file)
            st.session_state.files_encamin = None
            st.session_state.saved_files_paths = []
            st.rerun()

    with cool2:
        if st.button('Salvar', key='salvar'):
            if files_encam is not None:
                st.session_state.files_encamin = files_encam
                if os.listdir(UPLOAD_FOLDER):
                    for name_file in os.listdir(UPLOAD_FOLDER):
                        caminho_file = os.path.join(UPLOAD_FOLDER, name_file)
                        os.remove(caminho_file)
                for uploaded_file in files_encam:
                    # Salva o arquivo no sistema de arquivos local
                    file_path = os.path.join(UPLOAD_FOLDER, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    st.session_state.saved_files_paths.append(file_path)
                print('SALVAR')
            st.rerun()

# Configurações Iniciais
st.set_page_config(page_title="Intrince Web Automação Pedidos",
                   page_icon=":robot_face:",
                   layout="wide",)

st.title("Intrince Web Automação Pedidos")

col_registro, col_encam = st.columns(2)
with col_registro:
    # Campo de Arquivo Excel - REGISTRO DOS PACIENTES
    name_file = ''
    if st.session_state.arquivo_pacientes != None:
        name_file = st.session_state.arquivo_pacientes.name
    uploaded_file = st.file_uploader(f"Carregar Arquivo Excel - {name_file}", type="xlsx", key='arquivo_pacientes_excel')
    if uploaded_file is not None:
        st.session_state.arquivo_pacientes = uploaded_file
    # Campo para baixar os pdf dos encaminhamentos
with col_encam:
    st.write('')
    st.write('')
    st.button('Encaminhamentos', on_click=emcaminhamentos, type='primary')
    st.write(f'Arquivos: {len(st.session_state.saved_files_paths)} encaminhamentos')


if st.session_state.arquivo_pacientes is not None:
    try:
        # Lendo as Abas dp Arquivo Excel
        xls = pd.ExcelFile(st.session_state.arquivo_pacientes)

        # Obtendo a lista de nomes das planilhas (psicólogos)
        psicologos = xls.sheet_names
        #print(psicologos[1:])

        # Lista dos Psicologos
        index_psico = 0
        if st.session_state.psico_select is not None and st.session_state.psico_select in psicologos[1:]:
            index_psico = psicologos[1:].index(st.session_state.psico_select)
        st.session_state.psico_select = st.selectbox("Selecione a(o) Psicóloga(o)",
                                    psicologos[1:],
                                    index=index_psico, placeholder="Selecione a(o) Psicóloga(o)")

        # Lendo a planilha do psicólogo selecionado
        df = pd.read_excel(st.session_state.arquivo_pacientes, sheet_name=st.session_state.psico_select)

        # Tratar o DataFrame (remover linhas com col 'Nome do Paciente' vazia)
        df = df.dropna(subset=['Nome do Paciente'])

        # Converter colunas de data para string
        df['ENCAMINHAMENTO'] = df['ENCAMINHAMENTO'].apply(
            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (dt.datetime, pd.Timestamp)) else x)
        df['VENCIMENTO'] = df['VENCIMENTO'].apply(
            lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (dt.datetime, pd.Timestamp)) else x)

        # Tratar col 'Convênio' deixar somente paciente de convenio SUL AMERICA e S.A/ONLINE
        df = df[df['Convênio'].isin(['SUL AMERICA', 'S.A/ONLINE'])].reset_index(drop=True)

        st.subheader(f"Pacientes: {st.session_state.psico_select}")
        #st.dataframe(df[['Nome do Paciente', 'Convênio', 'N° Sessões', 'ENCAMINHAMENTO', 'VENCIMENTO']])


        tipo_select = st.selectbox('Tipo de Seleção', ['Todos', 'Incluir', 'Excluir'],
                                   index=0, key='tipo_select',)

        if tipo_select == 'Todos':
            st.session_state.pacientes_select = df['Nome do Paciente'].tolist()
        else:
            st.session_state.pacientes_select = []

        with st.container(border=True):
            cols = st.columns(3)
            for index, row in df.iterrows():
                id_paciente = f"{st.session_state.psico_select}_{index}"
                nome_paciente = row['Nome do Paciente']
                convenio_paciente = 'S.A' if row['Convênio'] == 'SUL AMERICA' else 'S.A/ON'
                info_paciente = f"{nome_paciente} - {convenio_paciente}"
                # Verifica se o paciente já está na lista de selecionados
                is_checked = nome_paciente in st.session_state.pacientes_select
                # if tipo_select == 'True':
                #     st.session_state[f"paciente_{id_paciente}"] = True
                # else:
                #     st.session_state[f"paciente_{id_paciente}"] = nome_paciente in st.session_state.pacientes_select

                with cols[index % 3]:
                    # Cria o checkbox
                    if st.checkbox(info_paciente, key=f"paciente_{id_paciente}", value=is_checked):
                        # Se marcado e ainda não está na lista, adiciona
                        if nome_paciente not in st.session_state.pacientes_select:
                            st.session_state.pacientes_select.append(nome_paciente)
                    else:
                        # Se desmarcado e está na lista, remove
                        if nome_paciente in st.session_state.pacientes_select:
                            st.session_state.pacientes_select.remove(nome_paciente)

        pacientes_automação = None
        if tipo_select == 'Excluir':
            pacientes_automação = [paciente for paciente in df['Nome do Paciente'] if paciente not in st.session_state.pacientes_select]
        elif tipo_select == 'Incluir':
            pacientes_automação = st.session_state.pacientes_select
        else:
            pacientes_automação = df['Nome do Paciente'].tolist()

        st.session_state.pedidos = df[df['Nome do Paciente'].isin(pacientes_automação)].copy().reset_index(drop=True)
        st.session_state.pedidos = st.session_state.pedidos[['Nome do Paciente', 'Convênio', 'Código', 'Dr(a)','CRM', 'CBO', 'N° Sessões']]
        st.session_state.pedidos.columns = ['paciente', 'convenio', 'codigo', 'dr', 'crm', 'cbo', 'n_sessoes']

        cols = st.columns(3)
        with cols[0]:
            @st.dialog('Gerar Pedidos')
            def gerar_pedido():
                if st.session_state.pedidos is None:
                    st.error('Nenhum paciente selecionado')
                elif st.session_state.files_encamin is None:
                    st.error('Nenhum arquivo de encaminhamento selecionado')
                else:
                    if st.button('▶ Gerar', key='gerar_pedidos', type='primary', disabled=st.session_state.process_running):
                        st.session_state.process_running = True
                        st.session_state.modo_auto = 'especifico'
                        st.rerun()
                    else:
                        st.write(st.session_state.pedidos)

            if st.button('Iniciar Automação', type='primary', on_click=gerar_pedido, use_container_width=True):
                st.session_state.status_log = {}
                st.session_state.process_running = False
            st.session_state.auto_incoberta = st.checkbox('Navegador Incoberta', value=True, key='auto_incoberta_check')

        with cols[1]:
            @st.dialog('Downloads')
            def dialog_download():
                try:
                    arquivos_na_pasta = [f for f in os.listdir('uploads/resultados') if f.endswith('.txt')]
                    if not arquivos_na_pasta:
                        st.warning("Nenhum arquivo .txt encontrado na pasta.")
                    else:
                        colis = st.columns(2)
                        arquivos_dir = os.listdir('uploads/resultados')
                        for idx, arquivo in enumerate(arquivos_dir):
                            file_path = os.path.join('uploads/resultados', arquivo)
                            if idx % 2 == 0:
                                with colis[0]:
                                    if st.download_button(f"Baixar {arquivo}",
                                                          type='tertiary',
                                                          data=open(file_path, 'rb').read(),
                                                          file_name=arquivo,
                                                          mime="text/plain",
                                                          icon=":material/download:", ):
                                        try:
                                            os.remove(file_path)
                                            st.success("Arquivo temporário removido com sucesso!")
                                        except Exception as e:
                                            st.error(f"Erro ao remover o arquivo: {e}")
                            else:
                                with colis[1]:
                                    if st.download_button(f"Baixar {arquivo}",
                                                          type='tertiary',
                                                          data=open(file_path, 'rb').read(),
                                                          file_name=arquivo,
                                                          mime="text/plain",
                                                          icon=":material/download:", ):
                                        try:
                                            os.remove(file_path)
                                            st.success("Arquivo temporário removido com sucesso!")
                                        except Exception as e:
                                            st.error(f"Erro ao remover o arquivo: {e}")

                        st.markdown("---")
                        culis = st.columns([8, 3.5])
                        with culis[0]:
                            if st.button('Voltar', type='secondary'):
                                st.rerun()
                        with culis[1]:
                            def make_zip_bytes(file_list):
                                buf = io.BytesIO()
                                with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
                                    for fname in file_list:
                                        path = os.path.join(UPLOAD_CODE, fname)
                                        # arcname evita colocar a estrutura "resultados/" dentro do zip
                                        z.write(path, arcname=fname)
                                buf.seek(0)
                                return buf.getvalue()

                            zip_bytes = make_zip_bytes(arquivos_na_pasta)
                            if st.download_button("Baixar Todos",
                                                  data=zip_bytes,
                                                  file_name=f"G_S dos Psicos.zip",
                                                  mime="application/zip"):
                                for idx, arquivo in enumerate(arquivos_dir):
                                    file_path = os.path.join('uploads/resultados', arquivo)
                                    try:
                                        os.remove(file_path)
                                        st.success("Arquivo temporário removido com sucesso!")
                                    except Exception as e:
                                        st.error(f"Erro ao remover o arquivo: {e}")

                except FileNotFoundError as e:
                    st.error(f"Erro ao baixar arquivos: {e}")
                except OSError as e:
                    st.error(f"Erro ao baixar arquivos: {e}")
                except Exception as e:
                    st.error("A pasta não foi encontrada. Por favor, crie-a com alguns arquivos.")

            arquivos_na_pasta = [f for f in os.listdir('uploads/resultados') if f.endswith('.txt')]
            if not arquivos_na_pasta:  # Se não houver arquivos na pasta
                st.session_state.download = True
            else:
                st.session_state.download = False

            st.button('Download', key='but_download', disabled=st.session_state.download, use_container_width=True, on_click=dialog_download)

        with cols[2]:
            with st.expander(f"📋 Ver {len(pacientes_automação)} paciente(s) selecionado(s)", expanded=False):
                for i, paciente in enumerate(pacientes_automação, 1):
                    st.write(f"{i}. {paciente}")


        # SCRIPT DA AUTOMAÇÃO --------------------------------------------------------------------
        message_queue = Queue()

        def stop_automation():
            st.session_state.stop_event.set()
            if st.session_state.run_thread and st.session_state.run_thread.is_alive():
                st.session_state.run_thread.join(timeout=5.0)
            st.session_state.process_running = False
            #st.session_state.run_thread = None

        if st.session_state.run_thread and st.session_state.run_thread.is_alive():
            st.info("Processo em execução, aguarde...")

        if st.session_state.process_running:
            with st.container(border=True, height=400):
                st.badge('Status do Processo', color='green')
                st.button('▶ Parar', key='parar_processo', type='primary', on_click=stop_automation)
                status_placeholder = st.empty()

            if st.session_state.run_thread and st.session_state.run_thread.is_alive():
                st.info("Processo em execução, aguarde...")
            else:
                st.session_state.stop_event.clear()
                # Inicia a thread de automação
                st.session_state.run_thread = threading.Thread(target=run_automation,
                    args=(message_queue, st.session_state.pedidos,
                          st.session_state.psico_select, st.session_state.files_encamin,
                          st.session_state.stop_event, st.session_state.modo_auto,
                          st.session_state.auto_incoberta,), daemon=True)
                st.session_state.run_thread.start()

        i = 0
        # Loop de monitoramento na thread principal do Streamlit
        while st.session_state.process_running:
            i += 1
            message = message_queue.get()
            if message['log'] == "--- FIM_DO_PROCESSO ---":
                st.session_state.process_running = False
                st.session_state.stop_event.set()
                print('Finalizar')
            else:
                paciente = message['paciente']
                log_message = message['log']

                if paciente not in st.session_state.status_log:
                    st.session_state.status_log[paciente] = []
                    st.session_state.status_log[paciente].append(log_message)
                else:
                    st.session_state.status_log[paciente].append(log_message)

            with status_placeholder.container():
                patients = list(st.session_state.status_log.keys())
                if patients:
                    tabs = st.tabs(patients)
                    for i, paciente_tab in enumerate(patients):
                        if paciente_tab in st.session_state.status_log:
                            with tabs[i]:
                                st.markdown("\n\n".join(st.session_state.status_log[paciente_tab]))
            time.sleep(1.5)


    except Exception as e:
        st.error(f"Erro ao ler o arquivo. Verifique se o arquivo.\n"
                 f"{e}")
        st.session_state.process_running = False
        st.session_state.stop_event.set()

# Campo dos Arquivos de Encaminhamento

# Checkbox berto/incoberto

# Selecionar Pacientes

# Button Iniciar

# Button Download dos resultados