import os
import time
import logging
import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

USUARIO_EMAIL = ""
USUARIO_SENHA = ""
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
VERIFY_SSL = '0'

def configurar_logger():
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "automacao_faturas.log")
    
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
        
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%d/%m/%Y %H:%M:%S'
    )
    return logging.getLogger()

def iniciar_navegador(logger):
    logger.info("Iniciando Chrome...")
    os.environ['WDM_SSL_VERIFY'] = VERIFY_SSL
    prefs = {"download.default_directory": PASTA_DOWNLOADS, "safebrowsing.enabled": True}
    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--start-maximized")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--remote-allow-origins=*")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--disable-features=NetworkService")
    opts.add_argument("--disable-web-security")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def pegar_planilha_aberta(logger):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        logger.error(f"O Microsoft Excel nao esta aberto! Erro: {e}")
        return None

    if excel.Workbooks.Count == 0:
        logger.warning("Excel aberto, mas sem planilhas.")
        return None

    wb_list = list(excel.Workbooks)
    if len(wb_list) == 1:
        print(f"Foi encontrada apenas uma planilha aberta: {wb_list[0].Name}. Ela sera utilizada!")
        logger.info(f"Planilha auto-selecionada: {wb_list[0].Name}")
        return wb_list[0]

    print("\n=== PLANILHAS ABERTAS NO EXCEL ===")
    for idx, wb in enumerate(wb_list):
        print(f"[{idx}] - {wb.Name}")
        
    try:
        escolha = int(input("\nDigite o número da planilha que deseja usar: "))
        if 0 <= escolha < len(wb_list):
            logger.info(f"Planilha selecionada manualmente: {wb_list[escolha].Name}")
            return wb_list[escolha]
    except ValueError:
        pass
        
    print("Escolha invalida!")
    logger.warning("Usuario fez escolha invalida de planilha.")
    return None

def processar_faturas():
    logger = configurar_logger()
    logger.info("=== INICIANDO SCRIPT DE AUTOMACAO ===")
    
    wb = pegar_planilha_aberta(logger)
    if not wb:
        return
        
    ws = wb.ActiveSheet
    logger.info(f"Aba ativada: {ws.Name}")
    
    primeira_conta = str(ws.Cells(2, 2).Value or "").strip()
    if not primeira_conta:
        logger.warning("A Coluna 2 (Conta) na linha 2 parece vazia. Verifique o layout da planilha.")

    driver = iniciar_navegador(logger)
    wait = WebDriverWait(driver, 20)
    
    try:
        logger.info("Acessando portal e fazendo Login...")
        driver.get("https://platform.hiperstream.com/login")
        try: 
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@onclick, 'aceitarCookie')]"))).click()
        except Exception: 
            pass
        wait.until(EC.visibility_of_element_located((By.ID, "Email"))).send_keys(USUARIO_EMAIL)
        driver.find_element(By.ID, "btnEntrarAuth0").click()
        wait.until(EC.element_to_be_clickable((By.ID, "Senha"))).send_keys(USUARIO_SENHA)
        time.sleep(1)
        driver.find_element(By.ID, "btnEntrar").click()
        logger.info("Login concluido!")
        time.sleep(3)

        logger.info("Acessando 'Biblioteca de Comunicacoes'...")
        try:
            btn_biblioteca = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Biblioteca de Comunicações')]")))
            try:
                btn_biblioteca.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn_biblioteca)
            logger.info("Clique em 'Biblioteca de Comunicacoes' realizado com sucesso!")
        except Exception as e:
            logger.warning(f"Nao achamos o botao 'Biblioteca de Comunicacoes'. Erro: {e}")
            
        time.sleep(3)
        
        i_linha = 2
        while True:
            nome_arquivo_novo = str(ws.Cells(i_linha, 1).Value or "").strip()
            conta = str(ws.Cells(i_linha, 2).Value or "").strip()
            status_atual = str(ws.Cells(i_linha, 3).Value or "").strip().upper()
            
            if not conta:
                logger.info(f"Fim da leitura da planilha atingido na linha {i_linha}.")
                break
                
            if status_atual == 'OK':
                i_linha += 1
                continue
                
            logger.info(f"Processando conta: {conta} (Salvar como: {nome_arquivo_novo}) [Linha {i_linha}]")
            
            try:
                elem_visao = wait.until(EC.element_to_be_clickable((By.ID, "VisaoId")))
                sel = Select(elem_visao)
                
                opcao_encontrada = False
                for opt in sel.options:
                    if "Consultar Fatura" in opt.text and "- afinz" not in opt.text.lower():
                        sel.select_by_visible_text(opt.text)
                        opcao_encontrada = True
                        break
                        
                if not opcao_encontrada:
                    sel.select_by_visible_text("Consultar Fatura")
                    
                time.sleep(1)
                
                data_inicio = driver.find_element(By.ID, "DataInicio")
                data_fim = driver.find_element(By.ID, "DataFim")
                
                val_inicio = data_inicio.get_attribute("value")
                val_fim = data_fim.get_attribute("value")
                
                if val_inicio != "01/12/2025" or val_fim != "31/12/2025":
                    driver.execute_script("arguments[0].value = '01/12/2025';", data_inicio)
                    driver.execute_script("arguments[0].value = '31/12/2025';", data_fim)
                    
                    if data_inicio.get_attribute("value") != "01/12/2025" or data_fim.get_attribute("value") != "31/12/2025":
                        raise ValueError("FALHA DE SEGURANCA: Nao foi possivel fixar a data em Dezembro/2025.")
                
                campo_consulta = driver.find_element(By.ID, "Consulta")
                campo_consulta.clear()
                campo_consulta.send_keys(conta)
                
                btn_filtrar = driver.find_element(By.ID, "btnFiltrar")
                driver.execute_script("arguments[0].click();", btn_filtrar)
                
                logger.info(f"Aguardando resultados para a conta {conta}...")
                time.sleep(4)
                
                icones_nuvem = driver.find_elements(By.XPATH, "//i[contains(@class, 'fa-cloud-download')]")
                
                if not icones_nuvem:
                    logger.warning(f"Nenhuma fatura de dezembro encontrada para a conta {conta}.")
                    ws.Cells(i_linha, 3).Value = 'Sem Faturas'
                    wb.Save()
                    i_linha += 1
                    continue
                else:
                    logger.info(f"{len(icones_nuvem)} faturas encontradas. Baixando a primeira...")
                    arquivos_antes = set(os.listdir(str(PASTA_DOWNLOADS)))
                    
                    botao_primeira_nuvem = icones_nuvem[0]
                    try:
                        botao_primeira_nuvem.click()
                    except Exception:
                        pai_a = botao_primeira_nuvem.find_element(By.XPATH, "..")
                        driver.execute_script("arguments[0].click();", pai_a)
                        
                    logger.info("Download Iniciado! Aguardando arquivo...")
                    
                    arquivo_baixado = None
                    tentativas = 0
                    while tentativas < 30:
                        time.sleep(1)
                        arquivos_agora = set(os.listdir(str(PASTA_DOWNLOADS)))
                        arquivos_novos = arquivos_agora - arquivos_antes
                        
                        for f in arquivos_novos:
                            if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                                arquivo_baixado = f
                                break
                        if arquivo_baixado:
                            break
                        tentativas += 1
                    
                    if not arquivo_baixado:
                        logger.error("O Chrome nao salvou o arquivo ou demorou demais (> 30s).")
                        ws.Cells(i_linha, 3).Value = 'Erro Download'
                        wb.Save()
                        i_linha += 1
                        continue
                        
                    caminho_original = os.path.join(str(PASTA_DOWNLOADS), str(arquivo_baixado))
                    _, extensao = os.path.splitext(str(arquivo_baixado))
                    
                    if not nome_arquivo_novo:
                        nome_arquivo_novo = f"Fatura_{conta}"
                        
                    nome_final = str(nome_arquivo_novo) + str(extensao)
                    caminho_novo = os.path.join(str(PASTA_DOWNLOADS), str(nome_final))
                    
                    logger.info(f"Renomeando '{arquivo_baixado}' para '{nome_final}'...")
                    
                    try:
                        if os.path.exists(caminho_novo):
                            os.remove(caminho_novo)
                        os.rename(caminho_original, caminho_novo)
                    except Exception as rn_err:
                        logger.warning(f"Nao foi possivel renomear automaticamente (Conta {conta}): {rn_err}")
            
            except Exception as loop_e:
                logger.error(f"Falha ao processar etapas na pagina para a conta {conta}: {loop_e}")
                ws.Cells(i_linha, 3).Value = 'Erro Sistema'
                try:
                    wb.Save()
                    driver.refresh() # Desobstrui modais travados ou falhas para a próxima conta
                except Exception:
                    pass
                i_linha += 1
                continue
            
            ws.Cells(i_linha, 3).Value = 'Ok'
            
            try:
                wb.Save()
            except Exception:
                pass
            
            logger.info(f"Sucesso: Conta {conta} processada. Planilha atualizada!")
            i_linha += 1
            
    except Exception as e:
        logger.critical(f"Erro inesperado durante a automacao: {e}")
    finally:
        logger.info("=== FINALIZANDO EXECUCAO DA AUTOMACAO ===")
        print("Automação finalizada! Consulte o arquivo automacao_faturas.log para visualizar tudo que ocorreu (O Chrome permanece aberto).")

if __name__ == "__main__":
    processar_faturas()
