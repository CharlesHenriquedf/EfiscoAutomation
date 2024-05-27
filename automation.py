import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pyautogui
import time
import os

# Configuração do logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler("automation.log"),
                              logging.StreamHandler()])

logger = logging.getLogger(__name__)

class SefazAutomation:
    def __init__(self):
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("detach", True)  # Esta opção serve para não fechar o navegador após a execução do script
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            # Remover "--headless" se você quiser ver o navegador funcionando
            # chrome_options.add_argument("--headless")  # Esta opção serve para rodar o Chrome em modo headless (sem interface gráfica)
            
            # Caminho para o ChromeDriver baixado manualmente
            chrome_driver_path = os.path.join(os.getcwd(), 'drivers', 'chromedriver')  # ou 'chromedriver.exe' no Windows
            
            service = Service(chrome_driver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)
            logger.info("Inicialização do WebDriver concluída com sucesso.")
        except Exception as e:
            logger.error(f"Erro na inicialização do WebDriver: {e}")
            raise
    
    def access_site(self, url):
        try:
            self.driver.get(url)
            self.driver.maximize_window()
            logger.info(f"Acessou o site: {url}")
        except Exception as e:
            logger.error(f"Erro ao acessar o site {url}: {e}")
            raise
    
    def login_with_certificate(self):
        try:
            self.wait.until(EC.element_to_be_clickable((By.NAME, 'btt_certificado'))).click()
            pyautogui.press('enter')  # Pressiona Enter para selecionar o certificado digital automaticamente
            logger.info("Login com certificado digital realizado com sucesso.")
        except Exception as e:
            logger.error(f"Erro no login com certificado digital: {e}")
            raise
    
    def issue_statement(self, period, inscricao_estadual):
        try:
            self.driver.get("https://efiscoi.sefaz.pe.gov.br/sfi_trb_cmt/PREmitirExtratosContribuinte")
            self.wait.until(EC.presence_of_element_located((By.ID, 'primeiro_campo'))).send_keys(period)
            self.driver.find_element(By.ID, 'nuDocumentoIdentificacao').send_keys(inscricao_estadual)
            self.driver.find_element(By.CSS_SELECTOR, "input[data-type='pdf']").click()
            logger.info(f"Preenchido formulário com período: {period} e Inscrição Estadual: {inscricao_estadual}")

            types = ['bttemitirRelatorioNaoCalculadas', 'bttemitirRelatorioCalculadas', 'bttemitirRelatorioDebitos']
            for type in types:
                self.driver.find_element(By.ID, type).click()
                self.driver.find_element(By.NAME, 'btt_pmu_acao_4').click()
                self.download_and_proceed()
                logger.info(f"Emitido relatório: {type}")
        except Exception as e:
            logger.error(f"Erro ao emitir extrato para Inscrição Estadual {inscricao_estadual}: {e}")
            raise
    
    def download_and_proceed(self):
        try:
            pyautogui.click(200, 200)  # Ajuste as coordenadas conforme necessário
            self.driver.find_element(By.ID, 'btt_prosseguir').click()
            logger.info("Download realizado e prosseguiu.")
        except Exception as e:
            logger.error(f"Erro no download e prosseguimento: {e}")
            raise
    
    def logout(self):
        try:
            self.driver.find_element(By.ID, 'a_usuario').click()
            self.wait.until(EC.element_to_be_clickable((By.ID, 'link_encerrar_sessao'))).click()
            logger.info("Logout realizado com sucesso.")
        except Exception as e:
            logger.error(f"Erro no logout: {e}")
            raise
    
    def close_browser(self):
        try:
            self.driver.quit()
            logger.info("Navegador fechado com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao fechar o navegador: {e}")
            raise
    
    def run_automation(self, inscricoes_estaduais, period):
        try:
            self.access_site("https://efisco.sefaz.pe.gov.br/sfi_com_sca/PRMontarMenuAcesso")
            self.login_with_certificate()
            for inscricao_estadual in inscricoes_estaduais:
                self.issue_statement(period, inscricao_estadual)
            self.logout()
            self.close_browser()
            logger.info("Automação executada com sucesso.")
        except Exception as e:
            logger.error(f"Erro na execução da automação: {e}")
            raise
