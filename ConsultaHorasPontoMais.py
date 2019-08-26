# -*- coding: utf-8 -*-
import os, sys
import smtplib
import time
import logging
from pyvirtualdisplay import Display
from selenium import webdriver
from selenium.webdriver.common.by import By
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

################ Configurações ##################

# Configurações pontomais

# Mínimo LogLevel
log_level = 20 # Log Level
try:
    log_level = logging._nameToLevel[os.environ['LOG_LEVEL']]
except:
    pass
logging.getLogger().setLevel(log_level)
logging.info("log_level: " + str(log_level))

pag_login_pontomais = "https://app.pontomaisweb.com.br/" # Página de Login do pontomais

login_pontomais = "" # Login do pontomais
try:
    login_pontomais = os.environ['LOGIN_PONTOMAIS']
except:
    pass
logging.info("login_pontomais: " + login_pontomais)

senha_pontomais = "" # Senha do pontomais
try:
    senha_pontomais = os.environ['SENHA_PONTOMAIS']
except:
    pass
logging.info("senha_pontomais: " + senha_pontomais)

# Configurações do email

email_from = "limamarcelo2017@outlook.com" # Endereço de email
try:
    email_from = os.environ['LOGIN_EMAIL']
except:
    pass
logging.info("email_from: " + email_from)

email_to = "limamarcelo2017@outlook.com" # Endereço de email
try:
    email_to = os.environ['LOGIN_EMAIL']
except:
    pass
logging.info("email_to: " + email_to)

email_subject = "PontoMais" # Assunto do email
try:
    email_subject = os.environ['ASSUNTO_EMAIL']
except:
    pass
logging.info("email_subject: " + email_subject)

smtp_server = "SMTP.office365.com: 587" # Servidor SMPT para o envio do email. Outlook: SMTP.office365.com: 587 | Gmail: smtp.gmail.com: 587
try:
    smtp_server = os.environ['SMPT_SERVER']
except:
    pass
logging.info("smtp_server: " + smtp_server)

email_senha = "Mar97908994" # Senha do email
try:
    email_senha = os.environ['SENHA_EMAIL']
except:
    pass
logging.info("email_senha: " + email_senha)

# Configuração do intervalo de tempo em minutos para cada verificação no site pontomais
intervalo_verificacao = 10
try:
    intervalo_verificacao = int(os.environ['INTERVALO_VERIFICACAO'])
except:
    pass
logging.info("intervalo_verificacao: " + str(intervalo_verificacao))

# Driver do navegador que deseja utilizar
driver_padrao = "Firefox" # Valores: "Firefox", "Chrome"
try:
    driver_padrao = os.environ['DRIVER_PADRAO']
except:
    pass
logging.info("driver_padrao: " + driver_padrao)

#################################################

class RegistroPonto:
    def __init__(self, primeira_entrada, primeira_saida, segunda_entrada, segunda_saida):
        self.primeira_entrada = primeira_entrada
        self.primeira_saida = primeira_saida
        self.segunda_entrada = segunda_entrada
        self.segunda_saida = segunda_saida

    def set_primeira_entrada(self, primeira_entrada):
        self.primeira_entrada = primeira_entrada
     
    def set_primeira_saida(self, primeira_saida):
        self.primeira_saida = primeira_saida
    
    def set_segunda_entrada(self, segunda_entrada):
        self.segunda_entrada = segunda_entrada

    def set_segunda_saida(self, segunda_saida):
        self.segunda_saida = segunda_saida
     
    def get_primeira_entrada(self):
        return self.primeira_entrada
         
    def get_primeira_saida(self):
        return self.primeira_saida
    
    def get_segunda_entrada(self):
        return self.segunda_entrada
    
    def get_segunda_saida(self):
        return self.segunda_saida

def buscar_registro_ponto():
    try:
        driver = iniciar_driver()
        entrar_pagina_historico(driver)
        registro_ponto = RegistroPonto(str(buscar_primeira_entrada(driver)), str(buscar_primeira_saida(driver)), \
                            str(buscar_segunda_entrada(driver)), str(buscar_segunda_saida(driver)))
        fechar_driver(driver)
        return registro_ponto
    except:
        logging.error("Erro ao buscar registro ponto")

def buscar_registro_ponto_mock():
    try:
        registro_ponto = RegistroPonto("07:30","13:10","","")
        return registro_ponto
    except:
        logging.error("Erro ao buscar registro ponto")

def montar_mensagem_email(registro_ponto):
    try:
        return "Primeira Entrada: " + registro_ponto.get_primeira_entrada() + \
                "\nPrimeira Saida: " + registro_ponto.get_primeira_saida() + \
                "\nSegunda Entrada: " + registro_ponto.get_segunda_entrada() + \
                "\nSegunda Saida: " + registro_ponto.get_segunda_saida()
    except:
        logging.error("Erro ao montar a mensagem de email")

def buscar_primeira_entrada(driver):
    try:
        return driver.find_element(By.XPATH,'//*[@id="content-wrapper"]/div[2]/div[2]/ng-view/ng-include[1]/div/div/table/tbody/tr/td[1]/span').text
    except:
        return "Primeira Entrada nao encontrada"

def buscar_primeira_saida(driver):
    try:
        return driver.find_element(By.XPATH,'//*[@id="content-wrapper"]/div[2]/div[2]/ng-view/ng-include[1]/div/div/table/tbody/tr/td[2]/span').text
    except:
        return "Primeira Saida nao encontrada"

def buscar_segunda_entrada(driver):
    try:
        return driver.find_element(By.XPATH,'//*[@id="content-wrapper"]/div[2]/div[2]/ng-view/ng-include[1]/div/div/table/tbody/tr/td[3]/span').text
    except:
        return "Segunda Entrada nao encontrada"

def buscar_segunda_saida(driver):
    try:
        return driver.find_element(By.XPATH,'//*[@id="content-wrapper"]/div[2]/div[2]/ng-view/ng-include[1]/div/div/table/tbody/tr/td[4]/span').text
    except:
        return "Segunda Saida nao encontrada"

def iniciar_driver():
    try:
        if driver_padrao == "Firefox":
            display = Display(visible=0, size=(800, 600))
            display.start()
            logging.info('Initialized virtual display..')

            firefox_profile = webdriver.FirefoxProfile()
            firefox_profile.set_preference('browser.download.folderList', 2)
            firefox_profile.set_preference('browser.download.manager.showWhenStarting', False)
            firefox_profile.set_preference('browser.download.dir', os.getcwd())
            firefox_profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')

            logging.info('Prepared firefox profile..')

            browser = webdriver.Firefox(firefox_profile=firefox_profile)
            logging.info('Initialized firefox browser..')
            return browser
        elif driver_padrao == "Chrome":
            display = Display(visible=0, size=(800, 600))
            display.start()
            logging.info('Initialized virtual display..')

            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument('--no-sandbox')

            chrome_options.add_experimental_option('prefs', {
                'download.default_directory': os.getcwd(),
                'download.prompt_for_download': False,
            })
            logging.info('Prepared chrome options..')

            browser = webdriver.Chrome(chrome_options=chrome_options)
            logging.info('Initialized chrome browser..')
        else:
            display = Display(visible=0, size=(800, 600))
            display.start()
            logging.info('Initialized virtual display..')

            firefox_profile = webdriver.FirefoxProfile()
            firefox_profile.set_preference('browser.download.folderList', 2)
            firefox_profile.set_preference('browser.download.manager.showWhenStarting', False)
            firefox_profile.set_preference('browser.download.dir', os.getcwd())
            firefox_profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')

            logging.info('Prepared firefox profile..')

            browser = webdriver.Firefox(firefox_profile=firefox_profile)
            logging.info('Initialized firefox browser..')
            return browser
    except:
        logging.error("Erro ao iniciar o driver")

def fechar_driver(driver):
    try:
        driver.quit()
    except:
        logging.error("Erro ao fechar o driver")

def entrar_pagina_historico(driver):
    try:
        entrar_pagina_login(driver)
        driver.get("https://app.pontomaisweb.com.br/#/meu_ponto")
        time.sleep(5)
        hoje_formatado = datetime.today().strftime('%d-%m-%Y')
        driver.find_element(By.XPATH,"//a[contains(@href, '#/meu_ponto/" + hoje_formatado + "/historico')]").click()
        time.sleep(5)
    except:
        logging.error("Erro ao entrar na pagina de historico do pontomais")
    
def entrar_pagina_login(driver):
    try:
        driver.get(pag_login_pontomais)
        login = driver.find_element_by_name('login')
        login.send_keys(login_pontomais)
        password = driver.find_element_by_name('password')
        password.send_keys(senha_pontomais)
        driver.find_element(By.CSS_SELECTOR, '.btn.btn-block.btn-bold.btn-primary.ng-binding').click()
        time.sleep(5)
    except:
        logging.error("Erro ao entrar na pagina de login")

def enviar_email(mensagem_email):
    try:
        email_message = f'''{mensagem_email}'''
        email_message = 'Subject: {}\n\n{}'.format(email_subject, email_message)
        server = smtplib.SMTP(smtp_server)
        server.starttls()
        server.login(email_from, email_senha)
        server.sendmail(email_from, email_to, email_message)
        server.quit()
        logging.info("Email enviado com sucesso")
    except:
        logging.error("Erro ao enviar email")

def main():
    logging.info("Iniciando a execucao do programa")

    data_primeira_entrada_enviada = datetime.min
    data_primeira_saida_enviada = datetime.min
    data_segunda_entrada_enviada = datetime.min
    data_segunda_saida_enviada = datetime.min

    while(True):
        registro_ponto = buscar_registro_ponto()
        hoje = datetime.today().strftime('%d-%m-%Y')

        if data_primeira_entrada_enviada != hoje and registro_ponto.primeira_entrada != "" and registro_ponto.primeira_entrada != "Primeira Entrada nao encontrada":
            mensagem_email = montar_mensagem_email(registro_ponto)
            enviar_email(mensagem_email)
            data_primeira_entrada_enviada = hoje
            logging.info("Primeira Entrada encontrada")
        else:
            logging.info("Primeira Entrada nao encontrada")

        if data_primeira_saida_enviada != hoje and registro_ponto.primeira_saida != "" and registro_ponto.primeira_saida != "Primeira Saida nao encontrada":
            mensagem_email = montar_mensagem_email(registro_ponto)
            enviar_email(mensagem_email)
            data_primeira_saida_enviada = hoje
            logging.info("Primeira Saida encontrada")
        else:
            logging.info("Primeira Saida nao encontrada")

        if data_segunda_entrada_enviada != hoje and registro_ponto.segunda_entrada != "" and registro_ponto.segunda_entrada != "Segunda Entrada nao encontrada":
            mensagem_email = montar_mensagem_email(registro_ponto)
            enviar_email(mensagem_email)
            data_segunda_entrada_enviada = hoje
            logging.info("Segunda Entrada encontrada")
        else:
            logging.info("Segunda Entrada nao encontrada")
            
        if data_segunda_saida_enviada != hoje and registro_ponto.segunda_saida != "" and registro_ponto.segunda_saida != "Segunda Saida nao encontrada":
            mensagem_email = montar_mensagem_email(registro_ponto)
            enviar_email(mensagem_email)
            data_segunda_saida_enviada = hoje
            logging.info("Segunda Saida encontrada")
        else:
            logging.info("Segunda Saida nao encontrada")

        time.sleep(intervalo_verificacao * 60)

    logging.info("Encerrando a execucao do programa")

if __name__ == "__main__":
    main()

