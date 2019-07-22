# -*- coding: utf-8 -*-
import os, sys
import smtplib
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

################ Configurações ##################

# Configurações pontomais
pag_login_pontomais = "https://app.pontomaisweb.com.br/" # Página de Login do pontomais
login_pontomais = "" # Login do pontomais
senha_pontomais = "" # Senha do pontomais

# Configurações do email
msg = MIMEMultipart()
msg['From'] = "" # Endereço de email
msg['To'] = "" # Endereço de email
msg['Subject'] = "PontoMais" # Assunto do email
smtp_server = "SMTP.office365.com: 587" # Servidor SMPT para o envio do email. Outlook: SMTP.office365.com: 587 | Gmail: smtp.gmail.com: 587
senha_email = "" # Senha do email

# Configuração do intervalo de tempo em minutos para cada verificação no site pontomais
intervalo_verificacao = 5

# Driver do navegador que deseja utilizar
driver_padrao = "Firefox" # Valores: "Firefox", "Chrome"

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
        print("Erro ao buscar registro ponto")

def montar_mensagem_email(registro_ponto):
    try:
        return "Primeira Entrada: " + registro_ponto.get_primeira_entrada() + \
                "\nPrimeira Saida: " + registro_ponto.get_primeira_saida() + \
                "\nSegunda Entrada: " + registro_ponto.get_segunda_entrada() + \
                "\nSegunda Saida: " + registro_ponto.get_segunda_saida()
    except:
        print("Erro ao montar a mensagem de email")

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
            return webdriver.Firefox()
        elif driver_padrao == "Chrome":
            return webdriver.Chrome()
        else:
            return webdriver.Firefox()
    except:
        print("Erro ao iniciar o driver")

def fechar_driver(driver):
    try:
        driver.quit()
    except:
        print("Erro ao fechar o driver")

def entrar_pagina_historico(driver):
    try:
        entrar_pagina_login(driver)
        driver.get("https://app.pontomaisweb.com.br/#/meu_ponto")
        time.sleep(5)
        hoje_formatado = datetime.today().strftime('%d-%m-%Y')
        driver.find_element(By.XPATH,"//a[contains(@href, '#/meu_ponto/" + hoje_formatado + "/historico')]").click()
        time.sleep(5)
    except:
        print("Erro ao entrar na pagina de historico do pontomais")
    
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
        print("Erro ao entrar na pagina de login")

def enviar_email(mensagem_email):
    try:
        msg.attach(MIMEText(mensagem_email, 'plain'))
        server = smtplib.SMTP(smtp_server)
        server.starttls()
        server.login(msg['From'], senha_email)
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
    except:
        print("Erro ao enviar email")

def main():
    print("Iniciando a execucao do programa" )

    primeira_saida_existe = False
    while(not primeira_saida_existe):

        registro_ponto = buscar_registro_ponto() 

        if registro_ponto.primeira_saida != "" and registro_ponto.primeira_saida != "Primeira Saida nao encontrada":
            mensagem_email = montar_mensagem_email(registro_ponto)
            enviar_email(mensagem_email)
            primeira_saida_existe = True
            print("Primeira Saida encontrada" )
        else:
            print("Primeira Saida nao encontrada" )
            time.sleep(intervalo_verificacao * 60)

    print("Encerrando a execucao do programa" )

if __name__ == "__main__":
    main()

