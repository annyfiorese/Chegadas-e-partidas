from asyncio import wait
import tkinter
import pyautogui
from tkinter.simpledialog import askstring
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from relatorio_terceiros import executar_reletorio_terceiros
from relatorio_terceiros import selecionar_opcao_com_tentativa
from relatorio_acesso import executar_relatorio_acesso 
from relatorio_acesso import selecionar_opcao_com_tentativa
from menu import execultar_menu_navegação

# Inicialize o ChromeDriver
driver = webdriver.Chrome()

# Aguarde um pouco
driver.implicitly_wait(10   )

try:
    # Abra o site
    driver.get("https://nexa-brasil.rainbowtec.com.br")
 
    # Aguarde um pouco
    driver.implicitly_wait(10)  # Espera 10 segundos

    # Pegando as informações de login

    root = tkinter.Tk()
    #root.withdraw()  # Oculta a janela principal

    usuario = askstring("Informações", "Digite seu usuário:")
    senha = askstring("Informações", "Digite sua senha:", show='*')
    datainicio = askstring("Informações", "Digite a data de início (DDMMYYYY):")
    datafim = askstring("Informações", "Digite a data de término (DDMMYYYY):")

    # Executar menu
    execultar_menu_navegação(driver,senha,usuario)
    
    # Executar o relatorio de terceiros
    executar_reletorio_terceiros(driver)
    
    # Executar relatiorio de acesso
    executar_relatorio_acesso(driver, datainicio, datafim)

finally:
    driver.implicitly_wait(10)
    # Feche o navegador
    driver.quit()

