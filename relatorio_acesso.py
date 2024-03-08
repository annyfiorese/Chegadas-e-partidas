from asyncio import wait
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from menu import execultar_menu_navegação
from selenium.webdriver.common.action_chains import ActionChains

def selecionar_opcao_com_tentativa(driver, element_id, valor, tentativas=3):
    for _ in range(tentativas):
        try:
            elemento_modulo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            dropdown = Select(elemento_modulo)
            dropdown.select_by_value(valor)
            return
        except StaleElementReferenceException:
            continue

# Função para formatar a data de "dia/mês/ano" para "mês/dia/ano"
def formatar_data(data):
    partes = data.split('/')
    return f"{partes[0]}/{partes[1]}/{partes[2]}"

def executar_relatorio_acesso(driver, datainicio, datafim):
    try:
        relatorios = ["7261", "7351", "7255", "7205", "7210"]
        for relatorio in relatorios:
            
            # Localize o campo de seleção de relatório Acesso
            selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphLeftColumn_ddlModulo_ddlCombo", "01")

            # Encontrar campo de reletorio para selecionar relatorio percuso de bloqueio
            selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "529")

            # Encotrar campo de grupo empresa pra selecionar grupo geral
            selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")

            # Encontrar campo empresa acesso para selecionar Nexa brasil
            selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_012_ddlCombo", "7005")
            
            wait = WebDriverWait(driver, 10)

            driver.implicitly_wait(10)
            # Adicione uma pausa de 3 segundos
            time.sleep(4)

            # Localize o campo_datainicio_nome
            campo_datainicio_nome = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_010$mskEdit"))
                )

            # Aplicar um clique triplo no campo_datainicio_nome
            action = ActionChains(driver)
            action.click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).release().perform()

            # Inserir data no campo_datainicio_nome
            campo_datainicio_nome.send_keys(datainicio)
            
            # Localize o campo_datafim_nome
            campo_datafim_nome = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_011$mskEdit"))
            )

            # Aplicar um clique triplo no campo_datafim_nome
            action = ActionChains(driver)
            action.click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).release().perform()

            # Inserir data no campo_datafim_nome
            campo_datafim_nome.send_keys(datafim)

            # Encontrar campo de filia acesso para selecionar a unidade da empresa
            selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_013_ddlCombo", relatorio)
            
            # Encontrar campo de gerar documento
            elemento_visualizador = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnShowOptions"))
            )
            elemento_visualizador.click()
            
            # Encontrar campo de formato excel
            elemento_visualizador = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
            )
            elemento_visualizador.click()
            
            # Encontrar o x para fechar o pop-up de download
            elemento_visualizador = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
            )
            elemento_visualizador.click()
            
            driver.implicitly_wait(10)

            localizar_home = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "rainbowtec-logo"))
            )
            localizar_home.click()

            execultar_menu_navegação(driver)
    except:
        print("erro no" + relatorios)
        input()