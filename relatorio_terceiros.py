
from asyncio import wait
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException


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


def executar_reletorio_terceiros(driver):
    
    wait = WebDriverWait(driver, 10)
    # Localize o campo de seleção de relatório terceiro
    selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphLeftColumn_ddlModulo_ddlCombo", "03")

    # Encontrar campo de relatorio
    selecionar_opcao_com_tentativa(
        driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "475")

    # Encontrar campo de grupo de empresa
    selecionar_opcao_com_tentativa(
        driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
    
    # Encontrar campo de ativos
    selecionar_opcao_com_tentativa(
        driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_005_ddlCombo", "0")
    
    # Localização das empresas que devem ser extraido o relatorio
    relatorios = ["22", "2", "1", "3", "8"]

    # Loop de extração de informações das empresas anteriores
    for relatorio in relatorios:
        # Encontrar campo de local de prestação e selecionar a unidade
        selecionar_opcao_com_tentativa(
            driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_006_ddlCombo", relatorio)

        # Encontrar campo de gerar documento
        elemento_visualizador = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_btnShowOptions"))
        )
        elemento_visualizador.click()
        
        # Encontrar campo de formato excel
        elemento_visualizador = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
        )
        elemento_visualizador.click()
        
        # Encontrar o x para fechar o pop up de dowload
        elemento_visualizador = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
        )
        elemento_visualizador.click()

