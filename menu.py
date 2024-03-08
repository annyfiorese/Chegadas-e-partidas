from asyncio import wait
import datetime
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException

def execultar_menu_navegação(driver):

    # Localize o elemento pelo ID "liLeftMenuModulo7" e clique nele
    elemento_modulo = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.ID, "liLeftMenuModulo7"))
    )
    elemento_modulo.click()

    # Após clicar em "liLeftMenuModulo7", localize e clique no link "Visualizador de Relatórios"
    elemento_visualizador = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located(
            (By.LINK_TEXT, "Visualizador de Relatórios"))
    )
    elemento_visualizador.click()
