import re, os,time, datetime
from PyPDF2 import PdfReader
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
from dotenv import load_dotenv

pagina_main = "https://sistemas.ambiente.gba.gob.ar/Establecimientos/acceso/index.php?v=0"

load_dotenv()
opds_password = os.getenv("OPDS_PASS")


def main():
    driver = webdriver.Chrome()
    driver.get(pagina_main)

    user = driver.find_element(By.NAME,"txtUsuario")
    user.send_keys("ragoa")

    key = driver.find_element(By.NAME, "txtClave")
    key.send_keys(opds_password)

    send = driver.find_element(By.NAME, "btnEnviar")
    send.click()

    driver.find_element(By.NAME,"txtCuit1").send_keys("30")
    driver.find_element(By.NAME,"txtCuit2").send_keys("57878067")
    driver.find_element(By.NAME,"txtCuit3").send_keys("6")

    buscar = driver.find_element(By.NAME,"btnBuscar")
    buscar.click()

    # tabla = driver.find_element(By.XPATH,"//table[@class ='table table-hover table-striped FontNormal')]")
    mytable = driver.find_element(By.CSS_SELECTOR,'table.table.table-table-hover.table-striped.FontNormal')
    for row in mytable.find_elements(By.CSS_SELECTOR,'tr'):
        for cell in row.find_elements(By.TAG_NAME,'td'):
            print(cell.text)
    time.sleep(5)


main()