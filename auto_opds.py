import re, os,time, datetime
from PyPDF2 import PdfReader
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd

pag_presentacion = "https://sistemas.ambiente.gba.gob.ar/Establecimientos/ASP/Presentacion.php?idPresentacion=1461241&presentaciones=todas&idEstablecimiento=18620&fecha=&idFabricante=0"
path_f = r"C:\Users\Cristian\Dropbox\OPDS Presentaciones\0-Air Liquide Tanques Estacionarios 2024\0-Contrato 1"
filename = "Estacionarios C1 2024 AirLiquide.xlsx"


def read(path, cols, names,header, sheetname) -> pd.DataFrame:
    return pd.read_excel(path,
            usecols=cols, 
            names=names, 
            header=header,
            sheet_name=sheetname)


def armar_modelo_f():
    path = r"C:\Users\Cristian\Dropbox\OPDS Presentaciones\0-Air Liquide Tanques Estacionarios 2024\0-Contrato 1"
    excel_equipos = read(os.path.join(path, "Equipos.xlsx"), "A", ["Nombre"], 0, "Hoja1")
    output = []

    for _,r in excel_equipos.iterrows():
        equipo = r["Nombre"].replace("Tanque-","").split("-")
        equipo = f"{equipo[1]}/{equipo[0]}"
        output.append(["-", equipo])
    cols = ("-","Equipo")
    df = pd.DataFrame(output)
    df.columns = cols
    df.to_excel(os.path.join(path,f"Estacionarios C1 2024 AirLiquide.xlsx"), index=False)


def main(equals:bool):

    with open("logger.txt",mode="a") as f:
        f.write("-"*15 + "\n")
    
    hechos = []
    formulario_f = read(os.path.join(path_f, filename), "B", ["ID"], 0, "page 1")
    driver = webdriver.Chrome()
    fecha_control = "26/04/2024"
    fecha_control_datetime = datetime.datetime.strptime(fecha_control, "%d/%m/%Y")
    fecha_proxima_calibracion = "25/04/2025"
    fecha_prox_control = "26/04/2029"

    date_prox_control = datetime.datetime.strptime(fecha_prox_control, "%d/%m/%Y")
    pending = []
    for i,r in formulario_f.iterrows():
        if i >= 20:
            break
        try:
            print(f"Renovando: {r['ID']}")
            driver.get(pag_presentacion)
            nueva_acta = driver.find_element(By.NAME, "btnNuevaA")
            nueva_acta.click()
            identificacion = r["ID"]
            if equals:
                renovar_equipo = driver.find_element(By.XPATH, f"//tr[td= '{identificacion}']//a")
            else:
                renovar_equipo = driver.find_element(By.XPATH, f"//tr[td[contains(.,'{identificacion}')]]//a")
            if renovar_equipo:
                renovar_equipo.click()
            else:
                continue
            fecha_inspeccion = driver.find_element(By.NAME, "txtFecha")
            fecha_inspeccion.clear()
            fecha_inspeccion.send_keys(fecha_control)

            save_fecha = driver.find_element(By.NAME, "guardar")
            save_fecha.click()

            tab_accesorios = driver.find_element(By.XPATH, "//div[a='Accesorios']//a[@href][5]")
            tab_accesorios.click()
            vto_vida_util = driver.find_element(By.NAME, "txt_gvencimientoVU")
            aux_vto_vu = vto_vida_util.get_attribute("value")
            aux_vto_vu = datetime.datetime.strptime(aux_vto_vu, "%d/%m/%Y")
            dif = aux_vto_vu - date_prox_control      

            if -182 <= dif.days < 0:
                vto_vida_util.clear()
                vto_vida_util.send_keys(fecha_prox_control)
                with open("logger.txt",mode="a") as f:
                    f.write(f"\nSe cambio la fecha de vida util del equipo {identificacion} de {aux_vto_vu} a {fecha_prox_control}\n")
                print(f"Se cambio la fecha de vida util del equipo {identificacion} de {aux_vto_vu} a {fecha_prox_control}")

            elif dif.days < -182:
                tab_acta = driver.find_element(By.XPATH, "//div[a='Accesorios']//a[@href][1]")
                tab_acta.click()
                evu = driver.find_element(By.XPATH, f"//tr[td = 'Extension de Vida Util']//input")
                evu.click()
                tab_accesorios.click()
                vto_vida_util.clear()
                nuevo_vto_vida_util = (fecha_control_datetime + datetime.timedelta(days=365*15)).strftime("%d/%m/%Y")
                with open("logger.txt",mode="a") as f:
                    f.write(f"\nSe cambio la fecha de vida util del equipo {identificacion} de {aux_vto_vu} a {nuevo_vto_vida_util} y se lo encuadro cono EVU\n")
                print(f"Se cambio la fecha de vida util del equipo {identificacion} de {aux_vto_vu} a {nuevo_vto_vida_util} y se lo encuadro cono EVU")
                if int(nuevo_vto_vida_util.split("/")[2]) > 2080:
                    nuevo_vto_vida_util = nuevo_vto_vida_util[:-4] + "2079"
                vto_vida_util.send_keys(nuevo_vto_vida_util)
                # ver si tambien hay que clickear donde dice habilitacion en la primer tab
            
            fecha_cali = driver.find_element(By.NAME, "txt_fechacali")
            fecha_cali.clear()
            fecha_cali.send_keys(fecha_control)

            fecha_prox_cali = driver.find_element(By.NAME, "txt_fechacaliP")
            fecha_prox_cali.clear()
            fecha_prox_cali.send_keys(fecha_proxima_calibracion)

            lado_cuerpo = driver.find_element(By.NAME, "txt_cuerpor")
            lado_cuerpo.clear()
            lado_cuerpo.send_keys(f"PROX. CONTROL {fecha_prox_control}")

            lado_camisa = driver.find_element(By.NAME, "txt_camisar")
            lado_camisa.clear()
            lado_camisa.send_keys(f"PROX. CONTROL {fecha_prox_control}")

            vto_espesores = driver.find_element(By.NAME, "txt_gvencimientoE")
            vto_espesores.clear()
            vto_espesores.send_keys(fecha_prox_control)
            save_accesorios = driver.find_element(By.NAME, "Guardar")
            save_accesorios.click()
            hechos.append(identificacion)
        except Exception as e:
            with open("logger.txt",mode="a") as f:
                f.write(f"\nError {e}, en equipo {r['ID']}\n")
            pending.append(r["ID"])
        
    driver.quit()
    with open("pending.txt",mode="a") as f:
        f.write(f"Equipos pendientes: {pending}\n")
    print("Equipos hechos:", hechos)

main(False)
