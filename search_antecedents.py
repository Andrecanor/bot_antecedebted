from logging import exception
from pydoc import doc
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import time as t
from os import mkdir,rename
from os.path import join, exists
import pandas 
from datetime import datetime
from cryptography import fernet

options = webdriver.ChromeOptions()


work_path = GetVar("work_path")
db_name = r"C:\Users\aprendiz8.automatiza\OneDrive - GANA S.A\Valoracion 360\Seleccion MVP - Plan Carrera\PlanCarrera.xlsx"
excel = pandas.read_excel(db_name,engine='openpyxl')
documentos= excel['Documento '].to_list()
fecha= excel['Fecha de postulacion'].to_list()
for i in range(len(fecha)):
    fecha[i] = datetime.strftime(fecha[i], '%d-%B-%Y')
dia = []
mes = []
year=[]
for i in range(len(fecha)):
    dia.append(fecha[i].split('-')[0])
    mes.append(fecha[i].split('-')[1])
    year.append(fecha[i].split('-')[2])
mesyear= []
for i in range(len(fecha)):
    mesyear.append(f"{year[i]}-{mes[i]}")
ruta= r"C:\Users\aprendiz8.automatiza\Documents\Bot Antecedentes\Antecedentes"
for i in range(len(fecha)):
    if not exists(join(ruta, mesyear[i])):
        mkdir(join(ruta, mesyear[i]))
    if not exists(join(ruta, mesyear[i], dia[i])):
        mkdir(join(ruta, mesyear[i], dia[i]))
    

options.add_experimental_option(

    'prefs', {
        "download.default_directory": ruta,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True

    }
)

def decryptUserPass(user, password):
    with open(r'"C:\Users\aprendiz8.automatiza\OneDrive - GANA S.A\Valoracion 360\key.key', 'rb') as key_file:
        key = key_file.read()
        f = fernet.Fernet(key)
        decryptedUser = f.decrypt(user).decode()
        decryptedPass = f.decrypt(password).decode()
        return decryptedUser, decryptedPass

work_path = GetVar("work_path")

siga_url = GetVar("siga_url")
siga_username = GetVar("siga_username")
siga_password = GetVar("siga_password")
siga_username, siga_password = decryptUserPass(siga_username, siga_password)
chrome_driver_path = r"C:\Users\aprendiz8.automatiza\Documents\Rocketbot\drivers\win\chrome\chromedriver.exe"
filter_text = "antecedentes"

path = []

for i in range(len(fecha)):

    path.append(join(ruta, mesyear[i], dia[i]))

def web_available(driver: webdriver.Chrome) -> bool:
    """
    Espera por la carga de la página de login de la plataforma Siga. 
    Retorna True si la página carga correctamente.
    False si la página no esta disponible antes del timeout.
    """
    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located(
            (By.NAME, "login")))
    except TimeoutException as e:
        return False
    else:
        return True

# def filter_available(driver: webdriver.Chrome) -> bool:
#     """
#     Espera por la carga de la página principal de la plataforma Siga y el input de filtro. 
#     Retorna True si la página carga correctamente.
#     False si la página no esta disponible antes del timeout.
#     """
#     try:
#         WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
#             (By.ID, "textFilter")))
#     except TimeoutException as e:
#         return False
#     else:
#         return True




if len(documentos):
    driver = webdriver.Chrome(
        executable_path = chrome_driver_path, options=options)
    driver.delete_all_cookies()
    driver.get(siga_url)
    if web_available(driver):
        # Insertar el username en la página de login
        driver.find_element(By.NAME, "login").send_keys(siga_username)
        # Insertar la password en la página de login
        driver.find_element(By.NAME, "password").send_keys(siga_password)
        # Hacer click en el botón "Aceptar"
        driver.find_element(By.CLASS_NAME, "btn.btn--primary").click()
        
        t.sleep(1)
        #driver.switch_to.default_content()
        #driver.switch_to.frame("frmNews")
        
        #input_filter = driver.find_element(By.ID, "textFilter")
       
        #input_filter.send_keys(filter_text)
        #input_filter.send_keys(Keys.ENTER)

        #driver.find_element(By.CSS_SELECTOR, 'div[title="Empleados"]').click()

        #driver.find_element(By.CSS_SELECTOR, 'div[title="Novedades disciplinarias"]').click()

        #driver.find_element(By.CSS_SELECTOR, 'div[title="Generar ficha de antecedentes disciplinarios"]').click()
        
        #driver.switch_to.default_content()
        #driver.switch_to.frame("frmContents")
        for documento in documentos:
                driver.get("https://ganaadmin.gruporeditos.com/sigaadmin/siga/servlet/ServletAntecedentesDisciplinariosCargar")
                input_filter = driver.find_element(By.ID, "ideUsuarioText")
                try:
                    input_filter.send_keys(documento)
                    input_filter.send_keys(Keys.TAB)
                    t.sleep(1)
                    driver.find_element(By.ID, "btnConsultar").click()
            


                    t.sleep(3)
            
                    rename(join(ruta, f"AntecedentesDisciplinariosPDF.pdf"), join(path[i], f"AntecedentesDisciplinariosPDF_{documento}.pdf"))
                except:
                    pass
               



        #driver.quit()