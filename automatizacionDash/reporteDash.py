import os
import time
import datetime
import schedule
import telebot
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import pychromedriver 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


ROOT_DIR = "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\"
DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CHROMEDRIVER_PATH = 'chromedriver.exe'
TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
CHAT_ID_ADMIN = 1397965849  # Asume que este es tu chat ID. Cambia según necesites.
CHAT_IDS = [1397965849, 5781054062, 6055198218,6512803121,5686702637,6564107654,6383100715]
# Instalar el browser de Chrome
pathChrom= os.path.expanduser('~')+'\\AppData\\Local\\Google\\Chrome\\User Data\\telcombas'
#chromedriver_autoinstaller.install()


def clear_all_images():
    for filename in os.listdir():
        if filename.endswith(".png"):
            try:
                os.remove(filename)
                print(f"Eliminada: {filename}")
            except Exception as e:
                print(f"Error al eliminar {filename}. Razón: {e}")

def set_chrome_options():
    options = Options()
    prefs = {"download.default_directory": DOWNLOADS_PATH}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument('--disable-extensions')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--no-sandbox')
    options.add_argument("--headless")
    options.add_argument("window-size=1400x1050")
    options.add_argument(f"--user-data-dir={pathChrom}")
    return options

def f5_dashboard_extraction():
    ChromeDriverManager().install()
    browserDash = webdriver.Chrome(set_chrome_options())
    # Ajusta el tamaño de la ventana del navegador para que se ajuste al contenido de la página
    total_width = browserDash.execute_script("return document.body.offsetWidth")
    total_height = browserDash.execute_script("return document.body.parentNode.scrollHeight")
    browserDash.set_window_size(total_width, total_height)
    try:
        servidores = ["130","253","243"]
        browserDash.maximize_window()
        for x in servidores:
            # URL para conectarse a la página web
            url = f'https://10.1.251.{x}/tmui/login.jsp?msgcode=2&'
            browserDash.get(url)
            time.sleep(2)
            username = browserDash.find_element(By.CSS_SELECTOR, "#username")
            username.clear()
            username.send_keys("nechever")
            print("ESCRIBIENDO USUARIO")
            time.sleep(2)
            password = browserDash.find_element(By.CSS_SELECTOR, "#passwd")
            password.clear()
            password.send_keys("Teleco062024")
            time.sleep(2)
            print("ESCRIBIENDO PASSWORD")
            btnlogin = browserDash.find_element(By.CSS_SELECTOR,"#loginform > button")
            btnlogin.click()
            print("ESCRIBIENDO REFRESH")
            time.sleep(3)
            page_source = browserDash.page_source
            # Parsear el código HTML con BeautifulSoup
            soup = BeautifulSoup(page_source, "html.parser")
            # Encontrar todos los enlaces
            links = soup.find_all("a")
            time.sleep(1)
            cont=0
            for link in links:
                url = link.get("href")  # Obtener el atributo href (la URL del enlace)
                if url:  # Verificar que el enlace no sea None  # Imprimir la URL (opcional)
                    if cont == 0:
                        if "dashboard" in url.lower():  # Buscar la palabra "dashboard" en la URL
                            # Aquí puedes hacer algo con el enlace, como abrirlo en Selenium
                            dash = f"https://10.1.251.{x}/"
                            print(url)
                            browserDash.get(dash +url)
                            time.sleep(3)
                            hour3 = browserDash.find_element(By.CSS_SELECTOR,"#grid > li:nth-child(5) > div > div.card-header > span.card-control.ng-scope > div > select > option:nth-child(2)")
                            hour3.click()
                            time.sleep(3)
                            hour3 = browserDash.find_element(By.CSS_SELECTOR,"#grid > li:nth-child(8) > div > div.card-header > span.card-control.ng-scope > div > select > option:nth-child(2)")
                            hour3.click()
                            time.sleep(5)
                            name = f"Dashboard{x}.png"
                            #browserDash.execute_script("window.scrollBy(0, 150);")  # Scroll hacia abajo por 150 píxeles
                            time.sleep(3)
                            browserDash.save_screenshot(name)
                            time.sleep(1)
                            browserDash.get_screenshot_as_file(name)
                            time.sleep(1)
                            cont +=1
                    if "243" in x :
                        if "local_traffic" in url.lower():
                            dash = f"https://10.1.251.243"
                            browserDash.get(dash +url)
                            time.sleep(3)
                            browserDash.switch_to.frame("contentframe")
                            # Espera hasta que el elemento esté presente en la página
                            select_element = Select(browserDash.find_element(By.NAME, "section"))
                            select_element.select_by_value(str("pool"))
                            time.sleep(3)
                            selectPool = browserDash.find_element(By.CSS_SELECTOR,"#stats_form > div > table > tbody > tr:nth-child(1) > td.settings > select > option:nth-child(7)")
                            selectPool.click()
                            time.sleep(3)
                            selectTop= browserDash.find_element(By.CSS_SELECTOR,"#list_header > tr.columnhead.color2 > td:nth-child(7) > a")
                            selectTop.click()
                            time.sleep(3)
                            selectPlus = browserDash.find_element(By.CSS_SELECTOR,"#list_header > tr.columnhead.color2 > td:nth-child(3) > div")
                            selectPlus.click()
                            time.sleep(3)
                            namePools = "Pools243.png"
                            browserDash.save_screenshot(namePools)
                            browserDash.switch_to.default_content()
                            pass
    finally:
        browserDash.quit()

def login(browser):
    nombre =browser.find_element(By.XPATH, '//*[@id="usernameField"]')
    nombre.send_keys("admin")

    passw = browser.find_element(By.ID, 'passwordField')
    passw.send_keys("Pacifico11")

    login = browser.find_element(By.CSS_SELECTOR, '#loginButtonContainer > input')
    login.click()

def riverbed_dashboard_extraction():
    ChromeDriverManager().install()
    browserRiver =  webdriver.Chrome(set_chrome_options())
    total_width = browserRiver.execute_script("return document.body.offsetWidth")
    total_height = browserRiver.execute_script("return document.body.parentNode.scrollHeight")
    browserRiver.set_window_size(total_width, total_height)
    try:
        url = 'https://10.225.200.25/pages/report_interactive?id=10079'
        browserRiver.get(url)
        time.sleep(3)
        main_window = browserRiver.current_window_handle
        browserRiver.maximize_window()
        time.sleep(3)

        # Obtener la hora actual
        now = datetime.datetime.now().hour

        # Definir los rangos
        ranges = {
            'early_morning': (5, 8),
            'morning': (9, 12),
            'afternoon': (13, 16),
            'evening': (17, 20)
        }
        # Determinar el rango anterior según la hora actual
        previous_range = None
        if now < 11:
            previous_range = ranges['early_morning']
        elif 11 <= now < 15:
            previous_range = ranges['morning']
        elif 15 <= now < 19:
            previous_range = ranges['afternoon']
        else:
            # Si es después de las 8 PM, captura el rango de la tarde
            previous_range = ranges['evening']

        # Convertir el rango a formato AM/PM
        start_time = datetime.datetime.strptime(f"{previous_range[0]}:00", "%H:%M").strftime("%I:%M %p")
        end_time = datetime.datetime.strptime(f"{previous_range[1]}:00", "%H:%M").strftime("%I:%M %p")




        # Iniciar sesión
        try:
            login(browserRiver)
        except:
            print("No se logeo")

        time.sleep(12)

        desac = browserRiver.find_element(By.ID,"timeToolbarAutoUpdateCheckbox")
        browserRiver.execute_script("arguments[0].click();", desac)
        #desac.click()

        time.sleep(3)

        horario = browserRiver.find_element(By.ID,"formatted-date-text")
        horario.click()

        input_elements = browserRiver.find_elements(By.CSS_SELECTOR,".yui3-rbt-field-time-picker-content")
        print(input_elements)
        time.sleep(3)
        # Asegurándose de que hay al menos dos elementos para interactuar
        if len(input_elements) >= 2:
            input_elements[0].clear()
            input_elements[0].send_keys(start_time)  # Enviar el inicio del rango
            time.sleep(2)
            input_elements[1].clear()
            input_elements[1].send_keys(end_time)    # Enviar el fin del rango
            time.sleep(2)
        else:
            print("No se encontraron suficientes elementos input")
        time.sleep(2)
        btnok = browserRiver.find_element(By.ID,"popOk")
        browserRiver.execute_script("arguments[0].click();", btnok)

        time.sleep(12)
        browserRiver.save_screenshot("DashboardOmnicanalidad1.png")

        time.sleep(4)
        pass
    finally:
        browserRiver.quit()


def send_reports_via_telegram_personal(chat_id):
    tb = telebot.TeleBot(TOKEN)
    fecha_actual = datetime.datetime.now()
    text = f"Graficos de Dashborad {str(fecha_actual)}"
    tb.send_message(chat_id, text)
    for archivo in os.listdir(ROOT_DIR):
        if archivo.lower().endswith('.png'):
            with open(os.path.join(ROOT_DIR, archivo), 'rb') as photo:
                tb.send_photo(chat_id, photo)
                tb.send_message(chat_id, archivo.split(".")[0])
                time.sleep(2)


def send_reports_via_telegram():
    tb = telebot.TeleBot(TOKEN)
    fecha_actual = datetime.datetime.now()
    try:
        for chat_id in CHAT_IDS:
            text = f"Graficos de Dashborad {str(fecha_actual)}"
            tb.send_message(chat_id, text)
            try:
                for archivo in os.listdir(ROOT_DIR):
                    if archivo.lower().endswith('.png'):
                        with open(os.path.join(ROOT_DIR, archivo), 'rb') as photo:
                            tb.send_photo(chat_id, photo)
                            tb.send_message(chat_id, archivo.split(".")[0])
                            time.sleep(2)
            except:
                tb.send_message(chat_id,f"No hay imagen {archivo} ")
    except:
        print(f"Usuario {chat_id} no se entrego")
    tb.close()


def delete_old_reports():
    for archivo in os.listdir(ROOT_DIR):
        if archivo.lower().endswith('.png') and archivo.lower().startswith("dashb"):
            try:
                os.remove(os.path.join(ROOT_DIR, archivo))
            except Exception as e:
                print(f"Error al eliminar {archivo}. Razón: {e}")
        if archivo.lower().endswith('.png') and archivo.lower().startswith("pool"):
            try:
                os.remove(os.path.join(ROOT_DIR, archivo))
            except Exception as e:
                print(f"Error al eliminar {archivo}. Razón: {e}")


def scheduled_task():
    f5_dashboard_extraction()
    riverbed_dashboard_extraction()
    send_reports_via_telegram()
    delete_old_reports()


def personal_report(chat_id):
    f5_dashboard_extraction()
    riverbed_dashboard_extraction()
    send_reports_via_telegram_personal(chat_id)
    delete_old_reports()
    

def main():
    schedule.every().day.at("07:50").do(scheduled_task)
    schedule.every().day.at("11:50").do(scheduled_task)
    schedule.every().day.at("15:50").do(scheduled_task)
    schedule.every().day.at("19:50").do(scheduled_task)

    while True:
        schedule.run_pending()
        time.sleep(20)


if __name__ == '__main__':
    scheduled_task()
