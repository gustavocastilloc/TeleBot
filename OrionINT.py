import os
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import datetime as dt
from selenium.webdriver.support.ui import Select
from PIL import Image
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import datetime
import numpy as np
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import telebot
from OrionAuto import reporteria
from pandas.plotting import table
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from fuzzywuzzy import process
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from concurrent.futures import ThreadPoolExecutor
import matplotlib.pyplot as plt
import matplotlib



archivoG = pd.read_excel("Base_Completa.xlsx")
DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CHROMEDRIVER_PATH = 'chromedriver.exe'
TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
CHAT_IDS = [1397965849, 5781054062, 6055198218,6512803121,5686702637,6564107654,6383100715,1176200280]
archivo = "Reporte.xlsx"


url = 'https://192.168.10.31/orion/netperfmon/events.aspx'
ROOT_DIR ="c:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\"
smtp_server = "smtp.office365.com"
smtp_port = 587
smtp_user = "ggcastil@pacifico.fin.ec"
smtp_password = "fedAVR96oasis"
correos = 'ggcastil@pacifico.fin.ec,mcaalvar@pacifico.fin.ec,sgellibe@pacifico.fin.ec,wdjara@pacifico.fin.ec,scastane@pacifico.fin.ec,cpcampos@pacifico.fin.ec,frevelo@pacifico.fin.ec,ajsuarez@pacifico.fin.ec'
correos = correos.split(",")
pathChrom= os.path.expanduser('~')+'\\AppData\\Local\\Google\\Chrome\\User Data\\telcombas'
def configurar_chrome():
    options = Options()
    arguments = [
        "--disable-gpu",
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.54 Safari/537.36",
        "--excludeSwitches", "enable-automation",
        "--disable-blink-features=AutomationControlled",
        "--useAutomationExtension", "False",
        "--disable-extensions",
        "--ignore-certificate-errors",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--headless",
        "window-size=1400x1050",
        f"--user-data-dir={pathChrom}"
    ]
    
    for arg in arguments:
        options.add_argument(arg)

    return options

def iniciar_navegador(options, url):
    ChromeDriverManager().install()
    browser = webdriver.Chrome(options)
    browser.get(url)
    browser.maximize_window()
    time.sleep(5)
    return browser

def login_navegador(browser, usuario, contrasena):
    try:
        campo_usuario = browser.find_element(By.CLASS_NAME, "password-policy-username")
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)
        time.sleep(2)

        campo_contrasena = browser.find_element(By.CLASS_NAME, "password-input")
        campo_contrasena.clear()
        campo_contrasena.send_keys(contrasena)
        time.sleep(2)

        boton_login = browser.find_element(By.ID, "ctl00_BodyContent_LoginButton")
        boton_login.click()
        time.sleep(10)
    except:
        print("Ya está logueado")

def periodo_hasta_ahora():
    """
    Devuelve el periodo desde las 8 pm del día anterior hasta ahora.
    """
    fecha_fin = datetime.date.today().strftime('%d/%m/%Y')
    hora_fin = datetime.datetime.now().strftime('%H:%M')
    
    fecha_inicio = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%d/%m/%Y')
    hora_inicio = "20:00"
    
    return fecha_inicio, hora_inicio, fecha_fin, hora_fin

def extraer_info_pagina(browser):
    """
    Extrae la tabla de eventos de la página web cargada en el navegador.
    
    Args:
        browser (webdriver object): El navegador donde se ha cargado la página.
        
    Returns:
        pd.DataFrame: DataFrame con las columnas 'Fecha' y 'Notificación'.
    """
    
    # Mediante BeautifulSoup, extraemos la información HTML de la página y buscamos la tabla de eventos
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    elementosTime = soup.find_all("td")
    
    fechaL = []
    notificacionL = []
    
    # Aquí recorremos los registros obtenidos, donde extraemos solo la fecha y la descripción
    try:
        for y in elementosTime:
            dato = str(y.text)
            dato = dato.strip().split("\n")
            if len(dato[0]) > 1:
                if "/" in dato[0] and len(dato[0]) < 20:
                    fechaL.append(dato[0])
                if len(dato[0]) > 28:
                    notificacionL.append(dato[0])
    except:
        print("No se encontró información o hubo un error al extraer los datos.")

    # Creamos un dataframe con la información obtenida
    diccionario = {"Fecha": fechaL, "Notificacion": notificacionL}
    
    # Encuentra el tamaño máximo entre todas las listas en el diccionario
    max_len = max(len(lst) for lst in diccionario.values())
    
    # Rellena las listas más cortas con NaN hasta que todas tengan el mismo tamaño
    for key in diccionario.keys():
        diccionario[key] += [np.nan] * (max_len - len(diccionario[key]))
        
    dataframe = pd.DataFrame(diccionario)
    print(dataframe)
    return dataframe

def limpiar_dataframe(df):
    df = df.dropna()
    filtro = df['Notificacion'].str.contains('has stopped|is responding again|reboot', case=False)
    return df[filtro]

def extraer_dispositivos(df):
    dispositivos = df['Notificacion'].str.extract(r'^(.+?) (?:has stopped)').dropna()
    print(dispositivos)
    return dispositivos[0].str.split("has stopped").str[0]

def procesar_notificaciones_por_dispositivo(df, dispositivo):
    df_filtered = df[df['Notificacion'].str.contains(dispositivo)]
    df_filtered = df_filtered.sort_values('Fecha', ascending=True)
    return evaluar_notificaciones(df_filtered, dispositivo)
def evaluar_notificaciones(df, dispositivo):
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%m/%d/%Y %H:%M %p')
    
    df_resultados = pd.DataFrame(columns=['Enlace', 'Fecha Down', 'Fecha Up', 'Tiempo', 'Estado'])

    df_reboot = df[df['Notificacion'].str.contains('reboot')]
    df_reboot = df_reboot.drop_duplicates(subset=['Notificacion'], keep='first')

    df_other = df[~df['Notificacion'].str.contains('reboot')]

    df = pd.concat([df_reboot, df_other], ignore_index=True)
    df = df.sort_values(by='Fecha', ascending=False)
    df.drop_duplicates(keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)
    print(df)

    indices = df[df['Notificacion'].str.contains('rebooted', na=False)].index.tolist()
    indices_para_eliminar = [indices[i] for i in range(1, len(indices)) if indices[i] == indices[i-1] + 1]
    df.drop(indices_para_eliminar, inplace=True)
    df.reset_index(drop=True, inplace=True)

    while len(df["Notificacion"]) != 0:
        indices = df[df['Notificacion'].str.contains('rebooted', na=False)].index.tolist()
        indices = indices[::-1]
        print(indices)
        print(df)

        indices_eliminar = []
        if len(indices) > 0:
            for i in indices:
                notificacion_actual = df.loc[i, 'Notificacion']
                fecha_actual = df.loc[i, 'Fecha']
                print(df)
                if 'rebooted' in notificacion_actual:
                    fecha_down, fecha_up = None, None

                    if i == 0 and len(df["Notificacion"]) <= 2:
                        fecha_down = "Reboot Duplicado"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i])

                    elif i == len(df)-1 and "again" in df.loc[i-1, "Notificacion"]:
                        print("EQUIVOCADO1")
                        fecha_down = "Reboot"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i-1, i])

                    elif i == len(df)-1 and "stopped" in df.loc[i-1, "Notificacion"]:
                        fecha_down = "Reboot Duplicado"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i])

                    elif i+1 < len(df) and 'again' in df.loc[i+1, 'Notificacion']:
                        print("EQUIVOCADO3")
                        if len(df["Fecha"]) > i+2:
                            fecha_down = df.loc[i+2, 'Fecha'] if i+2 < len(df) else None
                            indices_eliminar.extend([i, i+1, i+2])
                            fecha_up = df.loc[i+1, 'Fecha']
                        else:
                            fecha_up = df.loc[i+1, 'Fecha']
                            fecha_down = "No aparecio"
                            indices_eliminar.extend([i, i+1])

                    elif i+1 <= len(df["Notificacion"]) and 'stopped' in df.loc[i+1, 'Notificacion']:
                        fecha_down = df.loc[i+1, 'Fecha']
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i-1, i, i+1])

                    tiempoR = 0
                    try:
                        tiempo = abs((fecha_up - fecha_down).total_seconds() / 60)
                        tiempoR += tiempo
                        print(tiempoR)
                    except Exception as e:
                        print(f"Error en el cálculo del tiempo: {e}")
                        tiempoR = "revisar"
                        tiempo = 0

                    if fecha_down is not None and fecha_up is not None:
                        new_row = {
                            'Enlace': dispositivo,
                            'Fecha Down': fecha_down,
                            'Fecha Up': fecha_up,
                            'Tiempo': tiempo,
                            'Suma': tiempoR,
                            'Estado': 'reboot'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)

            try:
                indices_eliminar = list(set(indices_eliminar))
                df.drop(index=indices_eliminar, inplace=True)
                df.reset_index(drop=True, inplace=True)
            except:
                df.drop(df.index[-1], inplace=True)

        if len(df["Notificacion"]) > 2:
            notifyUlt = df.iloc[-2]['Notificacion']
            notifyPen = df.iloc[-1]['Notificacion']
            tiempoUlt = df.iloc[-2]['Fecha']
            tiempoPen = df.iloc[-1]['Fecha']

            if "again" in notifyUlt:
                if "stopped" in notifyPen:
                    fecha_up = tiempoUlt
                    fecha_actual = tiempoPen
                    print("ENTRO A LA RESTA DE Tiempo")
                    print(df)
                    tiempo = abs((fecha_up - fecha_actual).total_seconds() / 60)
                    tiempoH = tiempo
                    print(tiempoH)
                    if tiempoH > 2:
                        new_row = {
                            'Enlace': dispositivo,
                            'Fecha Down': fecha_actual,
                            'Fecha Up': fecha_up,
                            'Tiempo': tiempo,
                            'Suma': tiempoH,
                            'Estado': 'Caido y Recuperado'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                        ultimos_indices = df.index[-2:]
                        df.drop(index=ultimos_indices, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        print(df)
                    else:
                        ultimos_indices = df.index[-1:]
                        df.drop(index=ultimos_indices, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        print(df)

                elif "again" in notifyPen:
                    ultimos_indices = df.index[-1:]
                    df.drop(index=ultimos_indices, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    print(df)

            elif "stopped" in notifyUlt:
                ultimos_indices = df.index[-1:]
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)
                print(df)

        elif len(df["Notificacion"]) == 2:
            notifyUlt = df.iloc[-2]['Notificacion']
            notifyPen = df.iloc[-1]['Notificacion']
            tiempoUlt = df.iloc[-2]['Fecha']
            tiempoPen = df.iloc[-1]['Fecha']

            if "again" in notifyUlt:
                if "again" in notifyPen:
                    ultimos_indices = df.index[-2:]
                    df.drop(index=ultimos_indices, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                elif "stopped" in notifyPen:
                    fecha_up = tiempoUlt
                    fecha_actual = tiempoPen
                    print("ENTRO A LA RESTA DE Tiempo")
                    print(df)
                    tiempo = abs((fecha_up - fecha_actual).total_seconds() / 60)
                    tiempoH = tiempo
                    print(tiempoH)
                    if tiempoH > 0:
                        new_row = {
                            'Enlace': dispositivo,
                            'Fecha Down': fecha_actual,
                            'Fecha Up': fecha_up,
                            'Tiempo': tiempo,
                            'Suma': tiempoH,
                            'Estado': 'Caido y Recuperado'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                        ultimos_indices = df.index[-2:]
                        df.drop(index=ultimos_indices, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        print(df)
                    else:
                        ultimos_indices = df.index[-2:]
                        df.drop(index=ultimos_indices, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        print(df)

            elif "stopped" in notifyPen:
                print("QUE ESTA PASANDO")
                print(df)
                fecha_up = tiempoPen
                fecha_actual = tiempoUlt
                print("Evaluando")
                print(df)
                new_row = {
                    'Enlace': dispositivo,
                    'Fecha Down': fecha_actual,
                    'Fecha Up': None,
                    'Tiempo': None,
                    'Suma': None,
                    'Estado': 'Caido'
                }
                print(new_row)
                df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                ultimos_indices = df.index[-2:]
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)

            elif "responding" in notifyPen:
                ultimos_indices = df.index[-2:]
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)

            elif "rebooted" in notifyUlt:
                print("QUE ESTA PASANDO")
                print(df)
                fecha_actual = tiempoUlt
                print("Evaluando")
                print(df)
                new_row = {
                    'Enlace': dispositivo,
                    'Fecha Down': fecha_actual,
                    'Fecha Up': None,
                    'Tiempo': None,
                    'Suma': None,
                    'Estado': 'Caido'
                }
                print(new_row)
                df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                ultimos_indices = df.index[-1:]
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)

        elif len(df["Notificacion"]) == 1:
            notifyUlt = df.iloc[-1]['Notificacion']
            print(notifyUlt)
            tiempoUlt = df.iloc[-1]['Fecha']
            print(tiempoUlt)
            print("SOLO AQUI LLEGAN LOS VALIENTES")
            print(df)
            print(df)
            if "stopped" in notifyUlt:
                print(df)
                print("CAIDO")
                new_row = {
                    'Enlace': dispositivo,
                    'Fecha Down': tiempoUlt,
                    'Fecha Up': None,
                    'Tiempo': None,
                    'Suma': None,
                    'Estado': "Caido"
                }
            elif "again" in notifyUlt:
                print(df)
                print("CAIDO")
                new_row = {
                    'Enlace': dispositivo,
                    'Fecha Down': None,
                    'Fecha Up': tiempoUlt,
                    'Tiempo': None,
                    'Suma': None,
                    'Estado': "Revisar"
                }

            print(new_row)
            df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
            ultimos_indices = df.index[-1:]
            df.drop(index=ultimos_indices, inplace=True)
            df.reset_index(drop=True, inplace=True)

    return df_resultados

def crear_mensaje(correo_destino, fecha_actual, final):
    message = MIMEMultipart()
    message['From'] = "mcaalvar@pacifico.fin.ec"
    message['To'] = correo_destino.strip()
    message['Subject'] = f'INCIDENCIA DE ENLACE INTERNEEETTTT {fecha_actual} '
    #mensaje = f"Por favor revisar incidencia de este enlace \n \n {final['Enlace'].values} \n \n Capture adjunta en el correo"
    #message.attach(MIMEText(mensaje, 'plain'))
    # Seleccionar las columnas 'Enlace' y 'Tiempo'
    final = final[["Enlace", "Tiempo"]]
    html = final.to_html()
    cuerpo = MIMEText(html, 'html')
    message.attach(cuerpo)
    return message

def enviar_correo(message):
    smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
    smtp_connection.starttls()
    smtp_connection.login(smtp_user, smtp_password)
    smtp_connection.sendmail(message['From'], message['To'], message.as_string())
    smtp_connection.quit()



def send_reports_via_telegram_personal(chat_ids,enlace):
    tb = telebot.TeleBot(TOKEN)
    fecha_actual = datetime.datetime.now()
    text = f"Reporte Orion {str(fecha_actual)}"
    for chat_id in chat_ids:
        tb.send_message(chat_id, text)
        try:

            mensaje= f"Reporte Orion Internet Enlace caido {enlace}"
            tb.send_message(chat_id, mensaje)
            tb.send_message(chat_id, mensaje)
            tb.send_message(chat_id, mensaje)
            tb.send_message(chat_id, mensaje)
            tb.send_message(chat_id, mensaje)
            time.sleep(2)
        except:
            tb.send_message(chat_id, "No se encontro el archivo "+enlace)




def enviar_correo_con_excel(correos, fecha_actual, df_final):
    """
    Envia un correo con un archivo Excel adjunto a una lista de correos.

    Parámetros:
    - correos (list): Lista de direcciones de correo electrónico a las que enviar el correo.
    - fecha_actual (datetime): Fecha actual para referenciar en el correo.
    - df_final (DataFrame): DataFrame final que puede usarse para generar información en el correo.
    - archivo (str): Ruta del archivo Excel que se adjuntará.
    """
    for correo in correos:
        # Crear el mensaje
        mensaje = crear_mensaje(correo, fecha_actual, df_final)
        # Enviar el correo
        enviar_correo(mensaje)

# Ejemplo de uso:
# correos = ["email1@example.com", "email2@example.com"]
# fecha_actual = pd.Timestamp.now()
# df_final = pd.DataFrame()  # Asumiendo que tienes un DataFrame ya preparado
# archivo = "ruta_al_archivo.xlsx"
# enviar_correo_con_excel(correos, fecha_actual, df_final, archivo)



def main():
    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)
    # Iniciar sesión en el navegador
    login_navegador(browser,"mcaalvar","Silimarina97.")
    # Seleccionar el periodo de tiempo
    # Extraer información de la página
    time.sleep(3)
    btnTime = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList")
    btnTime.click()
    btnCustom = browser.find_element(By.XPATH, '//*[@id="ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList"]/option[3]')
    btnCustom.click()
    time.sleep(1)
    btnRefresh = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
    btnRefresh.click()
    time.sleep(5)  # Wait for the page to load, adjust as needed
    
    df_final = pd.DataFrame()
    time.sleep(2)

    # Refresh the page and check for "Node down" notifications
    while True:
        # Refresh the page
        # Extract information from the page
        df = extraer_info_pagina(browser)
        df = limpiar_dataframe(df)
        dispositivos_unicos = list(set(extraer_dispositivos(df)))

        if not dispositivos_unicos:  # Check if dispositivos_unicos is empty
            time.sleep(120)  # Wait for 5 minutes before trying again
        else:
            # Create a copy of the current dataframe
            df_copy = df_final.copy()

            for dispositivo in dispositivos_unicos:
                df_dispo = procesar_notificaciones_por_dispositivo(df, dispositivo)
                df_final = pd.concat([df_dispo, df_final], ignore_index=True)
            
            df_final.drop_duplicates(inplace=True)
            df_final.sort_values("Fecha Down", ascending=False, inplace=True)
            df_final.reset_index(drop=True, inplace=True)
            print(df_final)
            if not df_final.empty:
                fecha_up = df_final['Fecha Down'][0]
                fecha_actual = datetime.datetime.now()
                tiempo = abs((fecha_up - fecha_actual).total_seconds() / 60)
                print(tiempo)
                print(tiempo)
                print(df_final)
                print(df_copy)

                if tiempo > 3600:
                    df_final = pd.DataFrame()
                    print("Se ha reiniciado el dataframe")  
                # Check if the new dataframe is the same as the previous one
                if df_final.equals(df_copy):
                    print("son iguales")
                    time.sleep(60)  # Wait for 1 minute before trying again
                    continue

                # Check if the first record is within the current hour
                if fecha_up.hour == fecha_actual.hour:
                    print("First record is within the current hour")
                    time.sleep(60)  # Wait for 1 minute before trying again
                    continue

                # Check for "Node down" notifications
                if not df_final.empty:
                    # Send email alerts
                    df_diff = df_final[['Enlace', 'Fecha Down','Tiempo','Estado']]
                    if not df_diff.empty:
                        if df_diff.equals(df_copy):
                            print("Same dataframe, skipping email and telegram report")
                        else:
                            enviar_correo_con_excel(correos, datetime.datetime.now(), df_diff)
                            enlace = df_diff.to_string(index=False)
                            send_reports_via_telegram_personal(CHAT_IDS, enlace)
                    # Update the previous dataframe with the new one
                    df_copy = df_final.copy()

                    time.sleep(10)  # Wait for 15 minutes before trying again
                else:
                    time.sleep(60)  # Adjust the time interval as needed for regular refresh
                    print(df_final)

