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
import matplotlib.pyplot as plt
from OrionAuto import reporteria
from pandas.plotting import table
from datetime import datetime, timedelta

DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CHROMEDRIVER_PATH = 'chromedriver.exe'
TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
CHAT_IDS = [1397965849, 5781054062, 6055198218,6512803121,5686702637.6564107654,6383100715,6500653052,1176200280]


url = 'https://10.1.231.243/orion/netperfmon/events.aspx'
ROOT_DIR ="C:\\Users\\nicho\\OneDrive\\Documentos\\Gustavo\\ActivaciondeCheckPoint\\"
smtp_server = "smtp.office365.com"
smtp_port = 587
smtp_user = "mcaalvar@pacifico.fin.ec"
smtp_password = "Silimarina97."
correos = 'mcaalvar@pacifico.fin.ec,sgellibe@pacifico.fin.ec,wdjara@pacifico.fin.ec,cpcampos@pacifico.fin.ec,frevelo@pacifico.fin.ec,ajsuarez@pacifico.fin.ec,ggcastil@pacifico.fin.ec'
correos = correos.split(",")



def crearArchivoTotal(fecha_fin):
    try:
        df_final_total = pd.read_excel(f"ReporteTickets{str(fecha_fin).replace('/','-')}.xlsx")
        return df_final_total
    except:
        df_final_total = pd.DataFrame()
        return df_final_total
    


def configurar_chrome():
    options = webdriver.ChromeOptions()
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
        "window-size=1400x1050"
    ]
    
    for arg in arguments:
        options.add_argument(arg)

    return options


def iniciar_navegador(options, url):
    browser = webdriver.Chrome(options=options)
    browser.get(url)
    browser.maximize_window()
    time.sleep(15)
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





def fechas_tickets(fecha):

    # Convertir la cadena de texto a un objeto datetime
    start_date = datetime.strptime(fecha, '%d/%m/%Y')
    
    # Obtener la fecha actual
    current_date = datetime.now()
    
    # Crear una lista para almacenar las fechas
    date_list = []
    
    # Generar las fechas
    while start_date < current_date:
        date_list.append(start_date.strftime('%d/%m/%Y'))
        start_date += timedelta(days=1)
    
    return date_list



def selecc_horario(browser, fecha_inicio):
    # Generar lista de fechas
    date_list = fechas_tickets(fecha_inicio)
    print(date_list)
    total = pd.DataFrame()
    #Seleccionar el periodo de tiempo
    btnTime = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList")
    btnTime.click()
    btnCustom = browser.find_element(By.XPATH, '//*[@id="ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList"]/option[13]')
    btnCustom.click()
    time.sleep(2)
    for date in date_list:
        # Configurar la fecha y hora de inicio
        txtDateIn = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtDatePicker")
        txtDateIn.clear()
        txtDateIn.send_keys(date)
        time.sleep(2)
        
        txtDateInHora = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtTimePicker")
        txtDateInHora.clear()
        txtDateInHora.send_keys("00:00")
        time.sleep(1)
        
        # Configurar la fecha y hora de fin
        txtDateON = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtDatePicker")
        txtDateON.clear()
        txtDateON.send_keys(date)
        time.sleep(1)
        
        txtDateOnHora = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtTimePicker")
        txtDateOnHora.clear()
        txtDateOnHora.send_keys("23:59")
        time.sleep(1)
        print(date)
        print(date)
        print(date)
        print(date)
        # Refrescar la información
        btnRefresh = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
        btnRefresh.click()
        time.sleep(25)  # Espera para que la página cargue, ajusta según necesidad

        df = extraer_info_pagina(browser)

        df = limpiar_dataframe(df)

        dispositivos_unicos = extraer_dispositivos(df)
        dispositivos = []
        tiempos = []
        reportados = []
        detalles = []
        for dispositivo in dispositivos_unicos:
            dispositivo, tiempo, detalle  = procesar_notificaciones_por_dispositivo(df, dispositivo)
            if dispositivo:
                dispositivos.append(dispositivo)
                tiempos.append(tiempo)
                detalles.append(detalle)
                reportados.append("Na")
        time.sleep(25)
            # Crear el dataframe final
        df_final = crear_dataframe(dispositivos, tiempos,detalle,reportados)
        df_final.info()
        df_final = clean_column(df_final, "Enlace")
        total = pd.concat([total,df_final])
    return total

def clean_column(df, column_name):
    
    # Caracteres no permitidos en sistemas Windows y sistemas basados en Unix
    forbidden_characters = r'\/:*?"<>|'
    
    # Usamos el método `str.replace` de pandas para eliminar cada carácter no permitido
    for char in forbidden_characters:
        df[column_name] = df[column_name].str.replace(char, '', regex=False)
        df[column_name] = df[column_name].str.strip()
    print(df)
    return df

def limpiar_dataframe(df):
    df = df.dropna()
    filtro = df['Notificacion'].str.contains('has stopped|is responding again|reboot', case=False)
    return df[filtro]

def extraer_dispositivos(df):
    dispositivos = df['Notificacion'].str.extract(r'^(.+?) (?:has stopped)').dropna()
    return dispositivos[0].str.split("has stopped").str[0].unique()

def procesar_notificaciones_por_dispositivo(df, dispositivo):
    df_filtered = df[df['Notificacion'].str.contains(f"{dispositivo}")]
    enlace = dispositivo.split(" ")[0]

    if not str(enlace).lower().startswith("sw"):
        tiene_reboot = df_filtered['Notificacion'].str.contains("reboot").any()
        if tiene_reboot:
            index_reboot = df_filtered['Notificacion'].str.contains("reboot").tolist().index(True)
            df_filtered = df_filtered.drop(index=df_filtered.index[index_reboot:])
        
        df_filtered = df_filtered.sort_values('Fecha', ascending=False)
        
        return evaluar_notificaciones(df_filtered, dispositivo)
    print(dispositivo)
    return None, None, None

def evaluar_notificaciones(df, dispositivo):
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')
    print(df)
    ff = df['Notificacion'].str.contains("is responding again.")
    fi = df['Notificacion'].str.contains("has stopped responding")
    
    ff_count = ff.sum()
    fi_count = fi.sum()

    if ff_count == 0 and fi_count >= 1:
        fechaCaido = "Caido"
        return dispositivo, "Caido", fechaCaido


    if ff_count > 0 and fi_count > 0:
        df_ff = df[ff]
        df_fi = df[fi]

        # Ordenar por fechas en orden descendente
        df_ff = df_ff.sort_values(by="Fecha", ascending=False)
        df_fi = df_fi.sort_values(by="Fecha", ascending=False)
        print(dispositivo)
        
        while len(df_ff) > 0 and len(df_fi) > 0:
            fecha_ff = df_ff['Fecha'].iloc[-1]  # Última fecha en "is responding again."
            fecha_fi = df_fi['Fecha'].iloc[-1]  # Última fecha en "has stopped responding"
            if len(df_ff)< len(df_fi):
                df_fi = df_fi.iloc[:-1]
            if len(df_ff)> len(df_fi):   
                df_ff = df_ff.iloc[1:]

            diferencia_tiempo = (fecha_ff - fecha_fi).total_seconds() / 60
            
            if diferencia_tiempo > 5:
                print(diferencia_tiempo)
                notificaciones = f"{fecha_ff} Se recupero a esta hora, {fecha_fi} Se cayo a esta hora "
                return dispositivo, diferencia_tiempo, notificaciones
            else:
                # Eliminar los registros más recientes que no cumplen con el criterio
                df_ff = df_ff.iloc[:-1]
                df_fi = df_fi.iloc[:-1]

    

    return None, None, None


def crear_dataframe(dispositivos, tiempos,detalles,reportados):
    data = {
        'Enlace': dispositivos,
        'Tiempos': tiempos,
        'Detalle': detalles,
        'Reportado' : reportados
    }
    return pd.DataFrame(data)

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
    
    return dataframe

def main_personal():
    # Eliminar archivos antiguos

    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)

    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", "Silimarina97.")

    fecha_inicio = input("Ingrese la fecha DD/MM/AAAA")
    
    final = selecc_horario(browser, fecha_inicio)
    final.to_excel(f"ReporteFinal{str(fecha_inicio).replace('/','-')}.xlsx")


main_personal()