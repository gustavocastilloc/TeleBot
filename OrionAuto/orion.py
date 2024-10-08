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

import re



archivoG = pd.read_excel("Base_Completa.xlsx")
DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CHROMEDRIVER_PATH = 'chromedriver.exe'
TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
CHAT_IDS = [1397965849, 5781054062, 6055198218,6512803121,5686702637.6564107654,6383100715,1176200280]
archivo = "Reporte.xlsx"


url = 'https://10.1.231.243/orion/netperfmon/events.aspx'
ROOT_DIR ="c:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\"
smtp_server = "smtp.office365.com"
smtp_port = 587
smtp_user = "mcaalvar@pacifico.fin.ec"
smtp_password = "Silimarina97."
correos = 'mcaalvar@pacifico.fin.ec,sgellibe@pacifico.fin.ec,wdjara@pacifico.fin.ec,scastane@pacifico.fin.ec,cpcampos@pacifico.fin.ec,frevelo@pacifico.fin.ec,ajsuarez@pacifico.fin.ec,ggcastil@pacifico.fin.ec'
correos = correos.split(",")

def crearArchivoTotal(fecha_fin):
    try:
        df_final_total = pd.read_excel(f"ReporteDiario{str(fecha_fin).replace('/','-')}.xlsx")
        return df_final_total
    except:
        df_final_total = pd.DataFrame()
        return df_final_total
    

def pedir_dia_calendario(fecha, hora_inicio, hora_fin):
    """
    Establece el periodo para un día específico con horas de inicio y fin personalizadas.
    """
    fecha_inicio = (fecha - pd.Timedelta(days=1)).strftime('%d/%m/%Y')  # Día anterior
    fecha_fin = fecha.strftime('%d/%m/%Y')
    
    return fecha_inicio, hora_inicio, fecha_fin, hora_fin

def dividir_periodo_nocturno(fecha, num_segmentos=2):
    """
    Divide el periodo nocturno en varios segmentos.
    """
    segmentos = []
    horas_inicio = ["20:00", "02:00"]  # Ejemplo para 2 segmentos
    horas_fin = ["01:59", "08:00"]

    for inicio, fin in zip(horas_inicio, horas_fin):
        segmentos.append((inicio, fin))

    return segmentos


def pedir_dia_calendario_noche(fecha,inicio,fin,cont):
    # Convertir la fecha proporcionada en un objeto datetime usando el formato "%d/%m/%Y"
    fecha_dt = datetime.datetime.strptime(fecha, "%d/%m/%Y")
    # Restar un día a la fecha de inicio
    if cont == 0:
        fecha_inicio = fecha_dt - datetime.timedelta(days=1)
    else:
        fecha_inicio = fecha_dt
    # La fecha de fin es la misma que la proporcionada por el usuario
    fecha_fin = fecha_dt
    # Convertir las fechas de nuevo a formato de cadena usando el formato "%d/%m/%Y"
    fecha_inicio_str = fecha_inicio.strftime("%d/%m/%Y")
    print(fecha_inicio_str)
    fecha_fin_str = fecha_fin.strftime("%d/%m/%Y")
    print(fecha_fin_str)
    # Establecer la hora de inicio a las 8 am y la hora de fin a las 8 pm
    hora_inicio = inicio
    hora_fin = fin
    return fecha_inicio_str, hora_inicio, fecha_fin_str, hora_fin

def recolectar_datos_nocturnos(browser, fecha):
    segmentos = dividir_periodo_nocturno(fecha)
    dfs = []  # Lista para almacenar dataframes de cada segmento
    cont = 0 
    for inicio, fin in segmentos:
        fecha_inicio, hora_inicio, fecha_fin, hora_fin = pedir_dia_calendario_noche(fecha,inicio,fin,cont)
        selecc_horario(browser, fecha_inicio, hora_inicio, fecha_fin, hora_fin)
        df_segmento = extraer_info_pagina(browser)
        dfs.append(df_segmento)
        cont+=1

    df_final = pd.concat(dfs, ignore_index=True)
    return df_final

def eliminar_archivos_antiguos(ruta, extension='.png', dias=1):
    fecha_actual = datetime.date.today()
    un_dia_antes = fecha_actual - datetime.timedelta(days=dias)

    for archivo in os.listdir(ruta):
        ruta_archivo = os.path.join(ruta, archivo)
        if archivo.lower().endswith(extension) and datetime.datetime.fromtimestamp(os.path.getmtime(ruta_archivo)).date() < un_dia_antes:
            os.remove(ruta_archivo)
            print(f'Se eliminó: {ruta_archivo}')

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
        #"--headless",
        "window-size=1400x1050"
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
        time.sleep(3)
    except:
        print("Ya está logueado")


def selecc_horario(browser, fecha_inicio, hora_inicio, fecha_fin, hora_fin):
    # Seleccionar el periodo de tiempo
    browser.refresh()
    time.sleep(3)
    btnTime = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList")
    btnTime.click()
    btnCustom = browser.find_element(By.XPATH, '//*[@id="ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_TimePeriodList"]/option[13]')
    btnCustom.click()
    time.sleep(1)
    
    # Configurar la fecha y hora de inicio
    txtDateIn = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtDatePicker")
    txtDateIn.clear()
    txtDateIn.send_keys(fecha_inicio)

    time.sleep(1)
    
    txtDateInHora = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtTimePicker")
    txtDateInHora.clear()
    txtDateInHora.send_keys(hora_inicio)
    time.sleep(1)
    
    # Configurar la fecha y hora de fin
    txtDateON = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtDatePicker")
    txtDateON.clear()
    txtDateON.send_keys(fecha_fin)
    time.sleep(1)
    
    txtDateOnHora = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtTimePicker")
    txtDateOnHora.clear()
    txtDateOnHora.send_keys(hora_fin)
    time.sleep(1)
    
    # Refrescar la información
    btnRefresh = browser.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
    btnRefresh.click()
    time.sleep(5)  # Espera para que la página cargue, ajusta según necesidad


def periodo_hasta_ahora():
    """
    Devuelve el periodo desde las 8 pm del día anterior hasta ahora.
    """
    fecha_fin = datetime.date.today().strftime('%d/%m/%Y')
    hora_fin = datetime.datetime.now().strftime('%H:%M')
    
    fecha_inicio = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%d/%m/%Y')
    hora_inicio = "20:00"
    
    return fecha_inicio, hora_inicio, fecha_fin, hora_fin



def pedir_dia_calendario_dia(fecha):
    """
    Pide al usuario un día del calendario (fecha) y establece el periodo de ese día de 8 am a 8 pm.
    """
    # Pide al usuario la fecha del día específico en formato DD/MM/YYYY
    # La fecha de inicio y fin es la misma que la proporcionada por el usuario
    fecha_inicio = str(fecha)
    fecha_fin = str(fecha)
    # Establecer la hora de inicio a las 8 am y la hora de fin a las 8 pm
    hora_inicio = "08:00"
    hora_fin = "20:00"
    
    return fecha_inicio, hora_inicio, fecha_fin, hora_fin

def pedir_dia_calendarionew(fecha, hora_inicio, hora_fin):
    """
    Establece el periodo para un día específico con horas de inicio y fin personalizadas.
    """
    fecha_inicio = str(fecha)
    fecha_fin = str(fecha)
    
    return fecha_inicio, hora_inicio, fecha_fin, hora_fin


def dividir_dia_en_segmentos(fecha, num_segmentos=3):
    """
    Divide un día en varios segmentos.
    """
    segmentos = []
    horas_inicio = ["08:00", "12:00", "16:00"]  # Ejemplo para 3 segmentos
    horas_fin = ["11:59", "15:59", "20:00"]

    for inicio, fin in zip(horas_inicio, horas_fin):
        segmentos.append((inicio, fin))

    return segmentos

def recolectar_datos_por_segmentos(browser,fecha):
    segmentos = dividir_dia_en_segmentos(fecha)
    dfs = []  # Lista para almacenar dataframes de cada segmento

    for inicio, fin in segmentos:
        fecha_inicio, hora_inicio, fecha_fin, hora_fin = pedir_dia_calendarionew(fecha, inicio, fin)
        selecc_horario(browser, fecha_inicio, hora_inicio, fecha_fin, hora_fin)
        df_segmento = extraer_info_pagina(browser)

        dfs.append(df_segmento)

    df_final = pd.concat(dfs, ignore_index=True)
    return df_final



# def corregir_estados_reboot(df, time_margin='2min'):
#     """
#     Corrige los estados en un DataFrame de notificaciones para sincronizar eventos de reboot
#     en agencias con diferentes proveedores que tengan tiempos de caída y recuperación similares.

#     Parámetros:
#     - df (DataFrame): DataFrame con las columnas 'Enlace', 'Fecha Down', 'Fecha Up' y 'Estado'.
#     - time_margin (str): Margen de tiempo para considerar que dos fechas son similares (formato de timedelta).

#     Retorna:
#     - DataFrame: DataFrame con los estados ajustados.
#     """
    
#     # Convertir el margen de tiempo a Timedelta
#     time_margin = pd.Timedelta(time_margin)
    
#     # Asegurar que las columnas de fecha son datetime, manejando errores
#     df['Fecha Down'] = pd.to_datetime(df['Fecha Down'], errors='coerce')
#     df['Fecha Up'] = pd.to_datetime(df['Fecha Up'], errors='coerce')
#     # Separar la columna 'Enlace' para obtener 'Agencia' y 'Proveedor'
#     df['Agencia'] = df['Enlace'].apply(lambda x: ' '.join(x.split()[:-1]).lower().replace("backup","").replace(" ",""))
#     df['Proveedor'] = df['Enlace'].apply(lambda x: x.split()[-1])

#     # Crear una copia del DataFrame para realizar ajustes
#     data_adjusted = df.copy()

#     # Agrupar por 'Agencia' para procesar posibles correcciones
#     for agency, group in df.groupby('Agencia'):
#         print("###group en for###")
        
#         if group['Proveedor'].nunique() > 1:  # Más de un proveedor para la misma agencia
#             print("###Mas de un proveedor para la misma agencia")
#             print(group.head(20))
#             for i, row in group.iterrows():
#                 # Buscar entradas similares dentro del grupo
#                 similar_entries = group[
#                     (np.abs(group['Fecha Down'] - row['Fecha Down']) <= time_margin) &
#                     (np.abs(group['Fecha Up'] - row['Fecha Up']) <= time_margin) &
#                     (group['Proveedor'] != row['Proveedor'])
#                 ]
                
#                 # Si hay entradas similares y el estado actual no es 'reboot' pero otra entrada sí lo es
#                 if not similar_entries.empty:
#                     if 'reboot' in similar_entries['Estado'].values and row['Estado'] != 'reboot':
#                         data_adjusted.loc[i, 'Estado'] = 'reboot'
#         else:
#             print("###Solo un proveedor por agencia##")
#             print(group.head(20))
#     data_adjusted.reset_index(drop=True, inplace=True)
#     return data_adjusted
def corregir_estados_reboot(df, time_margin='1min'):
    # Convertir el margen de tiempo a Timedelta
    time_margin = pd.Timedelta(time_margin)
    print("######### Corregir estados ##########")
    print(df.head(30))
    # Asegurar que las columnas de fecha son datetime, manejando errores
    df['Fecha Down'] = pd.to_datetime(df['Fecha Down'], errors='coerce')
    df['Fecha Up'] = pd.to_datetime(df['Fecha Up'], errors='coerce')
    # Crear una copia del DataFrame para realizar ajustes
    data_adjusted = df.copy()

    # Agrupar por 'Agencia' o una parte común del nombre que no incluya 'Principal' o 'Backup'
    df['Agencia_base'] = df['Enlace'].apply(lambda x: ' '.join(x.split()[:-2]).lower().replace("backup", "").replace("principal", "").strip())

    for agency, group in df.groupby('Agencia_base'):
        # Filtrar enlaces principales en estado 'reboot', ignorando mayúsculas/minúsculas
        principal_reboot = group[(group['Estado'] == 'reboot') & (group['Enlace'].str.contains('Principal', case=False))]
        print(f"Este enlace es principal: \n", principal_reboot)
        # Si hay un principal en reboot
        if not principal_reboot.empty:
            for i, principal in principal_reboot.iterrows():
                # Buscar el backup correspondiente (misma agencia, mismo proveedor)
                backup = group[(group['Enlace'].str.contains('Backup', case=False)) ]
                
                print(f"Backups correspondientes:\n", backup)
                # Si existe el enlace backup
                if not backup.empty:
                    for j, backup_row in backup.iterrows():
                        # Verificar que las fechas coincidan entre el principal y el backup
                        fecha_down_principal = principal['Fecha Down']
                        fecha_up_principal = principal['Fecha Up']
                        fecha_down_backup = backup_row['Fecha Down']
                        fecha_up_backup = backup_row['Fecha Up']

                        # Comparar las fechas de down y up entre el principal y el backup con un margen de tiempo
                        if (np.abs(fecha_down_principal - fecha_down_backup) <= time_margin) and \
                           (np.abs(fecha_up_principal - fecha_up_backup) <= time_margin):

                            # Si el estado del backup es 'Caido y Recuperado', cambiar a 'reboot'
                            if backup_row['Estado'] == 'Caido y Recuperado':
                                data_adjusted.at[j, 'Estado'] = 'reboot'  # Cambiar el estado del backup a 'reboot'
                                print(f"El estado del enlace backup '{backup_row['Enlace']}' se cambió a 'reboot'.")
                
                

    # Reiniciar el índice del DataFrame ajustado
    data_adjusted.reset_index(drop=True, inplace=True)

    return data_adjusted


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
    print(dataframe.head())
    return dataframe



def limpiar_dataframe(df):
    df = df.dropna()
    filtro = df['Notificacion'].str.contains('has stopped|is responding again|reboot', case=False)
    return df[filtro]

def extraer_dispositivos(df):
    dispositivos = df['Notificacion'].str.extract(r'^(.+?) (?:has stopped)').dropna()
    print("extraer_dispositivos:")
    print(dispositivos)
    return dispositivos[0].str.split("has stopped").str[0]

def procesar_notificaciones_por_dispositivo(df, dispositivo):
    print(df)
    print(dispositivo)
        # Escapar caracteres especiales en la cadena dispositivo
    dispositivo_escaped = re.escape(dispositivo)

    # Filtrar el DataFrame usando la cadena escapada
    df_filtered = df[df['Notificacion'].str.contains(dispositivo_escaped)]
    df_filtered = df_filtered.sort_values('Fecha', ascending=True)
    return evaluar_notificaciones(df_filtered, dispositivo)



def evaluar_notificaciones(df, dispositivo):
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')
    
    df_resultados = pd.DataFrame(columns=['Enlace', 'Fecha Down', 'Fecha Up', 'Tiempo', 'Estado'])

        # Filtrar los registros que contienen 'reboot' en la columna 'Notificacion'
    df_reboot = df[df['Notificacion'].str.contains('reboot')]

    # Eliminar duplicados en la columna 'Notificacion' dentro de los filtrados
    df_reboot = df_reboot.drop_duplicates(subset=['Notificacion'], keep='first')

    # Opcionalmente, si deseas conservar los registros que no contienen 'reboot'
    df_other = df[~df['Notificacion'].str.contains('reboot')]

    # Concatenar los DataFrames si se decidió conservar los que no contienen 'reboot'
    df = pd.concat([df_reboot, df_other], ignore_index=True)  # Reasignando a df
    df = df.sort_values(by='Fecha', ascending=False)
    df.drop_duplicates( keep = "first", inplace=True )
    df = df.sort_values(by='Fecha', ascending=False)
    df.reset_index(drop=True, inplace=True)
    tiempoH = 0
    indices = df[df['Notificacion'].str.contains('rebooted', na=False)].index.tolist()
    indices_para_eliminar = []
    for i in range(1, len(indices)):
        if indices[i] == indices[i-1] + 1:
            indices_para_eliminar.append(indices[i])

    # Elimina los índices consecutivos.
    df.drop(indices_para_eliminar, inplace=True)

    df.reset_index(drop=True, inplace=True)
    while len(df["Notificacion"]) != 0:
        indices = df[df['Notificacion'].str.contains('rebooted', na=False)].index.tolist()
        indices = indices[::-1]
        indices_eliminar = []
        if len(indices) > 0:
            print(df)
            for i in indices:
                notificacion_actual = df.loc[i, 'Notificacion']
                fecha_actual = df.loc[i, 'Fecha']
                if 'rebooted' in notificacion_actual:

                    fecha_down, fecha_up = None, None
                    if i == 0 and len(df["Notificacion"]) == 2:
                        fecha_down = "Reboot Duplicado"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i])
                    if i == 0 and len(df["Notificacion"]) == 1:
                        fecha_down = "Reboot Duplicado"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i])
                    elif i == len(df)-1 and "again" in df.loc[i-1,"Notificacion"]:
                        print("EQUIVOCADO1")
                        fecha_down = "Reboot"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i-1, i])
                    # Verificar si el siguiente registro es 'is responding again'
                    elif i == len(df)-1 and "stopped" in df.loc[i-1,"Notificacion"]:
                        fecha_down = "Reboot Duplicado"
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i])

                    elif i+1 < len(df) and 'again' in df.loc[i+1, 'Notificacion']:
                        print("EQUIVOCADO3")
                        if len(df["Fecha"]) >i+2:
                            fecha_down = df.loc[i+2, 'Fecha'] if i+2 < len(df) else None
                            indices_eliminar.extend([i, i+1, i+2])
                            fecha_up = df.loc[i+1, 'Fecha']
                        else:
                            fecha_up = df.loc[i+1, 'Fecha']
                            fecha_down ="No aparecio"
                            indices_eliminar.extend([i, i+1 ])
                    # Verificar si el registro anterior es 'has stopped'
                    elif i+1 <= len(df["Notificacion"]) and 'stopped' in df.loc[i+1, 'Notificacion']:
                        fecha_down = df.loc[i+1, 'Fecha']
                        fecha_up = fecha_actual
                        indices_eliminar.extend([i-1, i, i+1])


                    tiempoR = 0
                    try:
                        tiempo = abs(( fecha_up - fecha_down ).total_seconds() / 60)  # Cambio aquí, se resta tiempo actual - tiempo pasado
                        tiempoR += tiempo  # Asumiendo que tiempoH está definido en alguna parte del código
                        print(tiempoR)
                    except Exception as e:
                        print(f"Error en el cálculo del tiempo: {e}")
                        tiempoR = "revisar"
                        tiempo = 0  # Asumir tiempo como 0 si hay un error
                                        # Añadir al resultado
                    if fecha_down is not None and fecha_up is not None:
                        new_row = {
                            'Enlace': dispositivo,  # Asegúrate de que 'dispositivo' está definido
                            'Fecha Down': fecha_down,
                            'Fecha Up': fecha_up,
                            'Tiempo': tiempo,
                            'Suma': tiempoR,  # Calcula el tiempo según lo necesario
                            'Estado': 'reboot'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
            try:
            # Eliminar las filas al final para evitar problemas durante la iteración
                indices_eliminar = list(set(indices_eliminar))  # Eliminar duplicados si los hay
                df.drop(index=indices_eliminar, inplace=True)
                df.reset_index(drop=True, inplace=True)
            except:
                df.drop(df.index[-1],inplace=True)

        if len(df["Notificacion"]) > 2:
            notifyUlt = df.iloc[-2]['Notificacion']      
            notifyPen = df.iloc[-1]['Notificacion']            
            tiempoUlt = df.iloc[-2]['Fecha'] 
            tiempoPen = df.iloc[-1]['Fecha']
            if "again" in notifyUlt:
                if "stopped" in notifyPen:
                    fecha_up = tiempoUlt
                    fecha_actual = tiempoPen
                    print(df)
                    # Ya no es necesario comprobar si son Timestamp porque ya se convirtieron
                    tiempo = abs(( fecha_up - fecha_actual ).total_seconds() / 60)  # Cambio aquí, se resta tiempo actual - tiempo pasado
                    tiempoH += tiempo  # Asumiendo que tiempoH está definido en alguna parte del código
                    print(tiempoH)
                    if tiempoH >= 5:
                        new_row = {
                            'Enlace': dispositivo,  # Asegúrate de que dispositivo está definido
                            'Fecha Down': fecha_actual,
                            'Fecha Up': fecha_up,
                            'Tiempo' : tiempo,
                            'Suma': tiempoH,
                            'Estado': 'Caido y Recuperado'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                        # Calcular los índices de las últimas dos filas
                        ultimos_indices = df.index[-2:]

                        # Usar estos índices para eliminar las filas
                        df.drop(index=ultimos_indices, inplace=True)

                        df.reset_index(drop=True, inplace=True)
                        print(df)
                    else:
                        ultimos_indices = df.index[-2:]

                        # Usar estos índices para eliminar las filas
                        df.drop(index=ultimos_indices, inplace=True)

                        df.reset_index(drop=True, inplace=True)
                        print(df)
                elif "again" in notifyPen:
                    ultimos_indices = df.index[-1:]

                    # Usar estos índices para eliminar las filas
                    df.drop(index=ultimos_indices, inplace=True)

                    df.reset_index(drop=True, inplace=True)
                    print(df)

            elif "stopped" in notifyUlt:
                ultimos_indices = df.index[-1:]

                # Usar estos índices para eliminar las filas
                df.drop(index=ultimos_indices, inplace=True)

                df.reset_index(drop=True, inplace=True)
                print(df)
        elif len(df["Notificacion"])== 2:
            notifyUlt = df.iloc[-2]['Notificacion']      
            notifyPen = df.iloc[-1]['Notificacion']            
            tiempoUlt = df.iloc[-2]['Fecha'] 
            tiempoPen = df.iloc[-1]['Fecha']
            if "again" in notifyUlt:
                if "again" in notifyPen:
                    ultimos_indices = df.index[-2:]
                    # Usar estos índices para eliminar las filas
                    df.drop(index=ultimos_indices, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                elif "stopped" in notifyPen:
                    fecha_up = tiempoUlt
                    fecha_actual = tiempoPen
                    print(df)
                    # Ya no es necesario comprobar si son Timestamp porque ya se convirtieron
                    tiempo = abs(( fecha_up - fecha_actual ).total_seconds() / 60)  # Cambio aquí, se resta tiempo actual - tiempo pasado
                    tiempoH += tiempo  # Asumiendo que tiempoH está definido en alguna parte del código
                    print(tiempoH)
                    if tiempoH > 5:
                        new_row = {
                            'Enlace': dispositivo,  # Asegúrate de que dispositivo está definido
                            'Fecha Down': fecha_actual,
                            'Fecha Up': fecha_up,
                            'Tiempo': tiempo,
                            'Suma': tiempoH,
                            'Estado': 'Caido y Recuperado'
                        }
                        print(new_row)
                        df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                        # Calcular los índices de las últimas dos filas
                        ultimos_indices = df.index[-2:]

                        # Usar estos índices para eliminar las filas
                        df.drop(index=ultimos_indices, inplace=True)

                        df.reset_index(drop=True, inplace=True)
                        print(df)
                    else:
                        # Calcular los índices de las últimas dos filas
                        ultimos_indices = df.index[-2:]
                        # Usar estos índices para eliminar las filas
                        df.drop(index=ultimos_indices, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        print(df)

            elif "stopped" in notifyPen:
                fecha_up = tiempoPen
                fecha_actual = tiempoUlt
                print(df)
                new_row = {
                            'Enlace': dispositivo,
                            'Fecha Down': fecha_actual,
                            'Fecha Up': None,
                            'Tiempo': None,
                            'Suma':None,
                            'Estado': 'Caido'
                        }
                print(new_row)
                df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                ultimos_indices = df.index[-2:]
                # Usar estos índices para eliminar las filas
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)

            elif "responding" in notifyPen:
                ultimos_indices = df.index[-2:]
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)


            elif "rebooted" in notifyUlt:
                print(df)
                fecha_actual = tiempoUlt
                new_row = {
                            'Enlace': dispositivo,
                            'Fecha Down': fecha_actual,
                            'Fecha Up': None,
                            'Tiempo': None,                            
                            'Suma':None,
                            'Estado': 'Caido'
                        }
                print(new_row)
                df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
                ultimos_indices = df.index[-2:]
                # Usar estos índices para eliminar las filas
                df.drop(index=ultimos_indices, inplace=True)
                df.reset_index(drop=True, inplace=True)

        elif len(df["Notificacion"]) == 1:
            notifyUlt = df.iloc[-1]['Notificacion']  
            print(notifyUlt)
            tiempoUlt = df.iloc[-1]['Fecha'] 

            print(df)
            if "stopped" in notifyUlt:
                new_row = {
                                'Enlace': dispositivo,
                                'Fecha Down': tiempoUlt,
                                'Fecha Up':None,
                                'Tiempo': None,                            
                                'Suma':None,
                                'Estado': "Caido"

                            }
            elif "again" in notifyUlt:
                new_row = {
                                'Enlace': dispositivo,
                                'Fecha Down': None,
                                'Fecha Up':tiempoUlt,
                                'Tiempo': None,                            
                                'Suma':None,
                                'Estado': "Revisar "

                            }

            print(new_row)
            df_resultados = pd.concat([df_resultados, pd.DataFrame([new_row])], ignore_index=True)
            ultimos_indices = df.index[-1:]

            # Usar estos índices para eliminar las filas
            df.drop(index=ultimos_indices, inplace=True)
            df.reset_index(drop=True, inplace=True)

    return df_resultados

def process_data_and_capture_screenshot(final, browserOrion):
    # With the dataframe we search the page for the events of the affected device within the schedule
    nodo = browserOrion.find_element(By.CSS_SELECTOR,"#ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects")
    nodo.click()

    #We obtain the information from the page where we look for each affected device
    soup = BeautifulSoup(browserOrion.page_source, 'html.parser')
    nodos = soup.find_all("option")
    final = final[final['Tiempo'] != 'reboot']

    txtDateInHora = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodBegin_txtTimePicker")
    txtDateInHora.clear()

    txtDateOnHora = browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_pickerTimePeriodControl_dtPeriodEnd_txtTimePicker")
    txtDateOnHora.clear()

    for x in nodos:
        if str(x.text) in final["Enlace"].values:
            nombre= str(x.text)
            try:
                partes = nombre.split(" ")
                indice = partes.index("P:")
                resultado = " ".join(partes[:indice+1])
                resultado = resultado.replace("/","").replace("\\","").replace(":","")
            except:
                print("No se pudo")
                resultado = nombre.replace("/","").replace("\\","").replace(":","")   
            enlace = x.attrs
            print("Encontrado  ", enlace, "  ", enlace["value"])
            wait = WebDriverWait(browserOrion, 10)
            select_element = Select(browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects"))
            select_element.select_by_value(str(enlace['value']))
            wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")))
            btnRefresh = browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
            btnRefresh.click()
            time.sleep(1)
            table_element = browserOrion.find_element(By.CLASS_NAME,"sw-pg-events")
            screenshot = browserOrion.get_screenshot_as_png()
            # Make a screenshot 
            location = table_element.location
            size = table_element.size
            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']
            table_screenshot = Image.open(BytesIO(screenshot)).crop((left, top, right*0.7, bottom))
            imgName = f'{resultado.strip()}.png'
            # Guarda la imagen recortada en  un archivo
            table_screenshot.save(imgName)
            img = Image.open(imgName)

            # Convertir la imagen a RGB (el formato JPG no soporta transparencia)
            img = img.convert('RGB')
            os.remove(imgName)
            nombre= f'{imgName.split(".")[0]}.jpg'
            # Guardar la imagen en formato JPG
            img.save(nombre)
            time.sleep(3)

def process_data_and_capture_screenshot_email(final, browserOrion):
    cont = 0 
    print(final.head(5))
    #We obtain the information from the page where we look for each affected device
    soup = BeautifulSoup(browserOrion.page_source, 'html.parser')
    nodos = soup.find_all("option")
    final_1 = final[final['Tiempo'] == 'reboot']
    final = final[final['Tiempo'] != 'reboot']
    final.reset_index(drop=True, inplace=True)
    # With the dataframe we search the page for the events of the affected device within the schedule
    nodo = browserOrion.find_element(By.CSS_SELECTOR,"#ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects")
    nodo.click()
    #We obtain the information from the page where we look for each affected device
    soup = BeautifulSoup(browserOrion.page_source, 'html.parser')
    nodos = soup.find_all("option")
    final = final[final['Tiempo'] != 'reboot']
    for x in nodos:
        if str(x.text) in final["Enlace"].values:
            nombre= str(x.text)
            print(nombre)

            try:
                partes = nombre.split(" ")
                indice = partes.index("P:")
                resultado = " ".join(partes[:indice+1])
                resultado = resultado.replace("/","").replace("\\","").replace(":","")
            except:
                print("No se pudo")
                resultado = nombre.replace("/","").replace("\\","").replace(":","")
                
            enlace = x.attrs
            print("Encontrado  ", enlace, "  ", enlace["value"])
            wait = WebDriverWait(browserOrion, 10)
            select_element = Select(browserOrion.find_element(By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_netObjects_netObjects"))
            select_element.select_by_value(str(enlace['value']))
            wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")))
            
            btnRefresh = browserOrion.find_element(By.ID,"ctl00_ctl00_BodyContent_ContentPlaceHolder1_RefreshButton")
            btnRefresh.click()
            time.sleep(1)
            table_element = browserOrion.find_element(By.CLASS_NAME,"sw-pg-events")
            screenshot = browserOrion.get_screenshot_as_png()
            # Make a screenshot 
            location = table_element.location
            size = table_element.size
            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']
            table_screenshot = Image.open(BytesIO(screenshot)).crop((left, top, right*0.7, bottom))
            imgName = f'{resultado.strip()}.png'
            # Guarda la imagen recortada en  un archivo
            table_screenshot.save(imgName)
            img = Image.open(imgName)
            # Convertir la imagen a RGB (el formato JPG no soporta transparencia)
            img = img.convert('RGB')
            os.remove(imgName)
            nombre= f'{imgName.split(".")[0]}.jpg'
            # Guardar la imagen en formato JPG
            img.save(nombre)
            time.sleep(5)
            print(cont)
            print(final["Enlace"][cont])
  
            df_final = reporteria.send_to_email_consola(final,cont)
            print(df_final)
            df_final = pd.concat([df_final, final_1])  # Asegúrate de que esta línea se ejecute correctamente
            cont += 1
    print("Salio de las imagenes")
    return df_final

def crear_mensaje(correo_destino, fecha_actual, final):
    message = MIMEMultipart()
    message['From'] = "mcaalvar@pacifico.fin.ec"
    message['To'] = correo_destino.strip()
    message['Subject'] = f'INCIDENCIA DE ENLACE {fecha_actual} '
    #mensaje = f"Por favor revisar incidencia de este enlace \n \n {final['Enlace'].values} \n \n Capture adjunta en el correo"
    #message.attach(MIMEText(mensaje, 'plain'))
    # Seleccionar las columnas 'Enlace' y 'Tiempo'
    final = final[["Enlace", "Tiempo"]]
    html = final.to_html()
    cuerpo = MIMEText(html, 'html')
    message.attach(cuerpo)
    return message

def adjuntar_imagen(message, enlace):
    archivoIMG = enlace
    try:
        with open(archivoIMG, 'rb') as f:
            img_data = f.read()
            img = MIMEImage(img_data, name=archivoIMG)
            message.attach(img)
    except:
        print("No se adjunto la imagen", archivoIMG)

def adjuntar_excel(message, archivo):
    try:
        with open(archivo, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='octet-stream')
            attachment.add_header('Content-Disposition', 'attachment', filename=archivo)
        message.attach(attachment)
    except:
        print("No se adjunto el archivo", archivo)


def enviar_correos_divididos(correos, fecha_actual, df_final, lista_enlaces, max_imgs_por_correo,archivo):
    # Dividir la lista de enlaces en subgrupos
    subgrupos_enlaces = [lista_enlaces[i:i + max_imgs_por_correo] for i in range(0, len(lista_enlaces), max_imgs_por_correo)]

    for correo in correos:
        # Para cada subgrupo de enlaces, crear y enviar un correo
        cont = 0
        for subgrupo in subgrupos_enlaces:
            mensaje = crear_mensaje(correo, fecha_actual, df_final,cont)
            adjuntar_todas_las_imagenes(mensaje, subgrupo)
            adjuntar_excel(mensaje, f"C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\{archivo}")
            enviar_correo(mensaje)
            cont +=1

# Tus otras funciones permanecen iguales...

 # Suponiendo que max_imgs_por_correo es 10
def enviar_correo_con_excel(correos, fecha_actual, df_final, archivo):
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
        
        # Adjuntar el archivo Excel
        adjuntar_excel(mensaje, archivo)
        
        # Enviar el correo
        enviar_correo(mensaje)

# Ejemplo de uso:
# correos = ["email1@example.com", "email2@example.com"]
# fecha_actual = pd.Timestamp.now()
# df_final = pd.DataFrame()  # Asumiendo que tienes un DataFrame ya preparado
# archivo = "ruta_al_archivo.xlsx"
# enviar_correo_con_excel(correos, fecha_actual, df_final, archivo)




def adjuntar_todas_las_imagenes(mensaje, lista_enlaces):
    for enlace in lista_enlaces:
        nombre_imagen = f"{str(enlace).strip()}.jpg"  # Asume que el nombre del archivo es el enlace + ".png"
        nombre_imagen = nombre_imagen.replace(":","")  # Limpiar el nombre del archivo
        adjuntar_imagen(mensaje, nombre_imagen)  # Asume que adjuntar_imagen toma el mensaje y el nombre del archivo como argumentos

def enviar_correo(message):
    smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
    smtp_connection.starttls()
    smtp_connection.login(smtp_user, smtp_password)
    smtp_connection.sendmail(message['From'], message['To'], message.as_string())
    smtp_connection.quit()

def delete_old_reports():
    for archivo in os.listdir(ROOT_DIR):
        if archivo.lower().endswith('.jpg'):
            try:
                os.remove(os.path.join(ROOT_DIR, archivo))
            except Exception as e:
                print(f"Error al eliminar {archivo}. Razón: {e}")

def send_reports_via_telegram_personal(chat_id):
    tb = telebot.TeleBot(TOKEN)
    fecha_actual = datetime.datetime.now()
    text = f"Reporte Orion {str(fecha_actual)}"
    tb.send_message(chat_id, text)
    try:
        archivo = "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\Reporte.png"
        with open(os.path.join(ROOT_DIR, archivo), 'rb') as photo:
            tb.send_photo(chat_id, photo)
            tb.send_message(chat_id, archivo.split(".")[0])
            time.sleep(2)
    except:
        tb.send_message(chat_id, "No se encontro el archivo "+archivo.split(".")[0])

def clear_all_images():
    for filename in os.listdir():
        if filename.endswith(".jpg"):
            try:
                os.remove(filename)
                print(f"Eliminada: {filename}")
            except Exception as e:
                print(f"Error al eliminar {filename}. Razón: {e}")

def eventosCaidos(chat_id):
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,"https://10.1.231.243/Orion/SummaryView.aspx?ViewID=1")

    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", "Silimarina97.")
    browserOrion = browser
    table_element = browserOrion.find_element(By.ID,"Resource1024_ctl00_ctl01_ResourceWrapper_resContent")
    screenshot = browserOrion.get_screenshot_as_png()

    # Make a screenshot 
    location = table_element.location
    size = table_element.size
    left = location['x']
    top = location['y']
    right = location['x'] + size['width']
    bottom = location['y'] + size['height']
    table_screenshot = Image.open(BytesIO(screenshot)).crop((left, top, right, bottom*0.6))
    archivo = "EventosActuales.png"
    # Guarda la imagen recortada en un archivo
    table_screenshot.save(archivo)
    tb = telebot.TeleBot(TOKEN)
    fecha_actual = datetime.datetime.now()
    text = f"Reporte Caidos {str(fecha_actual)}"
    tb.send_message(chat_id, text)

    with open(ROOT_DIR+archivo, 'rb') as photo:
        tb.send_photo(chat_id, photo)
    clear_all_images()

def analisis_reporte(df_1):
    df_1.info()
    df_1= df_1.drop_duplicates()
    cont = 0    
    df_reboot = df_1[df_1["Tiempo" ] == "reboot"].reset_index()

    df_1_validar = df_1.copy()
    df_1["Reportar"] = "Verificar" 
    df_1.dropna(subset=['Enlace'],inplace = True)
    df_1.dropna(subset=['Carpeta_Orion'], inplace=True)
    print("analisis_reporte: df1")
    print(df_1.head())
    for x in df_1.index:
        nombreL = str(df_1["Enlace"][x]).replace("Backup ","").split(" ")[:-1]
        enlace = ' '.join(nombreL)
        print(enlace)
        if "reboot" not in str(df_1["Tiempo"][x]) or "Caida" in str(df_1["Tiempo"][x]):
            if "Na" in str(df_1["Reportado"][x]):
                if "Caido" in str(df_1["Tiempo"][x]):
                    if "Cajeros B" in str(df_1["Carpeta_Orion"][x]) or "Ciudad " in str(df_1["Carpeta_Orion"][x]):
                        print("Verificar Backup o Principal")
                        df_1_validar.drop(x,inplace=True)
                        if enlace in df_1_validar["Enlace"]:
                            cont += 1
                            print(cont)
                            print("Reportar a Consola")
                            df_1.loc[x, 'Reportar'] = 'Consola'
                        else:
                            df_1.loc[x, 'Reportar'] = 'Proveedor'
                            print("Reportar a Proveedor")
                    else:
                        df_1.loc[x, 'Reportar'] = 'Consola'
                        print("Reportar Consola")
                else:
                    #revisar si tiene el enlace en reboot
                    if enlace in df_reboot["Enlace"]:
                        df_1.loc[x, 'Reportar'] = 'Reboot'
                    else:
                        df_1.loc[x, 'Reportar'] = 'Proveedor'
                    print("Reportar a proveedor")
            else:
                print("Esta reportado")
        else:
            df_1.loc[x, 'Reportar'] = 'Reboot'
            print("No se reporta")  
    return df_1     

def concatBase(df_Play,df_Base):
    df_Play.reset_index(drop=True, inplace=True)
    df_Base.reset_index(drop=True, inplace=True)
    df_con = pd.concat([df_Play,df_Base]).reset_index(drop=True)
    df_con.drop_duplicates(subset=["Enlace", "Fecha Down"],keep = "last", inplace=True)
    df_con.reset_index(drop=True, inplace=True)
    return df_con

def analisis_reporte_final(df_1):
    df_1= df_1.drop_duplicates()
    df_1["Reportar"]= "Verificar"
    df_1[['Enlace', 'Proveedor']] = df_1['Enlace'].str.rsplit(n=1, expand=True)
    df_1['Enlace'] = df_1['Enlace'].str.replace(" Backup","")
    df_1.reset_index(drop=True, inplace=True)
    df_1.dropna(subset=['Enlace'],inplace = True)
    df_1.dropna(subset=['Carpeta_Orion'], inplace=True)
    df_1['Tiempo'] = df_1['Tiempo'].astype(str)
    reboots = df_1[df_1['Tiempo']== "reboot"]
    df_1.reset_index(drop=True, inplace=True)

    # Procesar cada enlace
    for nombre in df_1['Enlace'].unique():
        enlaces_mismo_nombre = df_1[df_1['Enlace'] == nombre]
        indices = enlaces_mismo_nombre.index
        
        # Verificar condiciones
        for i in indices:
            # Si el enlace ya está marcado como 'Omitir', continuar al siguiente
            if df_1.loc[i, 'Reportar'] == 'Omitir':
                continue
            # Si hay 'reboot' en los Tiempo de los enlaces con el mismo nombre
            if 'reboot' in enlaces_mismo_nombre['Tiempo'].values:
                # Y además hay un número en los Tiempo, marcar todos como 'Omitir'
                if enlaces_mismo_nombre['Tiempo'].apply(lambda x: x.replace('.', '', 1).isdigit()).any():
                    df_1.loc[indices, 'Reportar'] = 'Omitir'
                    break  # No se necesitan más comprobaciones para estos enlaces
            # Aplicar otras condiciones solo si 'Reportar' no está marcado como 'Omitir'
            caidos = enlaces_mismo_nombre['Tiempo'].str.contains('Caido')
            if caidos.sum() > 1:  # Si hay más de un enlace caído
                df_1.at[i, 'Reportar'] = 'Consola'
                # Marcar todos menos uno de los enlaces caídos como 'Omitir'
                indices_para_omitir = caidos[caidos].index.difference([i])
                df_1.loc[indices_para_omitir, 'Reportar'] = 'Omitir'
            elif 'Ciudad' in df_1.at[i, 'Carpeta_Orion'] and caidos.sum() == 0:
                df_1.at[i, 'Reportar'] = 'Proveedor'
            
            elif 'Caido' in df_1.at[i, 'Tiempo']:
                df_1.at[i, 'Reportar'] = 'Consola'
            
            elif df_1.at[i, 'Tiempo'].replace('.', '', 1).isdigit():
                df_1.at[i, 'Reportar'] = 'Proveedor'

    # Aplicar las condiciones adicionales
    df_1.loc[df_1['Tiempo'].str.contains('Caido'), 'Reportar'] = 'Consola'
    df_1.loc[df_1['Tiempo'].apply(lambda x: x.replace('.','',1).isdigit()), 'Reportar'] = 'Proveedor' 
    df_1.reset_index(drop=True, inplace=True) 
    for nombre_enlace in reboots['Enlace'].unique():
        indices_con_mismo_nombre = df_1[df_1['Enlace'] == nombre_enlace].index
        df_1.loc[indices_con_mismo_nombre, 'Reportar'] = 'Omitir'
    df_1.reset_index(drop=True, inplace=True)
    return df_1     

def clean_column(df, column_name):
    
    # Caracteres no permitidos en sistemas Windows y sistemas basados en Unix
    forbidden_characters = r'\/:*?"<>|'
    
    # Usamos el método `str.replace` de pandas para eliminar cada carácter no permitido
    for char in forbidden_characters:
        df[column_name] = df[column_name].str.replace(char, '', regex=False)
        df[column_name] = df[column_name].str.strip()
    df.reset_index(drop=True, inplace=True)
    return df

def generar_horarios(fecha_inicio_str, fecha_fin_str):
    # Convertir las cadenas de texto a objetos datetime
    fecha_inicio = datetime.datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
    fecha_fin = datetime.datetime.strptime(fecha_fin_str, "%d/%m/%Y")
    horarios = []
    delta = datetime.timedelta(hours=4)
    fecha_actual = fecha_inicio

    while fecha_actual < fecha_fin:
        hora_inicio = fecha_actual
        while hora_inicio < fecha_actual + datetime.timedelta(days=1) and hora_inicio < fecha_fin:
            hora_fin = hora_inicio + delta
            if hora_fin > fecha_fin:
                hora_fin = fecha_fin
            horarios.append((hora_inicio, hora_fin))
            hora_inicio = hora_fin
        fecha_actual += datetime.timedelta(days=1)
    return horarios


def recolectar_datos_por_segmentos_mes(browser, horarios):
    dfs = []
    print(horarios)
    for inicio, fin in horarios:
        # Convertir fechas y horas a cadenas de texto
        fecha_inicio_str = inicio.strftime('%d/%m/%Y')
        hora_inicio_str = inicio.strftime('%H:%M')
        fecha_fin_str = fin.strftime('%d/%m/%Y')
        hora_fin_str = fin.strftime('%H:%M')
        print(inicio)
        print(fin)
        selecc_horario(browser, fecha_inicio_str, hora_inicio_str, fecha_fin_str, hora_fin_str)
        df_segmento = extraer_info_pagina(browser)
        dfs.append(df_segmento)

    df_final = pd.concat(dfs, ignore_index=True)
    df_final.drop_duplicates(keep="first",inplace=True)
    return df_final



def main(chat_id):
    # Eliminar archivos antiguos
    eliminar_archivos_antiguos(ruta=ROOT_DIR)
    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)
    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", smtp_password)
    # Seleccionar el periodo de tiempo
    fecha_inicio, hora_inicio, fecha_fin, hora_fin = periodo_hasta_ahora()  # o pedir_periodo()
    selecc_horario(browser, fecha_inicio, hora_inicio, fecha_fin, hora_fin)
    # Extraer información de la página
    df = extraer_info_pagina(browser)
    # Limpiar y procesar el dataframe
    df = limpiar_dataframe(df)
    browser.close()
    dispositivos_unicos = list(set(extraer_dispositivos(df)))
    df_final = pd.DataFrame()
    for dispositivo in dispositivos_unicos:
        df_dispo= procesar_notificaciones_por_dispositivo(df, dispositivo)
        df_final = pd.concat([df_final,df_dispo])
    
    df_final.info()
    time.sleep(2)
    df_final = clean_column(df_final, "Enlace")
    dataframe_toimag(df_final,filename="Reporte.png")
    archivo = "Reporte.xlsx"
    df_final=corregir_estados_reboot(df_final)
    df_final.to_excel(archivo,index=False)
  

    enlaces = list(set(list(df_final["Enlace"])[1:]))
    df_final.reset_index(drop=True,inplace=True)
    
    #process_data_and_capture_screenshot(df_final, browser)
    # Enviar el correo
    send_reports_via_telegram_personal(chat_id)
    #enviar_correos_divididos(correos, fecha_inicio, df_final, enlaces, 120,archivo) 
    enviar_correo_con_excel(correos, fecha_inicio, df_final,archivo)
    clear_all_images()

def main_personal(chat_id):
    # Eliminar archivos antiguos
    eliminar_archivos_antiguos(ruta=ROOT_DIR)

    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)

    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", smtp_password)

    # Seleccionar el periodo de tiempo
    fecha_inicio, hora_inicio, fecha_fin, hora_fin = periodo_hasta_ahora()  # o pedir_periodo()
    selecc_horario(browser, fecha_inicio, hora_inicio, fecha_fin, hora_fin)

    # Extraer información de la página
    df = extraer_info_pagina(browser)
    df.to_excel("historicofecha.xlsx")
    # Limpiar y procesar el dataframe
    df = limpiar_dataframe(df)
    browser.close()
    df.reset_index(drop=True, inplace=True)
    df.to_excel("historicofecha.xlsx")
    dispositivos_unicos = list(set(extraer_dispositivos(df)))
    
    
    df_final = pd.DataFrame()

    for dispositivo in dispositivos_unicos:
        df_dispo= procesar_notificaciones_por_dispositivo(df, dispositivo)
        df_final = pd.concat([df_final,df_dispo]).reset_index(drop=True)
        df_final.reset_index(drop=True, inplace=True)
    
    df_final.info()
    time.sleep(2)
    archivo = "Reporte.xlsx"
    df_final.info()
    df_final = clean_column(df_final, "Enlace")
    df_final=corregir_estados_reboot(df_final)
    fecha_actual = datetime.date.today().strftime('%d/%m/%Y')
    daily = f'ReporteDiario{str(fecha_actual).replace("/","-")}.xlsx'
    df_final.reset_index(drop=True, inplace=True)
    try:
        dailydf = pd.read_excel(daily)
    except:
        dailydf = pd.DataFrame()

    df_final = pd.merge(df_final,archivoG,how="left", left_on="Enlace", right_on="Nombre_Orion")
    df_final.to_excel(archivo,index=False)
    df_final.reset_index(drop=True, inplace=True)
    df_base = concatBase(df_final,dailydf)
    df_base.info()      
    df_base = analisis_reporte_final(df_base)
    df_base["Reportar"].replace('Na', np.nan, inplace=True)
    df_base.reset_index(drop=True, inplace=True)
    df_base.to_excel(daily)
    dataframe_toimag(df_base,filename="Reporte.png")
    send_reports_via_telegram_personal(chat_id)
    fecha = datetime.datetime.today()
    enviar_correo_con_excel(correos, fecha, df_final, archivo)
    #enviar_correos_divididos(correos, fecha_inicio, df_final, enlaces, 105,archivo) 
    clear_all_images()

def main_personal_calendario_dia(chat_id,fecha):
    # Eliminar archivos antiguos
    eliminar_archivos_antiguos(ruta=ROOT_DIR)

    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)

    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", smtp_password)
    
    df= recolectar_datos_por_segmentos(browser,fecha)
    # Convertir la columna 'Fecha' a datetime
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')

    # Ordenar el DataFrame por 'Fecha' de más reciente a más antiguo
    df.sort_values(by='Fecha', ascending=True, inplace=True)

    # Resetear el índice
    df.reset_index(drop=True, inplace=True)
    df.to_excel("historicofechaManana.xlsx")
    # Limpiar y procesar el dataframe
    df = limpiar_dataframe(df)
    browser.close()
    df.to_excel("historicofechaMananaLimpio.xlsx")
    dispositivos_unicos = list(set(extraer_dispositivos(df)))
    
    
    df_final = pd.DataFrame()

    for dispositivo in dispositivos_unicos:
        df_dispo= procesar_notificaciones_por_dispositivo(df, dispositivo)
        df_final = pd.concat([df_final,df_dispo])
    
    df_final.info()
    time.sleep(2)
    df_final = clean_column(df_final, "Enlace")
    df_final=corregir_estados_reboot(df_final)
    archivo = f'Reporte{str(fecha).replace("/","-")}.xlsx'
    df_final.to_excel(archivo,index=False)
    df_final.reset_index(drop=True, inplace=True)
    dataframe_toimag(df_final,filename="Reporte.png")
    send_reports_via_telegram_personal(chat_id)
    #enlaces = list(set(list(df_final["Enlace"])[1:]))
    enviar_correo_con_excel(correos, fecha, df_final, archivo)
    #enviar_correos_divididos(correos, fecha, df_final, enlaces, 105,archivo) 
    clear_all_images()

def main_personal_calendario_noche(chat_id,fecha):
    # Eliminar archivos antiguos
    eliminar_archivos_antiguos(ruta=ROOT_DIR)

    # Configurar y lanzar el navegador
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)

    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", smtp_password)

    df= recolectar_datos_nocturnos(browser,fecha)
    # Limpiar y procesar el dataframe
    df = limpiar_dataframe(df)
    browser.close()
    # Convertir la columna 'Fecha' a datetime
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')

    # Ordenar el DataFrame por 'Fecha' de más reciente a más antiguo
    df.sort_values(by='Fecha', ascending=False, inplace=True)

    # Resetear el índice
    df.reset_index(drop=True, inplace=True)
    df.to_excel("historicofechaNochelimpio.xlsx")
    dispositivos_unicos = list(set(extraer_dispositivos(df)))
    
    
    df_final = pd.DataFrame()

    for dispositivo in dispositivos_unicos:
        df_dispo= procesar_notificaciones_por_dispositivo(df, dispositivo)
        df_final = pd.concat([df_final,df_dispo])
    
    df_final.info()
    time.sleep(2)
    df_final = clean_column(df_final, "Enlace")
    df_final=corregir_estados_reboot(df_final)
    df_final.reset_index(drop=True, inplace=True)
    archivo = f'ReporteMadrugada{str(fecha).replace("/","-")}.xlsx'
    df_final.to_excel(archivo,index=False)
    dataframe_toimag(df_final,filename="Reporte.png")
    send_reports_via_telegram_personal(chat_id)

    #enlaces = list(set(list(df_final["Enlace"])[1:]))
    enviar_correo_con_excel(correos, fecha, df_final, archivo)
    # Ejemplo de uso:
    #enviar_correos_divididos(correos, fecha, df_final, enlaces, 20,archivo) 
    clear_all_images()

def main_mes(chat_id,fechaI,fechaF):
    chrome_options = configurar_chrome()
    browser = iniciar_navegador(chrome_options,url)
    horario = generar_horarios(fechaI,fechaF)
    # Iniciar sesión en el navegador
    login_navegador(browser, "mcaalvar@pacifico", smtp_password)

    df= recolectar_datos_por_segmentos_mes(browser,horario)
    browser.close()
    # Limpiar y procesar el dataframe
    df = limpiar_dataframe(df)
    # Convertir la columna 'Fecha' a datetime
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M')

    # Ordenar el DataFrame por 'Fecha' de más reciente a más antiguo
    df.sort_values(by='Fecha', ascending=False, inplace=True)
    df.drop_duplicates( keep = "first", inplace=True )
    # Resetear el índice
    df.reset_index(drop=True, inplace=True)
    df.to_excel("pruebasMesimp.xlsx")
    #df.to_excel("historicofechaMeslimpio.xlsx")
    dispositivos_unicos = list(set(extraer_dispositivos(df)))
    df_final = pd.DataFrame()

    for dispositivo in dispositivos_unicos:
        df_dispo= procesar_notificaciones_por_dispositivo(df, dispositivo)
        df_final = pd.concat([df_final,df_dispo])
    
    df_final.info()
    df_final.reset_index(drop=True, inplace=True)
    time.sleep(2)
    rag = str(f"{fechaI} - {fechaF}")
    df_final = clean_column(df_final, "Enlace")
    df_final=corregir_estados_reboot(df_final)
    archivo = f'Reporte{str(rag).replace("/","-")}.xlsx'
    df_final.reset_index(drop=True, inplace=True)
    df_final.to_excel(archivo,index=False)
    try:
        dataframe_toimag(df_final,filename="Reporte.png")
        dataframe_topdf(df_final, filename=f"Reporte{rag}.pdf")
    except:
        print("No hubo imagen o pdf")
    #dataframe_topdf(df_final, filename=f"Reporte{rag}.pdf")
    send_reports_via_telegram_personal(chat_id)

    #enlaces = list(df_final["Enlace"])[1:]
    enviar_correo_con_excel(correos, rag, df_final, archivo)
    # Ejemplo de uso:
    #enviar_correos_divididos(correos, rag, df_final, enlaces, 20,archivo) 
    clear_all_images()

def dataframe_toimag(df, filename="Reporte.png"):
    fig_width = min(20, 1 * len(df.columns))
    fig_height = max(4, 0.3 * len(df))
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis('off')
    tbl = table(ax, df, loc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)  # Reducir el tamaño de fuente
    plt.savefig(filename, bbox_inches='tight', dpi=150)  # Reducir los DPI
    plt.close(fig)

def dataframe_topdf(df, filename="Reporte.pdf"):

    with PdfPages(filename) as pdf:
        fig, ax = plt.subplots(figsize=(20, max(4, 0.3 * len(df))))
        ax.axis('off')
        tbl = table(ax, df, loc='center')
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(10)
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)




if __name__ == "__main__":
    main()




