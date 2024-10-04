import telebot
from automatizacionVPN import autovpn 
from automatizacionVPN import autowhilevpn
import time
from automatizacionDash import reporteDash
import os
import datetime
import pyautogui
from OrionAuto import orion
from datetime import datetime
import OrionINT



TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
bot = telebot.TeleBot(TOKEN)
usuarios_en_autenticacion = {}
user = "mcaalvar"
passw = "Silimarina97."
# DOMINIO = "sbp0100tcm14.pacifico.bpgf" 
DOMINIO = "https://10.1.231.243/" 

5451                                                                    
RUTA_PROGRAMA = "C:\\Program Files (x86)\\CheckPoint\\Endpoint Connect\\TrGUI.exe"
CHAT_ID_ADMIN =  1397965849  # Asume que este es tu chat ID. Cambia según necesites.
CHAT_IDS = [ 1397965849,5781054062, 6055198218,6512803121,5686702637,6564107654,6383100715,1176200280]
BREAK_FLAG = False
DOMINIO_2 = "https://10.1.231.243/#/login"
DOMINIO_3 = "https://10.225.200.25/login?next=/"



user_states = {}

user_dates = {}


def list_image_details():
    details = []
    for filename in os.listdir():
        if filename.endswith(".png"):
            timestamp = os.path.getmtime(filename)
            modified_date = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
            details.append(f"Nombre: {filename}, Última Modificación: {modified_date}")
    return details

def execute_mejora_with_ping_check():
    while True:
        if autovpn.hacer_ping(DOMINIO):  # Si el ping es exitoso
            reporteDash.main()
        time.sleep(60)  # Verifica cada 60 segundos

def send_screenshot_to_chat(chat_id):
    filename = get_latest_screenshot()
    if not filename:
        bot.send_message(chat_id, "No se encontró una captura con el nombre 'verificar'.")
        return

    tb = telebot.TeleBot(TOKEN)
    with open(filename, 'rb') as photo:
            tb.send_photo(chat_id, photo)

def send_screenshot_to_chat_loggin(chat_id):
    filenames = get_latest_screenshot_loggin()
    if not filenames:
        bot.send_message(chat_id, "No se encontró una captura con el nombre 'verificar'.")
        return

    tb = telebot.TeleBot(TOKEN)
    for filename in filenames:
        with open(filename, 'rb') as photo:
            tb.send_photo(chat_id, photo)
    delete_verification_images_loggin()

def get_latest_screenshot():
    # Listar todos los archivos en el directorio actual que comienzan con 'verificar'
    screenshots = [f for f in os.listdir() if os.path.isfile(f) and f.startswith('verificar')]
    
    # Si no se encuentra ninguna captura, retornar None
    if not screenshots:
        return None
    # Ordenar los archivos por tiempo de modificación
    latest_screenshot = max(screenshots, key=os.path.getmtime)
    return latest_screenshot

def get_latest_screenshot_loggin():
    # Listar todos los archivos en el directorio actual que comienzan con 'verificar'
    screenshotsloggin = [f for f in os.listdir() if os.path.isfile(f) and f.startswith('verificando')]
    
    # Si no se encuentra ninguna captura, retornar None
    if not screenshotsloggin:
        return None
    if len(screenshotsloggin) > 0:
        return screenshotsloggin

def unlock_computer(password):
    # Simula la tecla Enter
    pyautogui.press('enter')
    time.sleep(1)  # Espera un segundo
    
    # Escribe la contraseña/patrón de números
    pyautogui.write(password)

    
    time.sleep(1)  # Espera un segundo

    # Simula la tecla Enter nuevamente para intentar desbloquear
    pyautogui.press('enter')

    # Toma una captura de pantalla
    screenshot_filename = "screenshot.png"
    pyautogui.screenshot(screenshot_filename)

    # Envía la captura al administrador
    with open(screenshot_filename, 'rb') as photo:
        bot.send_photo(CHAT_ID_ADMIN, photo)

    # Opcionalmente, elimina la captura de pantalla del sistema de archivos
    os.remove(screenshot_filename)

def notify_ping_status():
    while True:
        time.sleep(1800)  # Espera 30 minutos
        if not autovpn.hacer_ping(DOMINIO) or not autovpn.hacer_ping(DOMINIO_2) or not autovpn.hacer_ping(DOMINIO_3):
            if not autovpn.hacer_ping(DOMINIO):
                bot.send_message(CHAT_ID_ADMIN, f"No se pudo hacer ping a {DOMINIO}.")
            elif not autovpn.hacer_ping(DOMINIO_2):
                bot.send_message(CHAT_ID_ADMIN, f"No se pudo hacer ping a {DOMINIO_2}.")
            elif not autovpn.hacer_ping(DOMINIO_3):
                bot.send_message(CHAT_ID_ADMIN, f"No se pudo hacer ping a {DOMINIO_3}.")

def cambioUser(username, password):
    global user, passw
    user, passw = username, password

def defaultUser():
    global user, passw
    user, passw = "mcaalvar", "Silimarina97."

def delete_verification_images():
    for filename in os.listdir():
        if filename.startswith("verificar") and filename.endswith(".png"):
            try:
                os.remove(filename)
                print(f"Eliminada: {filename}")
            except Exception as e:
                print(e)
                print(f"Error al eliminar {filename}. Razón: {e}")

def delete_verification_images_loggin():
    for filename in os.listdir():
        if filename.startswith("verificando") and filename.endswith(".png"):
            try:
                os.remove(filename)
                print(f"Eliminada: {filename}")
            except Exception as e:
                print(f"Error al eliminar {filename}. Razón: {e}")

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, """¡Hola ! Estas son las opciones disponibles:

    1. '/Ping': Para probar la conexión a VPN.
    2. '/ActivacionVPN': Si necesitas autenticarte.
    3. '/Reporte': Para el reporte de monitoreo Dashboard.
    4. '/Caidos' : Enlaces Caidos, SOLO IMAGEN DE CAIDOS '
    5. '/break': Cancela operaciones en curso.
    6. '/getid': Obtener el ID del Bot.
    7. '/Orion': Notificaion de Orion, CORREO PARA TODOS Dia ANTERIOR 8PM a hora actual.
    8.'/ReporteOrionNocheAnt' : Tabla de Enlaces (Dia anterior 20:00 al momento actual) ENVIO DE CORREOS
    10. '/ReporteOrionDia' : Reporte de Orion de 8 a 20 del dia que escojan
    11. '/ReporteOrionMadrugada' : Reporte de Orion de 20:00 del anterio dia hasta 8 AM dia escojito
    12. '/ReporteXRangoFecha': Tickets generados por rango de fechas.             
                 

Por favor, asegúrate de usar el signo '/' antes de cada comando. ¡Espero poder ayudarte!""")

@bot.message_handler(commands=['ActivacionVPN'])
def pedir_usuario(message):
    usuarios_en_autenticacion[message.chat.id] = {"step": "username"}
    bot.reply_to(message, "Por favor, envía tu nombre de usuario:")

@bot.message_handler(func=lambda message: message.chat.id in usuarios_en_autenticacion)
def recibir_datos(message):
    chat_id = message.chat.id
    if usuarios_en_autenticacion[chat_id]["step"] == "username":
        usuarios_en_autenticacion[chat_id] = {
            "step": "password",
            "username": message.text
        }
        bot.reply_to(message, "Gracias. Ahora, envía tu contraseña:")
    elif usuarios_en_autenticacion[chat_id]["step"] == "password":
        username = usuarios_en_autenticacion[chat_id]["username"]
        password = message.text
        del usuarios_en_autenticacion[chat_id]
        cambioUser(username, password)
        if autovpn.activar_vpn(username, password, DOMINIO):
            send_screenshot_to_chat_loggin(chat_id)
            defaultUser()
            respuesta = "SI SE PUDO LOCO, UN GALANAZO"
        else:
            defaultUser()
            respuesta = "No se pudo, vuelva a intentar"
        bot.reply_to(message, respuesta)

@bot.message_handler(commands=['getid'])
def send_chat_id(message):
    try:
        chat_id = message.chat.id
        bot.reply_to(message, f"Tu ID de chat es: {chat_id}")
    except Exception as e:
            print(e)
            bot.reply_to(message, f"Error al intentar desbloquear: {str(e)}")

@bot.message_handler(commands=['Reporte'])
def handle_reporte_request(message):
    chat_id = message.chat.id
    try:
        bot.reply_to(message, "Generando Reporte")
        reporteDash.personal_report(chat_id)
        bot.reply_to(message, "Reporte generado y enviado con éxito.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Hubo un error al generar el reporte: {str(e)}")


@bot.message_handler(commands=['Orion'])
def handle_reporte_request(message):
    chat_id = message.chat.id
    try:
        bot.reply_to(message, "Generando Reporte Orion")
        orion.main(chat_id)
        bot.reply_to(message, "Reporte generado y enviado con éxito.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Hubo un error al generar el reporte: {str(e)}")


@bot.message_handler(commands=['Caidos'])
def handle_reporte_request(message):
    chat_id = message.chat.id
    try:
        bot.reply_to(message, "Generando Reporte Caidos Orion")
        orion.eventosCaidos(chat_id)
        bot.reply_to(message, "Reporte generado y enviado con éxito.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Hubo un error al generar el reporte: {str(e)}")


@bot.message_handler(commands=['ReporteXRangoFecha'])
def handle_reporte_rango_request(message):
    try:
        chat_id = message.chat.id
        user_states[chat_id] = "WAITING_FOR_START_DATE"
        bot.reply_to(message, "Por favor, ingrese la fecha de inicio en formato DD/MM/YYYY.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Error al intentar procesar: {str(e)}")


@bot.message_handler(func=lambda message: message.chat.id in user_states and user_states.get(message.chat.id) == "WAITING_FOR_START_DATE")
def handle_start_date(message):
    chat_id = message.chat.id
    user_message = message.text
    if user_states.get(chat_id) == "WAITING_FOR_START_DATE":
        try:
            fecha_inicio = user_message
            user_dates[chat_id] = {'start': fecha_inicio}
            user_states[chat_id] = "WAITING_FOR_END_DATE"
            bot.reply_to(message, f"Fecha de inicio recibida: {fecha_inicio}. Por favor, ingrese la fecha de fin.")
        except Exception as e:
            print(e)
            bot.reply_to(message, f"Error al procesar la fecha de inicio: {str(e)}")

@bot.message_handler(func=lambda message: message.chat.id in user_states and user_states.get(message.chat.id) == "WAITING_FOR_END_DATE")
def handle_end_date(message):
    chat_id = message.chat.id
    user_message = message.text
    if user_states.get(chat_id) == "WAITING_FOR_END_DATE":
            fecha_fin = user_message
            fecha_inicio = user_dates[chat_id]['start']
            bot.reply_to(message, f"Fechas recibidas: inicio {fecha_inicio}, fin {fecha_fin}. Generando Reporte.")
            print(fecha_inicio, fecha_fin)
            # Ejecuta tu función para el rango de fechas aquí
            orion.main_mes(chat_id, fecha_inicio, fecha_fin)
            bot.reply_to(message, "Reporte generado y enviado con éxito.")

@bot.message_handler(commands=['ReporteOrionNocheAnt'])
def handle_reporte_request(message):
    try:
        chat_id = message.chat.id
        bot.reply_to(message, "Generando Reporte Orion")
        orion.main_personal(chat_id)
        bot.reply_to(message, "Reporte generado y enviado con éxito.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Error al intentar procesar: {str(e)}")

@bot.message_handler(commands=['ReporteOrionDia'])
def handle_reporte_request(message):
    try:
        chat_id = message.chat.id
        user_states[chat_id] = "WAITING_FOR_DATE"
        bot.reply_to(message, "Por favor, ingrese la fecha del día en formato DD/MM/YYYY.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Error al intentar procesar: {str(e)}")

@bot.message_handler(func=lambda message: message.chat.id in user_states and user_states.get(message.chat.id) == "WAITING_FOR_DATE")
def handle_all_message(message):
    chat_id = message.chat.id
    user_message = message.text
    if user_states.get(chat_id) == "WAITING_FOR_DATE":
        try:
            fecha = user_message
            bot.reply_to(message, f"Fecha recibida: {fecha}. Generando Reporte Orion.")
            
            # Ejecuta tu función orion.main_personal_calendario aquí
            orion.main_personal_calendario_dia(chat_id, fecha)
            
            bot.reply_to(message, "Reporte generado y enviado con éxito.")
            
        except Exception as e:
            print(e)
            bot.reply_to(message, f"Error al procesar la fecha: {str(e)}")
        finally:
            # Reiniciar el estado del usuario
            del user_states[chat_id]

@bot.message_handler(commands=['ReporteOrionMadrugada'])
def handle_reporte_request(message):
    try:
        chat_id = message.chat.id
        user_states[chat_id] = "WAITING_FOR_DATE_NIGHT"
        bot.reply_to(message, "Por favor, ingrese la fecha del día en formato DD/MM/YYYY.")
    except Exception as e:
        print(e)
        bot.reply_to(message, f"Error al intentar procesar: {str(e)}")

@bot.message_handler(func=lambda message: message.chat.id in user_states and user_states.get(message.chat.id) == "WAITING_FOR_DATE_NIGHT")
def handle_all_message(message):
    chat_id = message.chat.id
    user_message = message.text
    if user_states.get(chat_id) == "WAITING_FOR_DATE_NIGHT":

            fecha = user_message
            bot.reply_to(message, f"Fecha recibida: {fecha}. Generando Reporte Orion.")
            
            # Ejecuta tu función orion.main_personal_calendario aquí
            orion.main_personal_calendario_noche(chat_id, fecha)
            
            bot.reply_to(message, "Reporte generado y enviado con éxito.")

@bot.message_handler(commands=['Ping'])
def handle_ping(message): 
    if autovpn.hacer_ping(DOMINIO):
        bot.reply_to(message, f"Ping a {DOMINIO} fue exitoso.")
    else:
        bot.reply_to(message, f"No se pudo hacer ping a {DOMINIO}.")

@bot.message_handler(func=lambda m: True)
def echo_all(message):
    bot.reply_to(message, "No reconozco ese comando. Envía la palabra")


@bot.message_handler(commands=['Mostrar'])
def handle_show(message):
    details = list_image_details()
    for detail in details:
        bot.reply_to(message, detail)

def mainbot():
    for chat_id in CHAT_IDS:
        try:
            bot.send_message(chat_id, "¡Bienvenido! El bot se ha iniciado.")
        except e:
            print(f"Error al enviar mensaje a {chat_id}: {e}")
            
    while True:
        try:
            bot.polling(none_stop=True)
        except telebot.apihelper.ApiException as e:
            print("Error de API:", e)
            time.sleep(5)  # Espera 5 segundos antes de reintentar
            break  # Rompe el bucle si hay un error inesperado

import threading
def restart_thread(thread):
    if not thread.is_alive():
        thread.start()

if __name__ == "__main__":

    thread_mejora3 = threading.Thread(target=OrionINT.main, daemon=True)
    restart_thread(thread_mejora3)

    thread_mejora = threading.Thread(target=execute_mejora_with_ping_check, daemon=True)
    restart_thread(thread_mejora)
    mainbot()
    

    

