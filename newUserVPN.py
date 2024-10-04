import pyautogui
import time
import subprocess
import telebot
import datetime 
import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")


def abrir_programa(ruta_programa):
    pyautogui.hotkey('win', 'r')  # Presiona la tecla Win + R para abrir 'Ejecutar'
    time.sleep(0.5)
    pyautogui.write(ruta_programa)  # Escribe la ruta del programa
    pyautogui.press('enter')  # Presiona 'Enter' para ejecutar

def ping_dominio(dominio):
    # Dependiendo de tu sistema operativo, el comando y sus argumentos pueden variar
    # Este ejemplo es para Windows:
    respuesta = subprocess.run(['ping', '-n', '4', dominio], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # En Windows, si el ping es exitoso, retorna 0. Otros sistemas operativos pueden variar.
    return respuesta.returncode == 0


def ingresar_credenciales(usuario, contrasena):
    time.sleep(2)  # Espera 2 segundos para que el programa se abra completamente
      # Cambia al campo de la contraseña
    pyautogui.write(contrasena)  # Escribe la contraseña
    pyautogui.press('enter')  # Presiona 'Enter' para ingresar
def notificacion():
    TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
    chat_id = 1397965849
    tb = telebot.TeleBot(TOKEN)	#create a new Telegram Bot object

    for x in range(5):
        tb.send_message(chat_id,"Activacion de VPN PILAAAS")      
    time.sleep(5)
    tb.close()
# otro_script.py

def metodo_autenticacion(username, password):
    # Aquí va tu lógica para autenticar o hacer algo con el usuario y la contraseña
    return f"Usuario: {username}, Contraseña: {password}"



if __name__ == "__main__":
    RUTA_PROGRAMA = "C:\\Program Files (x86)\\CheckPoint\\Endpoint Connect\\TrGUI.exe"
    USUARIO = "nechever"
    CONTRASENA = "Teleco102023"
    DOMINIO = "sbp0100tcm14.pacifico.bpgf"
    count = 0
    while True:
        if not ping_dominio(DOMINIO):
            print(f"No se pudo hacer ping a {DOMINIO}. Ejecutando el programa...")
            abrir_programa(RUTA_PROGRAMA)
            ingresar_credenciales(USUARIO, CONTRASENA)
            count += 1
            time.sleep(120)
            
        else:
            print(f"Ping a {DOMINIO} exitoso.")
            time.sleep(1200)
