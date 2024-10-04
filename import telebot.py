import telebot
import requests
TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'
bot = telebot.TeleBot(TOKEN)

# Deshabilitar verificación SSL temporalmente
requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

def send_message_ssl(chat_id, text):
    bot.send_message(chat_id, text, parse_mode='HTML', disable_web_page_preview=True)

# Enviar mensaje sin verificar SSL (opción temporal)
response = requests.post(
    'https://api.telegram.org/bot5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-sN/sendMessage',
    data={'chat_id': 1176200280, 'text': '¡Bienvenido! El bot se ha iniciado.'},
    verify=False
)