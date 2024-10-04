import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Datos del correo
de = 'nechever@pacifico.fin.ec'
para = 'wdjara@pacifico.fin.ec'
asunto = 'Incidente de enlace'

# Crea un correo multipart
msg = MIMEMultipart()
msg['From'] = de
msg['To'] = para
msg['Subject'] = asunto

# Adjunta la imagen al correo
with open('C:\\Users\\nicho\\OneDrive\\Escritorio\\ActivaciondeCheckPoint\\Reporte.png', 'rb') as file:
    img = MIMEImage(file.read())
    img.add_header('Content-ID', '<imagen1>')  # El CID que usaremos en el HTML
    msg.attach(img)

# Crea el cuerpo del correo en HTML
body = """







<html>
<body>
<p>Aquí hay una imagen:</p>
<img src="cid:imagen1" alt="Una imagen">
</body>
</html>
"""

msg.attach(MIMEText(body, 'html'))
smtp_server = "smtp.office365.com"
smtp_port = 587
# Envía el correo
with smtplib.SMTP(smtp_server, 587) as server:
    server.starttls()  # Iniciar conexión segura
    server.login(de, 'Teleco102023')
    server.send_message(msg)

print("Correo enviado!")
