import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
from datetime import datetime
import os
import time
import re
from fuzzywuzzy import process


df_1_CONSOLA = pd.read_excel("Base_Consola.xlsx")


ROOT_DIR = "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\"
DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")


TOKEN = '5899493789:AAE316Rks8Y21wV6VfDgznoJEdBKC4n8n-s'

CHAT_IDS = [1397965849, 5781054062, 6055198218,6512803121,5686702637.6564107654,6383100715]

# Obtener la ruta del directorio donde se ejecuta el script
current_dir = os.path.dirname(os.path.abspath(__file__))
images_path = os.path.join(current_dir, '../ACTIVACIONDECHECKPOINT')

archivoG = pd.read_excel("Base_Completa.xlsx")
falta = pd.DataFrame()


def saludos():
    # Obtener la hora actual
    current_hour = datetime.now().hour

    # Determinar el saludo apropiado según la hora del día
    if 6 <= current_hour < 12:
        greeting = "Buenos días Estimados"
    elif 12 <= current_hour < 18:
        greeting = "Buenas tardes Estimados"
    elif 18 <= current_hour < 24:
        greeting = "Buenas noches Estimados"
    else:
        greeting = "Estimados"
    return greeting


def send_email_to(df_1):
    try:
        # Datos del correo
        de = 'nechever@pacifico.fin.ec'
        ccd= 'telecomunicaciones@pacifico.fin.ec'
        enlacesF = []
        for x in range(len(df_1["Enlace"])):
            try:
                print(df_1["Enlace"][x])

                print("LLEGO AQUI")
                if  "Reportado a proveedor" not in str(df_1.loc[x, "Reportar"]):
                    saludo = saludos()
                    destinatario = "CONSEGUR@pacifico.fin.ec; ConsolaSeguridadGuayaquil@pacifico.fin.ec; grpautoservicios@pacifico.fin.ec; soporteautoservicios@pacifico.fin.ec"
                    estado = "Node status is Down"
                    detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado. Para descartar problemas eléctricos."
                    proveedor = "Consola"

                    Wan = ""
                    login = ""
                    print(df_1["Enlace"][x])

                    if "Caido" not in str(df_1["Tiempos"][x]):
                        print(df_1["Enlace"][x])
                        if " CNT" in df_1["Enlace"][x]:
                            destinatario = "cntcorp@cnt.gob.ec ; soporte.cntcorp@cnt.gob.ec ; alex.mosquera@cnt.gob.ec; paola.lizano@cnt.gob.ec"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "CNT"

                        if " Telconet" in df_1["Enlace"][x]:
                            destinatario = "ipcc_l1_gye@telconet.ec; soporte@telconet.ec; glitardo@telconet.ec;"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "Telconet"
                            try:
                                Wan= df_1["WAN"][x]
                            except:
                                Wan = ""
                            try:
                                login = df_1["Login"][x]
                            except:
                                login = ""
                            
                
                        if " Puntonet" in df_1["Enlace"][x]:
                            destinatario = "cccorporativo@gye.puntonet.ec; soportebp@puntonet.ec"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "Puntonet"

                        if " Movistar" in df_1["Enlace"][x]:
                            destinatario = "csd.ec@telefonica.com"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "Movistar"

                        if " Cirion" in df_1["Enlace"][x]:
                            destinatario = "cpcampos@pacifico.fin.ec"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "Cirion"

                        if " Claro" in df_1["Enlace"][x]:
                            destinatario = "soportetecnicocorp@claro.com.ec; soportedatum@claro.com.ec"
                            estado = "Dentro de este lapso de tiempo"
                            detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                            estado = ""
                            proveedor = "Claro"
                    
                    
                    df_1.loc[x, "Reportado"] = f"Reportado a proveedor ({proveedor})"
                        # Crea un correo multipart
                    msg = MIMEMultipart()
                    msg['From'] = de
                    msg['To'] = destinatario
                    msg['Cc'] = ccd
                    msg['Subject'] = f'Incidente de enlace {df_1["Enlace"][x]}'
                    
                    archivo = f'{str(df_1["Enlace"][x]).strip()}.jpg'
                    firmaLg =  "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\OrionAuto\\ImgCorreo\\logobdp.jpg"
                    SeparadorLg =  "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\OrionAuto\\ImgCorreo\\Separador.jpg"

                    try:
                        with open(archivo, 'rb') as file:
                            img = MIMEImage(file.read())
                            img.add_header('Content-ID', '<imagen1>')  # El CID que usaremos en el HTML
                            msg.attach(img)
                    except:
                        print("No se encontro imagen", archivo)


                    try:
                        with open(firmaLg, 'rb') as file:
                            img2 = MIMEImage(file.read())
                            img2.add_header('Content-ID', '<imagen2>')  # El CID que usaremos en el HTML
                            msg.attach(img2)
                    except:
                        print("No se encontro imagen", firmaLg)
                    
                    try:
                        with open(SeparadorLg, 'rb') as file:
                            img3 = MIMEImage(file.read())
                            img3.add_header('Content-ID', '<imagen3>')  # El CID que usaremos en el HTML
                            msg.attach(img3)
                    except:
                        print("No se encontro imagen", SeparadorLg)
                    
                    # Puedes definir las variables previamente
                    enlace = str(df_1["Enlace"][x])
                    despedida = "Saludos cordiales,"
                    nombre = "Nicholas Echeverria."
                    cargo = "Ingeniero de Telecomunicaciones | MIT"
                    email = ' <a href="mailto:"nechever@pacifico.fin.ec">nechever@pacifico.fin.ec</a>'
                    direccion = "Dirección: P.Icaza y Pedro Carbo"
                    telefono = "Telf.: (04) 731500 Ext. 41802"
                    ciudad = "Guayaquil - Ecuador"
                    web = ' <a href="https://www.bancodelpacifico.com">www.bancodelpacifico.com</a>'

                    # Aquí creamos el cuerpo del correo con HTML

                    body = f'''
                    <html>
                    <body>
                    <h2>{saludo}</h2>
                    <p>{detalle}</p>
                    <p><strong>{enlace}</strong> - {estado}  WAN : {Wan} Login : {login}</p>
                    <p>Se adjunta imagen del evento:</p>
                    <img src="cid:imagen1" alt="Reporte">
                    <br>
                    <p>{despedida}</p>
                    <br>
                    <table style="width:20%">
                        <tr>
                            <td>
                                <img src="cid:imagen2" alt="LogoBdP" width="300" height="200">
                            </td>
                            <td>
                                <p><strong>{nombre}</strong><br>
                                {cargo}<br>
                                {email}</p>
                                <img src="cid:imagen3" alt="Separador" width="300" height="10">
                                <p>{direccion}<br>
                                {telefono}<br>
                                {ciudad}<br>
                                {web}</p>
                            </td>
                        </tr>
                    </table>

                    </body>
                    </html>
                    '''

                    msg.attach(MIMEText(body, 'html'))
                    smtp_server = "smtp.office365.com"
                    smtp_port = 587
                    # Envía el correo
                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls()  # Iniciar conexión segura
                        server.login(de, 'Teleco062024')
                        server.send_message(msg)
            except:
                print("SE cayo")
                enlacesF.append(df_1["Enlace"][x])
                pass
    except:
        print("Se Cayo toalmente")
        enlacesDIC = {"Faltantes": enlacesF}
        faltantes = pd.DataFrame(enlacesDIC)  
        faltantes.to_excel("REENVIARLISTA.xlsx")     
        df_1.to_excel("revisar.xlsx") 
        pass
    
    enlacesDIC = {"Faltantes": enlacesF}
    faltantes = pd.DataFrame(enlacesDIC)  
    faltantes.to_excel("REENVIARLISTA.xlsx")   
    df_1.to_excel("revisar.xlsx") 
    print("Se termino")   
    return df_1


def send_to_email_consola_m(df_1,x):
    saludo = saludos()
    name = str(df_1["Enlace"][x]).split(" ")
    proveedor = name[-1]
    enlace = str(name[:-1])
    ccd = 'telecomunicaciones@pacifico.fin.ec'
    try:
        if df_1["Reportar"][x] != "Reboot":
            if df_1["Reportar"][x] == "Consola":

                print(df_1["Reportar"][x])
                print("Caida")

                palabrasGYE = list(df_1_CONSOLA["Guayaquil"])
                palabrasUIO = list(df_1_CONSOLA["Quito"])
                palabrasCUE = list(df_1_CONSOLA["Cuenca"])
                destinatario = "CONSEGUR@pacifico.fin.ec;grpautoservicios@pacifico.fin.ec; soporteautoservicios@pacifico.fin.ec"

                if str(df_1["Carpeta_Orion"][x]) in palabrasUIO:
                    destinatario = 'consol05@pacifico.fin.ec; grpautoservicios@pacifico.fin.ec; soporteautoservicios@pacifico.fin.ec'
                    ccd='telecomunicaciones@pacifico.fin.ec'
                if str(df_1["Carpeta_Orion"][x]) in palabrasGYE:
                    destinatario = "CONSEGUR@pacifico.fin.ec; grpautoservicios@pacifico.fin.ec; soporteautoservicios@pacifico.fin.ec"
                    ccd='telecomunicaciones@pacifico.fin.ec'
                    #destinatario = "frevelo@pacifico.fin.ec; ajsuarez@pacifico.fin.ec; cpcampos@pacifico.fin.ec; mcaalvar@pacifico.fin.ec; sgellibe@pacifico.fin.ec"
                if str(df_1["Carpeta_Orion"][x]) in palabrasCUE:
                    destinatario = 'consol03@pacifico.fin.ec; grpautoservicios@pacifico.fin.ec; soporteautoservicios@pacifico.fin.ec'
                    ccd='telecomunicaciones@pacifico.fin.ec'


                saludo = "BUENAS TARDES"
                estado = "El enlace se encuentra caido"
                detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado. Para descartar problemas eléctricos."
                proveedor = "Consola"
                Wan = ""
                login = ""
                                    

                df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'
            
            elif df_1["Reportar"][x] == "Proveedor":
                    estado = ""
                    try:
                        Wan= df_1["WAN"][x]
                    except:
                        Wan = ""
                    try:
                        login = df_1["Login"][x]
                    except:
                        login = ""

                    ccd ="telecomunicaciones@pacifico.fin.ec"
                    if "CNT" in str(df_1["Enlace"][x]).split(" "):
                        print("ENTROOOOO  CNT")
                        print(df_1["Enlace"][x])
                        destinatario = "cntcorp@cnt.gob.ec; soporte.cntcorp@cnt.gob.ec; alex.mosquera@cnt.gob.ec; paola.lizano@cnt.gob.ec"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                        estado = "wdjara@pacifico.fin.ec"
                        proveedor = "CNT"
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'

                    elif "Telconet" in str(df_1["Enlace"][x]).split(" "):
                        print("ENTROOOOO  Telconet")
                        destinatario = " ipcc_l1@telconet.ec; soporte@telconet.ec; glitardo@telconet.ec"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                        estado = ""
                        proveedor = "Telconet"
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""

                    elif "Puntonet" in str(df_1["Enlace"][x]).split(" "):
                        print("ENTROOOOO  Puntonet")
                        destinatario = "cccorporativo@gye.puntonet.ec; soportebp@puntonet.ec;"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                        estado = ""
                        proveedor = "Puntonet"
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""

                    elif "Movistar" in str(df_1["Enlace"][x]).split(" "):
                        destinatario = "csd.ec@telefonica.com"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                        estado = ""
                        proveedor = "Movistar"
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'
                    
                    
                    elif "Cirion" in str(df_1["Enlace"][x]).split(" "):
                        destinatario = " nechever@pacifico.fin.ec"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                
                        proveedor = "Cirion"
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'

                    elif "Claro" in str(df_1["Enlace"][x]).split(" "):
                        destinatario = "soportetecnicocorp@claro.com.ec; soportedatum@claro.com.ec;"
                        estado = "Dentro de este lapso de tiempo"
                        detalle = "Favor su gentil ayuda revisando el siguiente enlace afectado en el horario descrito en la imagen. Se descartan problemas eléctricos."
                
                        proveedor = "Claro"
                        try:
                            Wan= df_1["WAN"][x]
                        except:
                            Wan = ""
                        try:
                            login = df_1["Login"][x]
                        except:
                            login = ""
                            
                        df_1.loc[x, 'Reportado'] = f'Reportado a {proveedor}'
                    
        else:
            df_1.loc[x, 'Reportado'] = "No Reportar"
            destinatario =""
            estado =""
            ccd=""
            proveedor=""
            detalle=""
    except:
        
        pass

    if str(df_1.loc[x, "Reportar"])  != "Reportar":
        msg = MIMEMultipart()
        de = "nechever@pacifico.fin.ec"
        msg['From'] = de
        msg['To'] = destinatario
        msg['Cc'] = ccd
        msg['Subject'] = f'Incidente de enlace {df_1["Enlace"][x]}'

        archivo = os.path.join(ROOT_DIR, f'{str(df_1["Enlace"][x]).strip()}.jpg')
        firmaLg =  "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\OrionAuto\\ImgCorreo\\logobdp.jpg"
        SeparadorLg =  "C:\\Users\\nicho\\OneDrive\\Escritorio\\Proyectos\\ActivaciondeCheckPoint\\OrionAuto\\ImgCorreo\\Separador.jpg"


        # Si se encontró una coincidencia, procede con la apertura y adjunto del archivo
        try:
            with open(archivo, 'rb') as file:
                img = MIMEImage(file.read())
                img.add_header('Content-ID', '<imagen1>')  # El CID que usaremos en el HTML
                msg.attach(img)
        except Exception as e:
            print("Ocurrió un error al intentar abrir la imagen", archivo)
            print(e)


        try:
            with open(firmaLg, 'rb') as file:
                img2 = MIMEImage(file.read())
                img2.add_header('Content-ID', '<imagen2>')  # El CID que usaremos en el HTML
                msg.attach(img2)
        except:
            print("No se encontro imagen", firmaLg)

        try:
            with open(SeparadorLg, 'rb') as file:
                img3 = MIMEImage(file.read())
                img3.add_header('Content-ID', '<imagen3>')  # El CID que usaremos en el HTML
                msg.attach(img3)
        except:
            print("No se encontro imagen", SeparadorLg)

        # Puedes definir las variables previamente
        enlace = str(df_1["Enlace"][x])
        despedida = "Saludos cordiales,"
        nombre = "Nicholas Echeverria."
        cargo = "Ingeniero de Telecomunicaciones | MIT"
        email = ' <a href="mailto:"nechever@pacifico.fin.ec">nechever@pacifico.fin.ec</a>'
        direccion = "Dirección: P.Icaza y Pedro Carbo"
        telefono = "Telf.: (04) 731500 Ext. 41802"
        ciudad = "Guayaquil - Ecuador"
        web = ' <a href="https://www.bancodelpacifico.com">www.bancodelpacifico.com</a>'

        # Aquí creamos el cuerpo del correo con HTML

        body = f'''
        <html>
        <body>
        <h2>{saludo}</h2>
        <p>{detalle}</p>
        <p><strong>{enlace}</strong> - {estado}  WAN : {Wan} Login : {login}</p>
        <p>Se adjunta imagen del evento:</p>
        <img src="cid:imagen1" alt="Reporte">
        <br>
        <p>{despedida}</p>
        <br>
        <table style="width:20%">
            <tr>
                <td>
                    <img src="cid:imagen2" alt="LogoBdP" width="300" height="200">
                </td>
                <td>
                    <p><strong>{nombre}</strong><br>
                    {cargo}<br>
                    {email}</p>
                    <img src="cid:imagen3" alt="Separador" width="300" height="10">
                    <p>{direccion}<br>
                    {telefono}<br>
                    {ciudad}<br>
                    {web}</p>
                </td>
            </tr>
        </table>

        </body>
        </html>
        '''

        msg.attach(MIMEText(body, 'html'))
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        # Envía el correo
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Iniciar conexión segura
            server.login(de, 'Teleco062024')
            server.send_message(msg)
        print("No se cayo")
        time.sleep(5)
    else:
        print ("No se envia")

        df_1.to_excel("revisar.xlsx")  
        print("Finalizooooooooo") 
    
    
    return df_1


def send_to_email_consola(df,cont):


    df.loc[cont, 'Reportado'] = f'Reportado a '
    
    return df