# GyG ENVIA RECIBOS DE PAGO
#
# Envía los Recibos de Pago "GyG Recibo_001.pdf", vía WhatsApp o correo electrónico, generados
# en las fases previas, a partir de la hoja de cálculo "1.1. GyG Recibos.xlsm".

# El envío de mensajes vía WhatsApp está basado en el código encontrado en
# http://techniknow.com/whatsapp-using-python/ y descargado el 21-Diciembre-2017

# POR HACER:
# ----------
#  Ok Enviar anexos a los mensajes por Whatsapp
#     ^--> (Esto implica que ya no es necesaria la actualización de los Shared Links en la hoja
#           de cálculo ni la copia de los recibos a Google Drive)
#  Ok Ordenar los recibos de pago por beneficiario y número de recibo para enviarlos en un
#     mismo correo o en conjunto en un mismo chat de WhatsApp
#  Ok Indicar a quién no se le envió el mensaje por habérselo enviado previamente
#  Ok Hacer referencia a números de recibo (ej. "Recibo 01207") en lugar de nombres de archivo en
#     el archivo de log
#  Ok Corregir: Los recibos a ser entregados "En Físico" son mostrados como a ser enviados por
#     Whatsapp
#  Ok Agrupar los recibos por beneficiario para realizar un solo envío
#     Ok Revisar qué hacer con los recibos que fueron enviados anteriormente
#        ^--> Emitir mensaje con los números de recibo de aquellos que fueron enviados previamente
#             y enviar sólo aquellos que no
#             ^--> Se estaban enviado correos erroneos a aquellos beneficiarios con todos los
#                  recibos ya enviados (lista de anexos vacía)
#  Ok Utilizar la librería YAGMAIL para el envío de los correos electrónicos, en lugar de
#     SMTPLIB
#  Ok Generar automáticamente el resumen de cuotas de vigilancia a partir de la misma hoja
#     de cálculo
#  Ok Enviar recibo de pago a múltiples destinatarios, e-mail o WhatsApp
#     ^--> Colocar los teléfonos o direcciones de correo, usando la barra vertical como separador:
#          me@mailserver.com | 0555-550.1234 | you@anothermailserver.com (19/04/2019)
#     ^--> Agrupar todas las direcciones de correo en una lista, a fin de enviar un único correo,
#          aprovechando las capacidades de yagmail (21/04/2019)
#     ^--> Corrección de errores introducidos (25/04/2019)
#  *  Enviar imagen del recibo por WhatsApp, en lugar del archivo PDF en sí (no todos los equipos
#     tienen como desplegar este tipo de archivos)


# Control de salidas
SEND_EMAIL      = True
SEND_WHATSAPP   = True


# Selecciona las librerías a utilizar
print('Cargando librerías...')
from pandas import read_excel, isnull, notnull
import re
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import time
import os
import sys
from locale import setlocale, LC_ALL

if SEND_EMAIL:
    import yagmail

if SEND_WHATSAPP:
    import win32gui
    import win32con
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from pdf2image import convert_from_path
#    import random, string

import credentials


# Define textos
attach_name     = "GyG Recibo_{seq:05d}.pdf"
# attach_path     = "C:/Users/MColosso/Google Drive/GyG Recibos/Recibos/"
attach_path     = "C:/Users/MColosso/Documents/GyG Recibos_temp/"
#temp_file       = attach_path + '~' + ''.join(random.sample(string.hexdigits, 7)) + '.png'

recibo_fmt      = "{seq:05d}"

excel_workbook  = '1.1. GyG Recibos.xlsm'
excel_worksheet = 'Vigilancia'
excel_resumen   = 'RESUMEN VIGILANCIA'

email_subject   = 'Cuadra Segura GyG - Recibos de Pago'
email_body_hdr  = '{saludo}<br><br>Anexamos {el_los} recibo{s} de pago de {beneficiario} correspondiente{s} a'
email_body_1    = ' "{concepto}"<br>'
email_body_n    = ':<ul>'
email_body_det  = '<li> {fecha:%d/%m/%Y}  "{concepto}"</li>'
email_cuotas    = 'Generado automáticamente...'
email_footer    = '<i>Equipo de cobranza, Junta Directiva GyG</i>'

wa_message      = 'Si desea recibir sus recibos de pago por correo electrónico, envíe un mensaje ' + \
                  'con sus datos a CuadraSeguraGyG@gmail.com'

msg_enviado_anteriormente = 'x  Recibo{s} {filename} enviado{s} anteriormente a {beneficiario}\n'
msg_no_email_or_phone     = 'x  Recibo{s} {filename} no enviado{s} a {beneficiario}: No se ha indicado e-mail o celular\n'

whatsapp_loaded = False
web_driver      = None

chrome_webdriver  = 'C:/Users/MColosso/Downloads/ChromeDriver_Win32/ChromeDriver'
firefox_webdriver = 'C:/Users/MColosso/Downloads/GeckoDriver-v0.24.0-Win64/GeckoDriver'

nMeses          = 6   # Se muestran los <nMeses> últimos meses en el resumen de cuotas

dummy = setlocale(LC_ALL, '')


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def saludo():
    return 'Hola vecino(a),'
#    now = datetime.now()
#    if now.hour in range(6, 13):      # 6:00 am -> 12:59 m
#        return 'Buenos días'
#    elif now.hour in range(13, 19):   # 1:00 pm -> 06:59 pm
#        return 'Buenas tardes'
#    else:                             # 7:00 pm -> 05:59 am del día siguiente
#        return 'Buenas noches'

def attachment_name(idx):
    return attach_path + attach_name.format(seq=idx)

def get_filename(filename):
    return os.path.basename(filename)

def convert_phone(phone_to):
    COUNTRY_CODE = '58'
    new_phone = ''.join(re.findall(r'\d+', phone_to))
    if new_phone[0] == '0':
        new_phone = COUNTRY_CODE + new_phone[1:]
    return new_phone

def is_email(send_to):
    return re.search(r'[\w.-]+@[\w.-]+', send_to)

def is_phone(send_to):
    return len(re.sub(r'[0-9\+\-\.\(\) ]*', r'', send_to)) == 0

def date_to_str(date):
    Meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
             'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    return '{dd} de {mm} de {yyyy}'.format(dd=date.day, mm=Meses[date.month-1], yyyy=date.year)

def decode_email_o_celular(email_o_celular):
    # return [str.strip(x) for x in re.split('\|', email_o_celular)]
    destinos = [str.strip(x) for x in re.split('\|', email_o_celular)]

    correos = list()
    otros_destinos = list()
    for destino in destinos:
        if is_email(destino):
            correos.append(destino)
        else:
            otros_destinos.append(destino)

    destinos = list()
    if len(correos) > 0:
        destinos.append(correos)
    for dest in otros_destinos:
        destinos.append(dest)

    return destinos

def list_to_string(email_o_celular):
    '''
    Convierte la lista 'email_o_celular' en un string separado por comas
    '''
    if not isinstance(email_o_celular, list):
        return email_o_celular
    elif len(email_o_celular) == 1:
        return email_o_celular[0]
    else:
        lista_separada_por_comas = ', '.join(email_o_celular)
        # convierte la última coma en ' y'
        última_coma = lista_separada_por_comas.rfind(',')
        return lista_separada_por_comas[: última_coma] + ' y' + lista_separada_por_comas[última_coma + 1: ]

def trunc_text(texto):
    MAX_WIDTH = 78

    txt = texto
    if len(texto) > MAX_WIDTH:
        txt = texto[0: MAX_WIDTH - 3] + '...'
    return txt


def resumen_de_cuotas():

    def rango_fechas(idx_inicial, idx_final, cuota, rango_final):
        # final = ', sujeto a revisiones mensuales' if rango_final else ''
        final = ''
        cuota_ed = f"{cuota:,.02f}".replace(',', 'x').replace('.', ',').replace('x', '.')
        if idx_inicial == idx_final:
            texto = f"<li> {columnas[idx_inicial]:%B/%Y}, Bs. {cuota_ed}{final}</li>"
        else:
            a_y = 'y' if idx_final - idx_inicial == 1 else 'a'
            de = 'de ' if idx_final - idx_inicial != 1 else ''
            año_inicio = '' if columnas[idx_inicial].year == columnas[idx_final].year else '/%Y'
            texto =  "<li>" + \
                    f" {de}{columnas[idx_inicial]:%B{año_inicio}}" + \
                    f" {a_y} {columnas[idx_final]:%B/%Y}, Bs. {cuota_ed}{final}" + \
                     "</li>"
        return texto

    def genera_resumen():
        global columnas

        columnas = df_cuotas.columns.tolist()
        cuotas = df_cuotas.iloc[0].tolist()
#        print(f'DEBUG: Columnas = {columnas}'); sys.stdout.flush()
#        print(f'DEBUG: dt_inicial = {dt_inicial}, dt_final = {dt_final}'); sys.stdout.flush()
        col_inicial = columnas.index(dt_inicial)
        col_final   = columnas.index(dt_final)
        idx = col_inicial
        cuota = cuotas[idx]
        last_notnull = idx
        txtCuotasMensuales = '<br>Recordamos que las cuotas mensuales quedaron establecidas de la siguiente forma:<ul>'
        for col in range(col_inicial + 1, col_final + 1):
            if notnull(cuotas[col]):
                last_notnull = col
                if cuotas[col] != cuota:
                    txtCuotasMensuales += rango_fechas(idx, col - 1, cuota, rango_final=False)
                    idx = col
                    cuota = cuotas[idx]
        txtCuotasMensuales += rango_fechas(idx, last_notnull, cuota, rango_final=True)
#            if cuotas[col] != cuota:
#                txtCuotasMensuales += rango_fechas(idx, col - 1, cuota, rango_final=False)
#                idx = col
#                cuota = cuotas[idx]
#        txtCuotasMensuales += rango_fechas(idx, col_final, cuota, rango_final=True)
        txtCuotasMensuales += '</ul><br>'

        return txtCuotasMensuales

    hoy = datetime.now()
    dt_final = datetime(hoy.year, hoy.month, 1)
    dt_inicial = dt_final - relativedelta(months=nMeses - 1)

    return genera_resumen()


#
# RUTINAS PARA ENVIAR CORREOS ELECTRONICOS O ENVIAR MENSAJES VIA WHATSAPP
#

def sendmail(toaddr, subject, body, attachments=[]):

    # Compose some of the basic message headers
    gmail_account = credentials.email_account
    gmail_password = credentials.email_password
    fromaddr = credentials.email_fromaddr

    email = yagmail.SMTP({gmail_account: fromaddr}, gmail_password)

    compact_body = ''.join([s.strip() for s in iter(body.splitlines())])

    try:
        email.send(to=toaddr, subject=subject, contents=compact_body, attachments=attachments)
        email_sent = True
        error_msg = ''
    except:
        email_sent = False
        error_msg  = str(sys.exc_info()[1])
        error_msg  = error_msg.replace('\\', '/')

    email.close()

    return email_sent, error_msg

def loadWhatsAppWeb():
    global whatsapp_loaded, web_driver

    # Replace below path with the absolute path
    # to chromedriver in your computer

    web_driver = webdriver.Chrome(executable_path=chrome_webdriver)
    #web_driver = webdriver.Firefox(executable_path=firefox_webdriver)

    #After opening browser open web.whatsapp.com through next command
    web_driver.get("https://web.whatsapp.com/")

    #Now you need to scan the QR CODE on browser through your mobile whatsapp
    web_driver.implicitly_wait(100)   # seconds
#    time.sleep(5)
    whatsapp_loaded = True

def unloadWhatsAppWeb():
    global whatsapp_loaded, web_driver

    # Open menu
    inp_xpath = '//*[@id="side"]/header/div[2]/div/span/div[3]/div/span'
    button = web_driver.find_elements_by_xpath(inp_xpath)
    button[0].click()

    # Close session
    inp_xpath = '//*[@id="side"]/header/div[2]/div/span/div[3]/span/div/ul/li[6]/div'
    button = web_driver.find_elements_by_xpath(inp_xpath)
    button[0].click()

    # Close web driver
    web_driver.quit()
    
    whatsapp_loaded = False

def sendWhatsApp(message_to, message, attachments=[]):

    """
    How to send a WhatsApp message without saving the contact

    * Type the following link in the search bar: https://api.whatsapp.com/send?phone=XXXXXXXXXXX (In place of the Xs
      type the phone number of the person you want to contact, including the country code, but without the + sign.)
    * That means that if the person has an American number (with the +1 prefix), it would look something like this:
      https://api.whatsapp.com/send?phone=19175550123
    * Press 'enter' on your smartphone.
    * A WhatsApp window will open asking if you want to send a message to that phone number. Press on 'send message'.
    * You will automatically be redirected to WhatsApp with the 'start chatting' window to the person you entered in
      your phone.

    """

    # By this you will give the location where to search your target or contact
    # So it will specify the place of message box on top
    # and than search inside that your contact name if found than move ahead

    try:
#        new_chat = web_driver.find_element_by_id('input-chatlist-search')
        new_chat = web_driver.find_elements_by_xpath('//*[@id=\"side\"]/div[1]/div/label/input')
    except:
        return False, 'No se encontró el placeholder para la búsqueda del destinatario'

    new_chat[0].send_keys(message_to, Keys.ENTER)

    time.sleep(2)

#    # Test if destinatary was found
#    inp_css = 'div.pluggable-input-body.copyable-text.selectable-text'
#    hwnd = web_driver.find_elements_by_css_selector(inp_css)
#    if hwnd == []:
#        # Clean destinatary field()
#        #button = web_driver.find_elements_by_xpath('//*[@id="side"]/div[2]/div/span/button/span')
#        #button[0].click()
#
#        return False, 'Destinatario no encontrado en la lista de contactos'
#-----------------
# Si, después de presionar <Enter> para buscar un destinatario, se muestra el texto "No se encontré ningún chat,
# contacto ni mensaje" en xpath=//*[@id="pane-side"]/div/div/span


    # Add the attachments to the message
    for attachment in attachments:

        # Press 'Adjuntar' button
        inp_xpath = '//*[@id=\"main\"]/header/div[3]/div/div[2]/div/span'
        button = web_driver.find_elements_by_xpath(inp_xpath)
        button[0].click()

        # Press 'Fotos y videos' button [1. Fotos y videos, 2. Cámara, 3. Documento, 4. Contacto]
        inp_xpath = '//*[@id=\"main\"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button'
        button = web_driver.find_elements_by_xpath(inp_xpath)
        button[0].click()

        # Loop until Open dialog is displayed (my Windows version is in Spanish)
        hdlg = 0
        while hdlg == 0:
            hdlg = win32gui.FindWindow(None, "Abrir")
    
        time.sleep(1)   # second
    
        # Set filename and press Enter key
        hwnd          = win32gui.FindWindowEx(hdlg, 0, 'ComboBoxEx32', None)
        hwnd          = win32gui.FindWindowEx(hwnd, 0, 'ComboBox', None)
        hwnd_filename = win32gui.FindWindowEx(hwnd, 0, 'Edit', None)

        hwnd_abrir    = win32gui.FindWindowEx(hdlg, 0, 'Button', '&Abrir')
    
        dirname_Win   = os.path.dirname(attachment).replace('/', '\\')
        basename_Win  = os.path.basename(attachment)
#        print(f'DEBUG: attachment = {dirname_Win} + {basename_Win}')

        # Stablish new default folder
        win32gui.SendMessage(hwnd_filename, win32con.WM_SETTEXT, None, dirname_Win)
        win32gui.SendMessage(hwnd_abrir,    win32con.BM_CLICK,   None, None)
#        print(f'DEBUG: sending message "{dirname_Win}" and pressing "Abrir"')

        time.sleep(3)   # second

        # Set image name
        win32gui.SendMessage(hwnd_filename, win32con.WM_SETTEXT, None, basename_Win)
        win32gui.SendMessage(hwnd_abrir,    win32con.BM_CLICK,   None, None)
#        print(f'DEBUG: sending message "{basename_Win}" and pressing "Abrir"')
    
        time.sleep(1)   # second

        # Press send button
        inp_xpath = '//*[@id=\"app\"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span[2]/div/div/span'
        button = web_driver.find_elements_by_xpath(inp_xpath)
        button[0].click()

    if len(attachments) > 0:
        # Wait 'till file is loaded (wait 'till download button is shown)
        inp_xpass = '//*[@id="main"]/div[2]/div/div/div[2]/div[9]/div/div[1]/div/a/div[2]/div[3]/span'
        web_driver.find_elements_by_xpath(inp_xpath)

    # Write aditional message
    inp_xpass = '//*[@id=\"main\"]/footer/div[1]/div[2]/div/div[2]'
    input_box = web_driver.find_elements_by_xpath(inp_xpass)
    input_box[0].send_keys(message, Keys.ENTER)
    
#    # Press Send message button
#    button = web_driver.find_elements_by_xpath('//*[@id="main"]/footer/div[3]/button/span')
#    button[0].click()
    
    time.sleep(1) # second

    return True, ''

# def envia_archivo(attachment, email_o_celular, beneficiario, fecha, shared_link, concepto):
def envia_archivo(beneficiario, email_o_celular, sub_lista):

    anexos = [recibo_fmt.format(seq=df.loc[i, 'Nro. Recibo']) for i in sub_lista]
    anexos = re.sub(r'(.*), (.*)', r'\1 y \2', ', '.join(anexos))

    if email_o_celular == None:
        logfile.write('-> Enviar recibo{} {} a {}\n'.format('s' if len(sub_lista) > 1 else '',
                                                            anexos, beneficiario))
        return
    es_correo  = isinstance(email_o_celular, list) or is_email(email_o_celular)
    es_celular = (not isinstance(email_o_celular, list)) and is_phone(email_o_celular)
    if es_correo:
        if not SEND_EMAIL:
            message = '-> E-Mail:   Enviar recibo{s} {filename} a {beneficiario} ({email})\n'
            logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                         filename=anexos,
                                         beneficiario=beneficiario,
                                         email=list_to_string(email_o_celular)))
            archivo_enviado = False
        else:
            to = dict()
            for send_to in email_o_celular:
                to.update({send_to: beneficiario})
            subject = email_subject
            body = email_body_hdr.format(saludo=saludo(), beneficiario=beneficiario,
                                         el_los='el' if len(sub_lista) == 1 else 'los',
                                         s='s' if len(sub_lista) > 1 else '')
            if len(sub_lista) == 1:
                body += email_body_1.format(concepto=df.loc[sub_lista[0], "Concepto"])
            else:
                body += email_body_n
                for idx in sub_lista:
                    body += email_body_det.format(fecha=df.loc[idx, "Fecha"],
                                                  concepto=df.loc[idx, "Concepto"])
                body += '</ul>'
            body += email_cuotas
            body += email_footer

            archivo_enviado, error_msg = sendmail(to, subject, body,
                                                  attachments=[attachment_name(df.loc[idx, 'Nro. Recibo']) for idx in sub_lista])
            if archivo_enviado:
                message = 'Ok Recibo{s} {filename} enviado{s} a {beneficiario} ({email})\n'
                logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                             filename=anexos,
                                             beneficiario=beneficiario,
                                             email=list_to_string(email_o_celular)))
            else:
                message = 'X  Recibo{s} {filename} no enviado{s} a {beneficiario}: {mensaje}\n'
                logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                             filename=anexos,
                                             beneficiario=beneficiario,
                                             mensaje=error_msg))
    elif es_celular:
        if not SEND_WHATSAPP:
            message = '-> WhatsApp: Enviar recibo{s} {filename} a {beneficiario} ({phone})\n'
            logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                         filename=anexos,
                                         beneficiario=beneficiario,
                                         phone=list_to_string(email_o_celular)))
            archivo_enviado = False
#        elif shared_link == None:
#            logfile.write(msg_no_shared_link.format(filename=attachment))
#            archivo_enviado = False
        else:
            if not whatsapp_loaded:
                loadWhatsAppWeb()

            message_to = beneficiario
            message = wa_message

            attachments = list()
            for idx in sub_lista:
                attachment = attachment_name(df.loc[idx, 'Nro. Recibo'])
                # Convert attachment into an image file
                temp_file, ext = os.path.splitext(attachment)
                temp_file += '.png'
                pages = convert_from_path(attachment)
                pages[0].save(temp_file, 'PNG')
                attachments.append(temp_file)

            archivo_enviado, error_msg = sendWhatsApp(message_to, message,
                                                      attachments=attachments)

            # Delete image files
            for attachment in attachments:
                os.remove(attachment)

            if archivo_enviado:
                message = 'Ok Recibo{s} {filename} enviado{s} a {beneficiario} ({phone}) vía WhatsApp\n'
                logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                             filename=anexos,
                                             beneficiario=beneficiario,
                                             phone=list_to_string(email_o_celular)))
            else:
                message = 'X  Recibo{s} {filename}" no enviado{s} a {beneficiario} ({phone}): {mensaje}\n'
                logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                             filename=anexos,
                                             beneficiario=beneficiario,
                                             phone=list_to_string(email_o_celular),
                                             mensaje=error_msg))
    else:
        message = '-> Enviar recibo{s} {filename} a {beneficiario} ({phone})\n'
        logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                     filename=anexos,
                                     beneficiario=beneficiario,
                                     phone=list_to_string(email_o_celular)))
        archivo_enviado = False


    return archivo_enviado


#
# PROCESO
#


# Abre la hoja de cálculo de Recibos de Pago
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))

df = read_excel(excel_workbook, sheet_name=excel_resumen)
df_cuotas = df[df['Beneficiario'] == 'CUOTAS MENSUALES']
email_cuotas = resumen_de_cuotas()

# Carga la hoja de cálculo con los pagos
df = read_excel(excel_workbook, sheet_name=excel_worksheet)

# Elimina registros que no fueron seleccionados para enviar
df.dropna(subset=['Archivo'], inplace=True)

# Convierte columnas 'Archivo' y 'Nro. Recibo' en enteros
df['Archivo'] = df['Archivo'].astype(int)
df['Nro. Recibo'] = df['Nro. Recibo'].astype(int)

# Ordena los pagos recibidos por beneficiario y recibo
df.sort_values(['Beneficiario', 'Nro. Recibo'], inplace=True)

# Genera una lista de pagos por beneficiario
indices = list(df.index)
pagos_por_beneficiario = list()
if len(indices) == 0:
    print('\n*** Proceso terminado: No hay recibos pendientes por enviar\n')
    sys.exit()

idx_anterior = indices[0]
sub_lista = [idx_anterior]
for idx in indices[1:]:
    if df.loc[idx_anterior, 'Beneficiario'] == df.loc[idx, 'Beneficiario']:
        sub_lista.append(idx)
    else:
        pagos_por_beneficiario.append(sub_lista)
        sub_lista = [idx]
    idx_anterior = idx
pagos_por_beneficiario.append(sub_lista)

# Crea el logfile
print('Creando logfile...')
# filename = '{}Recibos {:%Y-%m-%d %I.%M%p}.log'.format(attach_path, datetime.now())
filename = '{}Recibos {:%Y-%m-%d %H.%M}.log'.format(attach_path, datetime.now())
logfile = open(filename, 'w')

# Para cada recibo de pago, determina si se debe enviar por correo o como mensaje
print('Enviando mensajes', end="")

mensajes_enviados = 0
for sub_lista in pagos_por_beneficiario:
    r = df.loc[sub_lista[-1]]
    print('.', end='')   # Imprime un punto en la pantalla por cada mensaje
    sys.stdout.flush()   # Flush output to the screen
    beneficiario    = r['Beneficiario']
    email_o_celular = r['E-mail o celular']
    # concepto        = r['Concepto']
    # fecha           = r['Fecha'].date()
    # shared_link     = r['Shared Link']
    # attachment      = attachment_name(r['Nro. Recibo'])
    recibos = [recibo_fmt.format(seq=idx) for idx in sub_lista]
    recibos = re.sub(r'(.*), (.*)', r'\1 y \2', ', '.join(recibos))
    if isnull(email_o_celular):
        logfile.write(msg_no_email_or_phone.format(s='s' if len(sub_lista) > 1 else '',
                                                   filename=recibos,
                                                   beneficiario=beneficiario))
    else:
        enviados = [idx for idx in sub_lista if notnull(df.loc[idx, 'Enviado'])]
        no_enviados = [idx for idx in sub_lista if idx not in enviados]
        if len(enviados) > 0:
            recibos = [recibo_fmt.format(seq=idx) for idx in enviados]
            recibos = re.sub(r'(.*), (.*)', r'\1 y \2', ', '.join(recibos))
            logfile.write(msg_enviado_anteriormente.format(s='s' if len(enviados) > 1 else '',
                                                           filename=recibos,
                                                           beneficiario=beneficiario))
        if len(no_enviados) > 0:
            for send_to in decode_email_o_celular(email_o_celular):
                if envia_archivo(beneficiario, send_to, no_enviados):
                    mensajes_enviados += len(no_enviados)

logfile.close()

print('\n')

print(os.path.basename(filename))
with open(filename) as logfile:
    for cnt, line in enumerate(logfile):
        print(trunc_text(line.strip()))

print('\n*** Proceso terminado: {} de {} recibos enviados\n'.format(mensajes_enviados,
                                                                    df.shape[0]))

if whatsapp_loaded:
    unloadWhatsAppWeb()
