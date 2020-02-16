# GyG ENVIA RECIBOS DE PAGO
#
# Envía los Recibos de Pago "GyG Recibo_001.pdf", vía WhatsApp o correo electrónico, generados
# en las fases previas, a partir de la hoja de cálculo "1.1. GyG Recibos.xlsm".

"""
    PENDIENTE POR HACER
    -   Enviar como anexos archivos .png (en lugar de .pdf), si estuviesen presentes; ello implica
        que no se generará el archivo .png al enviar por WhatsApp, si estuviese presente.

    HISTORICO
    -   Se cambiaron las ubicaciones de los archivos resultantes a la carpeta GyG Recibos dentro de
        la carpeta actual para compatibilidad entre Windows y macOS (21/10/2019)
    -   Cambiar el manejo de cuotas para usar las rutinas en la clase Cuota (GyG_cuotas)
        (09/10/2019)
    -   Ajustar el resumen_de_cuotas() para indicar que, a partir de Septiembre 2019 se maneja una
        cuota mensual de 1 dólar, actualizada semanalmente. (04/09/2019)
    -   CORREGIR: En la hoja de cálculo '1.1. GyG Recibos.xlsm' se cambió el título de la columna 'B'
        de 'Archivo' a 'Generar' (corregido 08/07/2019)
    -   Los correos con los recibos de pagos por anticipado no muestran el monto de las nuevas cuotas
        (si estuvieran definidas). (Corregido 01/07/2019)
    -   Enviar una imagen del Recibo de Pago por WhatsApp en lugar del archivo .pdf generado
        (23/06/2019)
    -   CORREGIR: Cambiar referencias de "df['Archivo']" a "df['Nro. Recibo']" para evitar problemas
        de numerarión manual (corregido 21/06/2019)
    -   CORREGIR: Las cuotas mostradas para los colegios y otros inmuebles especiales son las
        mismas que la del resto de los vecinos, por lo que los montos desplegados no son correctos
        (corregido 10/06/2019)
    -   Enviar recibo de pago a múltiples destinatarios, e-mail o WhatsApp
        ^--> Colocar los teléfonos o direcciones de correo, usando la barra vertical como separador:
             me@mailserver.com | 0555-550.1234 | you@anothermailserver.com (19/04/2019)
        ^--> Agrupar todas las direcciones de correo en una lista, a fin de enviar un único correo,
             aprovechando las capacidades de yagmail (21/04/2019)
        ^--> Corrección de errores introducidos (25/04/2019)
    -   Generar automáticamente el resumen de cuotas de vigilancia a partir de la misma hoja
        de cálculo
    -   Utilizar la librería YAGMAIL para el envío de los correos electrónicos, en lugar de
        SMTPLIB
    -   Agrupar los recibos por beneficiario para realizar un solo envío
        -  Revisar qué hacer con los recibos que fueron enviados anteriormente
           ^--> Emitir mensaje con los números de recibo de aquellos que fueron enviados previamente
                y enviar sólo aquellos que no
                ^--> Se estaban enviado correos erroneos a aquellos beneficiarios con todos los
                     recibos ya enviados (lista de anexos vacía)
    -   Corregir: Los recibos a ser entregados "En Físico" son mostrados como a ser enviados por
        Whatsapp
    -   Hacer referencia a números de recibo (ej. "Recibo 01207") en lugar de nombres de archivo en
        el archivo de log
    -   Indicar a quién no se le envió el mensaje por habérselo enviado previamente
    -   Ordenar los recibos de pago por beneficiario y número de recibo para enviarlos en un
        mismo correo o en conjunto en un mismo chat de WhatsApp
    -   Enviar anexos a los mensajes por Whatsapp
        ^--> (Esto implica que ya no es necesaria la actualización de los Shared Links en la hoja
              de cálculo ni la copia de los recibos a Google Drive)
    
"""

# Control de salidas
SEND_EMAIL    =  True
SEND_WHATSAPP =  False

crop_image    =  True
tipo_imagen   = '.png'

# Selecciona las librerías a utilizar
print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *
from pandas import read_excel, isnull, notnull
import re
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import time
import os
import sys
import locale
from pdf2image import convert_from_path
if crop_image:
    from PIL import Image

if SEND_EMAIL:
    import yagmail

if SEND_WHATSAPP:
    import win32gui
    import win32con
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys

import credentials


# Define textos
attach_pdf_name = GyG_constantes.pdf_recibo            # "GyG Recibo_{recibo:05d}.pdf"
attach_img_name = GyG_constantes.img_recibo            # "GyG Recibo_{recibo:05d}.png"
attach_path     = GyG_constantes.ruta_recibos          # './GyG Recibos/Recibos de Pago'

recibo_fmt      = GyG_constantes.recibo_fmt            # "{recibo:05d}"

excel_workbook  = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_worksheet = GyG_constantes.pagos_ws_vigilancia   # 'Vigilancia'
excel_resumen   = GyG_constantes.pagos_ws_resumen      # 'RESUMEN VIGILANCIA'
excel_cuotas    = GyG_constantes.pagos_ws_cuotas       # 'CUOTA'

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

chrome_webdriver = 'C:/Users/MColosso/Downloads/ChromeDriver_Win32/ChromeDriver'

nMeses          = 6   # Se muestran los <nMeses> últimos meses en el resumen de cuotas

dummy = locale.setlocale(locale.LC_ALL, 'es_es')


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
    img_filename = os.path.join(attach_path, attach_img_name.format(recibo=idx))
    if os.path.lexists(img_filename):
        return img_filename
    else:
        return os.path.join(attach_path, attach_pdf_name.format(recibo=idx))

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
        if sys.platform.startswith('win'):
            error_msg  = error_msg.replace('\\', '/')

    email.close()

    return email_sent, error_msg

def loadWhatsAppWeb():
    global whatsapp_loaded, web_driver

    # Replace below path with the absolute path
    # to chromedriver in your computer
    web_driver = webdriver.Chrome(chrome_webdriver)
    #web_driver = webdriver.Firefox()
    web_driver.maximize_window()

    #After opening browser open web.whatsapp.com through next command
    web_driver.get("https://web.whatsapp.com/")

    #Now you need to scan the QR CODE on browser through your mobile whatsapp
    web_driver.implicitly_wait(100)   # seconds

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

    # By this you will give the location where to search your target or contact
    # So it will specify the place of message box on top
    # and than search inside that your contact name if found than move ahead

    try:
#        new_chat = web_driver.find_element_by_id('input-chatlist-search')
        inp_xpath = '/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/input'
        new_chat = web_driver.find_element_by_xpath(inp_xpath)
    except:
        return False, 'No se encontró el placeholder para la búsqueda del destinatario'

    new_chat.send_keys(message_to, Keys.ENTER)

    time.sleep(2)

    # Test if destinatary was found
    inp_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]'
    hwnd = web_driver.find_element_by_xpath(inp_xpath)
    if hwnd == []:
        return False, 'Destinatario no encontrado en la lista de contactos'

    # Add the attachments to the message
    for attachment in attachments:

        # Press Attach button
        button = web_driver.find_elements_by_xpath('//*[@id=\"main\"]/header/div[3]/div/div[2]/div/span')
        button[0].click()

        # Presiona botón 'Fotos y videos' (li[1]) 1. Fotos y videos, 2. Cámara, 3. Documento, 4. Contacto
        inp_xpath = '//*[@id=\"main\"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button'
        button = web_driver.find_elements_by_xpath(inp_xpath)
        button[0].click()
    
        # Loop until Open dialog is displayed (my Windows version is in Spanish)
        hdlg = 0
        while hdlg == 0:
            hdlg = win32gui.FindWindow(None, "Abrir")
        try:
            win32gui.SetForegroundWindow(hdlg)
        except:
            pass

        time.sleep(1)   # second
    
        # Set filename and press Enter key
        hwnd = win32gui.FindWindowEx(hdlg, 0, 'ComboBoxEx32', None)
        hwnd = win32gui.FindWindowEx(hwnd, 0, 'ComboBox', None)
        hwnd = win32gui.FindWindowEx(hwnd, 0, 'Edit', None)
    
        win_attachment = os.path.normcase(attachment) #.replace('/', '\\')

        # Send folder and press Save button
        win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)
        win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, os.path.dirname(win_attachment))
        btnOpen = win32gui.FindWindowEx(hdlg, 0, 'Button', '&Abrir')
        win32gui.SendMessage(btnOpen, win32con.BM_CLICK, None, None)
    
        time.sleep(40)   # second
    
        # Send filename and press Save button
        win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)
        win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, os.path.basename(win_attachment))
        btnOpen = win32gui.FindWindowEx(hdlg, 0, 'Button', '&Abrir')
        win32gui.SendMessage(btnOpen, win32con.BM_CLICK, None, None)
    
        time.sleep(25)   # second
    
        # Press send button
        inp_xpath = '//*[@id=\"app\"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span[2]/div/div/span'
        button = web_driver.find_elements_by_xpath(inp_xpath)
        button[0].click()

#    if len(attachments) > 0:
#        # Wait 'till file is loaded (wait 'till download button is shown)
#        inp_xpath = '//*[@id="main"]/div[2]/div/div/div[2]/div[9]/div/div[1]/div/a/div[2]/div[3]/span'
#        web_driver.find_elements_by_xpath(inp_xpath)

    # Write aditional message
    inp_css = 'div._3FeAD._1PRhq'
    input_box = web_driver.find_elements_by_css_selector(inp_css)
    input_box[0].send_keys(message, Keys.ENTER)
    
    # Press Send message button
#    button = web_driver.find_elements_by_css_selector('button._3M-N-')
#    button[0].click()
    
    time.sleep(1) # second

    return True, ''

# def envia_archivo(attachment, email_o_celular, beneficiario, fecha, shared_link, concepto):
def envia_archivo(beneficiario, email_o_celular, sub_lista):

    def convierte_en_imagenes(anexos):

        img_anexos = list()
        for anexo in anexos:
            img_anexo = os.path.join(os.path.dirname(anexo),
                                     os.path.basename(anexo).split('.')[0] + tipo_imagen)

            if not os.path.lexists(img_anexo):
                try:
                    pages = convert_from_path(anexo)
                except:
                    error_msg  = str(sys.exc_info()[1])
                    if sys.platform.startswith('win'):
                        error_msg  = error_msg.replace('\\', '/')
                    print(f'*** Error convirtiendo {os.path.basename(anexo)}: {error_msg}')
                    continue

                try:
                    pages[0].save(img_anexo)
                except:
                    error_msg  = str(sys.exc_info()[1])
                    if sys.platform.startswith('win'):
                        error_msg  = error_msg.replace('\\', '/')
                    print(f'*** Error generando {img_anexo}: {error_msg}')
                    continue

                if crop_image:
                    img_recibo = Image.open(img_anexo)
                    img_recibo = img_recibo.resize(size=(712, 363), box=(108, 84, 1570, 825), resample=Image.HAMMING)
                    img_recibo.save(img_anexo)

            img_anexos.append(img_anexo)

        return img_anexos


    anexos = [recibo_fmt.format(recibo=df.loc[i, 'Nro. Recibo']) for i in sub_lista]
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
            body += '<br>' + cuotas_obj.resumen_de_cuotas(beneficiario, fecha_referencia, formato=fmtHtml, lista=True)
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
        img_anexos = convierte_en_imagenes([attachment_name(df.loc[idx, 'Nro. Recibo']) for idx in sub_lista])
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
#            attachment = attachment.replace('/', '\\')  # Convertir Slashes en Backslashes (estándar de Windows)
            archivo_enviado, error_msg = sendWhatsApp(message_to, message,
                                                      attachments=img_anexos)
            if archivo_enviado:
                message = 'Ok Recibo{s} {filename} enviado{s} a {beneficiario} ({phone}) vía WhatsApp\n'
                logfile.write(message.format(s='s' if len(sub_lista) > 1 else '',
                                             filename=anexos,
                                             beneficiario=beneficiario,
                                             phone=list_to_string(email_o_celular)))
            else:
                message = 'X  Recibo{s} {filename} no enviado{s} a {beneficiario} ({phone}): {mensaje}\n'
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
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook ))

df = read_excel(excel_workbook , sheet_name=excel_resumen)

# Inicializa el handler para el manejo de las cuotas
cuotas_obj = Cuota(excel_workbook)

#df_cuotas = df[df['Beneficiario'] == 'CUOTAS MENSUALES']

df.loc[df[isnull(df['Beneficiario'])].index, 'Beneficiario'] = ' '
#df2 = df[df['Beneficiario'].str.startswith('CUOTA ')]

# Carga la hoja de cálculo con los pagos
df = read_excel(excel_workbook , sheet_name=excel_worksheet)

# Elimina registros que no fueron seleccionados para enviar
df.dropna(subset=['Generar'], inplace=True)

# Convierte columna 'Nro. Recibo' en enteros
df['Nro. Recibo'] = df['Nro. Recibo'].astype(int)

# Ordena los pagos recibidos por beneficiario y recibo
df.sort_values(['Beneficiario', 'Nro. Recibo'], inplace=True)

# Toma como fecha para los resúmenes de cuotas, el 1° del mes
fecha_referencia = datetime(datetime.today().year, datetime.today().month, 1)

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
filename = os.path.join(attach_path, f'Recibos {datetime.now():%Y-%m-%d %H.%M}.log')
logfile = open(filename, 'w', encoding='utf-8')

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
    recibos = [recibo_fmt.format(recibo=idx) for idx in sub_lista]
    recibos = re.sub(r'(.*), (.*)', r'\1 y \2', ', '.join(recibos))
    if isnull(email_o_celular):
        logfile.write(msg_no_email_or_phone.format(s='s' if len(sub_lista) > 1 else '',
                                                   filename=recibos,
                                                   beneficiario=beneficiario))
    else:
        enviados = [idx for idx in sub_lista if notnull(df.loc[idx, 'Enviado'])]
        no_enviados = [idx for idx in sub_lista if idx not in enviados]
        if len(enviados) > 0:
            recibos = [recibo_fmt.format(recibo=idx) for idx in enviados]
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
with open(filename, 'r', encoding='utf-8') as logfile:
    for cnt, line in enumerate(logfile):
        print(trunca_texto(line.strip(), max_width=78))

s = 's' if df.shape[0] != 1 else ''
print(f"\n*** Proceso terminado: {mensajes_enviados} de {df.shape[0]} recibo{s} enviado{s}\n")

if whatsapp_loaded:
    unloadWhatsAppWeb()
