# GyG Recibos - Separa páginas
#
# Genera un recibo por página a partir del archivo "1.3. GyG Recibos.pdf"
# Los recibos son extraidos en una carpeta temporal donde son renombrados utilizando el
# mismo offset con el cual fueron generados en Word


# Selecciona las librerías a utilizar
print('Cargando librerías...')
import os
import sys
import subprocess

from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBoxHorizontal

from pandas import read_excel


#
# TEXTOS
#

# Define la ubicación de las carpetas temporal (en el disco duro local), y la carpeta
# definitiva (la cual será sincronizada con una carpeta en la nube)
# El formato de las rutas es al 'estilo Windows' (backslash) y no Unix (slash)

temp_folder    = 'C:\\Users\\MColosso\\Documents\\GyG Recibos_temp\\'
output_folder  = 'C:\\Users\\MColosso\\Google Drive\\GyG Recibos\\Recibos\\'
aux_folder     = 'temp\\'
pdf_file       = '1.3. GyG Recibos.pdf'
patrón_pdfs    = 'GyG Recibo_%05d.pdf'

temp_path      = temp_folder + aux_folder
output_path    = temp_path + patrón_pdfs


attach_name    = 'GyG Recibo_{seq:05d}.pdf'
inicio_archivo = 11   # Posición de inicio de la numeración del recibo en el 
                      # nombre de archivo
long_archivo   =  5   # Longitud de la numeración del recibo en el nombre de
                      # archivo

patrón_recibo  = 'Recibo N° 00001'
prueba_long    =  9   # Cantidad de caracteres a probar para determinar si se trata
                      # de un número de recibo
inicio_recibo  = 11   # Posición de inicio de la numeración del recibo
long_recibo    =  5   # Longitud de la numeración del recibo


excel_workbook  = '1.1. GyG Recibos.xlsm'
excel_worksheet = 'Vigilancia'


#
# PROCEDIMIENTOS
#

def	renombra_recibos():
    # Al convertir el archivo .pdf generado por la combinación de la plantilla de
    # los recibos y la base de datos en Excel, el grupo seleccionado comienza su
    # numeración en "uno" (1), debiendo ser renombrados a su numeración definitiva.
    # Este proceso permite generar sólo los recibos necesarios, y no todos,
    # acelerando la emisión de los recibos.
    for filename in os.listdir(temp_path):
        file_starts_with_recibo = filename.startswith(attach_name[:attach_name.find('{')])
        file_ends_with_pdf = filename.endswith(attach_name[-4:])
        if file_starts_with_recibo and file_ends_with_pdf:
            nro_recibo = int(filename[inicio_archivo : inicio_archivo + long_archivo])
            new_filename = attach_name.format(seq=nro_recibo + offset - 1)
            os.replace(temp_path + filename,
                       temp_folder + new_filename)


def copia_recibos_de_pago():
    # Copia los recibos de pago seleccionados para ser enviados a la carpeta definitiva.
    #
    # Esto permite transferir a Internet sólo los recibos que van a ser enviados, y no
    # todos, acelerando el proceso.

    # Lee la hoja de cálculo de recibos
    recibos = read_excel(excel_workbook, sheet_name=excel_worksheet)

    # Elimina registros que no fueron seleccionados para enviar
    recibos.dropna(subset=['Archivo'], inplace=True)

    # Convierte columna 'Archivo' en enteros
    recibos['Archivo'] = recibos['Archivo'].astype(int)

    # Copia cada recibo de pago seleccionado
    print('Copiando archivos', end='')
    archivos_copiados = 0
    for archivo in recibos['Archivo']:
        r = os.system(' '.join(['copy',
                                '"' + temp_folder + attach_name.format(seq=archivo) + '"',
                                '"' + output_folder + '"',
                                '> dummy.txt']))
        if r == 0:
            archivos_copiados += 1
        print('.', end='')   # Imprime un punto en la pantalla por cada mensaje
        sys.stdout.flush()   # Flush output to the screen

    print('\n')
    print(f'*** Proceso terminado: {archivos_copiados} de {recibos.shape[0]} archivo(s) copiado(s)\n')


def get_first_receipt(pdf_file):
    # Code addapted from https://media.readthedocs.org/pdf/pdfminer-docs/latest/pdfminer-docs.pdf
    # page 13 - 2.3 Performing Layout Analysis
    document = open(pdf_file, 'rb')
    #Create resource manager
    rsrcmgr = PDFResourceManager()
    # Set parameters for analysis.
    laparams = LAParams()
    # Create a PDF page aggregator object.
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    recibo = ''
    for page in PDFPage.get_pages(document):
        interpreter.process_page(page)
        # receive the LTPage object for the page.
        layout = device.get_result()
        for element in layout:
            if isinstance(element, LTTextBoxHorizontal):
                texto = element.get_text()
                if texto[:prueba_long] == patrón_recibo[:prueba_long]:
                    receipt = texto[inicio_recibo : inicio_recibo + long_recibo]
                    break
        if receipt != '':
            break
    document.close()
    return receipt


#
#  PROCESO
#

# Separa los recibos en una carpeta temporal
print('Separando recibos...')
subprocess.run(["pdftk", pdf_file, "burst", "output", output_path, "encrypt_128bit", "allow", "Printing"])


# Obtiene el primer número de recibo
print('Obteniendo el número del primer recibo...')
receipt_number = get_first_receipt(pdf_file)
try:
    offset = int(receipt_number)
    Ok = True
except:
    error_msg = str(sys.exc_info()[1])
    print('ERROR:', error_msg)
    Ok = False

if Ok:
    # Renombra recibos de pago y los copia en su ubicación definitiva
    print(f'Renombrando recibos desde {receipt_number}...')
    renombra_recibos()
    copia_recibos_de_pago()
