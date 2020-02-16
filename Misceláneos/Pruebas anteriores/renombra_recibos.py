# GyG RENOMBRA RECIBOS DE PAGO
#
# Al convertir el archivo .pdf generado por la combinación de la plantilla de
# los recibos y la base de datos en Excel, el grupo seleccionado comienza su
# numeración en "uno" (1), debiendo ser renombrados a su numeración definitiva.
#
# Este proceso permite generar sólo los recibos necesarios, y no todos,
# acelerando la emisión de los recibos.

import GyG_constantes
import os
import sys

from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBoxHorizontal


#
# TEXTOS
#

attach_name    = GyG_constantes.pdf_recibo            # 'GyG Recibo_{recibo:05d}.pdf'
inicio_archivo = GyG_constantes.inicio_num_archivo    # Posición de inicio de la numeración del recibo en el 
                                                      # nombre de archivo
long_archivo   = GyG_constantes.long_num_archivo      # Longitud de la numeración del recibo en el nombre de
                                                      # archivo

patrón_recibo  = GyG_constantes.patrón_recibo         # 'Recibo N° 00001'
prueba_long    = GyG_constantes.long_prueba_patrón    # Cantidad de caracteres a probar para determinar si se trata
                                                      # de un número de recibo
inicio_recibo  = GyG_constantes.inicio_num_recibo     # Posición de inicio de la numeración del recibo
long_recibo    = GyG_constantes.long_num_recibo       # Longitud de la numeración del recibo


#
# PROCEDIMIENTOS
#

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

pdf_file = sys.argv[1]
default_path = sys.argv[2]
temp_path = default_path + '\\' + sys.argv[3]

print()
receipt_number = get_first_receipt(pdf_file)
prompt = '*** ¿Generar recibos desde? [' + receipt_number + '] '
offset = input(prompt)

try:
    if len(offset) == 0:
        offset = int(receipt_number)
    else:
        offset = int(offset)
except:
    error_msg = str(sys.exc_info()[1])
    print('ERROR:', error_msg)
    print('')
    exit(1)

for filename in os.listdir(temp_path):
    file_starts_with_recibo = filename.startswith(attach_name[:attach_name.find('{')])
    file_ends_with_pdf = filename.endswith(attach_name[-4:])
    if file_starts_with_recibo and file_ends_with_pdf:
        nro_recibo = int(filename[inicio_archivo : inicio_archivo + long_archivo])
        new_filename = attach_name.format(recibo=nro_recibo + offset - 1)
        os.replace(temp_path + '\\' + filename,
                   default_path + '\\' + new_filename)

print('')

exit(0)
