
# (basado en el código encontrado en:
#    https://simply-python.com/2018/11/15/convert-pdf-pages-to-jpeg-with-python/
#    https://stackoverflow.com/questions/46184239/python-extract-a-page-from-a-pdf-as-a-jpeg
# )

import sys, os

# Define textos
input_path  = 'C:/Users/MColosso/Documents/GyG Recibos_temp/'
filename    = 'GyG Recibo_{recibo:05d}.pdf'
output_path = 'C:/Users/MColosso/Dropbox/Vigilancia/Temporales/'
output_file = '{Beneficiario}, {Recibo}{Ext}'
output_ext  = '.jpg'   # .png   PNG
filetype    = 'JPEG'   # .jpg   JPEG

excel_workbook  = '1.1. GyG Recibos.xlsm'
excel_worksheet = 'Vigilancia'

crop_image  = True


#
#   P R O C E S O
#

input_path  = input_path.replace('/', '\\')
output_path = output_path.replace('/', '\\')

print()
print( 'Convierte Recibos de Pago en imágenes para su posterior envío')
print(f'   Los recibos son tomados de "{input_path}"')
print(f'   y generados en "{output_path}"\n')

if len(sys.argv) <= 1:
    print(f'Uso: python {sys.argv[0]} <Nro. de recibo> ...')
    sys.exit(1)

print('Cargando librerías...')
from pdf2image import convert_from_path
from pandas import read_excel
if crop_image:
    from PIL import Image

# Carga la hoja de cálculo con los pagos
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df = read_excel(excel_workbook, sheet_name=excel_worksheet)

for index, recibo in enumerate(sys.argv):

    if index == 0:
        continue # descarta el nombre del programa (primer argumento)

    # Verifica si el número de recibo indicado es numérico
    if not recibo.isdigit():
        print('\n*** Recibo "{recibo}" no es numérico')
        continue # con el siguiente recibo

    # Verifica si el número de recibo indicado existe
    nro_recibo = int(recibo)
    try:
#        r = df[df['Nro. Recibo'] == nro_recibo]   <--- cada columna en 'r' es una serie de Pandas
        r = df.iloc[nro_recibo - 1]
    except:
        print(f'\n*** Recibo {nro_recibo:05d} no es un número de recibo válido')
        continue  # con el siguiente recibo

    beneficiario = r['Beneficiario'].replace('Familia ', 'Fam. ')
    codRecibo = '{:05d}'.format(int(r['Nro. Recibo']))
    fecRecibo = r['Fecha'].strftime('%d/%m/%Y')

    respuesta = input("\n¿Recibo " + codRecibo + " del " + fecRecibo + \
                          " a nombre de " + str(r['Beneficiario']) + " [s/n]? ")
    while respuesta.upper() not in ['S', 'N', 'SÍ', 'SI', 'NO']:
        respuesta = input('   Indique Sí o No [s/n] ')
    if respuesta[0].upper() == 'N':
        continue

    try:
        pages = convert_from_path(input_path + filename.format(recibo=nro_recibo))
    except:
        error_msg  = str(sys.exc_info()[1])
        error_msg  = error_msg.replace('\\', '/')
        print(f'*** Error convirtiendo {filename.format(recibo=nro_recibo)}: {error_msg}')
        continue

    output_filename = output_file.format(Beneficiario=beneficiario, Recibo=codRecibo, Ext=output_ext)
    try:
        pages[0].save(output_path + output_filename, filetype)
    except:
        error_msg  = str(sys.exc_info()[1])
        error_msg  = error_msg.replace('\\', '/')
        print(f'*** Error generando {output_filename}: {error_msg}')
        continue

    if crop_image:
        img_recibo = Image.open(output_path + output_filename)
        img_recibo = img_recibo.resize(size=(712, 363), box=(108, 84, 1570, 825), resample=Image.BILINEAR)
        img_recibo.save(output_path + output_filename)

    print(f'   -> Generado "{output_filename}"')

sys.exit(0)
