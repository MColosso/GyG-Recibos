# COPIA RECIBOS DE PAGO
#
# Copia los recibos de pago indicados a la carpeta 'Vigilancia/Temporales' para ser enviados posteriormente
# vía WhatsApp

"""
    POR HACER
    -   

    HISTORICO
    -   Validar si el recibo existe antes de copiarlo; en caso contrario, generar mensaje de error (20/04/2020)
    -   Versión inicial (25/10/2019)

"""


import GyG_constantes
from GyG_utilitarios import input_si_no, genera_recibo
import sys, os

# Define textos
input_path  = GyG_constantes.ruta_recibos              # '../GyG Archivos/Recibos de Pago'
filename    = GyG_constantes.img_recibo                # 'GyG Recibo_{recibo:05d}.img'
output_path = GyG_constantes.ruta_imagen_recibos       # 'C:/Users/MColosso/Dropbox/Vigilancia/Temporales/'
output_file = '{Beneficiario}, {Recibo}{Ext}'

excel_workbook  = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_worksheet = GyG_constantes.pagos_ws_vigilancia   # 'Vigilancia'


#
#   P R O C E S O
#

# Selecciona la carpeta de destino
if len(sys.argv) > 1:
    output_path = sys.argv[1]

input_path  = os.path.normpath(input_path)
output_path = os.path.normpath(output_path)

print()
print( '   COPIA RECIBOS DE PAGO PARA SU POSTERIOR ENVIO\n')
print(f'   Los recibos son tomados de "{input_path}"')
print(f'   y copiados en "{output_path}"\n')

print('Cargando librerías...')
from pandas import read_excel
from shutil import copyfile

# Carga la hoja de cálculo con los pagos
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df = read_excel(excel_workbook, sheet_name=excel_worksheet)

while True:

    recibo = input('\nIndique el nro. de recibo a copiar (en blanco para terminar): ')

    if len(recibo) == 0:
        break   # Termina la ejecución

    # Verifica si el número de recibo indicado es numérico
    if not recibo.isdigit():
        print('*** Recibo "{recibo}" no es numérico')
        continue # con el siguiente recibo

    # Verifica si el número de recibo indicado existe
    nro_recibo = int(recibo)
    try:
#        r = df[df['Nro. Recibo'] == nro_recibo]   <--- cada columna en 'r' es una serie de Pandas
        r = df.iloc[nro_recibo - 1]
    except:
        print(f'*** "{nro_recibo:05d}" no es un número de recibo válido')
        continue  # con el siguiente recibo

    beneficiario = r['Beneficiario'].replace('Familia ', 'Fam. ')
    codRecibo = '{:05d}'.format(int(r['Nro. Recibo']))
    fecRecibo = r['Fecha'].strftime('%d/%m/%Y')

    respuesta = input("CONFIRME: ¿Recibo " + codRecibo + " del " + fecRecibo + \
                          ", " + beneficiario + " [s/n]? ")
    while respuesta.upper() not in ['S', 'N', 'SÍ', 'SI', 'NO']:
        respuesta = input('   Indique Sí o No [s/n] ')
    if respuesta[0].upper() == 'N':
        continue

    input_filename  = filename.format(recibo=nro_recibo)
    output_ext      = os.path.splitext(input_filename)[1]
    output_filename = output_file.format(Beneficiario=beneficiario, Recibo=codRecibo, Ext=output_ext)
    file_to_copy    = os.path.join(input_path, input_filename)

    try_copy = True
    if not os.path.exists(file_to_copy):
        print(f"*** Error: El recibo {nro_recibo:05d} probablemente fue eliminado previamente.")
        if input_si_no('¿Se regenera e intenta nuevamente?', 'Sí'):
            recibo_a_generar = {
                'Nro. Recibo':  r['Nro. Recibo'],
                'Fecha':        r['Fecha'],
                'Beneficiario': r['Beneficiario'],
                'Dirección':    r['Dirección'],
                'Monto':        r['Monto'],
                'Concepto':     r['Concepto'],
                'Categoría':    r['Categoría']
            }
            genera_recibo(recibo_a_generar)
        else:
            try_copy = False

    if try_copy:
        try:
            copyfile(file_to_copy, os.path.join(output_path, output_filename))
        except:
            error_msg  = str(sys.exc_info()[1])
            if sys.platform.startswith('win'):
                error_msg  = error_msg.replace('\\', '/')
            print(f'*** Error copiando {output_filename}: {error_msg}')
            continue

        print(f'   -> Copiado "{output_filename}"')

print()
