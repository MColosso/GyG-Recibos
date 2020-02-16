# GyG COPIA RECIBOS DE PAGO
#
# Copia los recibos de pago seleccionados para ser enviados a la carpeta definitiva.
#
# Esto permite transferir a Internet sólo los recibos que van a ser enviados, y no
# todos, acelerando el proceso.

"""
  PENDIENTES POR REVISAR

    - ¿Qué sucede con los recibos no marcados para ser enviados? ¿Se afecta la numeración
      de los recibos...?
      => 'renombra_recibos.py' deja en <temp_folder> todos los recibos, numerándolos a partir
         del número indicado ('offset') -> La numeración no se afecta
         'copia_recibos.py' (este programa) sólo copia en <output_folder> aquellos recibos
         marcados para ser emitidos
         Los archivos comprimidos DEBEN ser generados desde la carpeta <temp_folder>

  HISTÓRICO
  
    - CORREGIR: Cambiar referencias de "df['Archivo']" a "df['Nro. Recibo']" para evitar problemas
      de numerarión manual (corregido 28/06/2019)
    - CORREGIR: En la hoja de cálculo '1.1. GyG Recibos.xlsm' se cambió el título de la columna 'B'
      de 'Archivo' a 'Generar' (corregido 08/07/2019)

"""

print('Cargando librerías...')
import GyG_constantes
from pandas import read_excel
import os
import sys

# Define textos
attach_name     = GyG_constantes.pdf_recibo            # "GyG Recibo_{recibo:05d}.pdf"
excel_workbook  = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_worksheet = GyG_constantes.pagos_ws_vigilancia   # 'Vigilancia'

# Define la ubicación de las carpetas temporal (en el disco duro local), y la carpeta
# definitiva (la cual será sincronizada con una carpeta en la nube)
# El formato de las rutas es al 'estilo Windows' (backslash) y no Unix (slash)
temp_folder     = GyG_constantes.recibos_temp          # 'C:/Users/MColosso/Documents/GyG Recibos_temp/'
output_folder   = GyG_constantes.ruta_recibos          # 'C:/Users/MColosso/Google Drive/GyG Recibos/Recibos/'


#
# PROCESO
#

# Lee la hoja de cálculo de recibos
recibos = read_excel(excel_workbook, sheet_name=excel_worksheet)

# Elimina registros que no fueron seleccionados para enviar
recibos.dropna(subset=['Generar'], inplace=True)

# Convierte columna 'Nro. Recibo' en enteros
recibos['Nro. Recibo'] = recibos['Nro. Recibo'].astype(int)


# Copia cada recibo de pago seleccionado
print('Copiando archivos', end='')
archivos_copiados = 0
for archivo in recibos['Nro. Recibo']:
    r = os.system(' '.join(['copy',
                            '"' + temp_folder + attach_name.format(recibo=archivo) + '"',
                            '"' + output_folder + '"',
                            '> dummy.txt']))
    if r == 0:
        archivos_copiados += 1
    print('.', end='')   # Imprime un punto en la pantalla por cada mensaje
    sys.stdout.flush()   # Flush output to the screen

print('\n')
print('*** Proceso terminado: {} de {} archivo(s) copiado(s)\n'.format(archivos_copiados,
                                                                       recibos.shape[0]))
