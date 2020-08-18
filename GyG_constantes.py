#
#  GyG CONSTANTES
#

from re import search
from datetime import datetime
import os


# HOJAS DE CÁLCULO
pagos_wb_estandar           = '1.1. GyG Recibos.xlsm'
pagos_wb_historico          = '1.1. GyG Recibos - SOLO CONSULTA.xlsm'

pagos_ws_vigilancia         = 'Vigilancia'
pagos_ws_resumen            = 'RESUMEN VIGILANCIA'
pagos_ws_resumen_reordenado = 'R.VIGILANCIA (reordenado)'
pagos_ws_saldos             = 'Saldos'
pagos_ws_cuotas             = 'CUOTAS'
# pagos_ws_cobranzas          = 'Cobranzas'
pagos_ws_otros              = 'RESUMEN OTROS'
pagos_ws_vecinos            = 'Vecinos'

resumen_workbook            = '4.2. GyG Resúmenes - Control resúmenes.xlsx'
resumen_worksheet           = 'A solicitud'


# ARCHIVOS
long_num_archivo            =  5   # Longitud de la numeración del recibo en el nombre de
                                   # archivo
long_num_recibo             =  5   # Longitud de la numeración en el recibo   

pdf_GyG_Recibos             = '1.3. GyG Recibos.pdf'

patrón_pdfs                 = 'GyG Recibo_%0Xd.pdf'.replace('X', str(long_num_archivo))
                              # a ser usado en la separación de páginas del archivo pdf_GyG_Recibos
                              # mediante 'pdftk' en separa_paginas.py

pdf_recibo                  = 'GyG Recibo_{recibo:0Xd}.pdf'.replace('X', str(long_num_archivo))
img_recibo                  = 'GyG Recibo_{recibo:0Xd}.png'.replace('X', str(long_num_archivo))

# img_sello_GyG               = os.path.join('./recursos/imagenes', 'GyG_sello_negro.png')
img_sello_GyG               = os.path.join('./recursos/imagenes', 'GyG_sello_humedo.png')

pdf_resumen                 = 'GyG Resumen {resumen:0Xd}.pdf'.replace('X', str(long_num_archivo))
                              # pdf_file = 'GyG Resumen {}-{}T_{:03d}.pdf' en genera_ y envia_resumenes.py

txt_analisis_de_pago        = "GyG Analisis de Pagos {:%Y-%m (%B)}.txt"
txt_cartelera_virtual       = "GyG Cartelera Virtual {:%Y-%m (%B)}.txt"
txt_cambios_de_categoría    = "GyG Cambios de Categoría {:%Y-%m (%B)}.txt"

recibo_fmt                  = "{recibo:0Xd}".replace('X', str(long_num_archivo))

txt_relacion_ingresos       = "GyG Relación de Ingresos {:%Y-%m (%b)}"


# RUTAS
ruta_gyg_recibos            = '.'
                              # 'C:/Users/MColosso/Dropbox/GyG Recibos/'
                              # Carpeta de la aplicación

ruta_recibos                = '../GyG Archivos/Recibos de Pago'
                              # 'C:/Users/MColosso/Google Drive/GyG Recibos/Recibos/'
                              # Carpeta de recibos de pago

ruta_imagen_recibos         = '../GyG Archivos/Temporales/'
                              # 'C:/Users/MColosso/Dropbox/Vigilancia/Temporales/'
                              # Carpeta de recibos de pago como imágenes (.jpg | .png)

ruta_temporales             = '../GyG Archivos/Temporales/'
                              # 'C:/Users/MColosso/Dropbox/Vigilancia/Temporales/'
                              # Carpeta de recibos de pago como imágenes (.jpg | .png)

ruta_resumenes              = '../GyG Archivos/Resúmenes'
                              # 'C:/Users/MColosso/Google Drive/GyG Recibos/Resúmenes/'
                              # Carpeta de resúmenes de pagos

ruta_analisis_de_pagos      = '../GyG Archivos/Análisis de Pago'
                              # "C:/Users/MColosso/Google Drive/GyG Recibos/Análisis de Pago/"
                              # Carpeta de análisis de pagos

ruta_saldos_pendientes      = '../GyG Archivos/Saldos Pendientes'
                              # Carpeta de saldos de pagos

ruta_cartelera_virtual      = '../GyG Archivos/Cartelera Virtual'
                              # "C:/Users/MColosso/Google Drive/GyG Recibos/Cartelera Virtual/"
                              # Carpeta de carteleras_virtuales

ruta_graficas               = '../GyG Archivos/Graficas'
                              # Carpeta de carteleras_virtuales

ruta_relacion_ingresos       = '../GyG Archivos/Relación de Ingresos'
                              # "C:/Users/MColosso/Google Drive/GyG Recibos/Graficas/"
                              # Carpeta de carteleras_virtuales

ruta_cambios_de_categoría    = '../GyG Archivos/Otros'
                              # "C:/Users/MColosso/Google Drive/GyG Recibos/Graficas/"
                              # Carpeta de carteleras_virtuales
#recibos_temp                = 'C:/Users/MColosso/Documents/GyG Recibos_temp/'
                              # Carpeta en la cual se generan temporalmente (y desde el cual se envían)
                              # los recibos de pago

rec_imágenes                = './recursos/imagenes'
rec_fuentes                 = './recursos/fuentes'
rec_plantillas              = './recursos/plantillas'

plantilla_recibos           = os.path.join(rec_plantillas, 'plantilla_recibos.png')


# OTRAS CONSTANTES
fecha_de_corte              = datetime(2019, 9, 1)     # Fecha en la cual cambió la modalidad de cobro a un
                                                       # monto fijo en dólares

preferred_encoding          = 'cp1252'                 # Codificación de caracteres por defecto en Windows
                                                       # En macOS X, la codificación por defecto es 'UTF-8'
Apple_encoding              = 'UTF-8'
Windows_encoding            = 'cp1252'

#inicio_num_archivo          = pdf_recibo.find('{')    # Posición de inicio de la numeración del recibo en el 
                                                       # nombre de archivo
#patrón_recibo               = 'Recibo N° 00001'
#inicio_num_recibo           = search(r'\d', patrón_recibo).start() + 1
                                                       # Posición de inicio de la numeración del recibo
#long_prueba_patrón          = inicio_num_recibo - 2   # Cantidad de caracteres a probar para determinar si se trata
                                                       # de un número de recibo

sellos                       = ['ANULADO', 'SOLVENTE', 'REVERSADO', 'ANTICIPO']
                                                       # Sellos a imprimir sobre el recibo de pago, según la
                                                       # categoría de pago