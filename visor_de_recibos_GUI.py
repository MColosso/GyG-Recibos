# GYG VISOR DE RECIBOS DE PAGO
# 
# Muestra el recibo de pago seleccionado

"""
    POR HACER
    -   Revisar error en Windows:

        1.6 GyG Recibos - Visor de Recibos de Pago
        ------------------------------------------

        C:/Users/MColosso/Dropbox/GyG Recibos>python ./visor_de_recibos_GUI.py
        Cargando hoja de cálculo "1.1. GyG Recibos.xlsm"...
        Cargando la lista de recibos...
        * Warning... no image specified in Image Element! *

        *ERROR laying out form.... Image Element has no image specified*
        C:/Users/MColosso/AppData/Local/Programs/Python/Python38-32/lib/site-packages/PySimpleGUI/PySimpleGUI.py:2848: 
        UserWarning: You cannot Update element with key = -IMAGE- until the window has been Read or Finalized
          warnings.warn('You cannot Update element with key = {} until the window has been Read or Finalized'.format(self.Key), 
          UserWarning)

        Presione una tecla para continuar . . .


    HISTORICO
    -   Se agregó código para habilitar el botón «A Temporales» al seleccionar un recibo
        y se ajustó el despliegue de resultados de la copia (08/11/2020)
    -   Agregado botón «A Temporales» para copiar el recibo seleccionado en la carpeta
        "Temporales" (se utilizó un algoritmo similar al empleado en 'copia_recibos.py')
        (06/11/2020)
    -   Se intercambia el orden la imagen del recibo y su ubicación (04/11/2020)
    -   Versión inicial
"""


# print('Cargando librerías...')
import os.path
import re
import PySimpleGUI as sg
from pandas import read_excel, to_datetime
from shutil import copyfile
import GyG_constantes
from GyG_utilitarios import edita_beneficiario
import sys

# Define algunas constantes
EXCEL_WORKBOOK      = GyG_constantes.pagos_wb_estandar      # '1.1. GyG Recibos.xlsm'
EXCEL_WS_VIGILANCIA = GyG_constantes.pagos_ws_vigilancia
FILE_FOLDER         = GyG_constantes.ruta_recibos
INPUT_PATH          = GyG_constantes.ruta_recibos           # '../GyG Archivos/Recibos de Pago'
OUTPUT_PATH         = GyG_constantes.ruta_imagen_recibos    # 'C:/Users/MColosso/Dropbox/Vigilancia/Temporales/'


#
# RUTINAS DE UTILIDAD
#

def edita_categoría(str_categoría):
    """ Ajusta la categoría reduciendo su longitud y encerrándola entre corchetes """
    return '[' + str_categoría[:3] + ']'

def copia_recibo():
    input_path      = os.path.normpath(INPUT_PATH)
    output_path     = os.path.normpath(OUTPUT_PATH)

    try:
        nro_recibo = int(values["-FILE LIST-"][0][:GyG_constantes.long_num_archivo])
    except IndexError:  # No hay recibo seleccionado
        return
    r = df_recibos.iloc[nro_recibo - 1]

    beneficiario    = r['Beneficiario'].replace('Familia ', 'Fam. ')
    codRecibo       = f'{nro_recibo:05d}'
    input_filename  = GyG_constantes.img_recibo.format(recibo=nro_recibo)
    output_ext      = os.path.splitext(input_filename)[1]
    output_filename = f'{beneficiario}, {codRecibo}{output_ext}'
    file_to_copy    = os.path.join(input_path, input_filename)
    filename = os.path.join(
        FILE_FOLDER,
        GyG_constantes.img_recibo.format(recibo=nro_recibo)
    )
    if sys.platform.startswith('win'):
        filename  = filename.replace('/', '\\')

    try:
        copyfile(file_to_copy, os.path.join(output_path, output_filename))
        window["-TOUT-"].update(f'{filename} - COPIADO COMO "{output_filename}"')
        print(f' -> Copiado "{input_filename}" como "{output_filename}"')
        return True
    except:
        error_msg  = str(sys.exc_info()[1])
        if sys.platform.startswith('win'):
            error_msg  = error_msg.replace('\\', '/')
        window["-TOUT-"].update(f"{filename} - ERROR DE COPIA")
        print(f'*** Error copiando "{output_filename}": {error_msg}')
        return False


#
# PROCESO
#

print('Cargando hoja de cálculo "{filename}"...'.format(filename=EXCEL_WORKBOOK))
df_recibos = read_excel(EXCEL_WORKBOOK, sheet_name=EXCEL_WS_VIGILANCIA,
                        usecols=['Nro. Recibo', 'Beneficiario', 'Fecha', 'Categoría'])

print('Cargando la lista de recibos...')
file_list = []
for recibo in os.listdir(FILE_FOLDER):
    if recibo.endswith('.png'):
        nro_recibo = int(re.findall(r'.*_(.*)\.', recibo)[0])
        r = df_recibos.loc[df_recibos['Nro. Recibo'] == nro_recibo]
        n_recibo = f"{r['Nro. Recibo'].values[0]:0{GyG_constantes.long_num_archivo}d}"
        vecino = edita_beneficiario(r['Beneficiario'].values[0])
        categoría = r['Categoría'].values[0][:3]
        fecha = to_datetime(r['Fecha'].values[0]).strftime('%d/%m/%y')
        file_list.append(f"{n_recibo} - [{categoría}. {fecha}] {vecino}")

file_list.sort()

# First the window layout in 2 columns

file_list_column = [
    [
        sg.Listbox(
            values=file_list, enable_events=True, size=(40, 30), key="-FILE LIST-"
        )
    ],
    [sg.Button('Copia en Temporales', key="-COPIA RECIBO-", disabled=True), sg.Button('Finaliza')],
]

# For now will only show the name of the file that was chosen
gyg_logo = os.path.join(os.getcwd(), 'GyG logo', 'GyG logo.png')
image_viewer_column = [
    [sg.Image(filename=gyg_logo, key="-IMAGE-", size=(712, 0))],
    [sg.Text("<--  Seleccione un recibo a visualizar de la lista de la izquierda:", size=(100, 1), key="-TOUT-")],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(image_viewer_column),
    ]
]

window = sg.Window("GyG Visor de Recibos", layout)

while True:
    event, values = window.read()
    if event in ["Finaliza", None]:
        break
    elif event in ["-COPIA RECIBO-"]:       # Anteriormente ["A Temporales"]; corregido al agregar Key al botón
        copia_recibo()
    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        window['-COPIA RECIBO-'].update(disabled=False)
        num_recibo = int(values["-FILE LIST-"][0][:GyG_constantes.long_num_archivo])
        filename = os.path.join(
            FILE_FOLDER,
            GyG_constantes.img_recibo.format(recibo=num_recibo)
        )
        if sys.platform.startswith('win'):
            filename  = filename.replace('/', '\\')
        try:
            window["-TOUT-"].update(filename)
            window["-IMAGE-"].update(filename=filename)
        except:
            pass

window.close()
