# GyG Mantenimiento
#
# Borra los archivos anteriores a la fecha indicada, los cuales pueden ser
# reproducidos nuevamente, a fin de mantener espacio libre en disco


"""
    POR HACER
    -   

    HISTORICO
    -   Incluir el mantenimiento de la carpeta Graficas, eliminando archivos con más de
        tres meses (09/06/2020)
    -   Versión inicial (17/04/2020)


"""

print('Cargando librerías...')
import GyG_constantes
import PySimpleGUI as sg

from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')


# Constantes

FORMATO_MES =      '%b. %Y'
ANCHURA_TEXTOS =   30
OFFSET_IZQUIERDO =  3
CALENDAR_ICON =    os.path.join(GyG_constantes.rec_imágenes, '62925-spiral-calendar-icon.png')
# CALENDAR_ICON =    os.path.join(GyG_constantes.rec_imágenes, 'icons8-calendar-27-64.png')

#
# RUTINAS DE UTILIDAD
#

def maantenimiento(título, meses, carpeta):
    print(f" - Manteniendo {título}... ", end="")

    anterior_a = fecha_de_referencia - relativedelta(months=meses)
    archivos_eliminados = 0
    for path, dirs, files in os.walk(carpeta):
        for file in files:
            if file.startswith('.'):
                continue
            fecha_archivo = datetime.fromtimestamp(os.path.getmtime(os.path.join(path, file)))
            # print(f"{file:50} {fecha_archivo.strftime('%d-%m-%Y')}" + \
            #       f"{' <-- Eliminar' if fecha_archivo < anterior_a else ''}")
            if fecha_archivo < anterior_a:
                os.remove(os.path.join(path, file))
                archivos_eliminados += 1

    print(f"{archivos_eliminados} archivos eliminados")


#
# PROCESO
#

print("Generando layout ...")

ahora = datetime.now()
fecha_de_referencia = datetime(ahora.year, ahora.month, 1)

# Opciones: <Opción>, <Nro de meses>, <Carpeta>
opciones =  [['Recibos de pago',       '1',  GyG_constantes.ruta_recibos],
             ['Análiis de pago',       '3',  GyG_constantes.ruta_analisis_de_pagos],
             ['Cartelera virtual',     '3',  GyG_constantes.ruta_cartelera_virtual],
             ['Relación de ingresos',  '3',  GyG_constantes.ruta_relacion_ingresos],
             ['Resúmenes de pago',    '24',  GyG_constantes.ruta_resumenes],
             ['Saldos pendientes',     '3',  GyG_constantes.ruta_saldos_pendientes],
             ['Gráficas',              '3',  GyG_constantes.ruta_graficas],
             ['Cambios de categoría',  '3',  GyG_constantes.ruta_cambios_de_categoría],

             ['Archivos temporales',   '3',  GyG_constantes.ruta_temporales],
            ]

layout = [  [sg.Text("Mes de referencia:"),
             sg.InputText(key='_fecha_referencia_', size=(10,1),
                          default_text=fecha_de_referencia.strftime(FORMATO_MES)),
             sg.CalendarButton(button_text='Seleccione otra fecha',
                               image_filename=CALENDAR_ICON, image_size=(16, 16), image_subsample=8,
                               tooltip='Seleccione otra fecha',
                               default_date_m_d_y=(fecha_de_referencia.month, fecha_de_referencia.day, fecha_de_referencia.year),
                               target='_fecha_referencia_', format=FORMATO_MES,
                               close_when_date_chosen=True)],

            [sg.Text("")],
            [sg.Text("Indique la cantidad de meses remanentes luego del mantenimiento:")],

            # Punto de inserción: índice = 3

            [sg.Text('_'*60)],
            [sg.Button('Realiza mantenimientos'), sg.Button('Finaliza')],
         ]

# Inserta en el layout las distintas opciones
posición = 3
chk_opcion = '_chk_opcion_'
val_opcion = '_val_opcion_'
lista_cantidad_meses = [str(i) for i in range(1, 25)]
for idx, opcion in enumerate(opciones):
    layout.insert(posición + idx,
            [   sg.Text(" "*OFFSET_IZQUIERDO),
                sg.Checkbox(opcion[0], key=f"{chk_opcion}{idx}", size=(ANCHURA_TEXTOS, 1), default=True),
                sg.Spin(key=f"{val_opcion}{idx}", values=lista_cantidad_meses, initial_value=opcion[1], size=(3, 1)),
                sg.Text("meses")
            ])

# Create the Window
window = sg.Window('GyG Mantenimiento', layout)

# Event Loop to process "events"
#    while True:             
event, values = window.read()
if event in (None, 'Finaliza'):
    pass    # break
elif event == 'Realiza mantenimientos':
    fecha_de_referencia = datetime.strptime(values['_fecha_referencia_'], FORMATO_MES)

    # from pprint import pprint
    # pprint(values)

    for opcion in values.keys():
        if opcion.startswith(chk_opcion):
            if values[opcion]:
                indice = int(opcion[len(chk_opcion):])
#                print(f"DEBUG: {indice=}, {val_opcion+str(indice)=}, {values[val_opcion+str(indice)]=}")
                meses = int(values[val_opcion+str(indice)])
                maantenimiento(opciones[indice][0], meses, opciones[indice][2])

else:
    print(f'*** ERROR: Se recibió el evento "{event}", no esperado')
