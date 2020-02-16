# GyG RECONVERSIÓN MONETARIA
#
# Modifica los pagos en las hojas 'Vigilancia', 'Resumen Vigilancia' y 'Saldos'
# para llevarlos al nuevo cono monetario

"""
    POR HACER
    -   

    HISTORICO
    -   Incluir en la reconversión la hoja 'CUOTAS' (04/01/2020)
    -   Desplegar en la línea de status las acciones realizadas. Al inicio de cada rutina se
        actualiza el contenido de esta línea, pero la misma no se muestra sino hasta el final.
          . Se incluyó un Progress Bar para monitorear el avance (01/01/2020)
    -   Versión inicial GUI (28/12/2019)
    -   Versión inicial (26/12/2019)

"""

print("Cargando librerías ...")
from GyG_utilitarios import *
import PySimpleGUI as sg

import openpyxl
import GyG_constantes
import numbers
import warnings
warnings.simplefilter("ignore", category=UserWarning)
from pyparsing import Word, Regex, Literal, OneOrMore, ParseException
import os
import sys


# Gramática de expresiones aritméticas
operator  = Regex(r'(?<!--[\+\-\^\*/%])[\+\-]|[\^\*/%!]')
function  = Regex(r'[a-zA-Z_][a-zA-Z0-9_]*(?=([ \t]+)?\()')
variable  = Regex(r'[+-]?[a-zA-Z_][a-zA-Z0-9_]*(?!([ \t]+)?\()')
number    = Regex(r'[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?')       # Ej. -125,54e-5
lbrace    = Word('(')
rbrace    = Word(')')
linebreak = Word('\n')
skip      = Word(' \t')                                             # espacio o tabulador

lexOnly = operator | function | variable | number | lbrace \
    | rbrace | linebreak | skip
lexAllOnly = OneOrMore(lexOnly)

advertencia = """Revisar la conversión de cuotas de La Casita Encantada y
el Colegio El Trigal: Los redondeos aplicados en la hoja de
cálculo 'RESUMEN VIGILANCIA' pueden afectar el despliegue
de los resultados"""

default_fct      = 100000
default_cono     = 'VBS'

fct_reconversion = None
old_workbook     = None
new_workbook     = None


def convierte_Vigilancia():
    global workbook, step
#    event, values = window.read()
    window.Element('_status_line_').Update(value="Convirtiendo 'Vigilancia' ...")
    print("Convirtiendo 'Vigilancia' ...")

    worksheet = workbook['Vigilancia']

    for row in worksheet.rows:
        str_monto = row[4].value
        if str_monto in ['Monto', None]:
            continue
        row[4].value = convierte_monto(str_monto)

    progress_bar.UpdateBar(step := step + 1)


def convierte_RESUMEN_VIGILANCIA():
    global workbook, step
#    event, values = window.read()
    window.Element('_status_line_').Update(value="Convirtiendo 'RESUMEN VIGILANCIA' ...")
    print("Convirtiendo 'RESUMEN VIGILANCIA' ...")

    worksheet = workbook['RESUMEN VIGILANCIA']

    num_rows = 0
    ignora_filas = False

    row = list(worksheet.rows)[0]
    first_row = [col.value for col in row]
    first_col = first_row.index(2016)
    last_col = first_row.index('TOTAL')
    skip_conversion = False     # No aplica el análisis lexicográfico a las filas de Cuotas (final de la hoja)

    for row in worksheet.rows:
        num_rows += 1
        if num_rows == 1:
            continue
        elif row[0].value in [None, '']:
            skip_conversion = True
            continue
        for col in range(first_col, last_col):
            str_monto = row[col].value
            row[col].value = convierte_monto(str_monto, analisis_lexicografico=(not skip_conversion))

    progress_bar.UpdateBar(step := step + 1)


def convierte_Saldos():
    global workbook, step
#    event, values = window.read()
    window.Element('_status_line_').Update(value="Convirtiendo 'Saldos' ...")
    print("Convirtiendo 'Saldos' ...")

    worksheet = workbook['Saldos']

    columna_H = 7
    num_rows = 0

    for row in worksheet.rows:
        num_rows += 1
        if num_rows <= 3:   # Ignora las tres primeras lineas
            continue
        col = columna_H
        while col <= worksheet.max_column:
            row[col].value = convierte_monto(row[col].value)
            col += 5

    progress_bar.UpdateBar(step := step + 1)


def convierte_CUOTAS():
    global workbook, step
#    event, values = window.read()
    window.Element('_status_line_').Update(value="Convirtiendo 'CUOTAS' ...")
    print("Convirtiendo 'CUOTAS' ...")

    worksheet = workbook['CUOTAS']

    for row in worksheet.rows:
        moneda = row[2].value
        if moneda in ['Moneda', None]:
            continue
        if moneda == 'VEB':
            row[1].value = convierte_monto(row[1].value)    # 'Cantidad'
        elif moneda == 'USD':
            row[3].value = convierte_monto(row[3].value)    # 'Tasa Bs./US$'

    progress_bar.UpdateBar(step := step + 1)


def convierte_monto(str_monto, analisis_lexicografico=True):
    if isinstance(str_monto, numbers.Number):
        return f'={str_monto}/{fct_reconversion}'
    elif isinstance(str_monto, str) and str_monto[0] == '=':
        if analisis_lexicografico:
            return f"={aplica_factor_de_reconversion(str_monto[1:])}"
        else:
            return f"=({str_monto[1:]})/{fct_reconversion}"
    return str_monto


def aplica_factor_de_reconversion(test_value):
    try:
        parsed_expression = lexAllOnly.parseString(test_value, parseAll=True)
    except ParseException as err:
        print(f"{err.line}\n{' '*(err.column-1)}^\n{err}\n")
        return None

    result = ''
    count_braces = 0
    start_idx = 0
    curr_idx = 0
    for token in parsed_expression:
        if token == '(':
            count_braces += 1
        elif token == ')':
            count_braces -= 1
        elif token in ['+', '-'] and count_braces == 0:
            result += ''.join(parsed_expression[start_idx: curr_idx])
            if len(result) > 0:
                result += f"/{fct_reconversion}"
            result += token
            start_idx = curr_idx + 1
        curr_idx += 1
    result += f"{''.join(parsed_expression[start_idx:])}/{fct_reconversion}"
#    print(f"DEBUG: Expresión: {test_value}" + \
#          f"       Parsing:   {parsed_expression}" + \
#          f"       Resultado: {result}")

    return result


def valida_parametros():
    global fct_reconversion, cono_monetario, old_workbook, new_workbook

    window.Element('_fct_numeric_error_').Update(visible=False)
    fct_reconversion = values['_fct_conversion_']
    cono_monetario = values['_cono_monetario_']
    old_workbook = values['_filename_']
    if fct_reconversion == '':
        fct_reconversion = default_fct
    elif not fct_reconversion.isnumeric():
        window.Element('_fct_numeric_error_').Update(visible=True)
        return False
    if cono_monetario == '':
        cono_monetario = default_cono
    splitted_filename = os.path.splitext(old_workbook)
    new_workbook = splitted_filename[0] + ' (' + cono_monetario + ')' + splitted_filename[1]

    return True


def aplica_reconversion_monetaria():
    global old_workbook, new_workbook, workbook, step
#    event, values = window.read()
    window.Element('_status_line_').Update(value=f'Cargando hoja de cálculo "{old_workbook}" ...')
    print(f'Cargando hoja de cálculo "{old_workbook}" ...')

    workbook = openpyxl.load_workbook(old_workbook, read_only=False, keep_vba=True)
    progress_bar.UpdateBar(step := 1)

    convierte_Vigilancia()
    progress_bar.UpdateBar(2)
    convierte_RESUMEN_VIGILANCIA()
    progress_bar.UpdateBar(3)
    convierte_Saldos()
    progress_bar.UpdateBar(4)
    convierte_CUOTAS()
    progress_bar.UpdateBar(5)

#    event, values = window.read()
    window.Element('_status_line_').Update(value=f"Guardando en '{new_workbook}' ...")
    print(f"Guardando en '{new_workbook}' ...")
    workbook.save(new_workbook)
    window.Element('_status_line_').Update(value=f"Guardado en '{new_workbook}' ...")
    progress_bar.UpdateBar(step := step + 1)


#
# PROCESO
#

print("Generando layout ...")
sg.theme('DarkAmber')   # Add a little color to your windows

files = [x for x in os.listdir() if x.startswith("1.1. GyG Recibos")]
files.sort(reverse=True)

# All the stuff inside your window. This is the PSG magic code compactor...
layout = [  [sg.Text("Modifica los pagos en las hojas 'Vigilancia', 'Resumen Vigilancia' y 'Saldos'\n" +
                     "para llevarlos al nuevo cono monetario\n")],
            [sg.Text('Archivo de origen:'),
             sg.Spin(files, initial_value=files[0], size=(50, 10), key='_filename_')],
            [sg.Text(f"Factor de conversión [{default_fct}]:", size=(25,1)),
             sg.InputText(size=(10, 10), key='_fct_conversion_'),
             sg.Text('<- Número inválido', text_color='red', visible=False, size=(15,1), key='_fct_numeric_error_')],
            [sg.Text(f"Nuevo cono monetario [{default_cono}]:", size=(25, 1)),
             sg.InputText(size=(10, 10), key='_cono_monetario_')],
            [sg.Button('Reconvierte'), sg.Button('Cancela')],
            [sg.Text('_'*80)],
            [sg.ProgressBar(6, orientation='h', size=(20, 10), key='_progress_bar_'),
             sg.Text('', size=(50, 1), key='_status_line_')]]

# Create the Window
window = sg.Window('Reconversión Monetaria', layout)
progress_bar = window['_progress_bar_']

# Event Loop to process "events"
while True:             
    event, values = window.read()
    if event in (None, 'Cancela'):
        break
    elif event == 'Reconvierte':
        if valida_parametros():
            aplica_reconversion_monetaria()
            sg.Popup(advertencia, title="ADVERTENCIA")

window.close()
print()
