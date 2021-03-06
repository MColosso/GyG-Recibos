# GyG GENERA RECIBOS COMO IMAGENES
#
# Genera los recibos de pago pendientes por imprimir en formato .png

"""
    POR HACER
    -   

    NOTAS
    -   Generar el código de seguridad como código de barras
         -> El código generado ('code128' para permitir mayúsculas y minúsculas) es muy largo
            y abarcaría todo el recibo para que sea legible
             -> Ignorar este requerimiento.... (01/04/2021)

    HISTORICO
    -   "Sellar" los recibos de pago con el sello de la Asociación.
         -> Se agrega el parámetro 'sella_recibo' a la rutina 'genera_recibo()' de GyG_utilitarios
            para "imprimir" el sello de la Asociación en el Recibo de Pago (30/07/2020)
         -> Se añade una consulta al principio de la ejecución de este módulo para controlar la
            impresión del sello de la Asociación (30/07/2020)
         -> Falta por definir cuáles recibos estarán sellados: ¿todos?, ¿sólo algunos?
            Sabemos que los recibos del Colegio El Trigal deberán estar sellados, así como los
            recibos de solvencia, pero, ¿y los demás?
            Por indicaciones de Marisol Mendoza, todos los recibos serán generados con el sello de
            la Asociación (30/07/2020)
    -   Utiliza la versión unificada de 'genera_recibo()' en GyG_utilitarios, en la cual se ajustó
        el despliegue de montos "grandes" (27/04/2020)
    -   Unificar los archivos 'genera_recibos.py' y 'genera_recibos_seleccionados', separando
        la fuente de los recibos a generar en base a un indicador en la línea de comandos
        (09/12/2019)
    -   Sellar los recibos de pago según lo indicado en la columna 'Categoría' de la pestaña
        'Vigilancia'. Las categorías válidas están identificadas en la variable 'sellos' de
        'GyG_constantes' ('ANULADO', 'SOLVENTE' y 'REVERSADO', a la fecha) (09/12/2019)
    -   Se cambiaron las ubicaciones de los archivos resultantes a la carpeta GyG Recibos dentro de
        la carpeta actual para compatibilidad entre Windows y macOS (21/10/2019)
    -   Timbrar el recibo con el contenido de la columna 'Anulado' (20/10/2019)
    -   Se ajustó la ruta de destino de los recibos: de Google Drive a GyG Recibos_temp (20/10/2019)
    -   CORREGIR: En la hoja de cálculo '1.1. GyG Recibos.xlsm' se cambió el título de la columna 'B'
        de 'Archivo' a 'Generar' (corregido 08/07/2019)
    -   

"""

import GyG_constantes
from GyG_utilitarios import *
import sys, os

# Selecciona la carpeta de destino
# output_path = GyG_constantes.ruta_recibos    # './GyG Recibos/Recibos de Pago'

print('Cargando librerías...')
from pandas import read_excel
import locale

# Define textos
excel_workbook  = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_worksheet = GyG_constantes.pagos_ws_vigilancia   # 'Vigilancia'

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

# Fuente de los recibos a generar
selección_manual = False
if len(sys.argv) > 1:
    selección_manual = sys.argv[1] == '--seleccion_manual'


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def parse_range(str_range, min_value, max_value):
    """
    Returns a list of selected values from individual values and
    open or closed ranges.
    Also returns a list of invalid values and an indicator if
    any of the values corresponds to a range.

    Example: A string in the form:
                "6, -4, 9-"
             would return:
                <min_value>, ..., 4, 6, 9, ..., <max_value>
             as expected...
    """
    t_list = str_range.split(',')
    t_values = set()
    t_invalid = list()
    t_ranges = False
    for x in t_list:
        token = x.strip()
        if token.isnumeric():
            value = int(token)
            if value < min_value: value = min_value
            if value > max_value: value = max_value            
            t_values.add(value)
        else:
            values = token.split('-')
            if len(values) != 2:
                t_invalid.append(token)
                continue
            start_value, end_value = values[0].strip(), values[1].strip()
            if start_value == '': start_value = str(min_value)
            if end_value == '': end_value = str(max_value)
            try:
                start_value = int(start_value)
                end_value = int(end_value)
            except:
                t_invalid.append(token)
                continue
            if start_value < min_value: start_value = min_value
            if end_value > max_value: end_value = max_value
            if start_value <= end_value:
                for v in range(start_value, end_value+1):
                    t_values.add(v)
            t_ranges = True
    
    t_values = list(t_values)
    t_values.sort()
    
    return t_values, t_invalid, t_ranges

def encode_range(t_values, verbose=False):

    def show_values(prev_v, last_v):
        if prev_v == last_v:
            return f"{prev_v:0{npos}}"
        else:
            if verbose:
                return f"desde {prev_v:0{npos}} hasta {last_v:0{npos}}"
            else:
                return f"{prev_v:0{npos}}-{last_v:0{npos}}"

    npos = GyG_constantes.long_num_recibo
    last_x = None
    prev_x = None
    output = []
    
    for x in t_values:
        if last_x is None:
            last_x = x
        elif x != prev_x + 1:
            output.append(show_values(last_x, prev_x))
            last_x = x
        prev_x = x
    output.append(show_values(last_x, prev_x))
    
    return output

def muestra_rangos(t_values):
    recibos = ', '.join(encode_range(t_values, verbose=True))
    last_comma = recibos.rfind(',')
    if last_comma >= 0:
        recibos[:last_comma] + ' y' + recibos[last_comma+1:]
    return recibos


#
# PROCESO
#

# Selecciona si se generan los recibos del histórico o no
genera_historico = False
if selección_manual:
    print()
    genera_historico = input_si_no('Genera los recibos a partir del histórico', 'no')
    sella_recibos = input_si_no('Sella los recibos de pago', 'sí')
    codigo_de_seguridad = input_si_no('Con código de seguridad', 'sí')
    if codigo_de_seguridad:
        # codigo_de_barras = input_si_no(' ... como código de barras', 'no')
        codigo_de_barras = False
    print()

    excel_workbook = GyG_constantes.pagos_wb_historico if genera_historico else GyG_constantes.pagos_wb_estandar
    df = read_excel(excel_workbook, sheet_name=excel_worksheet, dtype={'Nro. Recibo': int})
    primer_recibo = 1
    ultimo_recibo = df.iloc[-1]['Nro. Recibo']
    #print(f"DEBUG: {primer_recibo=}, {ultimo_recibo=}")

    # Selecciona los recibos de pago a generar
    while True:
        rangos_recibos = input("*** Indique los recibos a generar: ")
        t_recibos, t_invalid, t_ranges = parse_range(rangos_recibos, primer_recibo, ultimo_recibo)
        if len(rangos_recibos) == 0 or len(t_invalid) == 0:
            break
        else:
            un, s = ('un ', '') if len(t_invalid) == 1 else ('', 's')
            print(f"    Hay {un}rango{s} inválido{s}: {', '.join(t_invalid)}")
    print()
else:
    sella_recibos = True
    codigo_de_seguridad = True
    codigo_de_barras = False

# Abre la hoja de cálculo de Recibos de Pago
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))

df = read_excel(excel_workbook, sheet_name=excel_worksheet)

if selección_manual:
    df = df[df['Nro. Recibo'].isin(t_recibos)]
else:
    # Elimina registros que no fueron seleccionados para enviar
    df.dropna(subset=['Generar'], inplace=True)

# Convierte columna 'Nro. Recibo' en enteros
df['Nro. Recibo'] = df['Nro. Recibo'].astype(int)

# Selecciona sólo las columnas a utilizar
df = df[['Nro. Recibo', 'Fecha', 'Monto', 'Monto $', '$', 'Beneficiario', 'Dirección', 'Concepto', 'Categoría']]

if df.shape[0] == 0:
    print(f'\n*** Proceso terminado: No hay recibos ' + \
          f'{"seleccionados para" if selección_manual else "pendientes por"} generar\n')
    sys.exit()

s = "s" if df.shape[0] != 1 else ""
if selección_manual:
    print()
    print(f"Generando recibo{s} {muestra_rangos(t_recibos)}: ", end="")
else:
    print('Generando recibos', end="")

recibos_convertidos = 0
for index, r in df.iterrows():
    print('.', end='')   # Imprime un punto en la pantalla por cada mensaje
    sys.stdout.flush()   # Flush output to the screen
    if genera_recibo(r, sella_recibos, codigo_de_seguridad, not codigo_de_barras):
        recibos_convertidos += 1

print()
print('\n*** Proceso terminado: {} de {} recibo{} generado{}\n'.format(recibos_convertidos,
                                                                       df.shape[0], s, s))
