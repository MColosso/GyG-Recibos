# GyG GENERA RECIBOS A SOLICITUD
#
# Genera los recibos de pago indicados a partir de la hoja de cálculo estándar o la
# hoja de cálculo con datos históricos antes de la reconversión.
# Esto permite borrar en cualquier momento los recibos generados y regenerarlos a
# voluntad.

"""
    POR HACER
    -   

    HISTORICO
    -   

"""

import GyG_constantes
import sys, os

# Selecciona la carpeta de destino
output_path = GyG_constantes.ruta_recibos    # './GyG Recibos/Recibos de Pago'
path_changed = False
if len(sys.argv) > 1:
    output_path = sys.argv[1]
    path_changed = True

print('Cargando librerías...')
from pandas import read_excel, isnull, notnull
from PIL import Image, ImageFont, ImageDraw, ImageEnhance
#from datetime import datetime
from monto_en_letras import MontoEnLetras
import locale

# Define textos
input_file      = GyG_constantes.plantilla_recibos     # './imagenes/plantilla_recibos.png'
output_file     = GyG_constantes.img_recibo            # 'GyG Recibo_{recibo:05d}.png'

excel_workbook  = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_worksheet = GyG_constantes.pagos_ws_vigilancia   # 'Vigilancia'

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

# Fuentes
calibri             = os.path.join(GyG_constantes.rec_fuentes, 'calibri.ttf')
calibri_italic      = os.path.join(GyG_constantes.rec_fuentes, 'calibrii.ttf')
calibri_bold        = os.path.join(GyG_constantes.rec_fuentes, 'calibrib.ttf')
calibri_bold_italic = os.path.join(GyG_constantes.rec_fuentes, 'calibriz.ttf')
stencil             = os.path.join(GyG_constantes.rec_fuentes, 'STENCIL.TTF')


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def edit_number(number, num_decimals=2):
    return locale.format_string(f'%.{num_decimals}f', number, grouping=True, monetary=True)

def convierte_recibo(r):

    def justifica_derecha(texto, anchura, font):
        return anchura - recibo.textsize(text=texto, font=font)[0]

    def multilineas(texto, anchura, font):
        words = texto.split()
        for x in reversed(range(len(words))):
            texto_inicial = ' '.join(words[:x+1])
            texto_final   = ' '.join(words[x+1:])
            if recibo.textsize(text=texto_inicial, font=font)[0] <= anchura:
                break
        return texto_inicial + ('\n' + texto_final if len(texto_final) > 0 else '')

    try:
        plantilla = Image.open(input_file)
#        plantilla = plantilla.convert('RGBA')
        cx, cy = plantilla.size[0] // 2, plantilla.size[1] // 2
    except:
        error_msg  = str(sys.exc_info()[1])
        if sys.platform.startswith('win'):
            error_msg  = error_msg.replace('\\', '/')
        print(f"*** Error cargando plantilla {input_file}: {error_msg}")
        return False
    recibo = ImageDraw.Draw(plantilla)

    font = ImageFont.truetype(font=calibri_bold, size=15)
    recibo.text(xy=(620,  64), text='{:05d}'.format(r['Nro. Recibo']), font=font, fill='black')

    font = ImageFont.truetype(font=calibri_bold, size=18)
#    monto = '{:,.2f}'.format(r['Monto']).replace(',', 'x').replace('.', ',').replace('x', '.')
    monto = edit_number(r['Monto'])
    recibo.text(xy=(571 + justifica_derecha(monto, 90, font),  91), text=monto, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    recibo.text(xy=(195, 169), text=r['Beneficiario'] + ', ' + r['Dirección'], font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    posicion = (195, 199)
    monto_en_letras = MontoEnLetras(r['Monto'])
    text_size = recibo.textsize(text=monto_en_letras, font=font)
    recibo.rectangle((posicion[0]-2, posicion[1]-2, posicion[0]+text_size[0]+2, posicion[1]+text_size[1]+2),
                     fill=(189, 215, 238))
    recibo.text(xy=posicion, text=monto_en_letras, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_bold_italic, size=15)
    recibo.text(xy=(195, 230), text=multilineas(r['Concepto'], 480, font), font=font, fill='black')

    font = ImageFont.truetype(font=calibri, size=14)
#    fecha = '{:%d de %B de %Y}'.format(r['Fecha'])
    fecha = f"{r['Fecha']:%d de %B de %Y}"
    recibo.text(xy=(121, 292), text=fecha, font=font, fill='black')

    if notnull(r['Anulado']):
        font = ImageFont.truetype(font=stencil, size=60)
        ancho, alto = recibo.textsize(text=r['Anulado'], font=font)
        tx, ty = cx - ancho // 2, cy - alto // 2
        angulo = 30
        transparente = (0, 0, 0, 0)
        opacidad = 0.5
        img_anulado = Image.new('RGBA', plantilla.size, color=transparente)
        anulado = ImageDraw.Draw(img_anulado)
        anulado.text(xy=(tx, ty), text=r['Anulado'], font=font, fill='red', align='center')
        img_anulado = img_anulado.rotate(angulo, center=(cx, cy), fillcolor=transparente)
        en = ImageEnhance.Brightness(img_anulado)
        mask = en.enhance(1.0 - opacidad)
        plantilla.paste(img_anulado, mask=mask)

    try:
        plantilla.save(os.path.join(output_path, output_file.format(recibo=r['Nro. Recibo'])))
        return True
    except:
        error_msg  = str(sys.exc_info()[1])
        if sys.platform.startswith('win'):
            error_msg  = error_msg.replace('\\', '/')
        print(f"*** Error guardando recibo {output_file.format(recibo=r['Nro. Recibo'])}: {error_msg}")
        return False


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
print()
while True:
    genera_historico = input("*** Genera los recibos a partir del histórico [no]: ")
    if len(genera_historico) == 0:
        genera_historico = "no"
    genera_historico = genera_historico.lower()
    if genera_historico[0] in 'sn':
        genera_historico = genera_historico[0] == 's'
        break
    else:
        print("    Indique 'sí' o 'no'")

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


# Abre la hoja de cálculo de Recibos de Pago
print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))

# Convierte columna 'Nro. Recibo' en enteros
#df['Nro. Recibo'] = df['Nro. Recibo'].astype(int)

# Selecciona sólo las columnas a utilizar
df = df[['Nro. Recibo', 'Fecha', 'Monto', 'Beneficiario', 'Dirección', 'Concepto', 'Anulado']]

if len(t_recibos) == 0:
    print('\n*** Proceso terminado: No hay recibos solicitados para generar\n')
    sys.exit()

print()
print(f"Generando recibos {muestra_rangos(t_recibos)}: ", end="")

recibos_convertidos = 0
for recibo in t_recibos:
    r = df[df['Nro. Recibo'] == recibo]
    r = r.iloc[0]        # 'r' es un dataframe que necesitamos convertir en una Series
    print('.', end='')   # Imprime un punto en la pantalla por cada mensaje
    sys.stdout.flush()   # Flush output to the screen
    if convierte_recibo(r):
        recibos_convertidos += 1

print()
print('\n*** Proceso terminado: {} de {} recibos generados\n'.format(recibos_convertidos,
                                                                     len(t_recibos)))
if path_changed:
    print(f'    Carpeta "{output_path}"')
