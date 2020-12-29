# GyG Utilitarios
#
# Rutinas utilitarias utilizadas en varios módulos
#  . input_mes_y_año(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . input_fecha(mensaje, valor_por_defecto)
#  . input_si_no(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . input_valor(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . edita_número(number, num_decimals=2) -> str
#  . trunca_texto(texto, max_width) -> str
#  . espacios(width=1, char=' ') -> str
#  . is_numeric(valor) -> bool
#  . remueve_acentos(text: str) -> str
#  . _obscure(data: bytes) -> bytes
#  . _unobscure(data: bytes) -> bytes
#  . valida_codigo_seguridad(recibo, fecha: datetime, codigo: str) -> bool
#  . MontoEnLetras(número, mostrar_céntimos=True, céntimos_en_letras=False)
#  . genera_recibo(r, sella_recibo=False, codigo_de_seguridad=False)
#      'r' es un diccionario, dataframe o serie con las siguientes claves:
#        . 'Nro. Recibo',  'Fecha',     'Beneficiario',  'Dirección',
#        . 'Monto',        'Concepto',  'Categoría',     'Monto $',    '$'
#

"""
    POR HACER
    -   Cambiar de nombre a la rutina 'is_numeric()' a 'es_numérico()':
            [ ] saldos_pendientes.py        [ ] resumen_saldos.py
            [ ] analisis_de_pagos.py        [ ] cambios_de_categorias.py
            [ ] cartelera_virtual.py
    -   Cambiar parámetro 'num_decimals' en la rutina 'edita_número()' a ???

     
    HISTORICO
    -   Se corrigió la rutina MontoEnLetras() para añadir la partícula 'de' cuando el monto es en millones.
        (Ejemplos: 1.000.000,00 = Un Millón de Bolivares con 00/100
                     870.000,00 = Ochocientos Setenta Mil Bolívares con 00/100) (22/11/2020)
    -   Se añadió la posibilidad de generar un recibo de pago con el monto expresado en dólares (19/11/2020)
    -   Se corrigió la rutina genera_recibo(): Al incorporar la impresión del sello de la Asociación, se
        produjo un error en la ubicación de los sellos en rojo («Anulado», etc.), quedando fuera del area
        visible (09/11/2020)
    -   Se agrega el código de validación al recibo de pago y las rutinas para la generación y validación del
        mismo (18/08/2020)
        Las rutinas _obscure() y _deobscure() fueron tomadas de:
        "Simple way to encode a string according to a password?"
        (https://stackoverflow.com/questions/2490334/simple-way-to-encode-a-string-according-to-a-password/16321853)
    -   Corregidos algunos acentos en rutina MontoEnLetras() (14/06/2020)
    
"""

import GyG_constantes
import re
from re import match
from datetime import datetime, timedelta
from unicodedata import normalize
import numbers
import locale

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

import zlib
from base64 import (
    urlsafe_b64encode as b64enc,
    urlsafe_b64decode as b64dec)


def _obscure(data: bytes) -> bytes:
    return b64enc(zlib.compress(data, level=zlib.Z_BEST_COMPRESSION))

def _unobscure(obscured: bytes) -> bytes:
    return zlib.decompress(b64dec(obscured))


def input_mes_y_año(mensaje, valor_por_defecto, toma_opción_por_defecto=False):
    """
    Lee del standard input un mes y año en el formato 'mm-yyyy', filtrando años entre 2017 y 2029
    """
    if toma_opción_por_defecto:
        print(f"*** {mensaje}: {valor_por_defecto}")
        return valor_por_defecto
    else:
        pattern = '(0[1-9]|1[012])-20(1[7-9]|2[0-9])'
        while True:
            valor_actual = input(f"*** {mensaje} [{valor_por_defecto}]: ")
            if len(valor_actual) == 0:
                valor_actual = valor_por_defecto
            if bool(match(pattern, valor_actual)):
                break
            else:
                print('  > Seleccione un mes y año correctos (2017+), con un guión como separador')
        return valor_actual

def input_fecha(mensaje, valor_por_defecto=None):
    """
    Lee del standard input una fecha en el formato 'dd-mm-yyyy' o 'dd/mm/yyyy'
    """
    pattern = '(0[1-9]|[12][0-9]|3[01])[/-](0[1-9]|1[012])[/-]20(1[7-9]|2[0-9])'
    while True:
        if valor_por_defecto is None:
            por_defecto = ''
        elif isinstance(valor_por_defecto, datetime):
            por_defecto = valor_por_defecto.strftime('%d/%m/%Y')
        else:
            por_defecto = valor_por_defecto
        valor_actual = input(f"*** {mensaje}{'' if valor_por_defecto is None else f'[{por_defecto}]'}: ")
        valor_actual = valor_actual.replace('-', '/')
        if len(valor_actual) == 0 and valor_por_defecto is not None:
            fecha = valor_por_defecto if isinstance(valor_por_defecto, datetime) \
                                      else datetime.strptime(valor_actual, '%d/%m/%Y')
            break
        if bool(match(pattern, valor_actual)):
            try:
                fecha = datetime.strptime(valor_actual, "%d/%m/%Y")
                break
            except:
                pass
        print("  > Seleccione una fecha correcta: 2017+, formato 'dd-mm-yyyy' o 'dd/mm/yyyy'")
    return fecha

def input_si_no(mensaje, valor_por_defecto, toma_opción_por_defecto=False):
    """
    Lee del standard input 'sí' o 'no' como respuesta a la pregunta formulada
    """
    if toma_opción_por_defecto:
        print(f"*** {mensaje}: {valor_por_defecto}")
        return valor_por_defecto[0] == 's'
    else:
        while True:
            valor_actual = input(f"*** {mensaje} [{valor_por_defecto}]: ")
            if len(valor_actual) == 0:
                valor_actual = valor_por_defecto
            valor_actual = valor_actual.lower()
            if valor_actual[0] in 'sn':
                valor_actual = valor_actual[0] == 's'
                break
            else:
                print("  > Indique 'sí' o 'no'")
        return valor_actual

def input_valor(mensaje, valor_por_defecto, toma_opción_por_defecto=False):
    """
    Lee del standard input un valor (int, float, str) como respuesta a la pregunta formulada
    """
    if toma_opción_por_defecto:
        print(f"*** {mensaje}: {valor_por_defecto}")
        return valor_por_defecto
    else:
        valor_actual = input(f"*** {mensaje} [{valor_por_defecto}]: ")
        if len(valor_actual) == 0:
            valor_actual = valor_por_defecto
        else:
            while True:
                if type(valor_actual) != type(valor_por_defecto):
                    typeof_valor_por_defecto = type(valor_por_defecto)
                    try:
                        valor_actual = typeof_valor_por_defecto(valor_actual)
                        break
                    except:
                        print(f"  > El valor ingresado no puede ser convertido a '{type(valor_por_defecto)}'")
                else:
                    break
            if isinstance(valor_actual, str):
                if   valor_por_defecto.islower(): valor_actual = valor_actual.lower()
                elif valor_por_defecto.isupper(): valor_actual = valor_actual.upper()
                elif valor_por_defecto.istitle(): valor_actual = valor_actual.title()
        return valor_actual


def edita_número(valor, num_decimals=2):
    return locale.format_string(f'%.{num_decimals}f', valor, grouping=True, monetary=True)

def trunca_texto(texto, max_width):
    return texto[0: max_width - 3] + '...' if len(texto) > max_width else texto

def espacios(width=1, char=' '):
    return char * width

def is_numeric(valor):
    return isinstance(valor, numbers.Number)

def valida_codigo_seguridad(recibo, fecha: datetime, codigo: str) -> bool:
    nro_recibo = f"{recibo:0{GyG_constantes.long_num_recibo}d}" if isinstance(recibo, int) else recibo
    str_fecha = fecha.strftime("%d/%m/%Y")
    return str(_obscure(bytes(' '.join([nro_recibo, str_fecha]), 'UTF-8')), 'UTF-8') == codigo

def remueve_acentos(texto: str) -> str:
    # -> NFD y eliminar diacríticos
    s = re.sub(
            r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1", 
            normalize( "NFD", texto), 0, re.I
        )
    # -> NFC
    return normalize( 'NFC', s)

def MontoEnLetras(número, mostrar_céntimos=True, céntimos_en_letras=False, moneda='Bolívar'):
    #
    # Constantes
    Vocales = ['a', 'e', 'i', 'o', 'u']
    Moneda = moneda          # Nombre de Moneda (Singular)
    Monedas = f"{moneda}{'s' if moneda[-1] in Vocales else 'es'}"
                             # Nombre de Moneda (Plural)
    Céntimo = "Céntimo"      # Nombre de Céntimos (Singular)
    Céntimos = "Céntimos"    # Nombre de Céntimos (Plural)
    Preposición = "con"      # Preposición entre Moneda y Céntimos
    Máximo = 1999999999.99   # Máximo valor admisible

    def _número_recursivo(número):
        UNIDADES = ("", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez",
                    "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho",
                    "Diecinueve", "Veinte", "Veintiun", "Veintidós", "Veintitrés", "Veinticuatro", "Veinticinco",
                    "Veintiséis", "Veintisiete", "Veintiocho", "Veintinueve")
        DECENAS  = ("", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta",
                    "Noventa", "Cien")
        CENTENAS = ("", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos",
                    "Setecientos", "Ochocientos", "Novecientos")

        if número == 0:
            resultado = 'Cero'
        elif número <= 29:
            resultado = UNIDADES[número]
        elif número <= 100:
            resultado = DECENAS[número // 10]
            if número % 10 != 0: resultado += ' y ' + _número_recursivo(número % 10)
        elif número <= 999:
            resultado = CENTENAS[número // 100]
            if número % 100 != 0: resultado += ' ' + _número_recursivo(número % 100)
        elif número <= 1999:
            resultado = 'Mil'
            if número % 1000 != 0: resultado += ' ' + _número_recursivo(número % 1000)
        elif número <= 999999:
            resultado = _número_recursivo(número // 1000) + ' Mil'
            if número % 1000 != 0: resultado += ' ' + _número_recursivo(número % 1000)
        elif número <= 1999999:
            resultado = 'Un Millón'
            if número % 1000000 != 0: resultado += ' ' + _número_recursivo(número % 1000000)
        elif número <= 1999999999:
#            resultado = CENTENAS(número // 1000000) + ' Millones'
            resultado = _número_recursivo(número // 1000000) + ' Millones'
            if número % 1000000 != 0: resultado += ' ' + _número_recursivo(número % 1000000)
        else:
#            resultado = format(número, ',.0f').replace(',', '.')
            resultado = edita_número(número, num_decimals=0)

        return resultado

    if 0 <= número <= Máximo:
        # Si el número es en millones, debe expresarse como: «monto en letras» de «moneda(s)»
        de = ' de ' if int(número) != 0 and int(número) % 1000000 == 0 else ' '

        # convertir la parte entera del número en letras
        letra = _número_recursivo(int(número))

        # Agregar la descripción de la moneda
        letra += de + (Moneda if int(número) == 1 else Monedas)

        # Obtener los céntimos del Numero
        num_céntimos = int(round((número - int(número)) * 100, 0))

        if mostrar_céntimos:
            if céntimos_en_letras:
                letra += ' ' + Preposición + ' ' + _número_recursivo(num_céntimos)
                letra += ' ' + (Céntimo if num_céntimos == 1 else Céntimos)
            else:
                letra += ' ' + Preposición + (' 0' if num_céntimos < 10 else ' ') + str(num_céntimos) + '/100'

        return letra
    else:
        return f'ERROR: El número {edita_número(número, num_decimals=2)} excede los límites admitidos.'


def genera_recibo(r, sella_recibo=False, codigo_de_seguridad=False):
    """ genera_recibo(r, sella_recibo=False)
        Recibe como parámetro con las siguientes claves:
          . 'Nro. Recibo',  'Fecha',    'Beneficiario',
          . 'Dirección',    'Monto',    'Concepto',
          . 'Categoría',    'Monto $',  '$'
    """

    # Librerías
    import GyG_constantes
    from PIL import Image, ImageFont, ImageDraw, ImageEnhance
    from math import ceil
    from random import random
    import sys
    import os

    # Define textos
    input_file  = GyG_constantes.plantilla_recibos     # './imagenes/plantilla_recibos.png'
    output_file = GyG_constantes.img_recibo            # 'GyG Recibo_{recibo:05d}.png'
    output_path = GyG_constantes.ruta_recibos          # '../GyG Archivos/Recibos de Pago'
    img_sello   = GyG_constantes.img_sello_GyG         # './recursos/imágenes/sello_GyG.png'

    moneda, moneda_abrev, monto_en_moneda = ('Dólar',   'US$', r['Monto $']) if r['$'] == 'ü' \
                                       else ('Bolívar', 'Bs.', r['Monto'])

    # Fuentes
    calibri             = os.path.join(GyG_constantes.rec_fuentes, 'calibri.ttf')
    calibri_italic      = os.path.join(GyG_constantes.rec_fuentes, 'calibrii.ttf')
    calibri_bold        = os.path.join(GyG_constantes.rec_fuentes, 'calibrib.ttf')
    calibri_bold_italic = os.path.join(GyG_constantes.rec_fuentes, 'calibriz.ttf')
    stencil             = os.path.join(GyG_constantes.rec_fuentes, 'STENCIL.TTF')


    def anchura_de_texto(text, font):
        return recibo.textsize(text=text, font=font)[0]

    def altura_de_texto(text, font):
        return recibo.textsize(text=text, font=font)[1]

    def justifica_derecha(texto, anchura, font):
        return anchura - anchura_de_texto(text=texto, font=font)

    def justifica_centro(texto, anchura, font):
        return int(ceil((anchura - anchura_de_texto(text=texto, font=font)) / 2.0))

    def multilineas(texto, anchura, font):
        words = texto.split()
        for x in reversed(range(len(words))):
            texto_inicial = ' '.join(words[:x+1])
            texto_final   = ' '.join(words[x+1:])
            if recibo.textsize(text=texto_inicial, font=font)[0] <= anchura:
                break
        return texto_inicial + ('\n' + texto_final if len(texto_final) > 0 else '')

    def aleatoriza(base, pct_variacion):
        """ Devuelve el valor de la base +/- el porcentaje de variación indicado """
        return base - pct_variacion/100 + 2 * pct_variacion/100 * random()


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
    recibo.text(xy=(620,  64), text=f"{r['Nro. Recibo']:0{GyG_constantes.long_num_recibo}d}", font=font, fill='black')

    font = ImageFont.truetype(font=calibri_bold, size=18)
    monto = edita_número(monto_en_moneda)
    if anchura_de_texto(monto, font) > 90:
        texto = f'Por {moneda_abrev} {monto}'
        recibo.text(xy=(498 + justifica_centro(texto, 670-498+1, font) + 1, 91),
                    text=texto, font=font, fill='black')
    else:
        recibo.text(xy=(515, 91), text=f'Por {moneda_abrev} ', font=font, fill='black')
        recibo.text(xy=(571 + justifica_derecha(monto, 90, font),  91),
                    text=monto, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    recibo.text(xy=(195, 169), text=r['Beneficiario'] + ', ' + r['Dirección'], font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    posicion = (195, 199)
    monto_en_letras = multilineas(MontoEnLetras(monto_en_moneda, moneda=moneda), 480, font)
    text_size = recibo.textsize(text=monto_en_letras, font=font)
    recibo.rectangle((posicion[0]-2, posicion[1]-2, posicion[0]+text_size[0]+2, posicion[1]+text_size[1]+2),
                     fill=(189, 215, 238))
    recibo.text(xy=posicion, text=monto_en_letras, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_bold_italic, size=15)
    recibo.text(xy=(195, 230), text=multilineas(r['Concepto'], 480, font), font=font, fill='black')

    font = ImageFont.truetype(font=calibri, size=14)
    fecha = f"{r['Fecha']:%d de %B de %Y}"
    recibo.text(xy=(121, 292), text=fecha, font=font, fill='black')

    transparente = (0, 0, 0, 0)

    if codigo_de_seguridad:
        font = ImageFont.truetype(font=calibri, size=10)
        código_a_convertir = f"{r['Nro. Recibo']:0{GyG_constantes.long_num_recibo}d} " + \
                             r['Fecha'].strftime('%d/%m/%Y')
        código_convertido = str(_obscure(bytes(código_a_convertir, 'UTF-8')), 'UTF-8')
        ancho_código = anchura_de_texto(código_convertido, font)

        # ----------------------------------------------------------------
        # CODIGO DE VALIDACION HORIZONTAL (justificado a la izquierda)
        #
        recibo.text(xy=(20, plantilla.size[1] - 25),
                    text=código_convertido, font=font, fill='black')
        # ----------------------------------------------------------------
        # CODIGO DE VALIDACION HORIZONTAL (justificado a la derecha)
        #
        # recibo.text(xy=(plantilla.size[0] - ancho_código - 38, plantilla.size[1] - 25),
        #             text=código_convertido, font=font, fill='black')
        # ----------------------------------------------------------------
        # CODIGO DE VALIDACION VERTICAL
        #
        # ancho_código = anchura_de_texto(código_convertido, font)
        # alto_código = altura_de_texto(código_convertido, font)
        # max_hw = max(ancho_código, alto_código)
        # canvas = Image.new(mode='RGBA', size=(max_hw, max_hw), color=transparente)
        # img_código = ImageDraw.Draw(canvas)
        # img_código.text(xy=(0, 0), text=código_convertido, font=font, fill='black')
        # canvas = canvas.rotate(90, expand=True, fillcolor=transparente)
        # opacidad = 0.5
        # en = ImageEnhance.Brightness(canvas)
        # mask = en.enhance(1.0 - opacidad)
        # plantilla.paste(canvas, box=(12, plantilla.size[1] - canvas.size[1] - 15), mask=mask)
        # ----------------------------------------------------------------

    if sella_recibo:
        # El fondo transparente de la imagen del sello de la Asociación fue logrado
        # con https://onlinepngtools.com/create-transparent-png
        try:
            sello_GyG = Image.open(img_sello)
            cx, cy = sello_GyG.size[0] // 2, sello_GyG.size[1] // 2
        except:
            error_msg  = str(sys.exc_info()[1])
            if sys.platform.startswith('win'):
                error_msg  = error_msg.replace('\\', '/')
            print(f"*** Error cargando sello {img_sello}: {error_msg}")
            return False
        angulo = 15
        angulo = aleatoriza(angulo, pct_variacion=500)
        opacidad = 0.5
        sello_GyG = sello_GyG.rotate(angulo, expand=True, center=(cx, cy), fillcolor=transparente)
        sello_GyG.thumbnail((200, 200))   # set the maximum width and height to 200 px
        en = ImageEnhance.Brightness(sello_GyG)
        mask = en.enhance(1.0 - opacidad)
        px, py = (0.60, 0.95)               # posición sobre el recibo (x=60%, y=95%)
        px = aleatoriza(px, pct_variacion=5)
        py = aleatoriza(py, pct_variacion=5)
        position = (int((plantilla.width - sello_GyG.width) * px), int((plantilla.height - sello_GyG.height) * py))
        plantilla.paste(sello_GyG, position, mask=mask)

    if r['Categoría'] in GyG_constantes.sellos:
        font = ImageFont.truetype(font=stencil, size=60)
        ancho, alto = recibo.textsize(text=r['Categoría'].capitalize(), font=font)
        cx, cy = plantilla.size[0] // 2, plantilla.size[1] // 2
        tx, ty = cx - ancho // 2, cy - alto // 2
        angulo = 30
        opacidad = 0.5
        img_anulado = Image.new('RGBA', plantilla.size, color=transparente)
        anulado = ImageDraw.Draw(img_anulado)
        anulado.text(xy=(tx, ty), text=r['Categoría'].capitalize(), font=font, fill='red', align='center')
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
