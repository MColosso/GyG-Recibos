# GyG Utilitarios
#
# Rutinas utilitarias utilizadas en varios módulos
#  . input_mes_y_año(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . input_si_no(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . input_valor(mensaje, valor_por_defecto, toma_opción_por_defecto=False)
#  . edita_número(number, num_decimals=2)
#  . trunca_texto(texto, max_width)
#  . espacios(width=1, char=' ')
#  . is_numeric(valor)
#  . MontoEnLetras(número, mostrar_céntimos=True, céntimos_en_letras=False)
#  . genera_recibo(r)
#      'r' es un diccionario, dataframe o serie con las siguientes claves:
#        . 'Nro. Recibo',  'Fecha',     'Beneficiario',  'Dirección',
#        . 'Monto',        'Concepto',  'Categoría'
#



from re import match
from datetime import datetime, timedelta
import numbers
import locale

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

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
                print('  > Seleccione un mes y año correctos (2017+), con un guión como separador...')
        return valor_actual

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


def edita_número(number, num_decimals=2):
    return locale.format_string(f'%.{num_decimals}f', number, grouping=True, monetary=True)

def trunca_texto(texto, max_width):
    return texto[0: max_width - 3] + '...' if len(texto) > max_width else texto

def espacios(width=1, char=' '):
    return char * width

def is_numeric(valor):
    return isinstance(valor, numbers.Number)


def MontoEnLetras(número, mostrar_céntimos=True, céntimos_en_letras=False):
    #
    # Constantes
    Moneda = "Bolívar"       # Nombre de Moneda (Singular)
    Monedas = "Bolívares"    # Nombre de Moneda (Plural)
    Céntimo = "Céntimo"      # Nombre de Céntimos (Singular)
    Céntimos = "Céntimos"    # Nombre de Céntimos (Plural)
    Preposición = "con"      # Preposición entre Moneda y Céntimos
    Máximo = 1999999999.99   # Máximo valor admisible

    def _número_recursivo(número):
        UNIDADES = ("", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez",
                    "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho",
                    "Diecinueve", "Veinte", "Veintiun", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco",
                    "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve")
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
        # convertir la parte entera del número en letras
        letra = _número_recursivo(int(número))

        # Agregar la descripción de la moneda
        letra += ' ' + (Moneda if int(número) == 1 else Monedas)

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


def genera_recibo(r):
    """ genera_recibo(r)
        Recibe como parámetro con las siguientes claves:
          . 'Nro. Recibo',  'Fecha',    'Beneficiario',
          . 'Dirección',    'Monto',    'Concepto0,
          . 'Categoría'
    """

    # Librerías
    import GyG_constantes
    from PIL import Image, ImageFont, ImageDraw, ImageEnhance
    from math import ceil
    import sys
    import os

    # Define textos
    input_file  = GyG_constantes.plantilla_recibos     # './imagenes/plantilla_recibos.png'
    output_file = GyG_constantes.img_recibo            # 'GyG Recibo_{recibo:05d}.png'
    output_path = GyG_constantes.ruta_recibos          # './GyG Recibos/Recibos de Pago'

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
    monto = edita_número(r['Monto'])
    if anchura_de_texto(monto, font) > 90:
        texto = 'Por Bs. ' + monto
        recibo.text(xy=(498 + justifica_centro(texto, 670-498+1, font) + 1, 91), text=texto, font=font, fill='black')
    else:
        recibo.text(xy=(515, 91), text='Por Bs. ', font=font, fill='black')
        recibo.text(xy=(571 + justifica_derecha(monto, 90, font),  91), text=monto, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    recibo.text(xy=(195, 169), text=r['Beneficiario'] + ', ' + r['Dirección'], font=font, fill='black')

    font = ImageFont.truetype(font=calibri_italic, size=15)
    posicion = (195, 199)
    monto_en_letras = multilineas(MontoEnLetras(r['Monto']), 480, font)
    text_size = recibo.textsize(text=monto_en_letras, font=font)
    recibo.rectangle((posicion[0]-2, posicion[1]-2, posicion[0]+text_size[0]+2, posicion[1]+text_size[1]+2),
                     fill=(189, 215, 238))
    recibo.text(xy=posicion, text=monto_en_letras, font=font, fill='black')

    font = ImageFont.truetype(font=calibri_bold_italic, size=15)
    recibo.text(xy=(195, 230), text=multilineas(r['Concepto'], 480, font), font=font, fill='black')

    font = ImageFont.truetype(font=calibri, size=14)
    fecha = f"{r['Fecha']:%d de %B de %Y}"
    recibo.text(xy=(121, 292), text=fecha, font=font, fill='black')

    if r['Categoría'] in GyG_constantes.sellos:
        font = ImageFont.truetype(font=stencil, size=60)
        ancho, alto = recibo.textsize(text=r['Categoría'].capitalize(), font=font)
        tx, ty = cx - ancho // 2, cy - alto // 2
        angulo = 30
        transparente = (0, 0, 0, 0)
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
