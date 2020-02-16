# GyG Utilitarios
#
# Rutinas utilitarias utilizadas en varios módulos

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
            resultado = 'cero'
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
