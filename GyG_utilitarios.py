# GyG Utilitarios
#
# Rutinas utilitarias utilizadas en varios módulos
#  . input_mes_y_año(mensaje: str, valor_por_defecto: str, toma_opción_por_defecto: bool=False) -> str
#  . input_fecha(mensaje: str, valor_por_defecto) -> datetime
#  . input_si_no(mensaje: str, valor_por_defecto: str, toma_opción_por_defecto: bool=False) -> str
#  . input_valor(mensaje: str, valor_por_defecto, toma_opción_por_defecto: bool=False)
#  . edita_beneficiario(beneficiario: str) -> str
#  . edita_dirección(dirección: str) -> str
#  . edita_categoría(categoría: str) -> str:
#  . alinea_texto(texto: str, anchura: int, alineación: str="derecha") -> str
#  . edita_número(number, num_decimals: int=2) -> str
#  . trunca_texto(texto: str, max_width: int) -> str
#  . espacios(width: int=1, char: str=' ') -> str
#  . bloque_de_texto(texto: str, anchura: int=80, margen: int=0, margen_1ra_linea: bool=False,
#                    continuacion:str=" > ", formato_lista:bool=False) -> list(str) | str
#  . is_numeric(valor) -> bool
#  . remueve_acentos(text: str) -> str
#  . separa_meses(mensaje: str, muestra_modificador: bool=False, as_string: bool=False) -> list(str) | str
#  . reagrupa_meses(mensaje: str, mes_completo: bool) -> str
#  . _obscure(data: bytes) -> bytes
#  . _unobscure(data: bytes) -> bytes
#  . valida_codigo_seguridad(recibo, fecha: datetime, codigo: str) -> bool
#  . MontoEnLetras(número: float, mostrar_céntimos: bool=True, céntimos_en_letras: bool=False) -> str
#  . genera_recibo(r, sella_recibo: bool=False, codigo_de_seguridad: bool=False)
#      'r' es un diccionario, dataframe o serie con las siguientes claves:
#        . 'Nro. Recibo',  'Fecha',     'Beneficiario',  'Dirección',
#        . 'Monto',        'Concepto',  'Categoría',     'Monto $',    '$'
#
# Clases
#  . MyTextWrapper - Basada en 'textwrap' para la separación en líneas, a la cual se le agrega el
#                    parámetro 'separa_palabras' (default: Falso), para separar en sílabas las
#                    palabras al final de las líneas y mejorar con ello el ajuste
#  . Silabeador    - Separa en sílabas la palabra pasada como argumento a la instancia
#

"""
    POR HACER
     -  Cambiar el nombre de la rutina 'is_numeric()' a 'es_numérico()':
            [ ] saldos_pendientes.py        [ ] resumen_saldos.py
            [ ] analisis_de_pagos.py        [ ] cambios_de_categorias.py
            [ ] cartelera_virtual.py
     -  Cambiar parámetro 'num_decimals' en la rutina 'edita_número()' a ???
     -  Revisar la rutina 'separa_meses()' en:
            [√] resumen_saldos          [√] saldos_pendientes.py    [√] carterlera_virtual.py
            [√] graficas_GUI.py         [√] cambios_de_categoria.py [X] analisis_de_pagos.py
            [√] estadistica_pagos.py
        y modificarla para usar la versión de 'GyG_utilitarios'
         -> 'analisis_de_pagos.py' utiliza una versión particular de esta rutina, por lo que,
            en primera instancia, no puede ser sustituida por la versión en 'GyG_constantes'
         -> Ajustada en base a la versión encontrada en 'cambios_de_categoria.py'

     
    HISTORICO
     -  Añadida función 'bloque_de_texto()' para separar texto en líneas de una anchura determinada
        (28/06/2021)
     -  Se ajustó la función 'trunca_texto()' para no finalizar el texto con puntos suspensivos si se
        puede cortar el mismo en ese espacio: "Chá..." -> "Chávez" (22/05/2021)
     -  Revisar la rutina "separa_meses()": El texto "Colaboración Vigilancia, meses de Agosto a Noviembre 
        2020" se descompuso en ['11-2020', '08-2020', '09-2020', '10-2020', '11-2020'] 
         -> Corregido (15/05/2021)
     -  Se incorporó la rutina 'separa_meses()' y se ajustó para permitir el manejo de las secuencias
        «Mes» '-' «Año», ignorando el guión, y no como «Mes» 'a' «Año» (13/04/2021)
     -  Se corrigió la rutina 'MontoEnLetras()' para añadir la partícula 'de' cuando el monto es en millones.
        (Ejemplos: 1.000.000,00 = Un Millón de Bolivares con 00/100
                     870.000,00 = Ochocientos Setenta Mil Bolívares con 00/100) (22/11/2020)
     -  Se añadió la posibilidad de generar un recibo de pago con el monto expresado en dólares (19/11/2020)
     -  Se corrigió la rutina genera_recibo(): Al incorporar la impresión del sello de la Asociación, se
        produjo un error en la ubicación de los sellos en rojo («Anulado», etc.), quedando fuera del area
        visible (09/11/2020)
     -  Se agrega el código de validación al recibo de pago y las rutinas para la generación y validación del
        mismo (18/08/2020)
        Las rutinas _obscure() y _deobscure() fueron tomadas de:
        "Simple way to encode a string according to a password?"
        (https://stackoverflow.com/questions/2490334/simple-way-to-encode-a-string-according-to-a-password/16321853)
     -  Corregidos algunos acentos en rutina MontoEnLetras() (14/06/2020)
    
"""

import GyG_constantes
import re
from re import match
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from unicodedata import normalize
from textwrap import TextWrapper
import numbers
import locale

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

import zlib
from base64 import (
    urlsafe_b64encode as b64enc,
    urlsafe_b64decode as b64dec)


# ---------------------------------------------------------

def _obscure(data: bytes) -> bytes:
    return b64enc(zlib.compress(data, level=zlib.Z_BEST_COMPRESSION))

def _unobscure(obscured: bytes) -> bytes:
    return zlib.decompress(b64dec(obscured))


# ---------------------------------------------------------

def input_mes_y_año(mensaje: str, valor_por_defecto: str, toma_opción_por_defecto: bool=False) -> str:
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


# ---------------------------------------------------------

def input_fecha(mensaje: str, valor_por_defecto: bool=None) -> datetime:
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


# ---------------------------------------------------------

def input_si_no(mensaje: str, valor_por_defecto: str, toma_opción_por_defecto: bool=False) -> str:
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


# ---------------------------------------------------------

def input_valor(mensaje: str, valor_por_defecto, toma_opción_por_defecto: bool=False):
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
            # Valida grupos numéricos
            if valor_actual.replace(".", "").replace(",", "", 1).isnumeric():
                valor_actual = locale.delocalize(valor_actual.replace(".", "").replace(",", "."))
            elif valor_actual.replace(",", "").replace(".", "", 1).isnumeric():
                valor_actual = locale.delocalize(valor_actual.replace(",", "").replace(".", "."))
            # Convierte el tipo de 'valor_actual' al del 'valor_por_defecto'
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
            # Convierte el 'case' de 'valor_actual' al del 'valor_por_defecto'
            if isinstance(valor_actual, str):
                if   valor_por_defecto.islower(): valor_actual = valor_actual.lower()
                elif valor_por_defecto.isupper(): valor_actual = valor_actual.upper()
                elif valor_por_defecto.istitle(): valor_actual = valor_actual.title()
        return valor_actual


# ---------------------------------------------------------

def edita_beneficiario(beneficiario: str) -> str:
    # STOPWORDS = ['Familia ', 'Dr. ', 'Dra. ', 'Sr. ', 'Sra. ']
    # for stopword in STOPWORDS:
    #     beneficiario = beneficiario.replace(stopword, '')
    # return beneficiario
    return beneficiario.replace('Familia ', '').replace('Dr. ', '').replace('Dra. ', '') \
                       .replace('Sr. ', '').replace('Sra. ', '')


# ---------------------------------------------------------

def edita_dirección(dirección: str) -> str:
    return dirección.replace('Calle ', '').replace('Nros. ', '').replace('Nro. ', '')


# ---------------------------------------------------------

def edita_categoría(categoría: str) -> str:
    return categoría


# ---------------------------------------------------------

def get_street(address: str) -> str:
    # return address.index(' ', address.index(' ') + 1)
    grupos = re.findall(r'\w+', address)
    if grupos[0].lower() == "av":
        return "Avenida"
    return grupos[1] if len(grupos) > 0 else ''


# ---------------------------------------------------------

def alinea_texto(texto: str, anchura: int, alineación: str="derecha") -> str:
    if alineación in ['derecha', '>']:
        return f"{texto:>{anchura}}"[-anchura:]
    elif alineación in ['centro', '^']:
        return f"{texto[:anchura]:^{anchura}}"
    elif alineación in ['izquierda', '<']:
        return f"{texto:<{anchura}}"[:anchura]
    else:
        return f"{trunca_texto('ALINEACION ERRADA', anchura)}"


# ---------------------------------------------------------

def edita_número(valor, num_decimals: int=2, anchura=None, alineación: str="derecha", sufijo: str=""):
    if valor is None or isinstance(valor, str):
        return valor
    valor_editado = ''.join([
                                locale.format_string(f'%.{num_decimals}f', valor, grouping=True, monetary=True),
                                sufijo
                            ])
    return alinea_texto(valor_editado, anchura=anchura, alineación=alineación) if anchura else valor_editado


# ---------------------------------------------------------

def trunca_texto(texto: str, max_width: int) -> str:
    last_space = texto[:max_width+1].rfind(' ')
    if last_space > max_width - 3:
        texto = texto[:last_space].strip()
    return texto[0: max_width - 3] + '...' if len(texto) > max_width else texto


# ---------------------------------------------------------

def espacios(width: int=1, char: str=' ') -> str:
    return char * width


# ---------------------------------------------------------

_wrapper = None

def bloque_de_texto(texto, anchura=80, margen=0, margen_1ra_linea=False, continuacion="  ",
                           formato_lista=False, separa_palabras=False):
    global _wrapper

    if _wrapper is None:
        _wrapper = MyTextWrapper()
    _wrapper.width = anchura
    _wrapper.initial_indent = espacios(margen) if margen_1ra_linea else ''
    _wrapper.subsequent_indent = espacios(margen) + continuacion
    _wrapper.separa_palabras = separa_palabras

    return _wrapper.wrap(texto) if formato_lista else _wrapper.fill(texto)


# ---------------------------------------------------------

def is_numeric(valor) -> bool:
    return isinstance(valor, numbers.Number)


# ---------------------------------------------------------

def valida_codigo_seguridad(recibo, fecha: datetime, codigo: str) -> bool:
    nro_recibo = f"{recibo:0{GyG_constantes.long_num_recibo}d}" if isinstance(recibo, int) else recibo
    str_fecha = fecha.strftime("%d/%m/%Y")
    return str(_obscure(bytes(' '.join([nro_recibo, str_fecha]), 'UTF-8')), 'UTF-8') == codigo


# ---------------------------------------------------------

def remueve_acentos(texto: str) -> str:
    # -> NFD y eliminar diacríticos
    s = re.sub(
            r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1", 
            normalize( "NFD", texto), 0, re.I
        )
    # -> NFC
    return normalize( 'NFC', s)


# ---------------------------------------------------------

def separa_meses(mensaje: str, muestra_modificador: bool=False, as_string: bool=False):     # -> list | str
 
    def convierte_separador(mensaje: str) -> str:
        # Convierte '-' en '/' en expresiones del tipo «mes»-«año»
        pattern = '([a-zA-Z\\.])-([0-9])'
        while True:
            m = re.search(pattern, mensaje)
            if m is None:
                return mensaje
            mensaje = re.sub(pattern, m.group(1) + "/" + m.group(2), mensaje, count=1)

    tokens_validos = GyG_constantes.meses + GyG_constantes.meses_abrev + GyG_constantes.conectores
    if muestra_modificador:
        tokens_validos += GyG_constantes.modificadores
    mensaje = convierte_separador(mensaje)
    mensaje = re.sub("\([^()]*\)", "", mensaje)
    # mensaje = re.sub(r"\W ", " ", mensaje.lower().replace('-', ' a ').replace('/', ' ')).split()
    mensaje = mensaje.lower().replace('-', ' a ').replace('/', ' ')
    if muestra_modificador:
        for token in GyG_constantes.textos_anticipos:
            mensaje.replace(token, GyG_constantes.modificadores[0])
        for token in GyG_constantes.textos_saldos:
            mensaje.replace(token, GyG_constantes.modificadores[1])
    mensaje = re.sub(r"\W ", " ", mensaje).split()
    mensaje_ed = [x for x in mensaje if (x in tokens_validos) or x.isdigit()]
    # print(f"DEBUG: {mensaje_ed = }")
    initial_year = last_year = None
    initial_month = last_month = None
    mensaje_anterior = None
    mensaje_final = list()
    maneja_conector = False

    for x in reversed(mensaje_ed):
        token = GyG_constantes.meses[GyG_constantes.meses_abrev.index(x)] if x in GyG_constantes.meses_abrev else x
        # print(f"----------\n{token = }, es conector = {maneja_conector}\n" + \
        #       f"Año: inicial = {initial_year}, final = {last_year}; " + \
        #       f"Mes: inicial = {initial_month}, final = {last_month}\n" + \
        #       f"{mensaje_anterior = },\n{mensaje_final =    }")
        if token.isdigit():
            # print(f" -> {token}: AÑO")
            if mensaje_anterior != None:
                mensaje_final = mensaje_anterior + mensaje_final
                mensaje_anterior = None
            if maneja_conector:
                initial_year = token
            else:
                last_year = token
            continue
        elif token in GyG_constantes.meses:
            # print(f" -> {token}: MES")
            if mensaje_anterior != None:
                mensaje_final = mensaje_anterior + mensaje_final
                mensaje_anterior = None
            if maneja_conector:
                if initial_year is None:
                    initial_year = last_year
                n_last_month = GyG_constantes.meses.index(last_month)
                initial_date = datetime(int(initial_year), GyG_constantes.meses.index(token) + 1, 1)
                curr_date = datetime(int(last_year), GyG_constantes.meses.index(last_month) + 1, 1)
                while curr_date > initial_date:
                    curr_date = curr_date - relativedelta(months=1)
                    mensaje_final = [curr_date.strftime("%m-%Y")] + mensaje_final
                    last_year = initial_year
                maneja_conector = False
                continue
            last_month = token
            mensaje_anterior = [f"{GyG_constantes.meses.index(last_month)+1:02d}-{last_year}"]
        elif x in GyG_constantes.conectores:
            # print(f" -> {token}: CONECTOR")
            maneja_conector = True
        elif x in GyG_constantes.modificadores and muestra_modificador:
            # print(f" -> {token}: MODIFICADOR")
            mensaje_final = [f"{GyG_constantes.meses.index(last_month)+1:02}-{last_year} {x}"] + mensaje_final
            mensaje_anterior = None

    if mensaje_anterior is not None:
        mensaje_final = mensaje_anterior + mensaje_final

    if as_string:
        mensaje_final = '|'.join(mensaje_final)

    return mensaje_final


# ---------------------------------------------------------

def reagrupa_meses(mensaje: str, mes_completo: bool=False) -> str:
    strMeses = GyG_constantes.meses if mes_completo else GyG_constantes.meses_abrev
    separador = ' ' if mes_completo else '-'

    def une_lista(lista):
        if not lista:
            return ''
        if len(lista) == 1:
            return lista[0]
        conjunción = ", " if " y " in lista[-1] else " y "
        return f"{', '.join(lista[:-1])}{conjunción}{lista[-1]}"

    def procesa_meses(fecha_1, fecha_2, año):
        if fecha_1 is None or fecha_2 is None:
            return año
        if fecha_1 == fecha_2:
            lista_meses.insert(0, ''.join([
                        strMeses[int(fecha_1[-2:])-1],
                        f"{separador}{fecha_1[:4]}" if fecha_1[:4] != año else "",
                ]))
        else:
            lista_meses.insert(0, ''.join([
                    strMeses[int(fecha_1[-2:])-1],
                    f"{separador}{fecha_1[:4]}" if fecha_1[:4] != fecha_2[:4] else "",
                    " y " if son_consecutivos(fecha_1, fecha_2) else " a ",
                    strMeses[int(fecha_2[-2:])-1],
                    f"{separador}{fecha_2[:4]}" if fecha_2[:4] != año else "",
                ]))
        return fecha_1[:4]

    def son_consecutivos(mes_inicial, mes_final) -> bool:
        if mes_inicial is None or mes_final is None:
            return False
        return datetime.strptime(mes_inicial, "%Y-%m") + relativedelta(months=1) == \
                    datetime.strptime(mes_final, "%Y-%m")

    # obtiene un string del tipo "05-2019|06-2019|07-2019|07-2019" con posibles duplicados
    meses = separa_meses(mensaje, muestra_modificador=True, as_string=True)
    # obtiene los grupos separados por '|' y elimina los duplicados
    meses = list(set(meses.split('|')))
    # convierte al formato 'yyyy-mm' y lo ordena de menor a mayor
    # meses = [f[3:] + '-' + f[:2] for f in reversed(meses)]
    meses = [f"{f[3:7]}-{f[:2]}{f[7:]}" for f in reversed(meses)]
    meses.sort()

    # convierte los meses al tipo "may., jun. y jul. 2019"
    lista_meses = list()
    último_año = None
    mes_actual = mes_siguiente = último_mes = None
    en_consecutivos = False
    for mes in reversed(meses):
        fecha = mes[:7]
        prefijo = mes[8:]
        año = fecha[:4]
        if len(prefijo) > 0:
            prefijo += ' '
            último_año = procesa_meses(mes_siguiente, último_mes, último_año)
            lista_meses.insert(0, ''.join([
                    f"{prefijo}{strMeses[int(fecha[-2:])-1]}",
                    f"{separador}{año}" if año != último_año else ''
                ]))
            último_año = año
            último_mes = mes_siguiente = None
        elif último_mes is not None:
            if son_consecutivos(mes, mes_siguiente):
                mes_siguiente = mes
            else:
                último_año = procesa_meses(mes_siguiente, último_mes, último_año)
                mes_siguiente = último_mes = mes
        else:
            mes_siguiente = último_mes = mes
        mes_actual = fecha

    último_año = procesa_meses(mes_actual, último_mes, último_año)

    return une_lista(lista_meses)


# ---------------------------------------------------------

def MontoEnLetras(número: float, mostrar_céntimos: bool=True, céntimos_en_letras: bool=False, moneda: str='Bolívar') -> str:
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

    def _número_recursivo(número) -> str:
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


# ---------------------------------------------------------

def genera_recibo(r, sella_recibo: bool=False, codigo_de_seguridad: bool=False, codigo_texto: bool=True):
    """ genera_recibo(r, sella_recibo=False)
        Recibe como parámetro con las siguientes claves:
          . 'Nro. Recibo',  'Fecha',    'Beneficiario',
          . 'Dirección',    'Monto',    'Concepto',
          . 'Categoría',    'Monto $',  '$'
    """

    # Librerías
    import GyG_constantes
    from PIL import Image, ImageFont, ImageDraw, ImageEnhance

    # En caso de querer reinstalar la librería para generar códigos de barra, ejecutar:
    #     python3 -m pip install python-barcode
    #
    # import barcode
    # from barcode.writer import ImageWriter

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
    dejavusansmono      = os.path.join(GyG_constantes.rec_fuentes, 'DejaVuSansMono.ttf')

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
        # font = ImageFont.truetype(font=calibri, size=10)
        font = ImageFont.truetype(font=dejavusansmono, size=10)
        código_a_convertir = f"{r['Nro. Recibo']:0{GyG_constantes.long_num_recibo}d} " + \
                             r['Fecha'].strftime('%d/%m/%Y')
        código_convertido = str(_obscure(bytes(código_a_convertir, 'UTF-8')), 'UTF-8')
        # ancho_código = anchura_de_texto(código_convertido, font)
        # print(f"DEBUG: {código_convertido = }")

        # ----------------------------------------------------------------
        # CODIGO DE VALIDACION HORIZONTAL (justificado a la izquierda)
        #
        # if codigo_texto:
        recibo.text(xy=(20, plantilla.size[1] - 25),
                    text=código_convertido, font=font, fill='black')
        # else:
        #     # Cambiado:  module_width: 0.065, module_height: 1.5, font_size: 6, text_distance: 0.5,
        #     #            posición Y: 60
        #     codigo_de_barras = barcode.get('code128', código_convertido, writer=ImageWriter())
        #     opciones = {
        #         'module_width': 0.085,  # Defaults to  0.2; min = 0.084666667 -> 0.085
        #         'module_height': 1.0,   # Defaults to 15.0
        #         'quiet_zone': 1.0,      # Defaults to  6.5
        #         'font_size': 0,         # Defaults to 10
        #         'text_distance': 0.5,   # Defaults to  5.0
        #     }
        #     plantilla.paste(codigo_de_barras.render(writer_options=opciones),
        #                     (20, plantilla.size[1] - 30))

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


# ---------------------------------------------------------

vocales = ['a', 'á', 'e', 'é', 'i', 'í', 'o', 'ó', 'u', 'ú', 'ü', \
           'A', 'Á', 'E', 'É', 'I', 'Í', 'O', 'Ó', 'U', 'Ú', 'Ü']

class MyTextWrapper(TextWrapper):
    """
    Divide en sílabas la palabra al final de la línea a fin de lograr un mejor ajuste
    """

    def __init__(self, *kargs, separa_palabras=False):
        self.separa_palabras = separa_palabras
        self._silabeador = Silabeador() # if self.separa_palabras else None
        TextWrapper.__init__(self, *kargs)

    def _wrap_chunks(self, chunks):
        """_wrap_chunks(chunks : [string]) -> [string]

        Wrap a sequence of text chunks and return a list of lines of
        length 'self.width' or less.  (If 'break_long_words' is false,
        some lines may be longer than this.)  Chunks correspond roughly
        to words and the whitespace between them: each chunk is
        indivisible (modulo 'break_long_words'), but a line break can
        come between any two chunks.  Chunks should not have internal
        whitespace; ie. a chunk is either all whitespace or a "word".
        Whitespace chunks will be removed from the beginning and end of
        lines, but apart from that whitespace is preserved.
        """

        lines = []
        if self.width <= 0:
            raise ValueError("invalid width %r (must be > 0)" % self.width)
        if self.max_lines is not None:
            if self.max_lines > 1:
                indent = self.subsequent_indent
            else:
                indent = self.initial_indent
            if len(indent) + len(self.placeholder.lstrip()) > self.width:
                raise ValueError("placeholder too large for max width")

        # Arrange in reverse order so items can be efficiently popped
        # from a stack of chucks.
        chunks.reverse()

        while chunks:

            # Start the list of chunks that will make up the current line.
            # cur_len is just the length of all the chunks in cur_line.
            cur_line = []
            cur_len = 0

            # Figure out which static string will prefix this line.
            if lines:
                indent = self.subsequent_indent
            else:
                indent = self.initial_indent

            # Maximum width for this line.
            width = self.width - len(indent)

            # First chunk on line is whitespace -- drop it, unless this
            # is the very beginning of the text (ie. no lines started yet).
            if self.drop_whitespace and chunks[-1].strip() == '' and lines:
                del chunks[-1]

            while chunks:
                l = len(chunks[-1])

                # Can at least squeeze this chunk onto the current line.
                if cur_len + l <= width:
                    cur_line.append(chunks.pop())
                    cur_len += l

                # Nope, this line is full.
                else:
                    break

            # The current line is full, and the next chunk is too big to
            # fit on *any* line (not just this one).
            if chunks and len(chunks[-1]) > width:
                self._handle_long_word(chunks, cur_line, cur_len, width)
                cur_len = sum(map(len, cur_line))

            # Verificar si se pueden agregar sílabas al final de la línea
            # (ignorarlo si se trata de un número)
            if chunks and self.separa_palabras:
                silabas = self._silabeador(chunks[-1])

                # Separación de palabras al final de una linea. Criterios a aplicar:
                #   √  La división se realiza entre sílabas, nunca se separan las letras de una misma sílaba
                #   √  Nunca se separa la primera sílaba si ésta es una sola vocal, ni la última sílaba si
                #      también es una sola vocal
                #   √  Tampoco se separan dos vocales juntas aunque pertenezcan a distintas sílabas (hiato)
                if len(silabas) > 1 and len(silabas[0]) == 1:
                    silabas[1] = ''.join(silabas[:2])
                    _ = silabas.pop(0)
                if len(silabas) > 1 and len(silabas[-1]) == 1:
                    silabas[-2] = ''.join(silabas[-2:])
                    _ = silabas.pop()

                silabas_copia = [_ for _ in silabas]
                for idx in reversed(range(len(silabas) - 1)):
                    if silabas[idx][-1] in vocales and silabas[idx + 1][0] in vocales:
                        silabas_copia[idx] = silabas_copia[idx] + silabas_copia[idx + 1]
                        _ = silabas_copia.pop(idx+1)
                silabas = silabas_copia

                silabas.reverse()
                any_pop = False
                while silabas:
                    l = len(silabas[-1])
                    if cur_len + l + 1 <= width:
                        cur_line.append(silabas.pop())
                        cur_len += l
                        any_pop = True
                    else:
                        if any_pop:
                            cur_line.append('-')
                            cur_len += 1
                            silabas.reverse()
                            chunks[-1] = ''.join(silabas)
                        break

            # If the last chunk on this line is all whitespace, drop it.
            if self.drop_whitespace and cur_line and cur_line[-1].strip() == '':
                cur_len -= len(cur_line[-1])
                del cur_line[-1]

            if cur_line:
                if (self.max_lines is None or
                    len(lines) + 1 < self.max_lines or
                    (not chunks or
                     self.drop_whitespace and
                     len(chunks) == 1 and
                     not chunks[0].strip()) and cur_len <= width):
                    # Convert current line back to a string and store it in
                    # list of all lines (return value).
                    lines.append(indent + ''.join(cur_line))
                else:
                    while cur_line:
                        if (cur_line[-1].strip() and
                            cur_len + len(self.placeholder) <= width):
                            cur_line.append(self.placeholder)
                            lines.append(indent + ''.join(cur_line))
                            break
                        cur_len -= len(cur_line[-1])
                        del cur_line[-1]
                    else:
                        if lines:
                            prev_line = lines[-1].rstrip()
                            if (len(prev_line) + len(self.placeholder) <=
                                    self.width):
                                lines[-1] = prev_line + self.placeholder
                                break
                        lines.append(indent + self.placeholder.lstrip())
                    break

        return lines


# ---------------------------------------------------------

# Silabeador: Basado en Mabodo's ipython notebook (https://github.com/mabodo/sibilizador)

class Char_Line():
    def __init__(self, word):
        self.word = word
        self.char_line = [(char, self.char_type(char)) for char in word]
        self.type_line = ''.join(chartype for char, chartype in self.char_line)
        
    def char_type(self, char):
        if char in set(['a', 'á', 'e', 'é', 'o', 'ó', 'í', 'ú']):
            return 'V' #strong vowel
        if char in set(['i', 'u', 'ü']):
            return 'v' #week vowel
        if char == 'x':
            return 'x'
        if char == 's':
            return 's'
        else:
            return 'c'
            
    def find(self, finder):
        return self.type_line.find(finder)
        
    def split(self, pos, where):
        return Char_Line(self.word[0: pos+where]), Char_Line(self.word[pos + where: ])
    
    def split_by(self, finder, where):
        split_point = self.find(finder)
        if split_point != -1:
            chl1, chl2 = self.split(split_point, where)
            return chl1, chl2
        return self, False
     
    def __str__(self):
        return self.word
    
    def __repr__(self):
        return repr(self.word)


class Silabeador():
    def __init__(self):
        self.grammar = []
        
    def split(self, chars):
        rules  = [
                    ('VV', 1), ('cccc', 2), ('xcc', 1), ('ccx', 2), ('csc', 2), ('xc', 1), ('cc', 1), ('vcc', 2), 
                    ('Vcc', 2), ('sc', 1), ('cs', 1), ('Vc', 1), ('vc', 1), ('Vs', 1), ('vs', 1)
                 ]
        for split_rule, where in rules:
            first, second = chars.split_by(split_rule, where)
            if second:
                if first.type_line in set(['c', 's', 'x', 'cs']) or second.type_line in set(['c', 's', 'x', 'cs']):
                    #'skip1', first.word, second.word, split_rule, chars.type_line
                    continue
                if first.type_line[-1] == 'c' and second.word[0] in set(['l', 'r']):
                    continue
                if first.word[-1] == 'l' and second.word[-1] == 'l':
                    continue
                if first.word[-1] == 'r' and second.word[-1] == 'r':
                    continue
                if first.word[-1] == 'c' and second.word[-1] == 'h':
                    continue
                return self.split(first) + self.split(second)
        return [chars]
        
    def __call__(self, word):

        def es_numerico(texto):
            """
            Utilitario para verificar si el string recibido corresponde
            a un valor numérico luego de haber eliminado los caracteres
            de formateo
            """
            for ch in ['.', ',', '+', '-', '(', ')']:
                if ch in texto:
                    texto = texto.replace(ch, '')
            return texto.isnumeric()

        if es_numerico(word):
            return [word]
        char_line_word = self.split(Char_Line(word))
        str_line_word = [str(_) for _ in char_line_word]
        return str_line_word
