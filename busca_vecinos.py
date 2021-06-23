# GyG CARTELERA VIRTUAL
#
# Elabora la Cartelera Virtual a partir de la hoja de cálculo de pagos

"""
    PENDIENTE POR HACER
    -   Convertir los caracteres dobles ('tt', 'ss', 'nn', 'cc', etc.) en sencillos

    NOTAS
    -   ¿Por qué la diferencia en el tratamiento de 'lopez'?
                Vecinos a buscar: xavier lopez
                DEBUG: resultado = {49, 110} (<class 'set'>)
                 - Familia González López
                 - Familia López Hermoso
                Vecinos a buscar: luis lopez
                DEBUG: resultado = set() (<class 'set'>)
                Vecinos a buscar: 
         -> En 'Xavier López', 'Xavier' no existe, por lo que el resultado corresponde a 'López'
            En 'Luis López', 'Luis' existe, y 'López' también, pero no la combinación de ambos
             -> ¿Debería hacerse la unión de resultados, y no la intersección?

         -> Probar con una selección "inteligente": si la «intersección» de resultados genera el
            conjunto vacío, hacer la «unión» de los mismos

    HISTORICO
    -   

"""

print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, NamedStyle
from pandas import read_excel, isnull, notnull
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

import sys
import re
import string
from unicodedata import normalize
import locale

dummy = locale.setlocale(locale.LC_ALL, 'es_es')


#
# CONSTANTES ============================================================================
# 

excel_workbook     = GyG_constantes.pagos_wb_estandar             # '1.1. GyG Recibos.xlsm'
# excel_workbook     = "1.1._GyG_Recibos_TEST.xlsm"
excel_worksheet    = GyG_constantes.pagos_ws_resumen_reordenado   # 'R.VIGILANCIA (reordenado)'
excel_resumen      = GyG_constantes.pagos_ws_resumen              # 'RESUMEN VIGILANCIA'
excel_pagos        = GyG_constantes.pagos_ws_vigilancia           # 'Vigilancia'

_índice_vecinos    = None
_índice_inverso    = None
_índice_dirección  = None


#
# RUTINAS UTILITARIAS ===================================================================
#

def analiza_texto(texto: str) -> list:
    STOPWORDS = ['familia', 'calle', 'av', 'avenida', 'numero', 'nro', 'numeros', 'nros', 'y', 'el']
    CVT_SEQUENCES = {'z': 's', 'h': '',
                     'ss': 's', 'cc': 'c', 'll': 'l', 'tt': 't', 'nn': 'n', 'bb': 'b'}

    def remueve_acentos(texto: str) -> list:
        # -> NFD y eliminar diacríticos
        s = re.sub(
                r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1", 
                normalize('NFD', texto), 0, re.I
            )
        # -> NFC
        return normalize('NFC', s)

    def remueve_stopwords(lista: list) -> list:
        return [token for token in lista if token not in STOPWORDS]

    def elimina_secuencias(texto: str) -> str:
        for key, value in CVT_SEQUENCES.items():
            texto = texto.replace(key, value)
        return texto

    # Remueve acentos, convierte a minúsculas, elimina signos de puntuación, convierte 'z's en 's's,
    # elimina las 'h's y convierte las secuencias dobles (s, c, l, t, n) en sencillas
    texto = remueve_acentos(texto) \
        .lower() \
        .translate(str.maketrans('', '', string.punctuation))
    lista = elimina_secuencias(texto) \
        .split()
    return remueve_stopwords(lista)


def indexa_listado_de_vecinos() -> list:
    global _índice_vecinos, _índice_inverso, _índice_dirección

    último_día_del_mes = datetime.today() + relativedelta(months=1, day=1) - relativedelta(days=1)

    # Abre la hoja de cálculo de Recibos de Pago
    df_worksheet = read_excel(excel_workbook, sheet_name=excel_worksheet)

    # Elimina los registros que no tienen una categoría definida
    df_worksheet.dropna(subset=['Categoría'], inplace=True)

    # Ajusta las columnas F.Desde y F.Hasta en aquellos en los que estén vacíos:
    # 01/01/2016 y 01/mes_siguiente/año
    df_worksheet.loc[df_worksheet[isnull(df_worksheet['F.Desde'])].index, 'F.Desde'] = datetime(2016, 1, 1)
    df_worksheet.loc[df_worksheet[isnull(df_worksheet['F.Hasta'])].index, 'F.Hasta'] = \
                                                                 último_día_del_mes + relativedelta(days=1)

    # Elimina aquellos vecinos que compraron (o se iniciaron) en fecha posterior
    # o que vendieron (o cambiaron su razón social) en fecha anterior a la fecha
    # de análisis
    df_worksheet = df_worksheet[df_worksheet['F.Desde'] < último_día_del_mes]
    df_worksheet = df_worksheet[df_worksheet['F.Hasta'] >= último_día_del_mes]

    # Cambia el nombre de la columna 2016 a datetimme(2016, 1, 1)
    df_worksheet.rename(columns={2016:datetime(2016, 1, 1)}, inplace=True)

    # Prepara lista de vecinos y el índice inverso del mismo
    _índice_vecinos = list()
    _índice_inverso = dict()
    _índice_dirección = list()

    vecino_ID = 0
    for _, r in df_worksheet.iterrows():
        texto_a_analizar = ''.join([
            r['Beneficiario'],
            '. ',
            r['Dirección']
        ])
        for token in analiza_texto(texto_a_analizar):
            if token not in _índice_inverso:
                _índice_inverso[token] = set()
            _índice_inverso[token].add(vecino_ID)
        vecino = r['Beneficiario']
        if vecino not in _índice_vecinos:
            _índice_vecinos.append(vecino)
            _índice_dirección.append(r['Dirección'])
            vecino_ID += 1

    # return _índice_vecinos, _índice_inverso


def busca_en_índices(texto: str, modo: str="AND") -> list:
    lista_generada = analiza_texto(texto)
    resultado = None
    for token in lista_generada:
        if token in _índice_inverso:
            if resultado is None:
                resultado = _índice_inverso[token]
            elif modo == "AND":
                resultado = resultado.intersection(_índice_inverso[token])
            elif modo == "OR":
                resultado.union(_índice_inverso[token])
            elif modo == "Smart":
                if resultado.intersection(_índice_inverso[token]) == set():
                    resultado = resultado.union(_índice_inverso[token])
                else:
                    resultado = resultado.intersection(_índice_inverso[token])
            else:
                raise ValueError(f"busca_en_indices(): error en modo: {modo}")

    # print(f"DEBUG: {resultado = } ({type(resultado)})")
    if resultado:
        resultado = [_índice_vecinos[idx] + ', ' + _índice_dirección[idx] for idx in resultado]
    elif resultado == set():
        resultado = None

    return resultado


def maneja_búsqueda(a_buscar: str):
    resultado = busca_en_índices(a_buscar, modo="Smart")
    if resultado is None:
        print(" * *** Búsqueda sin resultados")
    else:
        for vecino in resultado:
            print(f" - {vecino}")
    print()


#
# PROCESO ===============================================================================
#

# Genera el listado de vecinos y su índice inverso
print('Indexando lista de vecinos...')
indexa_listado_de_vecinos()

argc = len(sys.argv[1:])

if argc > 0:
    a_buscar = ' '.join(sys.argv[1:])
    print(f"\nVecinos a buscar: {a_buscar}")
    maneja_búsqueda(a_buscar)
else:
    print('Busca vecinos...\n')
    while True:
        a_buscar = input("Vecinos a buscar: ")
        if a_buscar:
            maneja_búsqueda(a_buscar)
        else:
            break
if argc == 0:
    print()
