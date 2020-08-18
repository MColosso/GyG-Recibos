# GyG CAMBIO DE CATEGORÍA
#
# 

"""
    POR HACER
    -   


    HISTORICO
    -   Versión inicial (07/06/2020)


    LAYOUT
        GyG PROPUESTAS DE CAMBIO DE CATEGORÍA 
        al {datetime.now():%d/%m/%Y %I:%M} {am_pm}

        Vecino               | Dirección      | Categoría actual | Propuesta
        ----------------------------------------------------------------------------
        xxxxxxxxxxxxxxxxxxxx | xxxxxxxxxxxxxx | xxxxxxxxxxxxxxxx | xxxxxxxxxxxxxxxx
                             | Último pago: anticipo jun/2020 (dd/mm/yyyy)
                             | --- No tiene pagos registrados

#       Vecino               | Dirección      | Último pago        | Categoría act.| Propuesta
#       -------------------------------------------------------------------------------------------
#       xxxxxxxxxxxxxxxxxxxx | xxxxxxxxxxxxxx | xxxxxxxxxxxxxxxxxx | xxxxxxxxxxxxx | xxxxxxxxxxxxx

"""


print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *
from pandas import read_excel, isnull
from numpy import NaN
from datetime import datetime
from dateutil.relativedelta import relativedelta
import swifter

import sys
import os
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

# Define algunas constantes
nombre_propuestas   = GyG_constantes.txt_cambios_de_categoría   # "GyG Cambio de Categorías {:%Y-%m (%B)}.txt"
attach_path         = GyG_constantes.ruta_cambios_de_categoría  # "../GyG Archivos/Otros"

excel_workbook      = GyG_constantes.pagos_wb_estandar          # '1.1. GyG Recibos.xlsm'
excel_ws_vigilancia = GyG_constantes.pagos_ws_vigilancia
excel_ws_resumen_r  = GyG_constantes.pagos_ws_resumen_reordenado

nMeses              = 5  # Meses a analizar para propuestas de cambio de categoría

patron_detalle = '{:<20} | {:<14} | {:<16} | {:<16}\n' + \
                 espacios(21) + '| {}\n'

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'

VERBOSE = False


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

# a_evaluar = ['Familia Triana González', ]

def genera_propuesta_categoría():

    def genera_propuesta(r):
        # comparar <r> con <cuotas_mensuales> en los últimos <nMeses> meses
        monto_cuotas = [cuotas_obj.cuota_vigente(r['Beneficiario'], mes_de_referencia - relativedelta(months=mes+1)) \
                                for mes in reversed(range(nMeses))]
        pagos = [r[mes_de_referencia - relativedelta(months=mes+1)] for mes in reversed(range(nMeses))]
        pagos = [p if is_numeric(p) else NaN for p in pagos]

        # No evaluable: menos de nMeses desde su inicio en la Asociación
        # print(f"DEBUG: {r.index = }, {no_evaluables = }")
        if r['F.Desde'] > mes_inicial:
            propuesta = 'No evaluable (*)'
        # Cuota completa: todos los pagos son mayores o iguales a la cuota del mes
        elif all(p >= m for p, m in zip(pagos, monto_cuotas)):
            propuesta = 'Cuota completa'
        # Colaboración: todos los pagos son inferiores a la cuota del mes
        elif all(p <  m for p, m in zip(pagos, monto_cuotas)):
            propuesta = 'Colaboración'
        # No colabora: todos los pagos son nulos
        elif all(isnull(p) for p in pagos):
            propuesta = 'No participa'
        # En cualquier otro caso, el resultado es indeterminado
        else:
            propuesta = ''
        if r['Categoría'] == propuesta:
            propuesta = ''
        if propuesta in ['Cuota completa', 'Colaboración']:
            propuesta = propuesta.upper()
        else:
            propuesta = propuesta.lower()
        # if r['Beneficiario'] in a_evaluar:
        #     print(f"Categoría: {r['Categoría']}, Propuesta de cambio: {propuesta}")
        return propuesta

    df = df_resumen_r[['Beneficiario', 'Dirección', 'Categoría']].copy()
    df['Propuesta'] = df_resumen_r.swifter.progress_bar(False).apply(genera_propuesta, axis=1)

    return df

# ----------------------------------------------------------------------------
nombre_meses     = ['enero',      'febrero', 'marzo',     'abril',
                    'mayo',       'junio',   'julio',     'agosto',
                    'septiembre', 'octubre', 'noviembre', 'diciembre']
meses_abrev      = ['ene', 'feb', 'mar', 'abr', 'may', 'jun',
                    'jul', 'ago', 'sep', 'oct', 'nov', 'dic']
conectores       = ['a', '-']
textos_anticipos = ['adelanto', 'anticipo'   ]
textos_saldos    = ['ajuste',   'complemento', 'diferencia', 'saldo']
modificadores    = ['anticipo', 'saldo']

tokens_validos = nombre_meses + meses_abrev + conectores

def separa_meses(mensaje, as_string=False, muestra_modificador=False):
    import re
    
    tokens_validos = nombre_meses + meses_abrev + conectores + modificadores

    mensaje = re.sub("\([^()]*\)", "", mensaje)
    mensaje = mensaje.lower().replace('-', ' a ').replace('/', ' ')
    for token in textos_anticipos:
        mensaje.replace(token, modificadores[0])
    for token in textos_saldos:
        mensaje.replace(token, modificadores[1])
    mensaje = re.sub(r"\W ", " ", mensaje).split()
    mensaje_ed = [x for x in mensaje if (x in tokens_validos) or x.isdigit()]
    last_year = None
    last_month = None
    acción = ''
    mensaje_anterior = None
    mensaje_final = list()
    maneja_conector = False
    for x in reversed(mensaje_ed):
        token = nombre_meses[meses_abrev.index(x)] if x in meses_abrev else x
        if token.isdigit():
            if mensaje_anterior != None:
                mensaje_final = mensaje_anterior + mensaje_final
            last_year = token
            last_month = None
            mensaje_anterior = None
        elif token in nombre_meses:
            if mensaje_anterior != None:
                mensaje_final = mensaje_anterior + mensaje_final
            if maneja_conector:
                try:
                    n_last_month = nombre_meses.index(last_month)
                except:
                    continue    # ignora los mensajes que contienen textos del tipo:
                                # "(saldo a favor: Bs. 69.862,95)"
                n_token = nombre_meses.index(token)
                for t in reversed(range(n_token + 1, n_last_month)):
                    mensaje_final = [f"{t+1:02}-{last_year}"] + mensaje_final
                maneja_conector = False
            last_month = token
            mensaje_anterior = [f"{nombre_meses.index(last_month)+1:02}-{last_year}"]
        elif x in conectores:
            maneja_conector = True
        elif x in modificadores and muestra_modificador:
            mensaje_final = [f"{nombre_meses.index(last_month)+1:02}-{last_year} {x}"] + mensaje_final
            mensaje_anterior = None

    if mensaje_anterior != None:
        mensaje_final = mensaje_anterior + mensaje_final

    if as_string:
        mensaje_final = '|'.join(mensaje_final)

    return mensaje_final
# ----------------------------------------------------------------------------

def edita_beneficiario(beneficiario):
    return beneficiario.replace('Familia ', '')

def edita_dirección(dirección):
    return dirección.replace('Calle ', '').replace('Nros. ', '').replace('Nro. ', '')

def edita_categoría(categoría):
    return categoría

def edita_último_pago(beneficiario):
    r = df_vigilancia[df_vigilancia['Beneficiario'] == beneficiario]
    if r.shape[0] == 0:
        return '--- No tiene pagos registrados'
    r = r.iloc[-1]
    fecha = r['Fecha'].strftime('%d/%m/%Y')
    mes = separa_meses(r['Concepto'], as_string=False, muestra_modificador=True)[-1].split()
    if len(mes) == 1: mes.append('')
    último_pago = f"{mes[1]}{' ' if len(mes[1])>0 else ''}{meses_abrev[int(mes[0][:2])-1]}{mes[0][2:]}"
    return f'Último pago: {último_pago} ({fecha})'

def dif_meses(d1, d2):
    return (d1.year - d2.year) * 12 + d1.month - d2.month


#
# PROCESO
#

fecha_de_referencia = datetime.now()
mes_de_referencia = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
mes_anterior = mes_de_referencia - relativedelta(months=1)
mes_inicial = mes_de_referencia - relativedelta(months=nMeses)
f_ref_último_día = mes_de_referencia + relativedelta(months=1, days=-1)


print()

# Número de meses a analizar para propuestas de cambio de categoría
nMeses = input_valor('Número de meses a analizar', nMeses, toma_opciones_por_defecto)

# Selecciona si se muestran sólo vecinos con propuestas de cambio o no
sólo_propuestas = input_si_no('Sólo vecinos con propuestas de cambio', 'sí', toma_opciones_por_defecto)

# # Selecciona si se ordenan alfabéticamente los vecinos
ordenado = input_si_no('Ordenados alfabéticamente', 'no', toma_opciones_por_defecto)


# Abre la hoja de cálculo de Recibos de Pago
print()
print(f'Cargando hoja de cálculo "{excel_workbook}"...')

cuotas_obj = Cuota(excel_workbook)

if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los pagos recibidos')
df_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

# Selecciona los pagos de vigilancia entre la fecha de referencia y num_meses_cuotas_equivalentes atrás
df_vigilancia = df_vigilancia[(df_vigilancia['Categoría'] == 'Vigilancia')  & \
                              (df_vigilancia['Fecha'] <  mes_de_referencia) & \
                              (df_vigilancia['Fecha'] >= datetime(2017, 1, 1))]
df_vigilancia = df_vigilancia[['Beneficiario', 'Fecha', 'Monto', 'Concepto', 'Día', 'Mes']]

if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Agrega la cuota vigente para la fecha de pago')
df_vigilancia['Cuota'] = df_vigilancia.apply(
                                lambda r: cuotas_obj.cuota_vigente(r['Beneficiario'], r['Fecha']), axis=1)
df_vigilancia['Num. Cuotas'] = df_vigilancia.apply(
                                lambda r: float(r['Monto']) / float(r['Cuota']), axis=1)


if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee el resumen de pago reordenado')
df_resumen_r = read_excel(excel_workbook, sheet_name=excel_ws_resumen_r)

# Ajusta las columnas F.Desde y F.Hasta en aquellos en los que estén vacíos:
# 01/01/2016 y día/mes/año+1
df_resumen_r.loc[df_resumen_r[df_resumen_r['F.Desde'].isnull()].index, 'F.Desde'] = datetime(2016, 1, 1)
df_resumen_r.loc[df_resumen_r[df_resumen_r['F.Hasta'].isnull()].index, 'F.Hasta'] = fecha_de_referencia + relativedelta(years=1)

# Elimina aquellos vecinos que compraron (o se iniciaron) en fecha posterior a la fecha de análisis,
# aquellos que vendieron (o cambiaron su razón social) en fecha anterior a la misma y aquellos que
# no tienen categoría asociada
df_resumen_r = df_resumen_r[(df_resumen_r['F.Desde'] <  f_ref_último_día) & \
                            (df_resumen_r['F.Hasta'] >= f_ref_último_día) & \
                            (df_resumen_r['Categoría'].notnull())]

no_evaluables = df_resumen_r[df_resumen_r['F.Desde'] > mes_inicial].index.to_list()

print('Generando propuestas de cambio de categoría...')

df_resumen_r = genera_propuesta_categoría()

if sólo_propuestas:
    df_resumen_r = df_resumen_r[df_resumen_r['Propuesta'] != '']

if ordenado:
    df_resumen_r.sort_values(by=['Beneficiario'], inplace=True)


# Define encabezado
am_pm = 'pm' if fecha_de_referencia.hour > 12 else 'm' if fecha_de_referencia.hour == 12 else 'am'
propuestas = f"GyG PROPUESTAS DE CAMBIO DE CATEGORÍA\nal {fecha_de_referencia:%d/%m/%Y %I:%M} {am_pm}\n\n\n" + \
              "Vecino               | Dirección      | Categoría actual | Propuesta\n" + \
              espacios(76, '-') + '\n'

# Detalle de las propuestas de cambio
for index, r in df_resumen_r.iterrows():
    propuestas += patron_detalle.format(
            trunca_texto(edita_beneficiario(r['Beneficiario']), 20),
            trunca_texto(edita_dirección(r['Dirección']), 14),
            trunca_texto(edita_categoría(r['Categoría']), 16),
            trunca_texto(edita_categoría(r['Propuesta']), 16),
            trunca_texto(edita_último_pago(r['Beneficiario']), 53))
    propuestas += espacios(76, '-') + '\n'

if len(no_evaluables) > 0:
    propuestas += f"\n(*) No evaluable: menos de {nMeses} meses desde su inicio en el sector.\n"

# Pié de reporte
m_inicial   = mes_anterior - relativedelta(months=nMeses-1)
propuestas += f"\nEvaluación realizada en base a los pagos recibidos entre " + \
              f"{m_inicial.strftime('%b')}" + \
              f"{m_inicial.strftime('/%Y') if m_inicial.year != mes_anterior.year else ''} " + \
              f"y {mes_anterior.strftime('%b/%Y')}.\n"


print(f'Grabando archivo "{nombre_propuestas.format(mes_anterior)}"...')

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_propuestas.format(mes_anterior))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(propuestas)

filename = os.path.join(attach_path, 'Windows', nombre_propuestas.format(mes_anterior))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(propuestas)
