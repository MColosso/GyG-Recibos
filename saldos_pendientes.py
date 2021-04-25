# GyG SALDOS PENDIENTES
#
# Determina los saldos pendientes por cancelar a la fecha y muestra todos,
# o sólo aquellos con saldos pendientes, ordenados alfabéticamente o no.

"""
    PENDIENTE POR HACER
    -   
    
    HISTORICO
    -   Se agregó una opción para aplicar o no el ajuste por inflación en el cálculo de los saldos
        pendientes (01/04/2021)
    -   Se corrige el ordenamiento alfabético para ignorar los acentos (18/09/2020)
    -   Cambiar la forma para la lectura de opciones desde el standard input, estandarizando su uso e
        implementando la opción '--toma_opciones_por_defecto' en la linea de comandos (08/12/2019)
    -   Incluir en el total de la deuda cualquier saldo pendiente del último mes pagado (tomar Resumen
        de Saldos a la fecha -resumen_saldos.py- como referencia) (14/11/2019)
    -   Ajustar los mensajes para hacer más claro su contenido (04/11/2019)
    -   Ajustar la selección de registros para igualarla a la utilizada en analisis_de_pagos.py
        (03/11/2019)
    -   Ajustar la codificación de caracteres para que el archivo de saldos sea legible en Apple y
        Windows
          -> Se generan archivos con codificaciones 'UTF-8' (Apple) y 'cp1252' (Windows) en carpetas
             separadas (29/10/2019)
    -   Se cambiaron las ubicaciones de los archivos resultantes a la carpeta GyG Recibos dentro
        de la carpeta actual para compatibilidad entre Windows y macOS (21/10/2019)
    -   Mostrar el saldo disponible para aquellos vecinos que tengan un depósito administrado
        por la Asociación (Del Negro Palermo, etc.) y su saldo sea mayor que cero (30/09/2019)
    -   Cambiar el manejo de cuotas para usar las rutinas en la clase Cuota (GyG_cuotas)
        (29/09/2019)
    
"""

print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *
from pandas import DataFrame, read_excel, isnull, notnull, to_numeric
from numpy import mean, std, NaN
from datetime import datetime, timedelta        # , date
from dateutil.relativedelta import relativedelta
import re
import sys
import os
import numbers

import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'


#
# DEFINE CONSTANTES
#

nombre_análisis  = "GyG Saldos pendientes {:%Y-%m (%B)}.txt"
attach_path      = GyG_constantes.ruta_saldos_pendientes   # ./GyG Recibos/Saldos Pendientes
# attach_path      = GyG_constantes.ruta_analisis_de_pagos   # "C:/Users/MColosso/Google Drive/GyG Recibos/Análisis de Pago/"

excel_workbook   = GyG_constantes.pagos_wb_estandar        # '1.1. GyG Recibos.xlsm'
excel_resumen    = GyG_constantes.pagos_ws_resumen_reordenado   # 'R.VIGILANCIA (reordenado)'
excel_detalle    = GyG_constantes.pagos_ws_vigilancia      # 'Vigilancia'
excel_cuotas     = GyG_constantes.pagos_ws_cuotas          # 'CUOTA'
excel_saldos     = GyG_constantes.pagos_ws_saldos          # 'Saldos'


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def esVigilancia(x):
    return x == 'Vigilancia'

def get_filename(filename):
    return os.path.basename(filename)

# def get_street(address):
#     return address.index(' ', address.index(' ') + 1)

def get_street(address):
    # return address.index(' ', address.index(' ') + 1)
    grupos = re.findall(r'\w+', address)
    if grupos[0].lower() == "av":
        return "Avenida"
    return grupos[1] if len(grupos) > 0 else ''

def edita_número(number, num_decimals=2):
#    return f"{number:,.{num_decimals}f}".replace(',', 'x').replace('.', ',').replace('x', '.')
    return locale.format_string(f'%.{num_decimals}f', number, grouping=True, monetary=True)

def edita_dirección(dirección):
    return dirección.replace('Calle ', '').replace('Nros. ', '').replace('Nro. ', '')

def edita_categoría(categoría):
    return "" if categoría in ['Cuota completa', 'Colaboración'] else f" ({categoría})"

def seleccionaRegistro(beneficiarios, categorías, montos):

    def list_and(l1, l2): return [a and b for a, b in zip(l1, l2)]
    def list_or(l1, l2):  return [a or  b for a, b in zip(l1, l2)]
    def list_not(l1):     return [not a for a in l1]

    def list_lt_cuota(l1):
#        l2 = [a if is_numeric(a) else cuotas_mensuales[last_col] for a in l1]
#        return [a < cuotas_mensuales[last_col] for a in l2]
        l2 = list()
        for beneficiario, monto in zip(beneficiarios, montos):
            f_ultimo_pago = fecha_ultimo_pago(beneficiario, columns[last_col].strftime('%m-%Y'), fecha_real=False)
            if f_ultimo_pago == None:
                l2.append(False)
            else:
                cuota = cuotas_obj.cuota_vigente(beneficiario, f_ultimo_pago)
                l2.append(monto < cuota)
        return l2

    # Selecciona aquellos que pagan cuota completa y el monto del mes analizado es inferior
    # al establecido o no lo han pagado
    list_1 = list_or(
                        list_and(categorías == 'Cuota completa', list_lt_cuota(montos)),
                        list_and(categorías == 'Cuota completa', isnull(montos))
                    )
    # Selecciona aquellos que colaboran, pero no han pagado el mes analizado
    list_2 = list_and(categorías == 'Colaboración', isnull(montos))
    # Selecciona aquellos que tienen una cuenta con saldo a favor
    df_saldos_gt_0 = df_saldos[df_saldos['Saldo'] > 0]
    list_3 = [len(df_saldos_gt_0[df_saldos_gt_0['Beneficiario'] == b]) > 0 for b in beneficiarios]

    return list_or(list_1, list_or(list_2, list_3))


def fecha_ultimo_pago(beneficiario, str_mes_año, fecha_real=True):
    try:
        df_fecha = df_pagos[(df_pagos['Beneficiario'] == beneficiario) & (df_pagos['Meses'].str.contains(str_mes_año))]
    except:
        print(f"ERROR: fecha_ultimo_pago({beneficiario}, {str_mes_año}): {str(sys.exc_info()[1])}")
        return None
#    if beneficiario in a_evaluar: print(f"  (fecha_ultimo_pago) {beneficiario=}, {fecha_real=}, \n{df_fecha[['Fecha', 'Meses']]} ")
    if df_fecha.shape[0] > 0:
        fecha_pago = df_fecha.iloc[-1]['Fecha'].to_pydatetime()
        if fecha_real:
            return fecha_pago
        fecha_objetivo = datetime(int(str_mes_año[3:]), int(str_mes_año[0:2]), 1)
        if fecha_objetivo < GyG_constantes.fecha_de_corte:
            return fecha_objetivo
        else:
            return datetime.today() if fecha_pago < fecha_objetivo else fecha_pago
    else:
        return None

#a_evaluar = ['Jocsan Orta Correa']

def no_participa_desde(r):
    """
        Busca a partir de qué fecha no se han recibido pagos
        (evalúa desde el mes y año indicado, hasta el 2016)
    """

#    if r['Beneficiario'] in a_evaluar: print(f"\nBeneficiario: {r['Beneficiario']}\n{'-'*(len('Beneficiario: ')+len(r['Beneficiario']))}")
#    if r['Beneficiario'] in a_evaluar: print(f"  (0): f_desde: {columns.index(r['F.Desde'])} ({columns[columns.index(r['F.Desde'])]}), last_col: {last_col} ({columns[last_col]})")

    x = last_col       # <-------- 'x' es la primera columna vacía
    ultimo_mes_con_pagos = None
    saldo_ultimo_mes = 0.00
    cuotas_pendientes = 0.00
    saldo_pendiente = False
    f_desde = columns.index(r['F.Desde'])

    for idx in reversed(range(f_desde, last_col+1)):
#        if r['Beneficiario'] in a_evaluar: print(f"  (1): idx: {idx}, cuota: {cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx])}, r.iloc[idx]: {r.iloc[idx]}")
        if notnull(r.iloc[idx]):
            ultimo_mes_con_pagos = idx
            break
        x = idx
        cuotas_pendientes += cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx], aplica_IPC=aplica_IPC) # cuotas[idx]
    if ultimo_mes_con_pagos == None:
        ultimo_mes_con_pagos = this_col = f_desde
        fecha_txt = '2016'
    else:
        this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
        fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
#    if r['Beneficiario'] in a_evaluar: print(f"  (1a): cuotas_pendientes: {cuotas_pendientes:,.2f}, this_col: {this_col}, fecha_txt: {fecha_txt}")

    # Determina el saldo del último mes
    if ultimo_mes_con_pagos == None:
        saldo_ultimo_mes = 0.00
    elif r.iloc[ultimo_mes_con_pagos] == 'ü':       # 'check'
#        if r['Beneficiario'] in a_evaluar: print(f"  (1c): x = {ultimo_mes_con_pagos}, columns[x] = {columns[ultimo_mes_con_pagos]} - DEUDA SALDADA")
        saldo_ultimo_mes = 0.00           # La mensualidad está saldada por completo
    else:
        f_ultimo_pago = fecha_ultimo_pago(r['Beneficiario'], columns[ultimo_mes_con_pagos].strftime('%m-%Y'))
#        if r['Beneficiario'] in a_evaluar: print(f"  (1b): fecha_ultimo_pago({r['Beneficiario']}, {columns[ultimo_mes_con_pagos].strftime('%m-%Y')}) = {f_ultimo_pago}, mes: '{columns[ultimo_mes_con_pagos]}'")
#        if r['Beneficiario'] in a_evaluar: print(f"  (1c): Fecha último pago: {f_ultimo_pago}, mes: '{columns[ultimo_mes_con_pagos]}'")
        if (f_ultimo_pago == None) or (r['Categoría'] == 'Colaboración'):
            saldo_ultimo_mes = 0.00   # Probablemente es un vecino que nunca ha pagado
        else:
            cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], f_ultimo_pago)
#            if r['Beneficiario'] in a_evaluar: print(f"  (3): Beneficiario: {r['Beneficiario']}, cuota actual: {cuota_actual}, pago: {r[columns[ultimo_mes_con_pagos]]}")
            saldo_ultimo_mes = cuota_actual - r[columns[ultimo_mes_con_pagos]]
            # Si el monto cancelado no cubre la cuota del período, recalcula el saldo del ultimo mes en base
            # a la última cuota
            if (saldo_ultimo_mes > 0.00) and (f_ultimo_pago >= GyG_constantes.fecha_de_corte):
                cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], datetime.now())
                saldo_ultimo_mes = cuota_actual - r[columns[ultimo_mes_con_pagos]]

    if saldo_ultimo_mes < 0.00:
        saldo_ultimo_mes = 0.00
    deuda_actual = cuotas_pendientes + saldo_ultimo_mes
#    if r['Beneficiario'] in a_evaluar: print(f"  (8): Deuda: actual: Bs. {edita_número(deuda_actual, num_decimals=0)}, " + \
#                                             f"Saldo último mes: Bs. {edita_número(saldo_ultimo_mes, num_decimals=0)}")

    info_deuda = ''
    if saldo_ultimo_mes == 0.00 and ultimo_mes_con_pagos < last_col and fecha_txt != '2016':
        ultimo_mes_con_pagos += 1
        this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
        fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
#    if r['Beneficiario'] in a_evaluar: print(f"  (9): x: {ultimo_mes_con_pagos}, last_col: {last_col}, r[{fecha_referencia}]: {r[fecha_referencia]}, deuda: {saldo_ultimo_mes}")
    if deuda_actual != 0.00:
        if saldo_ultimo_mes != 0.00:
            mensaje = 'Diferencia pendiente en ' + fecha_txt
            if ultimo_mes_con_pagos != last_col:
                mensaje += ' y meses subsiguientes'
            saldo_pendiente = True
        else:
            if ultimo_mes_con_pagos == last_col:
                mensaje = 'Pendiente ' + fecha_txt
            else:
                mensaje = f"{'Colaboraciones' if r['Categoría'] == 'Colaboración' else 'Cuotas'} " + \
                          f"pendientes desde {fecha_txt}"
                saldo_pendiente = True
    else:
        mensaje = f"No tiene {'colaboraciones' if r['Categoría'] == 'Colaboración' else 'cuotas'} pendientes"

    info_deuda = edita_número(deuda_actual, num_decimals=0) if (deuda_actual != 0.00) else ""
    if (deuda_actual != 0.00):
        if saldo_pendiente:
            sep1, sep2 = (' por ', '')
#            if x == last_col:
#                sep1 = ' de '
        else:
            sep1, sep2 = (' (', ')')
        info_deuda = f"{sep1}Bs. {edita_número(deuda_actual, num_decimals=0)}{sep2}"
    else:
        info_deuda = ""

    saldo_a_favor = df_saldos[df_saldos['Beneficiario'] == r['Beneficiario']]['Saldo']
    if saldo_a_favor.empty:
        info_saldo = ''
    else:
        saldo_a_favor = saldo_a_favor.iloc[0]
        if deuda_actual == 0.00:
            info_saldo = f"  y tiene un saldo a su favor de Bs. {edita_número(saldo_a_favor, num_decimals=2)}"
        else:
            comparacion = 'no cubre' if saldo_a_favor < deuda_actual else 'supera' if saldo_a_favor > deuda_actual else 'es igual a'
            info_saldo = f" (su saldo disponible -Bs. {edita_número(saldo_a_favor, num_decimals=2)}- {comparacion} esta cantidad)"

#    if r['Beneficiario'] in a_evaluar: print(f' (10): mensaje:    "{mensaje}",\n' + \
#                                             f'       info_deuda: "{info_deuda}",\n' + \
#                                             f'       info_saldo: "{info_saldo}"')

    return mensaje, info_deuda, info_saldo


#
# PROCESO
#

# Determina el mes actual, a fin de utilizarlo como opción por defecto
# mes_actual = datetime.now().strftime('%m-%Y')
hoy = datetime.now()
fecha_análisis = datetime(hoy.year, hoy.month, 1) - timedelta(days=1)
mes_actual = fecha_análisis.strftime('%m-%Y')
print()

# Selecciona el mes y año a procesar
mes_año = input_mes_y_año('Indique el mes y año a analizar', mes_actual, toma_opciones_por_defecto)

# Selecciona si se muestran sólo los saldos deudores o no
solo_deudores = input_si_no('Sólo vecinos con saldos pendientes', 'sí', toma_opciones_por_defecto)

# Selecciona si se aplica el ajuste por inflación (IPC - Indice de Precios al Consumidor)
aplica_IPC = input_si_no("Aplica ajuste por inflación (IPC)", 'sí', toma_opciones_por_defecto)

# Selecciona si se ordenan alfabéticamente los vecinos
ordenado = input_si_no('Ordenados alfabéticamente', 'no', toma_opciones_por_defecto)

año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = datetime(año, mes, 1)
f_ref_último_día = fecha_referencia + relativedelta(months=1) + relativedelta(days=-1)
fecha_análisis   = format(datetime(año, mes, 1), '%B/%Y').capitalize()

print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df_resumen = read_excel(excel_workbook, sheet_name=excel_resumen)

# Inicializa el handler para el manejo de las cuotas
cuotas_obj = Cuota(excel_workbook)

# Cambia el nombre de la columna 2016 a datetimme(2016, 1, 1)
df_resumen.rename(columns={2016:datetime(2016, 1, 1)}, inplace=True)

# Define algunas variables necesarias
columns = list(df_resumen.columns.values)
last_col = columns.index(datetime(año, mes, 1))
hoy = datetime.now().day

# Toma los saldos a favor de algunos vecinos para forzar su selección en el paso siguiente
df_saldos = read_excel(excel_workbook, sheet_name=excel_saldos, skiprows=[0, 1])
df_saldos = df_saldos[['Beneficiario', 'Dirección', 'Saldo']]
df_saldos.dropna(subset=['Beneficiario'], inplace=True)
df_saldos = df_saldos[df_saldos['Saldo'] > 0.0]


# Lee la hoja de cálculo con el detalle de los pagos
df_pagos = read_excel(excel_workbook, sheet_name=excel_detalle)

# Genera una columna con el resumen de los meses cancelados en cada pago
meses_cancelados = df_pagos['Concepto'].apply(lambda x: separa_meses(x, as_string=True))
df_pagos.insert(column='Meses', value=meses_cancelados, loc=df_pagos.shape[1])

# Elimina los registros que no no corresponden a pago de vigilancia
df_pagos = df_pagos.loc[esVigilancia(df_pagos['Categoría'])]

columns = df_resumen.columns.values.tolist()
last_col = columns.index(datetime(año, mes, 1))


# Elimina del resumen aquellos que no tienen una categoría definida, aquellos donde
# el mes a evaluar ya está cancelado, y aquellos en los cuales el beneficiario
# no participa en el pago de vigilancia
df_resumen.dropna(subset=['Categoría'], inplace=True)

if solo_deudores:
    df_resumen = df_resumen.loc[seleccionaRegistro(df_resumen['Beneficiario'],
                                                   df_resumen['Categoría'],
                                                   df_resumen[datetime(año, mes, 1)])]

# Elimina los registros con categoría "Sólo comida" si sólo se despliegan los deudores
if solo_deudores:
    df_resumen = df_resumen[df_resumen['Categoría'] != 'Sólo comida']

# Ajusta las columnas F.Desde y F.Hasta en aquellos en los que estén vacíos:
# 01/01/2016 y 01/mes/año+1
df_resumen.loc[df_resumen[isnull(df_resumen['F.Desde'])].index, 'F.Desde'] = datetime(2016, 1, 1)
df_resumen.loc[df_resumen[isnull(df_resumen['F.Hasta'])].index, 'F.Hasta'] = datetime(año + 1, mes, 1)

# Elimina aquellos vecinos que compraron (o se iniciaron) en fecha posterior
# a la fecha de análisis
df_resumen = df_resumen[df_resumen['F.Desde'] < f_ref_último_día]

# Elimina aquellos vecinos que vendieron (o cambiaron su razón social) en fecha anterior
# a la fecha de análisis
df_resumen = df_resumen[df_resumen['F.Hasta'] >= f_ref_último_día]

# Ordena los beneficiarios en orden alfabético
if ordenado:
    df_resumen['Benef_sort'] = df_resumen['Beneficiario'].apply(lambda benef: remueve_acentos(benef))
    df_resumen.sort_values(by='Benef_sort', inplace=True)


# ANÁLISIS DE SALDOS PENDIENTES

print(f"Creando archivo de saldos '{nombre_análisis.format(datetime(año, mes, 1))}'...")
print('')

mensaje = ' - {}, {}{}: {}{}{}\n'

# Encabezado
am_pm = 'pm' if datetime.now().hour > 12 else 'm' if datetime.now().hour == 12 else 'am'
ajuste_por_inflación = ' ajustados por inflación' if aplica_IPC else ''
análisis = f"GyG SALDOS PENDIENTES{ajuste_por_inflación}, {fecha_análisis}\n" + \
           f"({datetime.now():%d/%m/%Y %I:%M} {am_pm})\n\n"

# SALDOS
dirección_anterior = None
for index, r in df_resumen.iterrows():

    info_pago, info_deuda, info_saldo = no_participa_desde(r)

    if not(solo_deudores and len(info_deuda) == 0):
        if dirección_anterior == None:
            dirección_anterior = get_street(r['Dirección'])
        if (get_street(r['Dirección']) != dirección_anterior) and not ordenado:
            análisis += '\n'
            dirección_anterior = get_street(r['Dirección'])
        análisis += mensaje.format(
                r['Beneficiario'], edita_dirección(r['Dirección']), edita_categoría(r['Categoría']),
                info_pago, info_deuda, info_saldo)

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_análisis.format(datetime(año, mes, 1)))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(análisis)

filename = os.path.join(attach_path, 'Windows', nombre_análisis.format(datetime(año, mes, 1)))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(análisis)
