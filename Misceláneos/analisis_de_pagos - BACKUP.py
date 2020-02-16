# GyG ANALISIS DE PAGOS
#
# Determina los pagos no realizados en el mes actual 

"""
    NUEVOS ANALISIS A INCLUIR

    -   Agrupar en "mangos bajitos", "alto impacto" y el resto, para facilitar la gestión de cobranza
        (¿cómo poder establecer esta clasificación...?)
    -   En agunos casos se ubicó a un vecino en una categoría particular, a fin de que se mostrase perma-
        nentemente en ese segmento de la Cartelera Virtual. Para estos casos, no mostrar una propuesta
        de cambio de categoría
    -   ¿Hay vecinos cuyo comportamiento de pagos es similar al de otros? Ello pudiera ayudar para influen-
        ciar sobre algunos de ellos para lograr un cambio en el grupo
    -   Agregar al resumen:
          . Para aquellos vecinos que pagan la cuota completa, indicar los meses que están por debajo de
            la cuota (incluye meses no cancelados que quedaron "encerrados" entre otros)

    Ok  Evaluar cambios de categoría: Cuota completa, Colaboración o No participa
        ("más de xx meses cancelando {cuota completa | colaboración}. Cambiar su clasificación a 'xxx'")
    Ok  Destacar las propuestas de cambio de categoría cuando la resultante sea diferente a "No participa"
    Ok  Colocar la propuesta de cambio de categoría al final del análisis del vecino, si corresponde
    Ok  Mostrar la cantidad de cuotas recaudadas, el monto total y el monto de la cuota de los últimos
        <nMeses>
    Ok  No eliminar los registros de aquellos vecinos que pagan cuota completa y el monto cancelado en el
        mes de análisis sea inferior a la cuota del mes
        ( df_resumen = df_resumen.loc[isnull(df_resumen[datetime(año, mes, 1)])] )
    Ok  Manejar adecuadamente aquellos registros con Promedio o Variación == None (=> no se encontraron
        suficientes pagos en un período de <nMeses> meses)
    Ok  Mostrar el saldo a la fecha de aquellos vecinos quienes han hecho depósitos por adelantado
    Ok  Agregar al resumen:
          . Cantidad de pagos recibidos iguales o superiores a la cuota
          . Cantidad de pagos recibidos equivalentes a pagos al 100% de la cuota (total recaudado / cuota)
    Ok  Se muestran en el análisis los vecinos en fecha posterior a "F.Hasta" y "Categoría" no nula
        -- Corregido 02/02/2019
    Ok  Mostrar la lista de vecinos en orden alfabético (09/04/2019)
    Ok  En aquellos vecinos que parcicipan en el programa de comidas para los vigilantes, incluir un comen-
        tario al respecto (14/04/2019)

"""

# Selecciona las librerías a utilizar
print('Cargando librerías...')
from pandas import read_excel, isnull, notnull, to_numeric
from numpy import mean, std, NaN
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
import os
import numbers
# import codecs
from re import match
from locale import setlocale, LC_ALL

# Define textos
nombre_análisis = "{}GyG Analisis de Pagos {:%Y-%m (%B)}.txt"
attach_path     = "C:/Users/MColosso/Google Drive/GyG Recibos/Análisis de Pago/"

excel_workbook          = '1.1. GyG Recibos.xlsm'
excel_worksheet_resumen = 'R.VIGILANCIA (reordenado)'
excel_worksheet_detalle = 'Vigilancia'
excel_worksheet_Vecinos = 'Vecinos'
excel_worksheet_saldos  = 'Saldos'

nMeses = 5  # Últimos meses (pagos) a analizar

dummy = setlocale(LC_ALL, '')


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def esVigilancia(x):
    return x == 'Vigilancia'

def seleccionaRegistro(beneficiarios, categorías, montos):
    def list_and(l1, l2): return [a and b for a, b in zip(l1, l2)]
    def list_or(l1, l2):  return [a or  b for a, b in zip(l1, l2)]
    def list_not(l1):     return [not a for a in l1]
    def list_lt_cuota(l1):
        l2 = [a if is_numeric(a) else cuotas_mensuales[last_col] for a in l1]
        return [a < cuotas_mensuales[last_col] for a in l2]

    # Selecciona aquellos que pagan cuota completa y el monto del mes analizado es inferior
    # al establecido o no lo han pagado
    list_1 = list_or( list_and(categorías == 'Cuota completa', list_lt_cuota(montos)),
                      list_and(categorías == 'Cuota completa', isnull(montos)))
    # Selecciona aquellos que colaboran, pero no han pagado el mes analizado
    list_2 = list_and(categorías == 'Colaboración', isnull(montos))
    # Selecciona aquellos que tienen una cuenta con saldo a favor
    df_saldos_gt_0 = df_saldos[df_saldos['Saldo'] > 0]
    list_3 = [len(df_saldos_gt_0[df_saldos_gt_0['Beneficiario'] == b]) > 0 for b in beneficiarios]

    return list_or(list_1, list_or(list_2, list_3))

def dia_promedio(r):
    # Determina el día promedio de pago en los últimos 'nn' meses
    dias = [x.day for x in list(df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']]['Fecha'])]
    dias = dias[-nMeses:]
    if len(dias) == 0:
        promedio = None   # No se encontraron suficientes pagos en un período de <nMeses> meses
    else:
        promedio = int(round(mean(dias), 0))
    return promedio

def desviación(r):
    # Determina la desviación estándar del día promedio de pago en los últimos 'nn' meses
    """
    dias = [x.day for x in list(df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']]['Fecha'])]
    dias = dias[-nMeses:]
    return std(dias)
    """
    fechas = list(df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']]['Fecha'])
    fechas.sort()
    fechas = fechas[-nMeses:]
    dif_fechas = [(fechas[i+1] - fechas[i]).days for i in range(len(fechas) - 1)]
    if len(dif_fechas) == 0:
        desv_est = None   # No se encontraron suficientes pagos en un período de <nMeses> meses
    else:
        desv_est = std(dif_fechas)
    return desv_est
    
def no_participa_desde(r):
    """
        Busca a partir de qué fecha no se han recibido pagos
        (evalúa desde el mes y año indicado, hasta el 2016)
    """
    x = last_col
    for idx in reversed(range(3 + 2, last_col )):
        if isnull(r.iloc[idx]):
            x = idx
        else:
            break
    this_col = columns[x]
    if this_col == 2016:
        fecha_txt = '2016'
    else:
        fecha = this_col
        fecha_txt = fecha.strftime('%B/%Y')

    # Valida si sólo es un mes pendiente o no
    if x == last_col:
        if r[datetime(año, mes, 1)] < cuotas_mensuales[last_col - 2]:
            mensaje = 'tiene un saldo pendiente en ' + fecha_txt
        else:
            if isnull(r[datetime(año, mes, 1)]):
                mensaje = 'tiene pendiente ' + fecha_txt
            else:
                mensaje = 'no tiene saldo pendiente'
    else:
        mensaje = 'tiene pagos pendientes desde ' + fecha_txt

    # Valida si tiene un saldo a favor
    if df_saldos['Beneficiario'].str.contains(r['Beneficiario']).any():
        saldo = df_saldos[df_saldos['Beneficiario'] == r['Beneficiario']]['Saldo'].item()
        saldo = format(saldo, ',.0f').replace(',', '.')
        mensaje += f'\n  Dispone de un saldo a favor de Bs. {saldo}'
    return mensaje

def mas_de_un_mes(beneficiario):
    pagos_beneficiario = (df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']])['Concepto'].tolist()
    num_pagos = len(pagos_beneficiario)
    num_pagos_multiples = 0
    for pago in pagos_beneficiario:
        if 'meses' in pago:
            num_pagos_multiples += 1
    pct_pagos_multiples = num_pagos_multiples / num_pagos
    # convierte 'pct_pagos_multiples' en un entero múltiplo de 5
    pct_pagos_multiples = int(round(pct_pagos_multiples * 20) / 20 * 100)
    if num_pagos_multiples > 0:
        return f'\n  En {num_pagos_multiples} de {num_pagos} casos ({pct_pagos_multiples}%), ' + \
                'su pago correspondió a más de un mes'
    else:
        return "\n  Nunca ha pagado más de un mes simultáneamente"

def genera_pagos_mensuales():
    # Ubica la línea de totales
    linea_totales = df_resumen.loc[df_resumen['Dirección'] == 'TOTAL'].index[0]
    # Por columna, cuenta la cantidad de datos numéricos desde la primera línea hasta la línea
    # anterior a la línea de totales
    return [to_numeric(df_resumen[column].iloc[0:linea_totales], errors='coerce').count()
                for column in list(df_resumen.columns)]

def genera_resumen():
    análisis.write(f'Resumen de los últimos {nMeses} meses:\n')
    análisis.write(' ' * 19 + 'Pagos' + ' ' * 37 + 'Pagos  Equiv.\n')
    análisis.write('  Mes            | recib. |  Total Bs. | Promedio |   Cuota | 100% | 100%\n')
    análisis.write('  ' + '-' * 72 + '\n')
    for idx in range(last_col - nMeses - 1, last_col - 1):
        mes = format(columns[idx + 2], '%B/%Y').capitalize()
        num_pagos = pagos_mensuales[idx]
        tot_pagos = format(totales_mensuales[idx], ',.0f').replace(',', '.')
        num_pagos_100_pct = num_100_pct[idx]
        promedio  = format(round(totales_mensuales[idx] / pagos_mensuales[idx], -1), ',.0f').replace(',', '.')
        if notnull(cuotas_mensuales[idx]):
            cuota = cuotas_mensuales[idx]
            num_pagos_eqv = round(totales_mensuales[idx] / cuota, 1)
            # pct = round((totales_mensuales[idx] / pagos_mensuales[idx]) / cuotas_mensuales[idx] * 100)
        else:
            cuota = 0
            num_pagos_eqv = 0
            # pct = 0
        cuota = format(cuota, ',.0f').replace(',', '.')
        num_pagos_eqv = format(num_pagos_eqv, ',.1f').replace('.', ',')
        análisis.write(f'  {mes:16}{num_pagos:^8}{tot_pagos:>12}{promedio:>11}{cuota:>10}' + \
                       f'  {num_pagos_100_pct:^6} {num_pagos_eqv:^6}\n')
    análisis.write('\n\n')

def genera_propuesta_categoría():

    def genera_propuesta(r):
        # comparar <r> con <cuotas_mensuales> en los últimos <nMeses> meses [last_col - nMeses + 1: last_col]
        pagos = (r.iloc[last_col - nMeses + 1: last_col + 1]).tolist()
        pagos = [p if is_numeric(p) else NaN for p in pagos]
        # Cuota completa: todos los pagos son mayores o iguales a la cuota del mes
        if all(p >= m for p, m in zip(pagos, monto_cuotas)):
            propuesta = 'Cuota completa'
        # Colaboración: todos los pagos son inferiores a la cuota del mes
        elif all(p <  m for p, m in zip(pagos, monto_cuotas)):
            propuesta = 'Colaboración'
        # No colabora: todos los pagos son nulos
        elif all(isnull(p) for p in pagos):
            propuesta = 'No participa'
        # En cualquier otro caso, el resultado es indeterminado
        else:
            propuesta = NaN
        if r['Categoría'] == propuesta:
            propuesta = NaN
        return propuesta

    monto_cuotas = cuotas_mensuales[last_col - nMeses + 1: last_col + 1]
    df = df_resumen[['Beneficiario', 'Categoría', 'Comida']]
    df = df[notnull(df['Beneficiario'])]
    df = df[df['Beneficiario'] != 'CUOTAS MENSUALES']
    df = df[notnull(df['Categoría'])]
    df['Propuesta'] = df_resumen.apply(genera_propuesta, axis=1)
    df = df[notnull(df['Propuesta'])]
    return df

def propuesta_de_cambio(beneficiario):
    r = df_categoría[df_categoría['Beneficiario'] == beneficiario]
    if r.shape[0] == 0:
        return ''
    propuesta = r['Propuesta'].to_string(index=False)
    if propuesta != 'No participa':
        propuesta = propuesta.upper()
    return(f'\n  Se sugiere cambiar su categoría de pago a {propuesta}')

def programa_de_comida(comida):
    txt = ''
    if not(isnull(comida) or comida == ' '):
        txt = '\n  Actualmente participa en el programa de comidas para los vigilantes'
    return txt

def get_filename(filename):
    return os.path.basename(filename)

def is_numeric(valor):
    return isinstance(valor, numbers.Number)


#
# PROCESO
#

# Determina el mes actual, a fin de utilizarlo como opción por defecto
# mes_actual = datetime.now().strftime('%m-%Y')
hoy = datetime.now()
fecha_análisis = datetime(hoy.year, hoy.month, 1) - timedelta(days=1)
mes_actual = fecha_análisis.strftime('%m-%Y')
print('\n')

# Selecciona el mes y año a procesar
pattern = '(0[1-9]|1[012])-20(1[7-9]|2[0-9])'
while True:
    mes_año = input("*** Indique el mes y año a analizar [" + mes_actual + "]: ")
    if len(mes_año) == 0:
        mes_año = mes_actual
    if bool(match(pattern, mes_año)):
        break
    else:
        print('Seleccione un mes y año correctos (2017+), con un guión como separador...')

año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = date(año, mes, 1)
fecha_análisis = format(datetime(año, mes, 1), '%B/%Y').capitalize()

# Abre la hoja de cálculo de Recibos de Pago
print('\n')
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df_resumen = read_excel(excel_workbook, sheet_name=excel_worksheet_resumen)

# Define algunas variables necesarias
columns = list(df_resumen.columns.values)
last_col = columns.index(datetime(año, mes, 1))
hoy = datetime.now().day

# ------------------------------

# Ajusta las columnas F.Desde y F.Hasta en aquellos en los que estén vacíos:
# 01/01/2016 y 01/mes/año+1
df_resumen.loc[df_resumen[isnull(df_resumen['F.Desde'])].index, 'F.Desde'] = date(2016, 1, 1)
df_resumen.loc[df_resumen[isnull(df_resumen['F.Hasta'])].index, 'F.Hasta'] = date(año + 1, mes, 1)

# Elimina aquellos vecinos que compraron (o se iniciaron) en fecha posterior
# a la fecha de análisis
df_resumen = df_resumen[df_resumen['F.Desde'] < fecha_referencia]

# Elimina aquellos vecinos que vendieron (o cambiaron su razón social) en fecha anterior
# a la fecha de análisis
df_resumen = df_resumen[df_resumen['F.Hasta'] >= fecha_referencia]

# ------------------------------

# Conserva las líneas con las cuotas y los totales mensuales (Pandas DataFrame)
cuotas_mensuales  = (df_resumen[df_resumen['Beneficiario'] == 'CUOTAS MENSUALES'].values.tolist())[0]
totales_mensuales = (df_resumen[df_resumen['Dirección']    == 'TOTAL'           ].values.tolist())[0]
num_100_pct       = (df_resumen[df_resumen['Dirección']    == 'PAGOS 100% CUOTA'].values.tolist())[0]

# Determina la cantidad de pagos recibidos por mes
pagos_mensuales = genera_pagos_mensuales()
columnas = list(df_resumen.columns.values)

# Genera una propuesta de cambio de categoría de pago de cada vecino
df_categoría = genera_propuesta_categoría()

# Toma los saldos a favor de algunos vecinos para forzar su selección en el paso siguiente
df_saldos  = read_excel(excel_workbook, sheet_name=excel_worksheet_saldos, skiprows=[0, 1])
df_saldos  = df_saldos[['Beneficiario', 'Dirección', 'Saldo']]
df_saldos.dropna(subset=['Beneficiario'], inplace=True)

# Elimina los registros que no tienen una categoría definida, aquellos donde
# el mes a evaluar ya está cancelado, y aquellos en los cuales el beneficiario
# no participa en el pago de vigilancia
df_resumen.dropna(subset=['Categoría'], inplace=True)
df_resumen = df_resumen.loc[seleccionaRegistro(df_resumen['Beneficiario'],
                                               df_resumen['Categoría'],
                                               df_resumen[datetime(año, mes, 1)])]

# Ordena los beneficiarios en orden alfabético
df_resumen.sort_values(by='Beneficiario', inplace=True)

# Lee la pestaña con el detalle de pagos
df_detalle = read_excel(excel_workbook, sheet_name=excel_worksheet_detalle)

# Elimina los registros que no no corresponden a pago de vigilancia
df_detalle = df_detalle.loc[esVigilancia(df_detalle['Categoría'])]

# Inserta las columnas 'Promedio' y 'Variación' con el día promedio de pago y su
# desviación estándar
prom = df_resumen.apply(dia_promedio, axis=1)
desv = df_resumen.apply(desviación, axis=1)
df_resumen.insert(loc=2, column='Promedio',  value=prom)
df_resumen.insert(loc=3, column='Variación', value=desv)

df_vecinos = read_excel(excel_workbook, sheet_name=excel_worksheet_Vecinos)
df_vecinos = df_vecinos[['Beneficiario', 'SMS', 'WhatsApp']]

columns = list(df_resumen.columns.values)
last_col = columns.index(datetime(año, mes, 1))

# Imprime el archivo con los resultados del análisis

# Crea el archivo con el análisis
filename = nombre_análisis.format(attach_path, datetime(año, mes, 1))
print("Creando Análisis de Pago '{}'...".format(get_filename(filename)))
print('')

análisis = open(filename, 'w')
# análisis = codecs.open(filename, 'w', encoding='utf8')
mensaje = '* {}, {}{}\n' + \
          '  Paga {}{} usualmente el día {}\n' + \
          '  Su último pago fue el {} y {}{}{}{}\n\n'
mensaje_2 = '* {}, {}{}\n' + \
            '  Paga {}{}{}{}\n\n'
rango_fechas_txt = ' entre el {} y el {},'

# Encabezado
# análisis.write('GyG ANALISIS DE PAGOS\nal {:%d de %B de %Y}\n\n'.format(datetime.now()))
am_pm = 'pm' if datetime.now().hour > 12 else 'm' if datetime.now().hour == 12 else 'am'
análisis.write(f'GyG ANALISIS DE PAGOS, {fecha_análisis}\n({datetime.now():%d/%m/%Y %I:%M} {am_pm})\n\n')

# Resumen
genera_resumen()

# Análisis de pago
análisis.write('Análisis de pago:\n\n')
for index, r in df_resumen.iterrows():
    if round(r['Promedio'] + r['Variación'], 0) <= 100:   # 100: anteriormente 'hoy'
        min_day = min_d = int(r['Promedio']) - int(round(r['Variación']))
        max_day = max_d = int(r['Promedio']) + int(round(r['Variación']))
        if min_d <= 0:
            min_day = '{} del mes anterior'.format(30 + min_d)
            max_day = '{} del mes actual'.format(max_d)
        if max_d > 30:
            max_day = '{} del mes siguiente'.format(max_d - 30)
            if min_d > 0:
                min_day = '{} del mes actual'.format(min_d)
        if min_day == max_day:
            rango_fechas = ""
        else:
            rango_fechas = rango_fechas_txt.format(min_day, max_day)
        v = df_vecinos[df_vecinos['Beneficiario'] == r['Beneficiario']]
        tlf_sms = tlf_wapp = ''
        if not v.empty:
            if notnull(v['SMS']).bool():
                tlf_sms = ', SMS: ' + v['SMS'].to_string(index=False)
            if notnull(v['WhatsApp']).bool():
                tlf_wapp = ', WhatsApp: ' + v['WhatsApp'].to_string(index=False)
        telefonos = tlf_sms + tlf_wapp
        análisis.write(mensaje.format(
                r['Beneficiario'], r['Dirección'], telefonos,
                r['Categoría'].lower(), rango_fechas, int(r['Promedio']),
                '{:%d de %B}'.format(max(df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']]['Fecha'])),
                no_participa_desde(r),
                mas_de_un_mes(r['Beneficiario']),
                programa_de_comida(r['Comida']),
                propuesta_de_cambio(r['Beneficiario'])))
    else:
        v = df_vecinos[df_vecinos['Beneficiario'] == r['Beneficiario']]
        tlf_sms = tlf_wapp = ''
        if not v.empty:
            if notnull(v['SMS']).bool():
                tlf_sms = ', SMS: ' + v['SMS'].to_string(index=False)
            if notnull(v['WhatsApp']).bool():
                tlf_wapp = ', WhatsApp: ' + v['WhatsApp'].to_string(index=False)
        telefonos = tlf_sms + tlf_wapp
        if notnull(r['Promedio']):
            detalles_pago = ' y su único pago fue el {:%d de %B de %Y}'.format(
                                max(df_detalle[df_detalle['Beneficiario'] == r['Beneficiario']]['Fecha']))
        else:
            detalles_pago = ' pero no se tienen pagos registrados'
        análisis.write(mensaje_2.format(
                r['Beneficiario'], r['Dirección'], telefonos,
                r['Categoría'].lower(), detalles_pago,
                programa_de_comida(r['Comida']),
                propuesta_de_cambio(r['Beneficiario'])))
    df_categoría = df_categoría[df_categoría['Beneficiario'] != r['Beneficiario']]

# Otras propuestas de cambio de categoría
if df_categoría.shape[0] != 0:
    análisis.write('\n')

    # Imprime las propuestas de cambio de categoría
    análisis.write(f'Otras propuestas de cambio de categoría en base a los últimos {nMeses} meses:\n\n')

    for index, r in df_categoría.iterrows():
        categoría_actual = r['Categoría'].lower()
        propuesta = r['Propuesta']
        if categoría_actual != 'no participa':
            categoría_actual = 'paga ' + categoría_actual
        análisis.write(f'* {r["Beneficiario"]}. Actualmente {categoría_actual}. Se propone cambiar a {propuesta}' + \
                       f'{programa_de_comida(r["Comida"])}\n')


# Cierra el archivo de análisis
análisis.close()
print('\n')
