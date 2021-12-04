# GyG SALDOS PENDIENTES
#
# Determina los saldos pendientes por cancelar a la fecha y muestra todos,
# o sólo aquellos con saldos pendientes, ordenados alfabéticamente o no.

"""
    PENDIENTE POR HACER
     -  Revisar:
         - Wilson Escamilla, Güigüe 89-89: Cuotas pendientes desde septiembre 2021 por Bs. 42,00 y
           tiene un saldo a su favor de 2 dólares, ya descontado, el cual no cubre la deuda actual

        Como tiene un saldo a favor, el mensaje desplegado no debería ser 'Diferencia pendiente en
        agosto 2021 y meses subsiguientes', en lugar de 'Cuotas pendientes desde septiembre 2021'.
        Igualmente ocurriría si la deuda correspondiese al mes actual: 'Pendiente octubre 2021',
        debiéndo mostrarse 'Diferencia pendiente en octubre 2021'


    HISTORICO
     -  Se ajustó la rutina no_participa_desde() para que en aquellos vecinos que no han cancelado
        cuotas de vigilancia, se tome como fecha de inicio su fecha de ingreso al sector
        (02/12/2021)
     -  En la rutina 'no_participa_desde()', se añadió código para descontar los saldos en dólares
        y en bolívares del saldo deudor y se reescribió el texto a desplegar (07/10/2021)
     -  Corregido el monto de la cuota a considerar: En el caso de anticipos, si el monto del mismo
        superaba el monto de la cuota en el mes en el cual se hizo el pago, se entendía que el mes
        anticipado ya se encontraba pago, aunque fuese inferior al monto de la cuota para dicho mes
        (17/08/2021)
     -  Se agregó una opción para aplicar o no el ajuste por inflación en el cálculo de los saldos
        pendientes (01/04/2021)
     -  Se corrió el ordenamiento alfabético para ignorar los acentos (18/09/2020)
     -  Cambiar la forma para la lectura de opciones desde el standard input, estandarizando su uso e
        implementando la opción '--toma_opciones_por_defecto' en la linea de comandos (08/12/2019)
     -  Incluir en el total de la deuda cualquier saldo pendiente del último mes pagado (tomar Resumen
        de Saldos a la fecha -resumen_saldos.py- como referencia) (14/11/2019)
     -  Se ajustaron los mensajes para hacer más claro su contenido (04/11/2019)
     -  Se ajustó la selección de registros para igualarla a la utilizada en analisis_de_pagos.py
        (03/11/2019)
     -  Ajustar la codificación de caracteres para que el archivo de saldos sea legible en Apple y
        Windows
          -> Se generan archivos con codificaciones 'UTF-8' (Apple) y 'cp1252' (Windows) en carpetas
             separadas (29/10/2019)
     -  Se cambiaron las ubicaciones de los archivos resultantes a la carpeta GyG Recibos dentro
        de la carpeta actual para compatibilidad entre Windows y macOS (21/10/2019)
     -  Mostrar el saldo disponible para aquellos vecinos que tengan un depósito administrado
        por la Asociación (Del Negro Palermo, etc.) y su saldo sea mayor que cero (30/09/2019)
     -  Cambiar el manejo de cuotas para usar las rutinas en la clase Cuota (GyG_cuotas)
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

nombre_análisis   = GyG_constantes.txt_saldos_pendientes    # "GyG Saldos pendientes {:%Y-%m (%b)}.txt"
attach_path       = GyG_constantes.ruta_saldos_pendientes   # ./GyG Recibos/Saldos Pendientes

excel_workbook    = GyG_constantes.pagos_wb_estandar        # '1.1. GyG Recibos.xlsm'
excel_resumen     = GyG_constantes.pagos_ws_resumen         # 'RESUMEN VIGILANCIA'
excel_detalle     = GyG_constantes.pagos_ws_vigilancia      # 'Vigilancia'
excel_cuotas      = GyG_constantes.pagos_ws_cuotas          # 'CUOTA'
excel_saldos      = GyG_constantes.pagos_ws_saldos          # 'Saldos'
excel_saldos_usd  = GyG_constantes.pagos_ws_saldos_USD      # 'Saldos (US$)'
excel_vecinos     = GyG_constantes.pagos_ws_vecinos         # 'Vecinos'

epsilon           = 1.0e-6


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def esVigilancia(x):
    return x == GyG_constantes.CATEGORIA_VIGILANCIA

def get_filename(filename):
    return os.path.basename(filename)

def edita_correo(beneficiario):
    # - Familia Flores Montagne, Av. "M" 146-20, e-mail: 0414-530.0177 | oflormo@hotmail-
    #   .com | martaflo...: No tiene cuotas pendientes
    # Separa en componentes y reagrupa con barras verticales sólo aquellos que correspondan
    # a correos electrónicos

    # Revisar por qué unos correos electrónicos son de la forma:
    #   DEBUG: "hectorhernandezarmas@gmail.com | rosilbe.alzur..."
    # (49 caracteres, incluyendo los puntos suspensivos, en ambos casos)
    
    correo = df_vecinos[df_vecinos['Beneficiario'] == beneficiario]['E-mail']
    correo = correo.to_string(header=False, index=False).strip()
    if beneficiario == 'Familia Hernández Alzurutt':
        print(f'DEBUG: "{correo}"')
    correo = "sin correo electrónico" if correo == "NaN" else f"e-mail: {correo}"
    # NaN -> "sin correo electrónico"; not NaN -> "e-mail: {correo}"
    return f", {correo}"

def edita_categoría(categoría):
    # Cuota completa    -> 
    # Colaboración      ->
    # No participa      -> 
    # OCULTO            ->
    return "" if categoría in ['Cuota completa', 'Colaboración'] else f" ({categoría})"

def tabla_saldos(beneficiario, tabla):
    saldo = 0.00
    try:
        saldo = float(tabla[tabla['Beneficiario'] == beneficiario]['Saldo'])
    except:
        pass
    return saldo


def seleccionaRegistro(beneficiarios, categorías, montos):

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
    # Selecciona aquellos que tienen una cuenta en dólares con saldo a favor
    df_saldos_gt_0 = df_saldos_usd[df_saldos_usd['Saldo'] > 0]
    list_4 = [len(df_saldos_gt_0[df_saldos_gt_0['Beneficiario'] == b]) > 0 for b in beneficiarios]

    return list_or(list_1, list_or(list_2, list_or(list_3, list_4)))


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
        # if r['Beneficiario'] in a_evaluar: print(f"  (1): idx: {idx}, cuota: {cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx])}, r.iloc[idx]: {r.iloc[idx]}")
        if notnull(r.iloc[idx]):
            ultimo_mes_con_pagos = idx
            break
        x = idx
        cuotas_pendientes += cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx], aplica_INPC=aplica_INPC) # cuotas[idx]
    # if ultimo_mes_con_pagos == None:
    #     ultimo_mes_con_pagos = this_col = f_desde
    #     fecha_txt = '2016'
    # else:
    #     this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
    #     fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
    if ultimo_mes_con_pagos is None:
        ultimo_mes_con_pagos = f_desde
    this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
    fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
    # if r['Beneficiario'] in a_evaluar: print(f"  (1a): cuotas_pendientes: {cuotas_pendientes:,.2f}, this_col: {this_col}, fecha_txt: {fecha_txt}")

    # Determina el saldo del último mes
    if ultimo_mes_con_pagos == None:
        saldo_ultimo_mes = 0.00
    elif r.iloc[ultimo_mes_con_pagos] == 'ü':       # 'check'
        # if r['Beneficiario'] in a_evaluar: print(f"  (1c): x = {ultimo_mes_con_pagos}, columns[x] = {columns[ultimo_mes_con_pagos]} - DEUDA SALDADA")
        saldo_ultimo_mes = 0.00           # La mensualidad está saldada por completo
    else:
        f_ultimo_pago = fecha_ultimo_pago(r['Beneficiario'], columns[ultimo_mes_con_pagos].strftime('%m-%Y'))
        # if r['Beneficiario'] in a_evaluar: print(f"  (1b): fecha_ultimo_pago({r['Beneficiario']}, {columns[ultimo_mes_con_pagos].strftime('%m-%Y')}) = {f_ultimo_pago}, mes: '{columns[ultimo_mes_con_pagos]}'")
        # if r['Beneficiario'] in a_evaluar: print(f"  (1c): Fecha último pago: {f_ultimo_pago}, mes: '{columns[ultimo_mes_con_pagos]}'")
        if (f_ultimo_pago == None) or (r['Categoría'] == 'Colaboración'):
            saldo_ultimo_mes = 0.00   # Probablemente es un vecino que nunca ha pagado
        else:
            # ---------------------------------------------------------------------------------------
            # La cuota a utilizar el la del mes a evaluar, y no la del mes en el cual se hizo el pago
            # ---------------------------------------------------------------------------------------
            # cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], f_ultimo_pago)
            cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], columns[ultimo_mes_con_pagos])
            # ---------------------------------------------------------------------------------------
            # if r['Beneficiario'] in a_evaluar: print(f"  (3): Beneficiario: {r['Beneficiario']}, cuota actual: {cuota_actual}, pago: {r[columns[ultimo_mes_con_pagos]]}")
            saldo_ultimo_mes = cuota_actual - r[columns[ultimo_mes_con_pagos]]
            # Si el monto cancelado no cubre la cuota del período, recalcula el saldo del ultimo mes en base
            # a la última cuota
            if (saldo_ultimo_mes > 0.00) and (f_ultimo_pago >= GyG_constantes.fecha_de_corte):
                cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], datetime.now())
                saldo_ultimo_mes = cuota_actual - r[columns[ultimo_mes_con_pagos]]

    if saldo_ultimo_mes < 0.00:
        saldo_ultimo_mes = 0.00
    deuda_actual = cuotas_pendientes + saldo_ultimo_mes
    # if r['Beneficiario'] in a_evaluar: print(f"  (8): Deuda: actual: Bs. {edita_número(deuda_actual, num_decimals=0)}, " + \
    #                                          f"Saldo último mes: Bs. {edita_número(saldo_ultimo_mes, num_decimals=0)}")

    saldo_a_favor_VEB = tabla_saldos(r['Beneficiario'], df_saldos)
    saldo_a_favor_USD = tabla_saldos(r['Beneficiario'], df_saldos_usd)
    saldo_a_favor_USD_Bs = saldo_a_favor_USD * cuotas_obj.tasa_actual(datetime.today())
    saldo_a_favor_Bs = saldo_a_favor_VEB + saldo_a_favor_USD_Bs
    deuda_actual -= saldo_a_favor_Bs

    if deuda_actual < epsilon:
        deuda_actual = 0.00

    info_deuda = ''
    if saldo_ultimo_mes == 0.00 and ultimo_mes_con_pagos < last_col and fecha_txt != '2016':
        ultimo_mes_con_pagos += 1
        this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
        fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
    # if r['Beneficiario'] in a_evaluar: print(f"  (9): x: {ultimo_mes_con_pagos}, last_col: {last_col}, r[{fecha_referencia}]: {r[fecha_referencia]}, deuda: {saldo_ultimo_mes}")
    if deuda_actual > 0.00:
        if saldo_ultimo_mes > 0.00 or saldo_a_favor_Bs > 0.00:
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

    info_deuda = edita_número(deuda_actual, num_decimals=GyG_constantes.num_decimales) if (deuda_actual != 0.00) else ""
    if deuda_actual != 0.00:
        if saldo_pendiente:
            sep1, sep2 = (' por ', '')
            # if x == last_col:
            #     sep1 = ' de '
        else:
            sep1, sep2 = (' (', ')')
        info_deuda = f"{sep1}Bs. {edita_número(deuda_actual, num_decimals=GyG_constantes.num_decimales)}{sep2}"
    else:
        info_deuda = ""

#--------------------------------
    if saldo_a_favor_Bs < epsilon:
        info_saldo = ''
    else:
        es = '' if saldo_a_favor_VEB == 1.00 else 'es'
        txt_saldo_VEB = f'{edita_número(saldo_a_favor_VEB, num_decimals=2).replace(",00", "")} bolívar{es}'
        es = '' if saldo_a_favor_USD == 1.00 else 'es'
        txt_saldo_USD = f'{edita_número(saldo_a_favor_USD, num_decimals=2).replace(",00", "")} dólar{es}'
        info_saldo = ''.join([
                                f" y tiene un saldo a su favor de ",
                                txt_saldo_VEB if saldo_a_favor_VEB > 0.00 else '',
                                ' más ' if saldo_a_favor_VEB > 0.00 and saldo_a_favor_USD > 0.00 else '',
                                txt_saldo_USD if saldo_a_favor_USD > 0.00 else '',
                            ])
        deuda_neta = cuotas_pendientes + saldo_ultimo_mes
        saldo_neto = saldo_a_favor_VEB + saldo_a_favor_USD * cuotas_obj.tasa_actual(datetime.today())
        if deuda_actual > 0.00:
            info_saldo += ", ya descontado, el cual no cubre la deuda actual"
        elif deuda_actual == 0.00:
            pass
        elif saldo_neto > deuda_neta:
            info_saldo += ", el cual supera la deuda actual"
#--------------------------------
# comparacion = 'no cubre' if saldo_a_favor < deuda_actual else 'supera' if saldo_a_favor > deuda_actual else 'es igual a'
#  info_saldo = f" (su saldo disponible -Bs. {edita_número(saldo_a_favor, num_decimals=2)}- {comparacion} esta cantidad)"
#--------------------------------

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

# Selecciona si se muestran las colaboraciones pendientes
colaboraciones_pendientes = input_si_no('Muestra colaboraciones pendientes', 'no', toma_opciones_por_defecto)

# Selecciona si se ordenan alfabéticamente los vecinos
ordenado = input_si_no('Ordenados alfabéticamente', 'sí', toma_opciones_por_defecto)

# Selecciona si se aplica el ajuste por inflación (INPC - Indice Nacional de Precios al Consumidor)
aplica_INPC = input_si_no("Aplica ajuste por inflación (INPC)", 'sí', toma_opciones_por_defecto)


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
df_saldos = df_saldos[df_saldos['Saldo'] > epsilon]

df_saldos_usd = read_excel(excel_workbook, sheet_name=excel_saldos_usd, skiprows=[0, 1])
df_saldos_usd = df_saldos_usd[['Beneficiario', 'Dirección', 'Saldo']]
df_saldos_usd.dropna(subset=['Beneficiario'], inplace=True)
df_saldos_usd = df_saldos_usd[df_saldos_usd['Saldo'] > epsilon]


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
ajuste_por_inflación = ' ajustados por inflación' if aplica_INPC else ''
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
        análisis += bloque_de_texto(mensaje.format(
                        r['Beneficiario'], edita_dirección(r['Dirección']),
                        edita_categoría(r['Categoría']),
                        info_pago, '' if r['Categoría'] == 'Colaboración' else info_deuda, info_saldo),
                        anchura=85, continuacion=espacios(3), separa_palabras=True) + '\n'

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_análisis.format(datetime(año, mes, 1)))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(análisis)

filename = os.path.join(attach_path, 'Windows', nombre_análisis.format(datetime(año, mes, 1)))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(análisis)
