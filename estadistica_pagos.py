# GyG ESTADISTICA DE PAGOS
#
# Genera un archivo con el resumen de pagos en el mes, indicando aquellos inferiores,
# iguales o superiores a la cuota.

"""
    PENDIENTE POR HACER
      - 
    NOTAS
      - 

    HISTORICO
      - CORREGIDO: No se está separando adecuadamente el concepto "Cancelación Vigilancia, meses de
        Diciembre 2020 a Marzo 2021" en sus menses componentes: retorna '03-2021' en lugar de
        '12-2020|01-2021|02-2021|03-2021'
         -> Se ajustó la rutina 'separa_meses()' para contemplar rangos con años de inicio y final
            diferentes (12/04/2021)
      - Se agregó el "umbral de aceptación" para contemplar, como "montos adicionales", únicamente
        aquellos pagos por encima de este monto (10/04/2021)
      - Versión inicial (07/04/2021)
    
"""

print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *
from pandas import read_excel
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys

import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'


#
# DEFINE CONSTANTES
#

nombre_análisis  = GyG_constantes.txt_estadistica_de_pagos # "GyG Estadistica de Pagos {:%Y-%m (%b)}.txt"
attach_path      = GyG_constantes.ruta_analisis_de_pagos   # "./GyG Recibos/Análisis de Pago"

excel_workbook   = GyG_constantes.pagos_wb_estandar        # '1.1. GyG Recibos.xlsm'
excel_resumen    = GyG_constantes.pagos_ws_resumen_reordenado   # 'R.VIGILANCIA (reordenado)'
excel_detalle    = GyG_constantes.pagos_ws_vigilancia      # 'Vigilancia'

fmt_fecha        = "%m-%Y"

# tabla_encabezado = "Vecino               | Dirección      |      Pago   | Sobre cuota\n"
tabla_encabezado = "Vecino               | Dirección      |      Pago   |  Excedente\n"
tabla_separador  = "-" * 66 + "\n"
tabla_detalle    = "{:<20} | {:<14} | {:>11} | {:>11}\n"
tabla_encabezado_sin_sobrecuota = "Vecino               | Dirección      |      Pago\n"
tabla_separador_sin_sobrecuota  = "-" * 52 + "\n"
tabla_detalle_sin_sobrecuota    = "{:<20} | {:<14} | {:>11}{}\n"


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def get_street(address):
    # return address.index(' ', address.index(' ') + 1)
    grupos = re.findall(r'\w+', address)
    if grupos[0].lower() == "av":
        return "Avenida"
    return grupos[1] if len(grupos) > 0 else ''

def edita_beneficiario(beneficiario):
    return beneficiario.replace('Familia ', '')

def edita_dirección(dirección):
    return dirección.replace('Calle ', '').replace('Nros. ', '').replace('Nro. ', '')

def esVigilancia(x):
    return x == 'Vigilancia'

def alinea_texto(texto, anchura, alineación="derecha"):
    if alineación == 'derecha':
        return f"{texto:>{anchura}}"[-anchura:]
    elif alineación == 'centro':
        return f"{texto[:anchura]:^{anchura}}"
    elif alineación == 'izquierda':
        return f"{texto:<{anchura}}"[:anchura]
    else:
        return f"{trunca_texto('ALINEACION ERRADA', anchura)}"


# a_evaluar = ['Yuraima Rodríguez']

def fecha_ultimo_pago(beneficiario, str_mes_año, fecha_real=True):
    try:
        df_fecha = df_pagos[(df_pagos['Beneficiario'] == beneficiario) & (df_pagos['Meses'].str.contains(str_mes_año))]
    except:
        print(f"ERROR: fecha_ultimo_pago({beneficiario}, {str_mes_año}): {str(sys.exc_info()[1])}")
        return None
    # if beneficiario in a_evaluar: print(f" - {df_fecha['Fecha'] = }, {df_fecha['Meses'] = }")
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


def analiza_pagos():
    categorías = list()
    montos = list()
    montos_sobre_umbral = list()

    def gt_or_zero(value):
        return 0.00 if value < 0.00 else value

    # Barrer df_resumen para determinar si el pago es inferior, igual o superior a la cuota
    # y el monto por encima de la cuota, actualizando las listas «categorías» y «montos»

    for index, r in df_resumen.iterrows():

        # ¿Cómo manejar el monto del pago cuando la cuota es variable dentro del mes (como cuando se estuvo
        # manejando la cotización semanal del dólar)?
        #  -> Ubicar la última fecha de pago para el mes de referencia (fecha_referencia)
        #  -> Si el pago fue hecho ANTES de la fecha de referencia (01/04/2021, en este caso), debe ser tomado
        #     como un ANTICIPO, por lo que la cuota a utilizar es la CUOTA VIGENTE PARA LA FECHA DE INICIO DE
        #     MES (fecha_referencia).            
        #  -> Si el pago es realizado EN o DESPUÉS de la fecha de referencia, la cuota a utilizar es la CUOTA
        #     VIGENTE PARA LA FECHA DE PAGO.
        f_ultimo_pago = fecha_ultimo_pago(r['Beneficiario'], str_fecha_referencia)
        # print(f"DEBUG: {r['Beneficiario']}, {f_ultimo_pago}")
        f_pago = fecha_referencia if f_ultimo_pago < fecha_referencia else f_ultimo_pago
        # cuota = cuotas_obj.cuota_vigente(r['Beneficiario'], fecha_referencia)
        cuota = cuotas_obj.cuota_vigente(r['Beneficiario'], f_pago)
        umbral = cuota * (1.0 + umbral_de_aceptación / 100.0)
        monto = gt_or_zero(r[str_fecha_referencia] - cuota)
        monto_sobre_umbral = gt_or_zero(r[str_fecha_referencia] - umbral)

        if r[str_fecha_referencia] > umbral:
            categorías.append('Superior')
            montos_sobre_umbral.append(monto)
        elif r[str_fecha_referencia] < cuota:
            categorías.append('Inferior')
            montos_sobre_umbral.append(0.00)
        else:
            categorías.append('Igual')
            montos_sobre_umbral.append(0.00)
        montos.append(monto)

    return categorías, montos, montos_sobre_umbral


def genera_resumen():

    encabezado = f"GyG ESTADÍSTICA DE PAGOS, {fecha_referencia.strftime('%B/%Y').capitalize()}\n"
    am_pm = 'pm' if hoy.hour > 12 else 'm' if hoy.hour == 12 else 'am'
    encabezado += f"al {hoy.strftime('%d/%m/%Y %I:%M')} {am_pm}\n\n"

    análisis = encabezado
    if not muestra_detalle:
        análisis += "RESUMEN:\n"

    for categoría in ['Superior', 'Igual', 'Inferior']:
        df_subset = df_resumen[df_resumen['Categoría'] == categoría]
        num_pagos = df_subset.shape[0]
        mto_pagos = df_subset[str_fecha_referencia].sum()
        mto_sobrecuotas = df_subset['Sobrecuota'].sum()
        mto_sobreumbral = df_subset['Sobreumbral'].sum()
        s = 's' if num_pagos != 1 else ''
        str_pago = edita_número(mto_pagos, num_decimals=0)
        str_sobrecuota = edita_número(mto_sobrecuotas, num_decimals=0)
        str_sobreumbral = edita_número(mto_sobreumbral, num_decimals=0)
        if muestra_detalle:
            if umbral_de_aceptación > 0.00:
                if categoría == 'Superior':
                    análisis += f"VECINOS CON PAGOS SUPERIORES AL {str_umbral_de_aceptación}% DE LA CUOTA"
                elif categoría == 'Igual':
                    análisis += f"VECINOS CON PAGOS ENTRE LA CUOTA Y EL {str_umbral_de_aceptación}% DE LA MISMA"
                else:
                    análisis += f"VECINOS CON PAGOS INFERIORES A LA CUOTA"
            else:
               análisis += f"vecinos con pagos {categoría.lower()}es a la cuota".upper()
            análisis += '\n\n'
            if categoría == 'Superior':
                análisis += tabla_encabezado + tabla_separador
                for index, r in df_subset.iterrows():
                    análisis += tabla_detalle.format(
                            trunca_texto(edita_beneficiario(r['Beneficiario']), 20),
                            trunca_texto(edita_dirección(r['Dirección']), 14),
                            edita_número(r[str_fecha_referencia], num_decimals=0),
                            edita_número(r['Sobreumbral'], num_decimals=0) if categoría == 'Superior' else '')
                if num_pagos > 0:
                    análisis += tabla_separador
                análisis += tabla_detalle.format(
                        'Subtotales',
                        alinea_texto(f'{num_pagos} pago{s}', 14, 'derecha'),
                        str_pago,
                        str_sobreumbral)
            else:
                análisis += tabla_encabezado_sin_sobrecuota + tabla_separador_sin_sobrecuota
                for index, r in df_subset.iterrows():
                    señala_renglón = ' <-' if r['Sobrecuota'] > 0.00 else ''
                    análisis += tabla_detalle_sin_sobrecuota.format(
                            trunca_texto(edita_beneficiario(r['Beneficiario']), 20),
                            trunca_texto(edita_dirección(r['Dirección']), 14),
                            edita_número(r[str_fecha_referencia], num_decimals=0),
                            señala_renglón)
                if num_pagos > 0:
                    análisis += tabla_separador_sin_sobrecuota
                análisis += tabla_detalle_sin_sobrecuota.format(
                        'Subtotales',
                        alinea_texto(f'{num_pagos} pago{s}', 14, 'derecha'),
                        str_pago, '')
            análisis += '\n\n'
        else:
            if umbral_de_aceptación > 0.00:
                if categoría == 'Superior':
                    txt_descripción_categoría = f"con monto superior al {str_umbral_de_aceptación}% de la cuota" 
                elif categoría == 'Igual':
                    txt_descripción_categoría = f"entre la cuota y el {str_umbral_de_aceptación}% de la misma"
                else:
                    txt_descripción_categoría = "con monto inferior a la cuota"
            else:
                txt_descripción_categoría = f"con monto {categoría.lower()} a la cuota"
            txt_aporte_adicional = f"(de los cuales, Bs. {str_sobreumbral} corresponden a aportes adicionales)"
            análisis += ''.join([
                    '  - ',
                    f"{num_pagos:>3} pago{s} {txt_descripción_categoría} ",
                    f"por Bs. {str_pago}",
                    f"\n       {txt_aporte_adicional}" if categoría == "Superior" else "",
                    '\n'
                ])

    if not muestra_detalle:
        análisis += '\n'

    num_pagos = df_resumen.shape[0]
    mto_pagos = df_resumen[str_fecha_referencia].sum()
    s = 's' if num_pagos != 1 else ''
    análisis += f"Total: {num_pagos} pagos recibidos para un total de Bs. {edita_número(mto_pagos, num_decimals=0)}\n"

    if muestra_comparación:
        mto_sobrecuotas = df_resumen['Sobrecuota'].sum()
        mto_sobreumbral = df_resumen['Sobreumbral'].sum()
        pct_sobrecuotas = round(mto_sobrecuotas / mto_sobrecuotas * 100, 2)
        pct_sobreumbral = round(mto_sobreumbral / mto_sobrecuotas * 100, 2)

        análisis += ''.join([
            f"\n\nComparación del umbral de aceptación y el monto base:\n",
            f"  -  {alinea_texto('Montos mayores a la cuota:', 36, 'izquierda')}",
            f" {'Bs. ' + edita_número(mto_sobrecuotas, num_decimals=0):>16}",
            f"  ({edita_número(pct_sobrecuotas, 2):>6}%)\n",
            f"  -  {alinea_texto(f'Montos mayores al {str_umbral_de_aceptación}% de la cuota:', 36, 'izquierda')}",
            f" {'Bs. ' + edita_número(mto_sobreumbral, num_decimals=0):>16}",
            f"  ({edita_número(pct_sobreumbral, 2):>6}%)\n",
            f"{espacios(42)}{espacios(28, '-')}\n",
            f"  -  {alinea_texto('Diferencia:', 36, 'izquierda')}",
            f" {'Bs. ' + edita_número(mto_sobrecuotas - mto_sobreumbral, num_decimals=0):>16}",
            f"  ({edita_número(pct_sobrecuotas - pct_sobreumbral, 2):>6}%)\n"
        ])

    return análisis


#
# PROCESO
#

# Determina el mes actual, a fin de utilizarlo como opción por defecto
hoy = datetime.now()
fecha_análisis = datetime(hoy.year, hoy.month, 1)
mes_actual = (fecha_análisis - relativedelta(days=1)).strftime('%m-%Y')
print()

# Selecciona el mes y año a procesar
mes_año = input_mes_y_año('Indique el mes y año a analizar', mes_actual, toma_opciones_por_defecto)

# Selecciona el umbral de aceptación
print()
print("*** (umbral de aceptación: Porcentaje por encima de la cuota para")
print("***                        considerarlo un aporte adicional)")
umbral_de_aceptación = input_valor('Umbral de aceptación', 0.00, toma_opciones_por_defecto)
str_umbral_de_aceptación = edita_número(umbral_de_aceptación, num_decimals=2).replace(',00', '')

# Selecciona si se ordenan alfabéticamente los vecinos
if umbral_de_aceptación > 0.00:
    muestra_comparación = input_si_no(' -> muestra comparación con monto base', 'sí', toma_opciones_por_defecto)
else:
    muestra_comparación = False
print()

# Selecciona si se muestran sólo los saldos deudores o no
muestra_detalle = input_si_no('Muestra el detalle de los pagos', 'sí', toma_opciones_por_defecto)

# Selecciona si se ordenan alfabéticamente los vecinos
if muestra_detalle:
    ordenado = input_si_no(' -> vecinos ordenados alfabéticamente', 'sí', toma_opciones_por_defecto)
else:
    ordenado = False

año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = datetime(año, mes, 1)
str_fecha_referencia = fecha_referencia.strftime(fmt_fecha)

print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df_resumen = read_excel(excel_workbook, sheet_name=excel_resumen)
df_resumen.dropna(subset=['Categoría'], inplace=True)

# Inicializa el handler para el manejo de las cuotas
cuotas_obj = Cuota(excel_workbook)

# Cambia el nombre de la columna 2016 a datetimme(2016, 1, 1)
df_resumen.rename(columns={2016:datetime(2016, 1, 1)}, inplace=True)

# Cambia el formato del nombre de la columna a analizar por 'mm-aaaa'
df_resumen.rename(columns={fecha_referencia: str_fecha_referencia}, inplace=True)

# Selecciona las columnas a utilizar
df_resumen = df_resumen[['Beneficiario', 'Dirección', str_fecha_referencia]]

# Elimina los vecinos sin pagos en el período seleccionado
df_resumen = df_resumen[df_resumen[str_fecha_referencia] > 0.0]

# Ordena los beneficiarios en orden alfabético
if ordenado:
    df_resumen['_benef_sort_'] = df_resumen['Beneficiario'] \
        .apply(lambda benef: trunca_texto(edita_beneficiario(remueve_acentos(benef)), 20))
    df_resumen.sort_values(by='_benef_sort_', inplace=True)

# Lee la hoja de cálculo con el detalle de los pagos
df_pagos = read_excel(excel_workbook, sheet_name=excel_detalle)

# Elimina los registros que no no corresponden a pago de vigilancia
df_pagos = df_pagos.loc[esVigilancia(df_pagos['Categoría'])]

# Genera una columna con el resumen de los meses cancelados en cada pago
meses_cancelados = df_pagos['Concepto'].apply(lambda x: separa_meses(x, as_string=True))
df_pagos.insert(column='Meses', value=meses_cancelados, loc=df_pagos.shape[1])

# Genera columnas con un indicador si el pago es inferior, igual o superior a la cuota
# y el monto por encima de la cuota
categorías, montos, montos_sobre_umbral = analiza_pagos()
df_resumen.insert(column='Categoría', value=categorías, loc=df_resumen.shape[1])
df_resumen.insert(column='Sobrecuota', value=montos, loc=df_resumen.shape[1])
df_resumen.insert(column='Sobreumbral', value=montos_sobre_umbral, loc=df_resumen.shape[1])

# Crea el archivo con el análisis
print(f"Creando análisis '{nombre_análisis.format(fecha_referencia)}'...")
print('')

análisis = genera_resumen()

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(análisis)

filename = os.path.join(attach_path, 'Windows', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(análisis)
