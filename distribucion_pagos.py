# GyG DISTRIBUCIÓN DE PAGOS
#
# Genera un archivo con el resumen de pagos en el mes, indicando aquellos inferiores,
# iguales o superiores a la cuota, así como los correspondientes a otros pagos.

"""
    PENDIENTE POR HACER
      - 

    NOTAS
      - 

    HISTORICO
      - Se ajustó la descripción de los pagos con monto superior a la cuota en el resumen de
        vigilancia (27/06/2021)
      - Se incorporó una comparación del total de cuotas de vigilancia respecto al estimado de
        gastos del mes (25/05/2021)
      - Hay una discrepancia entre los resultados mostrados por la Relación de Ingresos y la gráfica
        Gestión de Cobranzas, las cuales toman en cuenta sólo los ingresos del mes, y los mostrados
        por Distribución de Pagos, los cuales incluyen los anticipos de meses anteriores.
         -> Tomar en cuenta sólo los pagos realizados en el mes para unificar resultados
            (23/05/2021)
      - Se separa la opción para la visualización del detalle de los pagos en dos: Vigilancia y
        Otros Pagos (19/05/2021)
      - Se añade la opción para generar las tablas de detalle en fuente monoespaciada para ser
        enviadas por WhatsApp (19/05/2921)
      - La indicación "VENDIDA" en una celda del mes a analizar generó el error "TypeError: '>' not
        supported between instances of 'str' and 'float'"
         -> Corregido usando pandas.to_numeric(errors='coerce') (15/05/2021)
      - Se agregó la información relacionada con Otros Pagos (pagos diferentes a Vigilancia: Aporte
        Vigilantes, Cesta de Navidad, etc.) (06/05/2021)
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
from pandas import read_excel, merge, to_numeric
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

nombre_análisis  = GyG_constantes.txt_estadistica_de_pagos      # "GyG Estadistica de Pagos {:%Y-%m (%b)}.txt"
attach_path      = GyG_constantes.ruta_analisis_de_pagos        # "./GyG Recibos/Análisis de Pago"

excel_workbook   = GyG_constantes.pagos_wb_estandar             # '1.1. GyG Recibos.xlsm'
excel_resumen    = GyG_constantes.pagos_ws_resumen_reordenado   # 'R.VIGILANCIA (reordenado)'
excel_vigilancia = GyG_constantes.pagos_ws_vigilancia           # 'Vigilancia'

fmt_fecha        = "%m-%Y"

# tabla_encabezado = "Vecino               | Dirección      |      Pago   | Sobre cuota\n"
tabla_encabezado = "Vecino               | Dirección      |    Monto    |   Excedente\n"
tabla_separador  = "-" * 66 + "\n"
tabla_detalle    = "{:<20} | {:<14} | {:>11} | {:>11}\n"
tabla_encabezado_sin_sobrecuota = "Vecino               | Dirección      |    Monto\n"
tabla_separador_sin_sobrecuota  = "-" * 52 + "\n"
tabla_detalle_sin_sobrecuota    = "{:<20} | {:<14} | {:>11}{}\n"

LONG_BENEFICIARIO = 20


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def esVigilancia(x, ifTrue=True):
    return x == GyG_constantes.CATEGORIA_VIGILANCIA if ifTrue else x != GyG_constantes.CATEGORIA_VIGILANCIA


# a_evaluar = ['Yuraima Rodríguez']

def fecha_ultimo_pago(beneficiario, str_mes_año, categoría=GyG_constantes.CATEGORIA_VIGILANCIA, fecha_real=True):
    try:
        df_fecha = df_pagos[(df_pagos['Beneficiario'] == beneficiario) & \
                            (df_pagos['Categoría'] == categoría)]
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


def analiza_pagos(df_subset):
    niveles_de_pago = list()
    montos = list()
    montos_sobre_umbral = list()

    def gt_or_zero(value):
        return 0.00 if value < 0.00 else value

    # Barrer df_resumen para determinar si el pago es inferior, igual o superior a la cuota
    # y el monto por encima de la cuota, actualizando las listas «niveles_de_pago» y «montos»

    for index, r in df_subset.iterrows():

        # ¿Cómo manejar el monto del pago cuando la cuota es variable dentro del mes (como cuando se estuvo
        # manejando la cotización semanal del dólar)?
        #  -> Ubicar la última fecha de pago para el mes de referencia (fecha_referencia)
        #  -> Si el pago fue hecho ANTES de la fecha de referencia (01/04/2021, en este caso), debe ser tomado
        #     como un ANTICIPO, por lo que la cuota a utilizar es la CUOTA VIGENTE PARA LA FECHA DE INICIO DE
        #     MES (fecha_referencia).            
        #  -> Si el pago es realizado EN o DESPUÉS de la fecha de referencia, la cuota a utilizar es la CUOTA
        #     VIGENTE PARA LA FECHA DE PAGO.
        f_ultimo_pago = fecha_ultimo_pago(r['Beneficiario'], str_fecha_referencia)
        f_pago = fecha_referencia if f_ultimo_pago < fecha_referencia else f_ultimo_pago
        # cuota = cuotas_obj.cuota_vigente(r['Beneficiario'], fecha_referencia)
        cuota = cuotas_obj.cuota_vigente(r['Beneficiario'], f_pago)
        umbral = cuota * (1.0 + umbral_de_aceptación / 100.0)
        monto = gt_or_zero(r['Monto'] - cuota)
        monto_sobre_umbral = gt_or_zero(r['Monto'] - umbral)

        if r['Monto'] > umbral:
            niveles_de_pago.append('Superior')
            montos_sobre_umbral.append(monto)
        elif r['Monto'] < cuota:
            niveles_de_pago.append('Inferior')
            montos_sobre_umbral.append(0.00)
        else:
            niveles_de_pago.append('Igual')
            montos_sobre_umbral.append(0.00)
        montos.append(monto)

    return niveles_de_pago, montos, montos_sobre_umbral


def genera_resumen_vigilancia():

    am_pm = 'pm' if hoy.hour > 12 else 'm' if hoy.hour == 12 else 'am'
    análisis = ''.join([
                    f"GyG DISTRIBUCIÓN DE PAGOS, {fecha_referencia.strftime('%B/%Y').capitalize()}\n",
                    f"al {hoy.strftime('%d/%m/%Y %I:%M')} {am_pm}\n\n"
                ])

    # Elimina los registros que no corresponden al pago de vigilancia
    df_subset = df_pagos.loc[esVigilancia(df_pagos['Categoría'])]

    # Genera columnas con un indicador si el pago es inferior, igual o superior a la cuota
    # y el monto por encima de la cuota
    niveles_de_pago, montos, montos_sobre_umbral = analiza_pagos(df_subset)
    df_subset.insert(column='Nivel', value=niveles_de_pago, loc=df_subset.shape[1])
    df_subset.insert(column='Sobrecuota', value=montos, loc=df_subset.shape[1])
    df_subset.insert(column='Sobreumbral', value=montos_sobre_umbral, loc=df_subset.shape[1])

    if not muestra_detalle_vigilancia:
        análisis += "RESUMEN VIGILANCIA:\n"

    for nivel in ['Superior', 'Igual', 'Inferior']:
        df_subset_nivel = df_subset[df_subset['Nivel'] == nivel]
        num_pagos = df_subset_nivel.shape[0]
        mto_pagos = df_subset_nivel['Monto'].sum()
        mto_sobrecuotas = df_subset_nivel['Sobrecuota'].sum()
        mto_sobreumbral = df_subset_nivel['Sobreumbral'].sum()
        s = 's' if num_pagos != 1 else ''
        str_pago = edita_número(mto_pagos, num_decimals=0)
        str_sobrecuota = edita_número(mto_sobrecuotas, num_decimals=0)
        str_sobreumbral = edita_número(mto_sobreumbral, num_decimals=0)
        if muestra_detalle_vigilancia:
            análisis += wa_bold
            if umbral_de_aceptación > 0.00:
                if nivel == 'Superior':
                    análisis += f"VECINOS CON PAGOS SUPERIORES AL {str_umbral_de_aceptación}% DE LA CUOTA"
                elif nivel == 'Igual':
                    análisis += f"VECINOS CON PAGOS ENTRE LA CUOTA Y EL {str_umbral_de_aceptación}% DE LA MISMA"
                else:
                    análisis += f"VECINOS CON PAGOS INFERIORES A LA CUOTA"
            else:
               análisis += f"vecinos con pagos {nivel.lower()}es a la cuota".upper()
            análisis += wa_bold + '\n'
            if nivel == 'Superior':
                análisis += wa_table + '\n' + tabla_encabezado + tabla_separador
                for index, r in df_subset_nivel.iterrows():
                    análisis += tabla_detalle.format(
                            trunca_texto(edita_beneficiario(r['Beneficiario']), LONG_BENEFICIARIO),
                            trunca_texto(edita_dirección(r['Dirección']), 14),
                            edita_número(r['Monto'], num_decimals=0),
                            edita_número(r['Sobreumbral'], num_decimals=0) if nivel == 'Superior' else '')
                if num_pagos > 0:
                    análisis += tabla_separador
                análisis += tabla_detalle.format(
                        'Subtotales',
                        alinea_texto(f'{num_pagos} pago{s}', 14, 'derecha'),
                        str_pago,
                        str_sobreumbral)
            else:
                análisis += wa_table + '\n' + tabla_encabezado_sin_sobrecuota + tabla_separador_sin_sobrecuota
                for index, r in df_subset_nivel.iterrows():
                    señala_renglón = ' <-' if r['Sobrecuota'] > 0.00 else ''
                    análisis += tabla_detalle_sin_sobrecuota.format(
                            trunca_texto(edita_beneficiario(r['Beneficiario']), 20),
                            trunca_texto(edita_dirección(r['Dirección']), 14),
                            edita_número(r['Monto'], num_decimals=0),
                            señala_renglón)
                if num_pagos > 0:
                    análisis += tabla_separador_sin_sobrecuota
                análisis += tabla_detalle_sin_sobrecuota.format(
                        'Subtotales',
                        alinea_texto(f'{num_pagos} pago{s}', 14, 'derecha'),
                        str_pago, '')
            análisis += wa_table + '\n\n'
        else:
            if umbral_de_aceptación > 0.00:
                if categoría == 'Superior':
                    txt_descripción_categoría = f"con monto superior al {str_umbral_de_aceptación}% de la cuota" 
                elif categoría == 'Igual':
                    txt_descripción_categoría = f"entre la cuota y el {str_umbral_de_aceptación}% de la misma"
                else:
                    txt_descripción_categoría = "con monto inferior a la cuota"
            else:
                txt_descripción_categoría = f"con monto {nivel.lower()} a la cuota"
            # txt_aporte_adicional = f"(de los cuales, Bs. {str_sobreumbral} corresponden a aportes adicionales,\n" + \
            #                         "        entre cuotas atrasadas y excedentes a la misma)"
            txt_aporte_adicional = bloque_de_texto(
                ' '.join([
                        f'(Bs. {edita_número(mto_pagos - mto_sobreumbral, num_decimals=0)} en cuotas del mes,',
                        f'más Bs. {str_sobreumbral}',
                        'entre cuotas atrasadas y excedentes a la misma)'
                    ]),
                anchura=70, margen=8, continuacion='')
            análisis += ''.join([
                    '  - ',
                    f"{num_pagos:>3} pago{s} {txt_descripción_categoría} ",
                    f"por Bs. {str_pago}",
                    f"\n       {txt_aporte_adicional}" if nivel == "Superior" else "",
                    '\n'
                ])

    if not muestra_detalle_vigilancia:
        análisis += '\n'

    num_pagos = df_subset.shape[0]
    mto_pagos = df_subset['Monto'].sum()
    s = 's' if num_pagos != 1 else ''
    # análisis += ''.join([
    #                 f"Total: {num_pagos} pagos recibidos para un total de Bs. {edita_número(mto_pagos, num_decimals=0)} ",
    #                 f"({int(mto_pagos / estimado_de_gastos * 100)}% del\n       estimado)\n"
    #             ])
    análisis += bloque_de_texto(' '.join([
            f"Total: {num_pagos} pagos recibidos para un total de",
            f"Bs. {edita_número(mto_pagos, num_decimals=0)}",
            f"({int(mto_pagos / estimado_de_gastos * 100)}% del estimado)"
        ]), anchura=71, margen=7) + '\n'

    if muestra_comparación:
        mto_sobrecuotas = df_subset['Sobrecuota'].sum()
        mto_sobreumbral = df_subset['Sobreumbral'].sum()
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


def genera_resumen_otros_pagos():
    análisis = ''
    if not muestra_detalle_otros_pagos:
        análisis += '\n\nRESUMEN OTROS PAGOS:\n'

    # Elimina los registros que corresponden al pago de vigilancia
    df_subset = df_pagos.loc[esVigilancia(df_pagos['Categoría'], ifTrue=False)]
    categorías = [cat for cat in list(set(df_subset['Categoría'])) if not cat.isupper()]
    lista_num_pagos = list()
    lista_txt_pagos = list()
    lista_montos = list()
    total_num_pagos = 0
    total_mto_pagos = 0.00

    if len(categorías) == 0:
        análisis += '  *** No hay otros pagos registrados\n'
    else:
        for categoría in categorías:
            df_this_category = df_subset[df_subset['Categoría'] == categoría]
            num_pagos = df_this_category.shape[0]
            s = '' if num_pagos == 1 else 's'
            mto_pagos = df_this_category['Monto'].sum()

            if muestra_detalle_otros_pagos:
                análisis += ''.join([
                                    '\n\n',
                                    wa_bold, categoría.upper(), wa_bold,
                                    '\n', wa_table, '\n',
                                    tabla_encabezado_sin_sobrecuota,
                                    tabla_separador_sin_sobrecuota
                                ])
                for idx, r in df_this_category.iterrows():
                    análisis += tabla_detalle_sin_sobrecuota.format(
                                    trunca_texto(edita_beneficiario(r['Beneficiario']), 20),
                                    trunca_texto(edita_dirección(r['Dirección']), 14),
                                    edita_número(r['Monto'], num_decimals=0),
                                    ''
                                )
                análisis += ''.join([
                                    tabla_separador_sin_sobrecuota,
                                    tabla_detalle_sin_sobrecuota.format(
                                        'Subtotales',
                                        alinea_texto(f'{num_pagos} pago{s}', 14, 'derecha'),
                                        edita_número(mto_pagos, num_decimals=0),
                                        ''
                                    ),
                                    wa_table
                                ])
            else:
                lista_num_pagos.append(edita_número(num_pagos, num_decimals=0))
                lista_txt_pagos.append(f"pago{s}")
                lista_montos.append(edita_número(mto_pagos, num_decimals=0))
                total_num_pagos += num_pagos
                total_mto_pagos += mto_pagos
        if not muestra_detalle_otros_pagos:
            maxlen_niveles_de_pago = len(max(niveles_de_pago, key=len))
            maxlen_num_pagos = len(max(lista_num_pagos, key=len))
            maxlen_txt_pagos = len(max(lista_txt_pagos, key=len))
            maxlen_montos = len(max(lista_montos, key=len))
            for categoría, num_pago, txt_pago, monto in zip(niveles_de_pago, lista_num_pagos, lista_txt_pagos, lista_montos):
                análisis += ''.join([
                                    '  -  ',
                                    alinea_texto(categoría + ':', maxlen_niveles_de_pago + 1, 'izquierda'),
                                    '  ',
                                    alinea_texto(num_pago, maxlen_num_pagos, 'derecha'),
                                    ' ',
                                    alinea_texto(txt_pago + ' por', maxlen_txt_pagos + 4, 'izquierda'),
                                    ' Bs. ',
                                    alinea_texto(monto, maxlen_montos, 'derecha'),
                                    '\n'
                    ])
            if len(categorías) > 1:
                s = 's' if total_num_pagos != 1 else ''
                análisis += ''.join([
                                    '\n',
                                    f"Total: {total_num_pagos} pagos recibidos para un total de ",
                                    f"Bs. {edita_número(total_mto_pagos, num_decimals=0)}\n"
                    ])

    return análisis


#
# PROCESO
#

# Determina el mes actual, a fin de utilizarlo como opción por defecto
hoy = datetime.now()
fecha_análisis = datetime(hoy.year, hoy.month, 1)
# mes_actual = (fecha_análisis - relativedelta(days=1)).strftime('%m-%Y')
mes_actual = fecha_análisis.strftime('%m-%Y')
print()

# Selecciona el mes y año a procesar
mes_año = input_mes_y_año('Indique el mes y año a analizar', mes_actual, toma_opciones_por_defecto)

# # Selecciona el umbral de aceptación
# print()
# print("*** (umbral de aceptación: Porcentaje por encima de la cuota para")
# print("***                        considerarlo un aporte adicional)")
# umbral_de_aceptación = input_valor('Umbral de aceptación', 0.00, toma_opciones_por_defecto)
# str_umbral_de_aceptación = edita_número(umbral_de_aceptación, num_decimals=2).replace(',00', '')

# # Selecciona si se ordenan alfabéticamente los vecinos
# if umbral_de_aceptación > 0.00:
#     muestra_comparación = input_si_no(' -> muestra comparación con monto base', 'sí', toma_opciones_por_defecto)
# else:
#     muestra_comparación = False
# print()
umbral_de_aceptación = 0.00
muestra_comparación = False

# Selecciona si se muestran sólo los saldos deudores o no
muestra_detalle_vigilancia = input_si_no('Muestra el detalle de los pagos de vigilancia', 'no', toma_opciones_por_defecto)
muestra_detalle_otros_pagos = input_si_no('Muestra el detalle de los otros pagos', 'sí', toma_opciones_por_defecto)

# Selecciona si se ordenan alfabéticamente los vecinos
if muestra_detalle_vigilancia or muestra_detalle_otros_pagos:
    ordenado = input_si_no(' -> vecinos ordenados alfabéticamente', 'sí', toma_opciones_por_defecto)
    # Selecciona si se colocarán marcas adicionales para ser interpretadas por WhatsApp
    whatsapp = input_si_no(" -> para ser enviado por WhatsApp", 'no', toma_opciones_por_defecto)
else:
    ordenado = False
    whatsapp = False

# Prepara los caracteres de negrita, italizado y monoespaciado para los textos, en caso de WhatsApp
wa_bold, wa_italic, wa_table = ('*', '_', '```') if whatsapp else ('', '', '')


año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = datetime(año, mes, 1)
str_fecha_referencia = fecha_referencia.strftime(fmt_fecha)

# Inicializa el handler para el manejo de las cuotas
cuotas_obj = Cuota(excel_workbook)

print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))

# Lee la hoja de cálculo con el detalle de los pagos
df_pagos = read_excel(excel_workbook, sheet_name=excel_vigilancia)

# Elimina los pagos que no corresponden al mes de referencia
df_pagos = df_pagos[(df_pagos['Enviado'] == 'ü')  & \
                    (df_pagos['Mes'] == fecha_referencia)]

# Restringe las columnas a utilizar
df_pagos = df_pagos[['Beneficiario', 'Dirección', 'Monto', 'Fecha', 'Categoría']]

# Sumariza los pagos por beneficiario y categoría, dejando como fecha de pago la
# última fecha
df_subset_monto = df_pagos[['Beneficiario', 'Dirección', 'Categoría', 'Monto']]
df_subset_monto = df_subset_monto.groupby(['Beneficiario', 'Dirección', 'Categoría']).sum().reset_index()
df_subset_fecha = df_pagos[['Beneficiario', 'Dirección', 'Categoría', 'Fecha']]
df_subset_fecha = df_subset_fecha.groupby(['Beneficiario', 'Dirección', 'Categoría']).max().reset_index()

df_pagos = merge(df_subset_fecha, df_subset_monto)

# Ordena los beneficiarios en orden alfabético
if ordenado:
    df_pagos['_benef_sort_'] = df_pagos['Beneficiario'] \
        .apply(lambda benef: trunca_texto(edita_beneficiario(remueve_acentos(benef)), LONG_BENEFICIARIO))
    df_pagos.sort_values(by='_benef_sort_', inplace=True)

# Lee la hoja de cálculo con el resumen de los pagos de vigilancia
df_resumen = read_excel(excel_workbook, sheet_name=excel_resumen)

# Toma el estimado de gastos del mes de referencia
estimado_de_gastos = float(df_resumen[df_resumen['Beneficiario'] == 'PAGO ESTIMADO DE VIGILANCIA'][fecha_referencia])

# Elimina el area de memoria ocupado por la hoja de cálculo con el resumen de los pagos de vigilancia
del(df_resumen)

# Crea el archivo con el análisis
print(f"Creando análisis '{nombre_análisis.format(fecha_referencia)}'...")
print()

análisis = genera_resumen_vigilancia()
análisis += genera_resumen_otros_pagos()

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(análisis)

filename = os.path.join(attach_path, 'Windows', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(análisis)

