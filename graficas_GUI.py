# GyG GRAFICAS
#
# Genera las gráficas de 'Gestión de Cobranzas', 'Pagos 100% Equivalentes', 'Pagos recibidos en el
# mes equivalentes a cuotas completas', y 'Cuotas Recibidas en el Mes' en base a la última información
# cargada en la hoja de cálculo.

"""
    POR HACER
    -   

    HISTORICO
    -   Mostrar el número de pagos recibidos directamente sobre la gráfica y no como un comentario
        aparte.
         -> Se actualizaron las gráficas 'Gestión de Cobranzas' y 'Cuotas Recibidas en el Mes' con las
            cantidades de cuotas en el mes y en el promedio de meses anteriores en la fecha de referencia
            (27/05/2020)
    -   Agregar una gráfica donde se comparen la cantidad de cuotas completas recibidas día a día con
        el promedio de los últimos 'n' meses, con el objeto de informar a los vecinos sobre la gestión
        de cobranzas. (23/05/2020)
    -   La gráfica 'Pagos recibidos en el mes equivalentes a cuotas completas' no está tomando en cuenta
        el día del análisis en meses anteriores al actual.
         -> En distribución_de_pagos() y gráfica_3() se filtraron aquellos pagos posteriores a la fecha
            de referencia (23/05/2020)
    -   Ajustado el nombre de la curva 'Promedio 5 últimos meses' a 'Promedio de dic/2019 a abr/2020'
        en la gráfica 1 (Gestión de Cobranzas)
        (15/05/2020)
    -   Corregido el error en la gráfica de Cuotas Equivalentes en la opción horizontal donde se
        mostraban distribuciones y textos erroneos en las barras (12/05/2020)
    -   Ajustadas las gráficas Pagos 100% Equivalentes y Cuotas Equivalentes para que puedan ser
        desplegadas como barras horizontales (01/05/2020)
    -   Dado que los pagos de meses atrasados se hece en base a la última cuota, la gráfica de Cuotas
        100% Equivalentes podrá mostrar cifras "infladas" al utilizar como cuota de referencia la vigente
        para dicho mes, especialmente cuando ha habido cambios en el monto.
         -> Hacer un análisis similar al realizado en "Análisis de Pagos", desmembrando el pago en cada
            uno de los meses afectados y dividiéndolo por la cuota vigente para la fecho de pago.
        Se incorporó la grafica_3() para tal efecto (29/04/2020)
    -   Incluir opción para generar las gráficas del mes anterior, sin tener que seleccionar la fecha
        correspondiente:
         -> fecha_de_referencia = <último día del mes inmediato anterior> √
         -> Los gráficos son guardados en archivos del tipo <nombre de gráfica> <año>-<mes> (<mes>).png √
         -> toma_opciones_por_defecto genera archivos <nombre de gráfica>.png o <nombre de grafica>-
            <mes> (<mes>).png dependiendo si se ejecuta en último de mes o no √
         -> mantenimiento.py debe ajustarse para incluir la posibilidad de borrar las gráficas generadas
            correspondientes a los tres últimos meses √
        Cambios realizados (20/04/2020)
    -   Se agregar opciones para seleccionar la cantidad de meses a mostrar o promediar
        en las gráficas en modo GUI (11-04-2020)
    -   Se independizaron las gráficas de la información contenida en la pestaña
        'Cobranzas', tomándolas directamente de 'Vigilancia' y 'Resumen Vigilancia',
        permitiendo generar gráficas con una fecha de referencia diferente
        (10/04/2020)
    -   Versión inicial en base a la información cargada en la pestaña 'Cobranzas'
        (01/04/2020)

    REVISIONES
    -   Revisar gráfica_3(): Se muestran 62 pagos para el mes de abril, mientras que gráfica_2()
        muestra 69, que coincide con el total de abril / cuota de abril.
         -> Gráfica_3() muestra que 26 PAGOS PARA ABRIL SE HICIERON EN ABRIL. Eso quiere decir que
            el resto (69 cuotas) debieron hacerse en marzo. En efecto, 6 pagos fueron para meses
            posteriores a marzo
        Revisado 03/05/2020

"""
print('Cargando librerías...')
import GyG_constantes
from GyG_utilitarios import *
import PySimpleGUI as sg

# import plotly
import plotly.graph_objs as go
from plotly.offline import plot
from pandas import read_excel, pivot_table
from scipy import stats
from numpy import mean
from datetime import datetime
from dateutil.relativedelta import relativedelta
# import swifter        # .swifter.progress_bar(False).apply()

import sys
import os
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

import warnings
warnings.simplefilter("ignore", category=RuntimeWarning)


# Define algunas constantes
excel_workbook      = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_ws_cobranzas  = GyG_constantes.pagos_ws_cobranzas
excel_ws_vigilancia = GyG_constantes.pagos_ws_vigilancia
excel_ws_resumen    = GyG_constantes.pagos_ws_resumen

num_meses_gestion_cobranzas   =  5      # Nro de meses a promediar en gráfica 'Gestión de Cobranzas'
num_meses_cuotas_equivalentes = 12      # Nro de meses a mostrar en gráfica 'Cuotas Equivalentes'

PUNTO_DE_EQUILIBRIO           = 55      # Cantidad de familias usadas en el cálculo de la cuota

FORMATO_MES         = '%b %Y'           # <mes abreviado> '.' <año>

CALENDAR_ICON       = os.path.join(GyG_constantes.rec_imágenes, '62925-spiral-calendar-icon.png')
CALENDAR_SIZE       = (16, 16)
CALENDAR_SUBSAMPLE  = 8

VERBOSE             = False             # Muestra mensajes adicionales

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'
    if not toma_opciones_por_defecto:
        print(f"Uso: {sys.argv[0]} [--toma_opciones_por_defecto]")
        sys.exit(1)


#
# RUTINAS DE UTILIDAD
#

def es_fin_de_mes(fecha):
    fecha_aux = fecha + relativedelta(months=1)
    fin_de_mes = datetime(fecha_aux.year, fecha_aux.month, 1) - relativedelta(days=1)
    return fecha == fin_de_mes


#
# GRAFICAS
#

def genera_gráficas():

    print('Generando gráficas...')

    if g1:
        gráfica_1()     # Gestión de Cobranzas
    if g2:
        gráfica_2()     # Pagos 100% Equivalentes (montos distribuidos a lo largo de los meses)
    if g3:
        gráfica_3()     # Cuotas Equivalentes (montos recibidos en el mes)
    if g4:
        gráfica_4()     # Cuotas recibidas

def gráfica_1():
    global num_meses_gestion_cobranzas

    grafica_nombre = 'Gestión de Cobranzas'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"1. Gestion_de_cobranzas {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '1. Gestion_de_cobranzas.png'


    if in_GUI_mode:
        num_meses_gestion_cobranzas = int(values['_g1_nro_meses_'])
        print(f'  - {grafica_nombre}...')

    # Define las fechas de referencia
    dia_de_referencia = fecha_de_referencia.day
    mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
    fecha_inicial = fecha_de_referencia - relativedelta(months=num_meses_gestion_cobranzas)
    mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)
    # --------------------
    mes_final = mes_actual - relativedelta(months=1)
    str_mes_inicial = mes_inicial.strftime('%b'+('/%Y' if mes_inicial.year != mes_final.year else ''))
    str_mes_final = mes_final.strftime('%b/%Y')
    # --------------------

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los estimados del servicio de vigilancia')
    ws_resumen = read_excel(excel_workbook, sheet_name=excel_ws_resumen)
    ws_cuotas_mensuales = ws_resumen[(ws_resumen['Beneficiario'] == 'CUOTAS MENSUALES')]
    ws_pago_estimado = ws_resumen[(ws_resumen['Beneficiario'] == \
                                                    'PAGO ESTIMADO DE VIGILANCIA (Vigilantes, Pasivos y Consumibles)')]


    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los pagos realizados')
    ws_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

    # Selecciona los pagos de vigilancia entre la fecha de referencia y num_meses_gestion_cobranzas atrás
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Filtra los pagos realizados por Categoría y Fecha')
    ws_vigilancia = ws_vigilancia[(ws_vigilancia['Categoría'] == 'Vigilancia') & \
                                  (ws_vigilancia['Fecha'] <= fecha_de_referencia) & \
                                  (ws_vigilancia['Fecha'] >= mes_inicial)]
    ws_vigilancia = ws_vigilancia[['Mes', 'Día', 'Monto', 'Fecha']]

    # Valida que los datos estén completos y agrega registros en cero para los datos faltantes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Valida la completitud de valores en el período ' + \
                                                                     'y rellena con ceros')
    fecha = mes_inicial
    a_agregar = []
    while fecha <= fecha_de_referencia:
        if ws_vigilancia[ws_vigilancia['Fecha'] == fecha].empty:
            a_agregar.append({'Mes':datetime(fecha.year, fecha.month, 1), 'Día':fecha.day, 'Monto':0.0, 'Fecha':fecha})
        fecha = fecha + relativedelta(days=1)

    ws_vigilancia = ws_vigilancia.append(a_agregar, ignore_index=True)

    # Totaliza los pagos de vigilancia por día y mes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Totaliza los pagos de vigilancia por día y mes')
    ws_vigilancia = pivot_table(ws_vigilancia, values='Monto', index=['Día', 'Mes'], aggfunc=sum)   # , fill_value=0
    ws_vigilancia = ws_vigilancia.reset_index() 

    # Normalizar los montos dividiéndolo por el estimado del mes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Normaliza los pagos de vigilancia en base al estimado ' + \
                                                                     'de pago del mes')
    ws_vigilancia['Monto_normalizado'] = ws_vigilancia.apply(
                                            lambda ws: ws['Monto'] / ws_pago_estimado[ws['Mes']], axis=1)
    ws_vigilancia['_Cuotas_'] = ws_vigilancia.apply(
                                            lambda ws: ws['Monto'] / ws_cuotas_mensuales[ws['Mes']], axis=1)

    # Genera columnas a graficar
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los datos del mes actual')
    ws_mes_actual = ws_vigilancia[ws_vigilancia['Mes'] == mes_actual].copy().reset_index()
    ws_mes_actual['Acumulado'] = ws_mes_actual['Monto_normalizado'].cumsum()
    ws_mes_actual['Cuotas'] = ws_mes_actual['_Cuotas_'].cumsum()
    total_recaudado = ws_mes_actual['Monto'].sum()
    pct_del_estimado = total_recaudado / ws_pago_estimado[mes_actual] * 100

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los datos del promedio de los meses anteriores')
    ws_meses_anteriores = ws_vigilancia[ws_vigilancia['Mes'] < mes_actual]
    ws_meses_anteriores = pivot_table(ws_meses_anteriores, values='Monto_normalizado', index=['Día'], fill_value=0)
    ws_meses_anteriores = ws_meses_anteriores.reset_index()
    ws_meses_anteriores['Acumulado'] = ws_meses_anteriores['Monto_normalizado'].cumsum()
    ws_cuotas_anteriores = ws_vigilancia[ws_vigilancia['Mes'] < mes_actual]
    ws_cuotas_anteriores = pivot_table(ws_cuotas_anteriores, values='_Cuotas_', index=['Día'], fill_value=0)
    ws_cuotas_anteriores = ws_cuotas_anteriores.reset_index()
    ws_meses_anteriores['Cuotas'] = ws_cuotas_anteriores['_Cuotas_'].cumsum()


    # Mes actual
    trace_mes_actual = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = ws_mes_actual['Acumulado'],
                mode = 'lines',
                name = fecha_de_referencia.strftime(FORMATO_MES),
                line = dict(
                            width = 3,
                            color = 'orange',
                        )
            )

    # Regresión lineal
    x = ws_meses_anteriores['Día']
    y = ws_mes_actual['Acumulado']
    slope, intercept, r_value, p_value, std_err = stats.linregress(x[:dia_de_referencia], y[:dia_de_referencia])
    linea_lr = [slope * xi + intercept for xi in x]

    trace_regresion = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = linea_lr,
                mode = 'lines',
                name = 'Ajuste lineal ' + fecha_de_referencia.strftime(FORMATO_MES),
                line = dict(
                            width = 2,
                            color = 'orange',
                            dash = 'dot',
                        )
            )

    # Promedio de los últimos 5 meses
    trace_prom_5_ultimos_meses = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = ws_meses_anteriores['Acumulado'],
                mode = 'lines',
                # name = f"Promedio {num_meses_gestion_cobranzas} últimos meses",
                name = f"Promedio de {str_mes_inicial} a {str_mes_final}",
                line = dict(
                            width = 3,
                            color = 'blue',
                        )
            )

    # Pago Vigilantes (estimado)
    trace_pago_vigilantes = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = [1.0 for _ in range(31)],
                mode = 'lines',
                name = 'Pago Vigilantes (estimado)',
                line = dict(
                            color = 'gray',
                            dash = 'solid',
                            width = 2,
                        )
            )

    data = [trace_mes_actual, trace_prom_5_ultimos_meses, trace_pago_vigilantes, trace_regresion]

    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b %Y")
    comentario = f"Recaudado: Bs. {edita_número(total_recaudado, num_decimals=0)} " + \
                 f"({edita_número(pct_del_estimado, num_decimals=0)}% del estimado)"
    día = fecha_de_referencia.day

    layout = go.Layout(
                title = dict(
                            text = f'<b>Gestión de Cobranzas</b><br>{subtítulo}',
                        ),
                width = 1000,
                height = 600,
                xaxis = dict(
                            nticks = 31,
                            showgrid = True, gridwidth=1, gridcolor='lightgray',
                            showline=True, linewidth=1, linecolor='black',
                        ),
                yaxis = dict(
                            title = '% Cobertura de Pagos de Vigilantes, Pasivos y Consumibles',
                            tickformat = ',.0%',
                            showline=True, linewidth=1, linecolor='black',
                        ),
                legend = dict(
                            orientation = 'h',
                            tracegroupgap = 20,
                        ),
                annotations = [
                        dict(
                            text = comentario,
                            font = dict(
                                    size = 14,
                                ),
                            showarrow = False,
                            xref = 'paper',
                            yref = 'paper',
                            x = 0.98,
                            y = 0.05,
                            align = 'right',
                            valign = 'bottom',
                        ),
                    ],
#                paper_bgcolor = 'white',
                plot_bgcolor = 'white',
            )

    fig = go.Figure(data = data, layout = layout)

    ultimo_indice = ws_mes_actual['Acumulado'].index[-1]
    dx, dy = 20, 40
    actual_ge_anteriores = ws_mes_actual['Acumulado'][ultimo_indice] >= ws_meses_anteriores['Acumulado'][ultimo_indice]

    delta_x = -dx if actual_ge_anteriores else dx
    if ultimo_indice == 0:
        delta_x = abs(delta_x)
    elif ultimo_indice == 30:
        delta_x = -abs(delta_x)
    fig.add_annotation(
        x = ws_meses_anteriores['Día'][ultimo_indice],
        y = ws_mes_actual['Acumulado'][ultimo_indice],
        text = f"{int(round(ws_mes_actual['Cuotas'][día - 1]))} cuotas",
        xref = "x",
        yref = "y",
        ax = delta_x,
        ay = -dy if actual_ge_anteriores else dy,
        showarrow = True,
        arrowhead = 3,
    )
    delta_x = -dx if not actual_ge_anteriores else dx
    if ultimo_indice == 0:
        delta_x = abs(delta_x)
    elif ultimo_indice == 30:
        delta_x = -abs(delta_x)
    fig.add_annotation(
        x = ws_meses_anteriores['Día'][ultimo_indice],
        y = ws_meses_anteriores['Acumulado'][ultimo_indice],
        text = f"{int(round(ws_meses_anteriores['Cuotas'][día - 1]))} cuotas",
        xref = "x",
        yref = "y",
        ax = delta_x,
        ay = -dy if not actual_ge_anteriores else dy,
        showarrow = True,
        arrowhead = 3,
    )

    if g1_show:
        print(f'    . Desplegando imagen...')
        fig.show()

    if g1_save:
        print(f'    . Grabando "{grafica_png}"...')
        img_bytes = fig.to_image(format="png")
        with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
            f.write(img_bytes)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Gráfica concluida')


def gráfica_2():
    global num_meses_cuotas_equivalentes

    VERTICAL_BARS = not g2_horizontal

    grafica_nombre = 'Pagos 100% equivalentes'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"2. Pagos_100pct_equivalentes {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '2. Pagos_100pct_equivalentes.png'


    if in_GUI_mode:
        num_meses_cuotas_equivalentes = int(values['_g2_nro_meses_'])
        print(f'  - {grafica_nombre}...')

    dia_de_referencia = fecha_de_referencia.day
    mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los estimados del servicio de vigilancia ' + \
                                                                     'y total de pagos por mes')
    ws_resumen = read_excel(excel_workbook, sheet_name=excel_ws_resumen)
    ws_cuotas_mensuales =  ws_resumen[(ws_resumen['Beneficiario'] == 'CUOTAS MENSUALES')]
    ws_totales_mensuales = ws_resumen[(ws_resumen['Dirección']    == 'TOTAL')]

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los rangos de la imagen')
    meses = [mes_actual - relativedelta(months=mes) for mes in reversed(range(num_meses_cuotas_equivalentes +1))]
    if VERTICAL_BARS:
        meses_ed = [mes.strftime(FORMATO_MES) for mes in meses]
        pagos_100pct_equivalentes = [int(round(float(ws_totales_mensuales[mes]) / float(ws_cuotas_mensuales[mes]), ndigits=0)) \
                                        for mes in meses]
        promedio = int(round(mean(pagos_100pct_equivalentes[:-1]), ndigits=0))
    else:
        meses_ed = [mes.strftime(FORMATO_MES) for mes in reversed(meses)]
        pagos_100pct_equivalentes = [int(round(float(ws_totales_mensuales[mes]) / float(ws_cuotas_mensuales[mes]), ndigits=0)) \
                                        for mes in reversed(meses)]
        promedio = int(round(mean(pagos_100pct_equivalentes[1:]), ndigits=0))
    promedios = [promedio for _ in meses]
    pto_equilibrio = [PUNTO_DE_EQUILIBRIO for _ in meses]

    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b %Y")


    # 100% Equivalentes
    trace_100pct_equiv = go.Bar(
                x = meses_ed if VERTICAL_BARS else pagos_100pct_equivalentes,
                y = pagos_100pct_equivalentes if VERTICAL_BARS else meses_ed,
                orientation = 'v' if VERTICAL_BARS else 'h',
                name = 'Pagos 100% eqv.',
                text = [f"{y:.0f}" for y in pagos_100pct_equivalentes],
                textposition = "outside",
                marker_color = 'cornflowerblue',        # '#5b9bd5',
                # marker = dict(
                #             color = 'mediumblue',
                #         ),
            )

    # Promedio últimos 12 meses
    trace_prom_12_ultimos_meses = go.Scatter(
                x = meses_ed if VERTICAL_BARS else promedios,
                y = promedios if VERTICAL_BARS else meses_ed,
                mode = 'lines',
                name = f"Promedio últimos {num_meses_cuotas_equivalentes} meses = {promedio}; " + \
                       f"actual = {pagos_100pct_equivalentes[-1 if VERTICAL_BARS else 0]}",
                line = dict(
                            color = '#ed7d31',          # 'orange',
                            width = 2,
                        ),
            )

    # Punto de equilibrio
    trace_punto_equilibrio = go.Scatter(
                x = meses_ed if VERTICAL_BARS else pto_equilibrio,
                y = pto_equilibrio if VERTICAL_BARS else meses_ed,
                mode = 'lines',
                name = f"Punto de equilibrio = {pto_equilibrio[0]} pagos 100% eqv.",
                line = dict(
                            color = 'yellow',           # '#ffc000',
                            dash = 'dash',
                            width = 2,
                        ),
            )

    data = [trace_100pct_equiv, trace_prom_12_ultimos_meses, trace_punto_equilibrio]

    layout = dict(
                title = f'<b>Pagos 100% Equivalentes</b><br>{subtítulo}',
                width = 1000,
                height = 600,
                legend = dict(
                            orientation = 'h',
                            tracegroupgap = 20,
                        ),
                xaxis_tickangle = -30 if VERTICAL_BARS else 0,
#                paper_bgcolor = 'white',
                plot_bgcolor = 'white',
            )

#    subtítulo = 'al ' + celda('A36').strftime("%d %b %Y")
#    comentario = celda('B18')

    fig = go.Figure(data = data, layout = layout)

    if g2_show:
        print(f"    . Desplegando imagen {'vertical' if VERTICAL_BARS else 'horizontal'}...")
        fig.show()

    if g2_save:
        print(f'    . Grabando "{grafica_png}"...')
        img_bytes = fig.to_image(format="png")
        with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
            f.write(img_bytes)


# def gráfica_3():
#     from GyG_cuotas import Cuota

#     grafica_nombre = 'Pagos 100% equivalentes'
#     if es_fin_de_mes(fecha_de_referencia):
#         grafica_png = f"3. Cuotas_equivalentes {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
#     else:
#         grafica_png =  '3. Cuotas_equivalentes.png'


#     if in_GUI_mode:
#         num_meses_cuotas_equivalentes = int(values['_g3_nro_meses_'])
#         print(f'  - {grafica_nombre}...')

#     dia_de_referencia = fecha_de_referencia.day
#     mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
#     fecha_inicial = fecha_de_referencia - relativedelta(months=num_meses_cuotas_equivalentes)
#     mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)

#     cuotas_obj = Cuota(excel_workbook)

#     if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los pagos por mes del período seleccionado')
#     ws_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

#     # Selecciona los pagos de vigilancia entre la fecha de referencia y num_meses_cuotas_equivalentes atrás
#     ws_vigilancia = ws_vigilancia[(ws_vigilancia['Categoría'] == 'Vigilancia')  & \
#                                   (ws_vigilancia['Mes'] <= mes_actual) & \
#                                   (ws_vigilancia['Mes'] >= mes_inicial)]
#     ws_vigilancia = ws_vigilancia[['Mes', 'Beneficiario', 'Fecha', 'Monto']]

#     if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Agrega la cuota vigente para la fecha de pago')
#     ws_vigilancia['Cuota'] = ws_vigilancia.apply(
#                                   lambda r: cuotas_obj.cuota_vigente(r['Beneficiario'], r['Fecha']), axis=1)
#     ws_vigilancia['Pagos eqv'] = ws_vigilancia.apply(
#                                   lambda r: float(r['Monto']) / float(r['Cuota']), axis=1)

#     if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Totalizando Pagos Equivalentes por Mes...')
#     ws_vigilancia = pivot_table(ws_vigilancia, values=['Pagos eqv'], index=['Mes'], aggfunc=sum).reset_index()

#     if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los rangos de la imagen')
#     meses  = ws_vigilancia['Mes'].to_list()
#     meses_ed = [mes.strftime(FORMATO_MES) for mes in meses]
#     pagos_100pct_equivalentes = ws_vigilancia['Pagos eqv'].to_list()
#     promedio = int(round(mean(pagos_100pct_equivalentes[:-1]), ndigits=0))
#     promedios = [promedio for _ in meses]
#     pto_equilibrio = [PUNTO_DE_EQUILIBRIO for _ in meses]

#     subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b %Y")


#     # 100% Equivalentes
#     trace_100pct_equiv = go.Bar(
#                 x = meses_ed,
#                 y = pagos_100pct_equivalentes,
#                 name = '100% eqv.',
#                 text = [f"{y:.0f}" for y in pagos_100pct_equivalentes],
#                 textposition = "outside",
#                 marker_color = 'cornflowerblue',        # '#5b9bd5',
#                 # marker = dict(
#                 #             color = 'mediumblue',
#                 #         ),
#             )

#     # Promedio últimos 12 meses
#     trace_prom_12_ultimos_meses = go.Scatter(
#                 x = meses_ed,
#                 y = promedios,
#                 mode = 'lines',
#                 name = f"Promedio últimos {num_meses_cuotas_equivalentes} meses = {promedio}; " + \
#                        f"actual = {int(round(pagos_100pct_equivalentes[-1], 0))}",
#                 line = dict(
#                             color = '#ed7d31',          # 'orange',
#                             width = 2,
#                         ),
#             )

#     # Punto de equilibrio
#     trace_punto_equilibrio = go.Scatter(
#                 x = meses_ed,
#                 y = pto_equilibrio,
#                 mode = 'lines',
#                 name = f"Punto de equilibrio = {pto_equilibrio[0]} pagos 100% eqv.",
#                 line = dict(
#                             color = 'yellow',           # '#ffc000',
#                             dash = 'dash',
#                             width = 2,
#                         ),
#             )

#     data = [trace_100pct_equiv, trace_prom_12_ultimos_meses, trace_punto_equilibrio]

#     layout = dict(
#                 title = f'<b>Pagos recibidos en el mes equivalentes a cuotas completas</b><br>{subtítulo}',
#                 width = 1000,
#                 height = 600,
#                 legend = dict(
#                             orientation = 'h',
#                             tracegroupgap = 20,
#                         ),
#                 xaxis_tickangle = -30,
# #                paper_bgcolor = 'white',
#                 plot_bgcolor = 'white',
#             )

# #    subtítulo = 'al ' + celda('A36').strftime("%d %b %Y")
# #    comentario = celda('B18')

#     fig = go.Figure(data = data, layout = layout)

#     if g3_show:
#         print(f'    . Desplegando imagen...')
#         fig.show()

#     if g3_save:
#         print(f'    . Grabando "{grafica_png}"...')
#         img_bytes = fig.to_image(format="png")
#         with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
#             f.write(img_bytes)


def distribución_de_pagos():
    from openpyxl import load_workbook
    from pyparsing import Word, Regex, Literal, OneOrMore, ParseException
    import warnings
    warnings.simplefilter("ignore", category=UserWarning)

    meses            = ['enero',      'febrero', 'marzo',     'abril',
                        'mayo',       'junio',   'julio',     'agosto',
                        'septiembre', 'octubre', 'noviembre', 'diciembre']
    meses_abrev      = ['ene', 'feb', 'mar', 'abr', 'may', 'jun',
                        'jul', 'ago', 'sep', 'oct', 'nov', 'dic']
    conectores       = ['a', '-']
    textos_anticipos = ['adelanto', 'anticipo'   ]
    textos_saldos    = ['ajuste',   'complemento', 'diferencia', 'saldo']
    modificadores    = ['anticipo', 'saldo']

    tokens_validos = meses + meses_abrev + conectores


    def ajusta_fecha(ref):
        """ ajusta_fecha toma una fecha 'ref' de la forma '%m-%Y' (ej. 'may.2019') y la convierte en
            '%Y-%m' ('2019-05') para facilitar su comparación """
        return f"{ref[3:7]}-{ref[0:2]}"

    def separa_meses(mensaje, as_string=False, muestra_modificador=False):
        import re
        
        tokens_validos = meses + meses_abrev + conectores + modificadores

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
            token = meses[meses_abrev.index(x)] if x in meses_abrev else x
            if token.isdigit():
                if mensaje_anterior != None:
                    mensaje_final = mensaje_anterior + mensaje_final
                last_year = token
                last_month = None
                mensaje_anterior = None
            elif token in meses:
                if mensaje_anterior != None:
                    mensaje_final = mensaje_anterior + mensaje_final
                if maneja_conector:
                    try:
                        n_last_month = meses.index(last_month)
                    except:
                        continue    # ignora los mensajes que contienen textos del tipo:
                                    # "(saldo a favor: Bs. 69.862,95)"
                    n_token = meses.index(token)
                    for t in reversed(range(n_token + 1, n_last_month)):
    #                        mensaje_final = [f"{meses_abrev[t]}.{last_year}"] + mensaje_final
                        mensaje_final = [f"{t+1:02}-{last_year}"] + mensaje_final
                    maneja_conector = False
                last_month = token
    #                mensaje_anterior = [f"{meses_abrev[meses.index(last_month)]}.{last_year}"]
                mensaje_anterior = [f"{meses.index(last_month)+1:02}-{last_year}"]
            elif x in conectores:
                maneja_conector = True
            elif x in modificadores and muestra_modificador:
    #                mensaje_final = [f"{meses_abrev[meses.index(last_month)]}.{last_year} {x}"] + mensaje_final
                mensaje_final = [f"{meses.index(last_month)+1:02}-{last_year} {x}"] + mensaje_final
                mensaje_anterior = None

        if mensaje_anterior != None:
            mensaje_final = mensaje_anterior + mensaje_final

        if as_string:
            mensaje_final = '|'.join(mensaje_final)

        return mensaje_final


    def análisis_de_pagos(fecha_referencia):

        # Gramática de expresiones aritméticas
        operator  = Regex(r'(?<!--[\+\-\^\*/%])[\+\-]|[\^\*/%!]')
        function  = Regex(r'[a-zA-Z_][a-zA-Z0-9_]*(?=([ \t]+)?\()')
        variable  = Regex(r'[+-]?[a-zA-Z_][a-zA-Z0-9_]*(?!([ \t]+)?\()')
        number    = Regex(r'[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?')       # Ej. -125,54e-5
        lbrace    = Word('(')
        rbrace    = Word(')')
        linebreak = Word('\n')
        skip      = Word(' \t')                                             # espacio o tabulador

        lexOnly = operator | function | variable | number | lbrace \
            | rbrace | linebreak | skip
        lexAllOnly = OneOrMore(lexOnly)

        def evalua_celda(formula, modificador):
            try:
                parsed_expression = lexAllOnly.parseString(formula[1:], parseAll=True)
            except ParseException as err:
                print(f"{err.line}\n{' '*(err.column-1)}^\n{err}\n")
                return None
            result = ''
            count_braces = 0
            start_idx = 0
            curr_idx = 0
            for token in parsed_expression:
                if token == '(':
                    count_braces += 1
                elif token == ')':
                    count_braces -= 1
                elif token in ['+', '-'] and count_braces == 0:
                    result += ''.join(parsed_expression[start_idx: curr_idx])
                    if len(result) > 0 and modificador in [modificadores[0], '']:   # Cancela un Anticipo
                        return eval(result)
                    result = ''
                    start_idx = curr_idx
                curr_idx += 1
            result = ''.join(parsed_expression[start_idx:])
            return eval(result)

        pagos = df_pagos.drop(df_pagos.index[df_pagos['Fecha'] != fecha_referencia])
        total_meses_anteriores = total_mes_actual = total_meses_siguientes = 0
        num_pagos = 0
        for index, pago in pagos.iterrows():
            beneficiario = pago['Beneficiario']
            meses_pagados = separa_meses(pago['Concepto'], muestra_modificador=True)
            num_pagos += 1
            r = df_resumen[df_resumen['Beneficiario'] == beneficiario]
            for mes in separa_meses(pago['Concepto'], muestra_modificador=False):
                modificador = ''
                for mes_pago in meses_pagados:
                    if mes_pago.startswith(mes):
                        modificador = mes_pago[len(mes) + 1:]
                try:
                    pago_mes = r[mes].values[0]
                except:
                    print(f"\n*** ERROR: {str(sys.exc_info()[1])} (Beneficiario: {pago['Beneficiario']}, mes: {mes})")
#                    print(f"DEBUG: df_resumen['Beneficiario']=\n{df_resumen['Beneficiario']}, {beneficiario=}")
#                    print(f"DEBUG: r=\n{r}")
                    sys.exit()
                if pago_mes == None:
                    monto = 0
                elif isinstance(pago_mes, str):
                    try:
                        _ = eval(pago_mes[1:])   # Comprueba si es una fórmula válida
                    except:
                        print(f'*** Error evaluando "{pago_mes}", ({beneficiario}, {mes})')
                        monto = 0
                    else:
                        monto = evalua_celda(pago_mes, modificador)
                else:
                    monto = pago_mes
                if ajusta_fecha(mes) < ajusta_fecha(fecha_referencia):
                    total_meses_anteriores += monto
                elif ajusta_fecha(mes) > ajusta_fecha(fecha_referencia):
                    total_meses_siguientes += monto
                else:   # mes == fecha_referencia
                    total_mes_actual += monto

        return (num_pagos, total_meses_anteriores, total_mes_actual, total_meses_siguientes)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Calcula cómo se distribuyen los pagos por mes')

    # Lee la hoja con el resumen de pagos. Usa la libreria OpenPyXL para preservar
    # las fórmulas en las columnas y genera un dataframe de Pandas como resultado

    # Maneja la primera columna ('Beneficiario') como su valor asociado
    df_resumen = read_excel(excel_workbook, sheet_name=excel_ws_resumen, usecols=['Beneficiario'])

    # Y el resto de las columnas como fórmulas
    wb = load_workbook(filename=excel_workbook)
    ws = wb[excel_ws_resumen]
    col_value = None
    for idx, col in enumerate(ws.columns):
        if idx == 0:
            continue
        curr_col = [expr.value for expr in col]
        col_name = curr_col.pop(0)
        try:
            if idx == 2:
                col_name = 'Check'
            elif col_name[0] == '=':
                col_name = col_value + relativedelta(months=1)
        except:
            pass
        col_value = col_name
        df_resumen[col_name] = curr_col

    # Ajusta los nombres de columnas
    new_cols = list()
    for col in df_resumen.columns.values.tolist():
        if isinstance(col, datetime):
            new_cols.append(col.strftime('%m-%Y'))
        else:
            new_cols.append(col)
    df_resumen.columns = new_cols

    # Lee la hoja con el detalle de los pagos recibidos, elimina aquellos cuya Referen-
    # cia no corresponda al pago de Vigilancia, descarta aquellos posteriores a la fecha
    # de referencia y estandariza la fecha de pago
    df_pagos = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)
    #df_pagos = df_pagos[df_pagos['Categoría'] == 'Vigilancia']
    df_pagos.drop(df_pagos.index[df_pagos['Categoría'] != 'Vigilancia'], inplace=True)
    df_pagos = df_pagos[['Beneficiario', 'Dirección', 'Fecha', 'Monto', 'Concepto', 'Mes', 'Nro. Recibo']]
    df_pagos = df_pagos[df_pagos['Fecha'] <= fecha_de_referencia]
    df_pagos.sort_values(by=['Beneficiario', 'Fecha'], inplace=True)
    #df_pagos.dropna(subset=['Fecha'], inplace=True)
    df_pagos['Fecha'] = df_pagos['Fecha'].apply(lambda x: f'{x:%m-%Y}')


    def edit(valor, width=9, decimals=0):
        valor_ed = edita_número(valor, num_decimals=decimals)
        return f"{valor_ed:>{width}}"

    def edit_pct(valor, total, width=5, decimals=1):
        return edit(valor/total*100, width-1, decimals) + '%'


    lista_resultados = list()
    if VERBOSE: print()
    for offset in reversed(range(num_meses_cuotas_equivalentes+1)):
        f_ref = f"{mes_actual - relativedelta(months=offset) + relativedelta(day=1):%m-%Y}"
        resultado = análisis_de_pagos(f_ref)
        total = sum(resultado[1:])
        lista_resultados.append(resultado)

        if VERBOSE:
            print(f"  {f_ref}: {edit(total, width=10)} [" + \
                  f"{edit(resultado[1])} ({edit_pct(resultado[1], total)}), " + \
                  f"{edit(resultado[2], width=10)} ({edit_pct(resultado[2], total)}), " + \
                  f"{edit(resultado[3])} ({edit_pct(resultado[3], total)})"   + "]")

    if VERBOSE: print()

    return lista_resultados


def gráfica_3():
    global num_meses_cuotas_equivalentes
    from GyG_cuotas import Cuota

    VERTICAL_BARS = not g3_horizontal

    grafica_nombre = 'Cuotas equivalentes'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"3. Cuotas_equivalentes {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '3. Cuotas_equivalentes.png'


    if in_GUI_mode:
        num_meses_cuotas_equivalentes = int(values['_g3_nro_meses_'])
        print(f'  - {grafica_nombre}...')

    dia_de_referencia = fecha_de_referencia.day
    mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
    fecha_inicial = fecha_de_referencia - relativedelta(months=num_meses_cuotas_equivalentes)
    mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)

    distribución = distribución_de_pagos()
    pctMesAnt = list()
    pctMesAct = list()
    pctMesSig = list()
    str_pctMesAnt = list()
    str_pctMesAct = list()
    str_pctMesSig = list()
    for nPagos, mesAnt, mesAct, mesSig  in distribución:
        total_mes = mesAnt + mesAct + mesSig
        if total_mes == 0.00:
            pctMesAnt.append(pMesAnt := 0.00)
            pctMesAct.append(pMesAct := 0.00)
            pctMesSig.append(pMesSig := 0.00)
        else:
            pctMesAnt.append(pMesAnt := mesAnt / total_mes)
            pctMesAct.append(pMesAct := mesAct / total_mes)
            pctMesSig.append(pMesSig := mesSig / total_mes)
        str_pctMesAnt.append(edita_número(pMesAnt * 100, num_decimals=1) + '%')
        str_pctMesAct.append(edita_número(pMesAct * 100, num_decimals=1) + '%')
        str_pctMesSig.append(edita_número(pMesSig * 100, num_decimals=1) + '%')

    cuotas_obj = Cuota(excel_workbook)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los pagos por mes del período seleccionado')
    ws_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

    # Selecciona los pagos de vigilancia entre la fecha de referencia y num_meses_cuotas_equivalentes atrás
    ws_vigilancia = ws_vigilancia[(ws_vigilancia['Categoría'] == 'Vigilancia')  & \
                                  (ws_vigilancia['Fecha'] <= fecha_de_referencia) & \
                                  (ws_vigilancia['Fecha'] >= mes_inicial)]
    ws_vigilancia = ws_vigilancia[['Mes', 'Beneficiario', 'Fecha', 'Monto']]

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Agrega la cuota vigente para la fecha de pago')
    ws_vigilancia['Cuota'] = ws_vigilancia.apply(
                                    lambda r: cuotas_obj.cuota_vigente(r['Beneficiario'], r['Fecha']), axis=1)
    ws_vigilancia['Pagos eqv'] = ws_vigilancia.apply(
                                    lambda r: float(r['Monto']) / float(r['Cuota']), axis=1)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Totalizando Pagos Equivalentes por Mes...')
    ws_vigilancia = pivot_table(ws_vigilancia, values=['Pagos eqv'], index=['Mes'], aggfunc=sum).reset_index()

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los rangos de la imagen')
    meses = ws_vigilancia['Mes'].to_list()
    pagos_100pct_equivalentes = ws_vigilancia['Pagos eqv'].to_list()
    if VERTICAL_BARS:
        meses_ed = [mes.strftime(FORMATO_MES) for mes in meses]
        promedio = int(round(mean(pagos_100pct_equivalentes[:-1]), ndigits=0))
    else:
        meses_ed = [mes.strftime(FORMATO_MES) for mes in reversed(meses)]
        pagos_100pct_equivalentes.reverse()
        pctMesAnt.reverse()
        pctMesAct.reverse()
        pctMesSig.reverse()
        promedio = int(round(mean(pagos_100pct_equivalentes[1:]), ndigits=0))
    bar_mesAnt = [num_pagos * pctMesAnt[idx] for idx, num_pagos in enumerate(pagos_100pct_equivalentes)]
    bar_mesAct = [num_pagos * pctMesAct[idx] for idx, num_pagos in enumerate(pagos_100pct_equivalentes)]
    bar_mesSig = [num_pagos * pctMesSig[idx] for idx, num_pagos in enumerate(pagos_100pct_equivalentes)]
    sep_1, sep_2 = ('<br />', '') if VERTICAL_BARS else (' pagos (', ')')
    iterator = enumerate(pagos_100pct_equivalentes) if VERTICAL_BARS \
                    else reversed(list(enumerate(pagos_100pct_equivalentes)))
    str_pctMesAct = [f"{int(round(numPagos * pctMesAct[idx], 0))}{sep_1}" + \
                     f"{edita_número(pctMesAct[idx] * 100, num_decimals=1)}%{sep_2}" \
                            for idx, numPagos in iterator]
    promedios = [promedio for _ in meses]
    pto_equilibrio = [PUNTO_DE_EQUILIBRIO for _ in meses]

    título =    'Pagos recibidos en el mes equivalentes a cuotas completas'
    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b %Y")


    fig = go.Figure()

    # Distribución de cuotas
    fig.add_trace(
                    go.Bar(
                        name = 'Meses anteriores',
                        x = meses_ed if VERTICAL_BARS else bar_mesAnt,
                        y = bar_mesAnt if VERTICAL_BARS else meses_ed,
                        orientation = 'v' if VERTICAL_BARS else 'h',
                        text = str_pctMesAnt if VERTICAL_BARS else [text for text in reversed(str_pctMesAnt)],
                        textposition = "inside",
                        marker_color = '#4475cd',            # '#5485dd',
                        legendgroup = 'bar',
                    ))
    fig.add_trace(
                    go.Bar(
                        name = 'Mes actual',
                        x = meses_ed if VERTICAL_BARS else bar_mesAct,
                        y = bar_mesAct if VERTICAL_BARS else meses_ed,
                        orientation = 'v' if VERTICAL_BARS else 'h',
                        text = str_pctMesAct if VERTICAL_BARS else [text for text in reversed(str_pctMesAct)],
                        textposition = "inside",
                        marker_color = 'cornflowerblue',     # '#6495ed',
                        legendgroup = 'bar',
                    ))
    fig.add_trace(
                    go.Bar(
                        name = 'Meses subsiguientes',
                        x = meses_ed if VERTICAL_BARS else bar_mesSig,
                        y = bar_mesSig if VERTICAL_BARS else meses_ed,
                        orientation = 'v' if VERTICAL_BARS else 'h',
                        text = str_pctMesSig if VERTICAL_BARS else [text for text in reversed(str_pctMesSig)],
                        textposition = "inside",
                        marker_color = '#84b5ff',           # '#74a5fd',
                        legendgroup = 'bar',
                    ))
    fig.add_trace(
                    go.Bar(
                        name = '',
                        x = meses_ed if VERTICAL_BARS else [0 for _ in meses],
                        y = [0 for _ in meses] if VERTICAL_BARS else meses_ed,
                        orientation = 'v' if VERTICAL_BARS else 'h',
                        text = [f"{y:.0f}" for y in pagos_100pct_equivalentes],
                        textposition = "outside",
                        marker_color = 'white',             # '#5b9bd5',
                        showlegend=False,
                    ))
    fig.update_layout(
                    title = f'<b>{título}</b><br>{subtítulo}',
                    width = 1000,
                    height = 600,
                    legend = dict(
                                orientation = 'h',
                                tracegroupgap = 20,
                            ),
                    xaxis_tickangle = -30 if VERTICAL_BARS else 0,
    #                xaxis_title = 'Cantidad de cuotas completas',
    #                paper_bgcolor = 'white',
                    plot_bgcolor = 'white',
                    barmode='stack',
                )

    # # Punto de equilibrio
    fig.add_trace(
                    go.Scatter(
                        x = meses_ed if VERTICAL_BARS else pto_equilibrio,
                        y = pto_equilibrio if VERTICAL_BARS else meses_ed,
                        mode = 'lines',
                        name = f"Punto de equilibrio = {pto_equilibrio[0]} cuotas completas",
                        line = dict(
                                    color = 'red',          # 'yellow',           # '#ffc000',
                                    dash = 'dash',
                                    width = 2,
                                ),
                       legendgroup = 'lines',
                    ))

    # Promedio últimos 12 meses
    fig.add_trace(
                    go.Scatter(
                        x = meses_ed if VERTICAL_BARS else promedios,
                        y = promedios if VERTICAL_BARS else meses_ed,
                        mode = 'lines',
                        name = f"Promedio últimos {num_meses_cuotas_equivalentes} meses = {promedio}; " + \
                               f"actual = {int(round(pagos_100pct_equivalentes[-1 if VERTICAL_BARS else 0], 0))}",
                        line = dict(
                                    color = 'blue',         # '#ed7d31',          # 'orange',
                                    width = 2,
                                ),
                       legendgroup = 'lines',
                    ))

    if g3_show:
        print(f"    . Desplegando imagen {'vertical' if VERTICAL_BARS else 'horizontal'}...")
        fig.show()

    if g3_save:
        print(f'    . Grabando "{grafica_png}"...')
        img_bytes = fig.to_image(format="png")
        with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
            f.write(img_bytes)


def gráfica_4():
    global num_meses_gestion_cobranzas

    grafica_nombre = 'Cuotas Recibidas en el Mes'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"4. Cuotas_recibidas {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '4. Cuotas_recibidas.png'


    if in_GUI_mode:
        num_meses_gestion_cobranzas = int(values['_g1_nro_meses_'])
        print(f'  - {grafica_nombre}...')

    # Define las fechas de referencia
    dia_de_referencia = fecha_de_referencia.day
    mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
    fecha_inicial = fecha_de_referencia - relativedelta(months=num_meses_gestion_cobranzas)
    mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)
    # --------------------
    mes_final = mes_actual - relativedelta(months=1)
    str_mes_inicial = mes_inicial.strftime('%b'+('/%Y' if mes_inicial.year != mes_final.year else ''))
    str_mes_final = mes_final.strftime('%b/%Y')
    # --------------------

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee las cuotas mensuales')
    ws_cuotas_mensuales = read_excel(excel_workbook, sheet_name=excel_ws_resumen)
    ws_cuotas_mensuales = ws_cuotas_mensuales[(ws_cuotas_mensuales['Beneficiario'] == 'CUOTAS MENSUALES')]


    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los pagos realizados')
    ws_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

    # Selecciona los pagos de vigilancia entre la fecha de referencia y num_meses_gestion_cobranzas atrás
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Filtra los pagos realizados por Categoría y Fecha')
    ws_vigilancia = ws_vigilancia[(ws_vigilancia['Categoría'] == 'Vigilancia') & \
                                  (ws_vigilancia['Fecha'] <= fecha_de_referencia) & \
                                  (ws_vigilancia['Fecha'] >= mes_inicial)]
    ws_vigilancia = ws_vigilancia[['Mes', 'Día', 'Monto', 'Fecha']]

    # Valida que los datos estén completos y agrega registros en cero para los datos faltantes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Valida la completitud de valores en el período ' + \
                                                                     'y rellena con ceros')
    fecha = mes_inicial
    a_agregar = []
    while fecha <= fecha_de_referencia:
        if ws_vigilancia[ws_vigilancia['Fecha'] == fecha].empty:
            a_agregar.append({'Mes':datetime(fecha.year, fecha.month, 1), 'Día':fecha.day, 'Monto':0.0, 'Fecha':fecha})
        fecha = fecha + relativedelta(days=1)

    ws_vigilancia = ws_vigilancia.append(a_agregar, ignore_index=True)

    # Totaliza los pagos de vigilancia por día y mes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Totaliza los pagos de vigilancia por día y mes')
    ws_vigilancia = pivot_table(ws_vigilancia, values='Monto', index=['Día', 'Mes'], aggfunc=sum)   # , fill_value=0
    ws_vigilancia = ws_vigilancia.reset_index() 

    # Normalizar los montos dividiéndolo por el estimado del mes
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Normaliza los pagos de vigilancia en base a la cuota ' + \
                                                                     'del mes')
    ws_vigilancia['Monto_normalizado'] = ws_vigilancia.apply(
                                            lambda ws: ws['Monto'] / ws_cuotas_mensuales[ws['Mes']], axis=1)

    # Genera columnas a graficar
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los datos del mes actual')
    ws_mes_actual = ws_vigilancia[ws_vigilancia['Mes'] == mes_actual].copy().reset_index()
    ws_mes_actual['Acumulado'] = ws_mes_actual['Monto_normalizado'].cumsum()
    total_recaudado = ws_mes_actual['Monto'].sum()
    pct_del_estimado = total_recaudado / ws_cuotas_mensuales[mes_actual] * 100

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los datos del promedio de los meses anteriores')
    ws_meses_anteriores = ws_vigilancia[ws_vigilancia['Mes'] < mes_actual]
    ws_meses_anteriores = pivot_table(ws_meses_anteriores, values='Monto_normalizado', index=['Día'], fill_value=0)
    ws_meses_anteriores = ws_meses_anteriores.reset_index()
    ws_meses_anteriores['Acumulado'] = ws_meses_anteriores['Monto_normalizado'].cumsum()


    # Mes actual
    trace_mes_actual = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = ws_mes_actual['Acumulado'],
                mode = 'lines',
                name = fecha_de_referencia.strftime(FORMATO_MES),
                line = dict(
                            width = 3,
                            color = 'orange',
                        )
            )

    # Regresión lineal
    x = ws_meses_anteriores['Día']
    y = ws_mes_actual['Acumulado']
    slope, intercept, r_value, p_value, std_err = stats.linregress(x[:dia_de_referencia], y[:dia_de_referencia])
    linea_lr = [slope * xi + intercept for xi in x]

    trace_regresion = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = linea_lr,
                mode = 'lines',
                name = 'Ajuste lineal ' + fecha_de_referencia.strftime(FORMATO_MES),
                line = dict(
                            width = 2,
                            color = 'orange',
                            dash = 'dot',
                        )
            )

    # Promedio de los últimos 5 meses
    trace_prom_5_ultimos_meses = go.Scatter(
                x = ws_meses_anteriores['Día'],
                y = ws_meses_anteriores['Acumulado'],
                mode = 'lines',
                # name = f"Promedio {num_meses_gestion_cobranzas} últimos meses",
                name = f"Promedio de {str_mes_inicial} a {str_mes_final}",
                line = dict(
                            width = 3,
                            color = 'blue',
                        )
            )

    data = [trace_mes_actual, trace_prom_5_ultimos_meses, trace_regresion]

    título = 'Cuotas Recibidas en el Mes'
    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b %Y")
    día = fecha_de_referencia.day

    layout = go.Layout(
                title = dict(
                            text = f'<b>{título}</b><br>{subtítulo}',
                        ),
                width = 1000,
                height = 600,
                xaxis = dict(
                            nticks = 31,
                            showgrid = True, gridwidth=1, gridcolor='lightgray',
                            showline=True, linewidth=1, linecolor='black',
                        ),
                yaxis = dict(
                            title = 'Cantidad de cuotas recibidas en el mes',
                            tickformat = ',.0',
                            showline=True, linewidth=1, linecolor='black',
                        ),
                legend = dict(
                            orientation = 'h',
                            tracegroupgap = 20,
                        ),
                # annotations = [
                #         dict(
                #             text = f"{int(round(ws_mes_actual['Acumulado'][día - 1]))} cuotas recibidas en el mes; " + \
                #                    f"{int(round(ws_meses_anteriores['Acumulado'][día - 1]))} cuotas promedio a la fecha",
                #             font = dict(
                #                     size = 14,
                #                 ),
                #             showarrow = False,
                #             xref = 'paper',
                #             yref = 'paper',
                #             x = 0.98,
                #             y = 0.05,
                #             align = 'right',
                #             valign = 'bottom',
                #         ),
                #     ],
#                paper_bgcolor = 'white',
                plot_bgcolor = 'white',
            )

    fig = go.Figure(data = data, layout = layout)

    ultimo_indice = ws_mes_actual['Acumulado'].index[-1]
    dx, dy = 20, 40
    actual_ge_anteriores = ws_mes_actual['Acumulado'][ultimo_indice] >= ws_meses_anteriores['Acumulado'][ultimo_indice]

    delta_x = -dx if actual_ge_anteriores else dx
    if ultimo_indice == 0:
        delta_x = abs(delta_x)
    elif ultimo_indice == 30:
        delta_x = -abs(delta_x)
    fig.add_annotation(
        x = ws_meses_anteriores['Día'][ultimo_indice],
        y = ws_mes_actual['Acumulado'][ultimo_indice],
        text = f"{int(round(ws_mes_actual['Acumulado'][día - 1]))} cuotas",
        ax = delta_x,
        ay = -dy if actual_ge_anteriores else dy
    )
    delta_x = -dx if not actual_ge_anteriores else dx
    if ultimo_indice == 0:
        delta_x = abs(delta_x)
    elif ultimo_indice == 30:
        delta_x = -abs(delta_x)
    fig.add_annotation(
        x = ws_meses_anteriores['Día'][ultimo_indice],
        y = ws_meses_anteriores['Acumulado'][ultimo_indice],
        text = f"{int(round(ws_meses_anteriores['Acumulado'][día - 1]))} cuotas",
        ax = delta_x,
        ay = -dy if not actual_ge_anteriores else dy
    )
    fig.update_annotations(dict(
        xref = "x",
        yref = "y",
        showarrow = True,
        arrowhead = 3,
    ))

    if g4_show:
        print(f'    . Desplegando imagen...')
        fig.show()

    if g4_save:
        print(f'    . Grabando "{grafica_png}"...')
        img_bytes = fig.to_image(format="png")
        with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
            f.write(img_bytes)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Gráfica concluida')


#
# PROCESO
#

# print(f'Cargando hoja de cálculo "{excel_workbook}"...')
# ws = read_excel(excel_workbook, sheet_name=excel_ws_cobranzas, header=None)

fecha_de_referencia = datetime.today()
mes_anterior = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1) - relativedelta(days=1)

if toma_opciones_por_defecto:

    in_GUI_mode = False

    ahora = datetime.now()
    fecha_de_referencia = datetime(ahora.year, ahora.month, ahora.day)
    mes_actual = datetime(ahora.year, ahora.month, 1)
    mes_anterior = mes_actual - relativedelta(days=1)

    print()

    # Selecciona si se muestran las gráficas generadas
    se_muestran_las_gráficas = input_si_no('Se muestran las gráficas generadas', 'no', toma_opciones_por_defecto)

    # Selecciona si se graban las imágenes generadas
    se_graban_las_imágenes = input_si_no('Se graban las imágenes generadas', 'sí', toma_opciones_por_defecto)

    # Selecciona la forma de desplegar las gráficas de barras
    graficas_de_barra_en_horizontal = input_si_no('Se muestran las barras en horizontal', 'no', toma_opciones_por_defecto)

    print()

    g1 = True
    g1_show = se_muestran_las_gráficas
    g1_save = se_graban_las_imágenes

    g2 = True
    g2_show = se_muestran_las_gráficas
    g2_save = se_graban_las_imágenes
    g2_horizontal = graficas_de_barra_en_horizontal

    g3 = True
    g3_show = se_muestran_las_gráficas
    g3_save = se_graban_las_imágenes
    g3_horizontal = graficas_de_barra_en_horizontal

    g4 = True
    g4_show = se_muestran_las_gráficas
    g4_save = se_graban_las_imágenes

    fecha_de_referencia = mes_anterior

    genera_gráficas()

else:

    in_GUI_mode = True

    print("Generando layout ...")
#    sg.theme('DarkAmber')   # Add a little color to your windows

    sg.SetOptions(
            icon=os.path.join(GyG_constantes.rec_imágenes, 'GyG_logo.png'),
        )

    fechas_layout = [   
                        [sg.Radio("", size=(2, 1), default=True,
                                  group_id='fechas', key='_fecha_manual_'),
                         sg.Text("Gráficas al:"),
                         sg.InputText(key='_fecha_referencia_', size=(10,1),
                                      default_text=fecha_de_referencia.strftime("%d-%m-%Y")),
                         sg.CalendarButton(button_text='Seleccione otra fecha',
                                           image_filename=CALENDAR_ICON, image_size=(16, 16), image_subsample=8,
                                           default_date_m_d_y=(fecha_de_referencia.month, fecha_de_referencia.day, fecha_de_referencia.year),
                                           target='_fecha_referencia_', format='%d-%m-%Y',
                                           close_when_date_chosen=True)],
                        [sg.Radio("", size=(2, 1), 
                                  group_id='fechas', key='_fecha_mes_anterior_'),
                         sg.Text(f"Gráficas del mes anterior:  {mes_anterior.strftime('%b %Y')}", size=(30, 1)),
                         sg.Text("", size=(36, 1))],    # espaciador
                    ]
    ANCHURA_NOMBRE_GRAFICO = 28
    graficas_layout = [
                        [sg.Text("", size=(ANCHURA_NOMBRE_GRAFICO - 3, 1)), sg.Text("Visualiza"), sg.Text("Graba"), sg.Text("Horiz.")],
                        [sg.Checkbox("Gestión de Cobranzas", key="_g1_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g1_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g1_save_", size=(5, 1)),
                         sg.Text("", size=(5, 1)),      # espaciador (gráfica siempre vertical)
                         sg.Text("Nº meses:"), sg.Spin(key="_g1_nro_meses_", values=(str(i) for i in range(2, 13)),
                                                       initial_value='5', size=(3, 1)),
                         sg.Text("", size=(1, 1))],
                        [sg.Checkbox("Pagos 100% equivalentes", key="_g2_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g2_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g2_save_", size=(5, 1)),
                         sg.Checkbox("", key="_g2_horizontal_", size=(5, 1)),
                         sg.Text("Nº meses:"), sg.Spin(key="_g2_nro_meses_", values=(str(i) for i in range(3, 25)),
                                                       initial_value='12', size=(3, 1))],
                        [sg.Checkbox("Cuotas equivalentes", key="_g3_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g3_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g3_save_", size=(5, 1)),
                         sg.Checkbox("", key="_g3_horizontal_", size=(5, 1)),
                         sg.Text("Nº meses:"), sg.Spin(key="_g3_nro_meses_", values=(str(i) for i in range(3, 25)),
                                                       initial_value='12', size=(3, 1))],
                        [sg.Checkbox("Pagos Recibidos", key="_g4_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g4_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g4_save_", size=(5, 1)),
                         sg.Text("", size=(5, 1)),      # espaciador (gráfica siempre vertical)
                         sg.Text("Nº meses:"), sg.Spin(key="_g4_nro_meses_", values=(str(i) for i in range(2, 13)),
                                                       initial_value='5', size=(3, 1)),
                         sg.Text("", size=(1, 1))],
                    ]
    layout = [  [sg.Frame(" Fecha ",    fechas_layout,   size=(60, None))],
                [sg.Frame(" Gráficas ", graficas_layout, size=(60, None))],
                [sg.Button('Genera gráficas'), sg.Button('Finaliza')],
             ]

    # Create the Window
    window = sg.Window('GyG Gráficas', layout)

    # Event Loop to process "events"
    while True:             
        event, values = window.read()

        # from pprint import pprint
        # pprint(values)
        
        if event in (None, 'Finaliza'):
            break
        elif event == 'Genera gráficas':
            if values['_fecha_manual_']:
                fecha_de_referencia = datetime.strptime(values['_fecha_referencia_'], '%d-%m-%Y')
            else:
                fecha_de_referencia = mes_anterior
            if fecha_de_referencia > datetime.today():
                fecha_de_referencia = datetime.today()

            mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)

            g1 = values['_g1_']
            g1_show = values['_g1_show_']
            g1_save = values['_g1_save_']

            g2 = values['_g2_']
            g2_show = values['_g2_show_']
            g2_save = values['_g2_save_']
            g2_horizontal = values['_g2_horizontal_']

            g3 = values['_g3_']
            g3_show = values['_g3_show_']
            g3_save = values['_g3_save_']
            g3_horizontal = values['_g3_horizontal_']

            g4 = values['_g4_']
            g4_show = values['_g4_show_']
            g4_save = values['_g4_save_']

            genera_gráficas()

    window.close()

# print()
# print('Proceso terminado . . .')
