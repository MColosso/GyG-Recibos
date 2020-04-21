# GyG GRAFICAS
#
# Genera las gráficas de Gestión de Cobranzas y Cuotas 100% Equivalentes
# en base a la última información cargada en la hoja de cálculo

"""
    POR HACER
    -   

    HISTORICO
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

import sys
import os
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

# import warnings
# warnings.simplefilter("ignore", category=RuntimeWarning)


# Define algunas constantes
excel_workbook      = GyG_constantes.pagos_wb_estandar     # '1.1. GyG Recibos.xlsm'
excel_ws_cobranzas  = GyG_constantes.pagos_ws_cobranzas
excel_ws_vigilancia = GyG_constantes.pagos_ws_vigilancia
excel_ws_resumen    = GyG_constantes.pagos_ws_resumen

num_meses_gestion_cobranzas   =  5      # Nro de meses a promediar en gráfica 'Gestión de Cobranzas'
num_meses_cuotas_equivalentes = 12      # Nro de meses a mostrar en gráfica 'Cuotas Equivalentes'

PUNTO_DE_EQUILIBRIO           = 55      # Cantidad de familias usadas en el cálculo de la cuota

CALENDAR_ICON       = os.path.join(GyG_constantes.rec_imágenes, '62925-spiral-calendar-icon.png')
CALENDAR_SIZE       = (16, 16)
CALENDAR_SUBSAMPLE  = 8

VERBOSE             = False                         # Muestra mensajes adicionales

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
        grafica_1()     # Gestión de Cobranzas
    if g2:
        gráfica_2()     # Cuotas Equivalentes


def grafica_1():

    grafica_nombre = 'Gestión de Cobranzas'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"1. Gestion_de_cobranzas {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '1. Gestion_de_cobranzas.png'


    if in_GUI_mode:
        print(f'  - {grafica_nombre}...')

    # Obtiene los datos de las pestañas 'Vigilancia' y 'RESUMEN VIGILANCIA' de la hoja de cálculo '1.1. GyG Recibos.xlsm'
    dia_de_referencia = fecha_de_referencia.day
    mes_actual = datetime(fecha_de_referencia.year, fecha_de_referencia.month, 1)
    fecha_inicial = fecha_de_referencia - relativedelta(months=num_meses_gestion_cobranzas)
    mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)

    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Lee los estimados del servicio de vigilancia')
    ws_estimados_mensuales = read_excel(excel_workbook, sheet_name=excel_ws_resumen)
    ws_estimados_mensuales = ws_estimados_mensuales[(ws_estimados_mensuales['Beneficiario'] == \
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
    ws_vigilancia['Monto_normalizado'] = ws_vigilancia.apply(lambda ws: ws['Monto'] / ws_estimados_mensuales[ws['Mes']], axis=1)

    # Genera columnas a graficar
    if VERBOSE: print(f'    . [{datetime.now().strftime("%H:%M:%S")}] Genera los datos del mes actual')
    ws_mes_actual = ws_vigilancia[ws_vigilancia['Mes'] == mes_actual].copy()
    ws_mes_actual['Acumulado'] = ws_mes_actual['Monto_normalizado'].cumsum()
    total_recaudado = ws_mes_actual['Monto'].sum()
    pct_del_estimado = total_recaudado / ws_estimados_mensuales[mes_actual] * 100

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
                name = fecha_de_referencia.strftime("%b.%Y"),
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
                name = 'Ajuste lineal ' + fecha_de_referencia.strftime("%b.%Y"),
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
                name = f"Promedio {num_meses_gestion_cobranzas} últimos meses",
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

    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b. %Y")
    comentario = f"Recaudado: Bs. {edita_número(total_recaudado, num_decimals=0)} " + \
                 f"({edita_número(pct_del_estimado, num_decimals=0)}% del estimado)"

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

    grafica_nombre = 'Pagos 100% equivalentes'
    if es_fin_de_mes(fecha_de_referencia):
        grafica_png = f"2. Pagos_100pct_equivalentes {fecha_de_referencia.strftime('%Y-%m (%b)')}.png"
    else:
        grafica_png =  '2. Pagos_100pct_equivalentes.png'


    if in_GUI_mode:
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
    meses_ed = [mes.strftime("%b.%Y") for mes in meses]
    pagos_100pct_equivalentes = [int(round(float(ws_totales_mensuales[mes]) / float(ws_cuotas_mensuales[mes]), ndigits=0)) \
                                    for mes in meses]
    promedio = int(round(mean(pagos_100pct_equivalentes[:-1]), ndigits=0))
    promedios = [promedio for _ in meses]
    pto_equilibrio = [PUNTO_DE_EQUILIBRIO for _ in meses]

    subtítulo = 'al ' + fecha_de_referencia.strftime("%d %b. %Y")


    # 100% Equivalentes
    trace_100pct_equiv = go.Bar(
                x = meses_ed,
                y = pagos_100pct_equivalentes,
                name = '100% eqv.',
                text = [f"{y:.0f}" for y in pagos_100pct_equivalentes],
                textposition = "outside",
                marker_color = 'cornflowerblue',        # '#5b9bd5',
                # marker = dict(
                #             color = 'mediumblue',
                #         ),
            )

    # Promedio últimos 12 meses
    trace_prom_12_ultimos_meses = go.Scatter(
                x = meses_ed,
                y = promedios,
                mode = 'lines',
                name = f"Promedio últimos {num_meses_cuotas_equivalentes} meses = {promedio}; " + \
                       f"actual = {pagos_100pct_equivalentes[-1]}",
                line = dict(
                            color = '#ed7d31',          # 'orange',
                            width = 2,
                        ),
            )

    # Punto de equilibrio
    trace_punto_equilibrio = go.Scatter(
                x = meses_ed,
                y = pto_equilibrio,
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
                xaxis_tickangle = -30,
#                paper_bgcolor = 'white',
                plot_bgcolor = 'white',
            )

#    subtítulo = 'al ' + celda('A36').strftime("%d %b. %Y")
#    comentario = celda('B18')

    fig = go.Figure(data = data, layout = layout)

    if g2_show:
        print(f'    . Desplegando imagen...')
        fig.show()

    if g2_save:
        print(f'    . Grabando "{grafica_png}"...')
        img_bytes = fig.to_image(format="png")
        with open(os.path.join(GyG_constantes.ruta_graficas, grafica_png), 'wb') as f:
            f.write(img_bytes)


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

    print()

    # Selecciona si se muestran las gráficas generadas
    se_muestra_la_gráfica = input_si_no('Se muestran las gráficas generadas', 'no', toma_opciones_por_defecto)

    # Selecciona si se graban las imágenes generadas
    se_graba_la_imagen = input_si_no('Se graban las imágenes generadas', 'sí', toma_opciones_por_defecto)

    g1 = True
    g1_show = se_muestra_la_gráfica
    g1_save = se_graba_la_imagen

    g2 = True
    g2_show = se_muestra_la_gráfica
    g2_save = se_graba_la_imagen

    print()
    genera_gráficas()

else:

    in_GUI_mode = True

    print("Generando layout ...")
#    sg.theme('DarkAmber')   # Add a little color to your windows

    sg.SetOptions(
            icon='GyG Logo/GyG Logo.ico',
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
                         sg.Text(f"Gráficas del mes anterior:  {mes_anterior.strftime('%b. %Y')}", size=(30, 1)),
                         sg.Text("", size=(29,1))],
                    ]
    ANCHURA_NOMBRE_GRAFICO = 28
    graficas_layout = [
                        [sg.Text("", size=(ANCHURA_NOMBRE_GRAFICO - 3, 1)), sg.Text("Visualiza"), sg.Text("Graba")],
                        [sg.Checkbox("Gestión de Cobranzas", key="_g1_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g1_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g1_save_", size=(5, 1)),
                         sg.Text("Nº meses:"), sg.Spin(key="_g1_nro_meses_", values=(str(i) for i in range(2, 13)), initial_value='5', size=(3, 1)),
                         sg.Text("", size=(1, 1))],
                        [sg.Checkbox("Pagos 100% equivalentes", key="_g2_", size=(ANCHURA_NOMBRE_GRAFICO, 1), default=True),
                         sg.Checkbox("", key="_g2_show_", size=(7, 1), default=True),
                         sg.Checkbox("", key="_g2_save_", size=(5, 1)),
                         sg.Text("Nº meses:"), sg.Spin(key="_g2_nro_meses_", values=(str(i) for i in range(3, 25)), initial_value='12', size=(3, 1))],
                    ]
    layout = [  [sg.Frame(" Fecha ",    fechas_layout,   size=(60, None))],
                [sg.Frame(" Gráficas ", graficas_layout, size=(60, None))],
                [sg.Button('Genera gráficas'), sg.Button('Finaliza')],
             ]

    # Create the Window
    window = sg.Window('GyG Gráficas', layout)

    # Event Loop to process "events"
#    while True:             
    event, values = window.read()

    # from pprint import pprint
    # pprint(values)
    
    if event in (None, 'Finaliza'):
        pass    # break
    elif event == 'Genera gráficas':
        if values['_fecha_manual_']:
            fecha_de_referencia = datetime.strptime(values['_fecha_referencia_'], '%d-%m-%Y')
        else:
            fecha_de_referencia = mes_anterior
        num_meses_gestion_cobranzas   = int(values['_g1_nro_meses_'])
        num_meses_cuotas_equivalentes = int(values['_g2_nro_meses_'])

        g1 = values['_g1_']
        g1_show = values['_g1_show_']
        g1_save = values['_g1_save_']

        g2 = values['_g2_']
        g2_show = values['_g2_show_']
        g2_save = values['_g2_save_']

        genera_gráficas()

    window.close()

# print()
# print('Proceso terminado . . .')
