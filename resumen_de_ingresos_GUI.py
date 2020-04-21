# RESUMEN DE INGRESOS
#
# Genera el resumen de ingresos en base al contenido de la pestaña 'Vigilancia' de la
# hoja de cálculo '1.1. GyG Recibos' y los aportes individuales de la pestaña 'Saldos'

"""
    POR HACER
    -   

    HISTORICO
    -   Generar la relación de ingresos en formato PDF en lugar de texto (16/04/2020)
    -   Versión inicial (14/04/2020)

"""

print('Cargando librerías...')

import GyG_constantes
from GyG_utilitarios import *
import PySimpleGUI as sg

from pandas import DataFrame, read_excel, pivot_table
from numpy import isnan
from datetime import datetime
from dateutil.relativedelta import relativedelta
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

import os, sys


# Define textos
nombre_análisis     = GyG_constantes.txt_relacion_ingresos      # "GyG Relación de Ingresos {:%Y-%m (%B)}.txt"
attach_path         = GyG_constantes.ruta_relacion_ingresos     # "./GyG Recibos/Análisis de Pago"

excel_workbook      = GyG_constantes.pagos_wb_estandar          # '1.1. GyG Recibos.xlsm'
excel_ws_cobranzas  = GyG_constantes.pagos_ws_cobranzas
excel_ws_vigilancia = GyG_constantes.pagos_ws_vigilancia
excel_ws_saldos     = GyG_constantes.pagos_ws_saldos

CALENDAR_ICON       = os.path.join(GyG_constantes.rec_imágenes, '62925-spiral-calendar-icon.png')
CALENDAR_SIZE       = (16, 16)
CALENDAR_SUBSAMPLE  = 8

NUM_MESES           =  12

# VERBOSE = False                         # Muestra mensajes adicionales

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'
    if not toma_opciones_por_defecto:
        print(f"Uso: {sys.argv[0]} [--toma_opciones_por_defecto]")
        sys.exit(1)


#
# RUTINAS
#

def despliega_resumen_txt():
    nChars_números = 14      # '000.000.000,00'
    nChars_meses   =  9      # 'Xxx. 0000'

    def wrap_encabezado(texto):
        lista_final = list()
        if len(texto) > nChars_números:
            lista_palabras = texto.split(' ')
            desde = 0
            while desde < len(lista_palabras):
                hasta = len(lista_palabras)
                while len(' '.join(lista_palabras[desde: hasta])) > nChars_números:
                    hasta -= 1
                lista_final.append(' '.join(lista_palabras[desde: hasta]))
                desde = hasta
        else:
            lista_final.append(texto)

        return lista_final

    textos_encabezados = [wrap_encabezado(col) for col in ws_relacion_ingresos.columns]
    max_rows = max([len(x) for x in textos_encabezados])
    encabezados = list()
    for encabezado in textos_encabezados:
        while len(encabezado) < max_rows:
            encabezado.insert(0, '')
        encabezados.append(encabezado)

    resumen = 'GyG - TOTAL DE INGRESOS MENSUALES\n'
    resumen += f"{mes_inicial.strftime('%B').title()} {mes_inicial.strftime('%Y')+' ' if mes_inicial.year != mes_final.year else ''}" + \
               f"a {mes_final.strftime('%B %Y').title()}\n\n\n"
    for row in range(max_rows):
        linea = list()
        linea.append(f" {encabezados[0][row]:^{nChars_meses}} ")
        for encabezado in encabezados[1:]:
            linea.append(f" {encabezado[row]:^{nChars_números}} ")
        resumen += '|'.join(linea) + '\n'
    linea = list()
    linea.append('-' * (nChars_meses + 2))
    for col in encabezados[1:]:
        linea.append('-' * (nChars_números + 2))
    resumen += '|'.join(linea) + '\n'

    num_rows, num_columns = ws_relacion_ingresos.shape
    columns = ws_relacion_ingresos.columns
    for row in range(num_rows):
        linea = list()
        linea.append(f" {ws_relacion_ingresos.iloc[row]['Mes'].strftime('%b. %Y').title()} ")
        for col in range(1, num_columns):
            monto = ws_relacion_ingresos.iloc[row][columns[col]]
            str_monto = '' if isnan(monto) else edita_número(monto, num_decimals=2)
            linea.append(f" {str_monto:>{nChars_números}} ")
        resumen += '|'.join(linea) + '\n'

    resumen += f"\n\n(resumen generado el {ahora.strftime('%d/%m/%Y %I:%M')}{am_pm})\n"

    # Graba los archivos de resumen (encoding para Windows y para macOS X)
    filename = nombre_análisis.format(datetime(mes_final.year, mes_final.month, 1)) + '.txt'
    print(f'Generando "{filename}"...')
    resumen_path = os.path.join(attach_path, 'Apple')
    if not os.path.exists(resumen_path):
        os.mkdir(resumen_path)
    with open(os.path.join(resumen_path, filename), 'w', encoding=GyG_constantes.Apple_encoding) as output:
        output.write(resumen)

    resumen_path = os.path.join(attach_path, 'Windows')
    if not os.path.exists(resumen_path):
        os.mkdir(resumen_path)
    with open(os.path.join(resumen_path, filename), 'w', encoding=GyG_constantes.Windows_encoding) as output:
        output.write(resumen)

    # Despliega el resumen en la cónsola
    # print(resumen)


def despliega_resumen_pdf():
    from reportlab.lib import colors, pdfencrypt
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

    filename = nombre_análisis.format(datetime(mes_final.year, mes_final.month, 1)) + '.pdf'
    outputfile = os.path.join(attach_path, filename)

    # Security of .pdf files
    allow_print = pdfencrypt.StandardEncryption("",             # sin contraseña
                                                canPrint=1, canModify=0, canCopy=0, canAnnotate=0,
                                                strength=128)
    doc = SimpleDocTemplate(outputfile, pagesize=(11*72, 8.5*72),
                                        leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=18,
                                        encrypt=allow_print)

    FONT_NAME = 'Times-Roman'
    FONT_NAME_HELVETICA = 'Helvetica'
    styleSheet = getSampleStyleSheet()
    normalStyle = styleSheet['Normal']
    normalStyle.fontName = FONT_NAME
    normalStyle.alignment = 4                       # TA_JUSTIFY
    bodytextStyle = styleSheet['BodyText']
    bodytextStyle.fontName = FONT_NAME
    bodytextStyle.alignment = 4                     # TA_JUSTIFY


    # Genera la tabla de valores a desplegar
    num_rows, num_columns = ws_relacion_ingresos.shape
    columns = list(ws_relacion_ingresos.columns)
    # tabla = [[Paragraph(f'<para wordwrap=CJK>{col}</para>', bodytextStyle) for col in columns]]
    tabla = [[Paragraph(f"<para align='CENTER'>{col}</para>", bodytextStyle) for col in columns]]
    # tabla = [[columns]]
    for row in range(num_rows):
        linea = list()
        linea.append(f"{ws_relacion_ingresos.iloc[row]['Mes'].strftime('%b. %Y').title()}")
        for col in range(1, num_columns):
            monto = ws_relacion_ingresos.iloc[row][columns[col]]
            str_monto = '' if isnan(monto) else edita_número(monto, num_decimals=2)
            linea.append(str_monto)
        tabla.append(linea)

    primera_fila = ['' for _ in columns]
    primera_fila[1] = Paragraph(' Categorias', bodytextStyle)
    tabla.insert(0, primera_fila)

    colwidths = [72 for _ in columns[1:]]
    colwidths.insert(0, 60)


    # Crea la maqueta del archivo .PDF
    body = list()
    body.append(Paragraph('<font size=14><b>GyG - Total de Ingresos Mensuales</b></font>', normalStyle))
    período = f"{mes_inicial.strftime('%B').title()} {mes_inicial.strftime('%Y')+' ' if mes_inicial.year != mes_final.year else ''}" + \
              f"a {mes_final.strftime('%B %Y').title()}"
    body.append(Spacer(1, 6))
    body.append(Paragraph(f'<font size=10><i>{período}</i></font>', normalStyle))
    body.append(Spacer(1,24))

    tabla_resumen = Table(tabla, colWidths=colwidths, hAlign='LEFT')
    # Table coordinates are given as (column, row) which follows the spreadsheet 'A1' model,
    # but not the more natural (for mathematicians) 'RC' ordering. The top left cell is (0, 0)
    # the bottom right is (-1, -1)
    tabla_resumen.setStyle(TableStyle([('VALIGN',     ( 0, 0), (-1,  1), 'BOTTOM'),                 # todas las columnas, 1ra y 2da fila
    #                                  ('TEXTCOLOR',  ( 0, 0), (-1,  1), colors.white),             # todas las columnas, 1ra y 2da fila
                                       ('BACKGROUND', ( 0, 0), (-1,  1), colors.cornflowerblue),    # todas las columnas, 1ra y 2da fila
                                       ('ALIGN',      ( 1, 0), ( 1,  0), 'LEFT'),                   # columna 'B1': 'Categorías'
                                       ('ALIGN',      ( 0, 1), (-1,  1), 'CENTER'),                 # todas las columnas, 2da fila
                                       ('LINEABOVE',  ( 1, 1), (-2,  1), 0.25, colors.white),       # desde 'B2' hasta '<end-1>2'
                                       ('LINEBEFORE', ( 1, 0), ( 1,  0), 0.25, colors.white),       # antes de ' Categoría'
                                       ('LINEBEFORE', (-1, 0), (-1,  0), 0.25, colors.white),       # después de la última categoría
                                       ('LINEBEFORE', ( 0, 1), (-1,  1), 0.25, colors.white),       # entre columnas, 2da línea
                                       ('ALIGN',      ( 0, 2), ( 0, -1), 'CENTER'),                 # todas las filas a partir de la 3ra fila
                                       ('ALIGN',      ( 1, 2), (-1, -1), 'RIGHT'),                  # bloque 'B3:<end><end>'
    #                                  ('VALIGN',     ( 1, 1), (-1, -1), 'TOP'),
    #                                  ('FONTNAME',   ( 0, 0), (-1, -1), FONT_NAME_HELVETICA),
    #                                  ('FONTSIZE',   ( 0, 0), (-1, -1), 10)
                                    ]))

    body.append(tabla_resumen)
    body.append(Spacer(1,24))

    body.append(Paragraph(f"<font size=10><i>(resumen generado el {ahora.strftime('%d/%m/%Y %I:%M')}{am_pm})</i></font>", normalStyle))

    print(f'Generando "{filename}"...')
    doc.build(body)


#
# PROCESO
#

# Selecciona el período a reportar
ahora = datetime.now()
am_pm = 'pm' if ahora.hour > 12 else 'm' if ahora.hour == 12 else 'am'
fecha_de_referencia = datetime(ahora.year, ahora.month, ahora.day)
mes_actual = datetime(ahora.year, ahora.month, 1)
mes_final = mes_actual - relativedelta(months=1)
fecha_inicial = fecha_de_referencia - relativedelta(months=NUM_MESES)
mes_inicial = datetime(fecha_inicial.year, fecha_inicial.month, 1)

if not toma_opciones_por_defecto:
    frame_layout = [    [sg.Checkbox("PDF",   key='_formato_PDF_', default=True),
                         sg.Checkbox("Texto", key='_formato_texto_')]
                   ]
    layout = [  [sg.Text("Mes inicial:", size=(10,1)),
                 sg.InputText(key='_mes_inicial_', size=(10,1),
                              tooltip='Indique el mes inicial a desplegar',
                              default_text=mes_inicial.strftime("%b.%Y")),
                 sg.CalendarButton(button_text='Seleccione fecha',
                                   image_filename=CALENDAR_ICON, image_size=CALENDAR_SIZE, image_subsample=CALENDAR_SUBSAMPLE,
                                   tooltip='Presione para seleccionar otra fecha',
                                   default_date_m_d_y=(mes_inicial.month, mes_inicial.day, mes_inicial.year),
                                   target='_mes_inicial_', format='%b.%Y',
                                   close_when_date_chosen=True),
                 sg.Text(" "*5),
                 sg.Frame("Formato de salida:", frame_layout)],
                [sg.Text("Mes final:", size=(10,1)),
                 sg.InputText(key='_mes_final_', size=(10,1),
                              tooltip='Indique el mes final a desplegar',
                              default_text=mes_final.strftime("%b.%Y")),
                 sg.CalendarButton(button_text='Seleccione fecha',
                                   image_filename=CALENDAR_ICON, image_size=CALENDAR_SIZE, image_subsample=CALENDAR_SUBSAMPLE,
                                   tooltip='Presione para seleccionar otra fecha',
                                   default_date_m_d_y=(mes_final.month, mes_final.day, mes_final.year),
                                   target='_mes_final_', format='%b.%Y',
                                   close_when_date_chosen=True)],
                [sg.Text('_'*70)],
                [sg.Button('Genera resumen'), sg.Button('Finaliza')]
             ]

    # Create the Window
    window = sg.Window('Resumen de Ingresos', layout)

    event, values = window.read()
    if event in (None, 'Finaliza'):
        sys.exit()
    elif event != 'Genera resumen':
        print(f'\n*** ERROR: evento recibido "{event}"\n')
        sys.exit()

    mes_inicial = datetime.strptime(values['_mes_inicial_'], '%b.%Y')
    mes_final   = datetime.strptime(values['_mes_final_'], '%b.%Y')
    mes_actual  = mes_final + relativedelta(months=1)

print()


### SALDOS ---------------------------------------------------------------------------------------

# Lee pestaña de 'Saldos' y extrae las columnas relativas a fechas y montos
print(f'Cargando hoja de cálculo "{excel_ws_saldos}" del libro "{excel_workbook}"...')
ws = read_excel(excel_workbook, sheet_name=excel_ws_saldos, header=None, skiprows=3) #762548
col_actual = 5
col_offset = 5
ws_saldos = DataFrame()
while col_actual < ws.shape[1]:
    data = {'Fecha': list(ws[col_actual]), 'Monto': list(ws[col_actual+2])}
    ws_saldos = ws_saldos.append(DataFrame(data))
    col_actual += col_offset

ws_saldos = ws_saldos.dropna()

# Agrega columna 'Mes' para, posteriormente, filtrar por ella los valores requeridos, así como
# la columna 'Categoría' (siempre igual a 'Vigilancia')
ws_saldos['Mes'] = ws_saldos['Fecha'].apply(lambda x: datetime(x.year, x.month, 1))
ws_saldos['Categoría'] = ['Vigilancia' for _ in range(ws_saldos.shape[0])]

ws_saldos = ws_saldos[(ws_saldos['Mes'] < mes_actual) & \
                              (ws_saldos['Mes'] >= mes_inicial)]
ws_saldos = ws_saldos[['Mes', 'Monto', 'Categoría']]


### VIGILANCIA -----------------------------------------------------------------------------------

print(f'Cargando hoja de cálculo "{excel_ws_vigilancia}" del libro "{excel_workbook}"...')
ws_vigilancia = read_excel(excel_workbook, sheet_name=excel_ws_vigilancia)

# Selecciona los pagos de vigilancia entre la fecha de referencia y NUM_MESES_GESTION_COBRANZAS atrás
ws_vigilancia = ws_vigilancia[(ws_vigilancia['Enviado'] == 'ü')  & \
                              (ws_vigilancia['Mes'] < mes_actual) & \
                              (ws_vigilancia['Mes'] >= mes_inicial)]
ws_vigilancia = ws_vigilancia[['Mes', 'Monto', 'Categoría']]


### RESUMEN POR MES ------------------------------------------------------------------------------

print('Totalizando por Categoría y Mes...')
# Concatena las tabla de pagos a vigilancia y de saldos
ws_relacion_ingresos = ws_vigilancia.append(ws_saldos)

# Totaliza los pagos por Categoría y Mes
ws_relacion_ingresos = pivot_table(ws_relacion_ingresos, values='Monto', index=['Categoría', 'Mes'], aggfunc=sum) \
                            .reset_index()

# Convierte el resumen de ingreso en una tabla tipo worksheet (Monto en Meses x Categoría)
ws_relacion_ingresos = ws_relacion_ingresos.pivot(index='Mes', columns='Categoría')
ws_relacion_ingresos.columns = ws_relacion_ingresos.columns.droplevel(0)
ws_relacion_ingresos = ws_relacion_ingresos.reset_index()

# Totaliza las categorías por mes
ws_relacion_ingresos['Total general'] = list(ws_relacion_ingresos.sum(axis=1))

# Genera resumen por mes
if toma_opciones_por_defecto:
    despliega_resumen_pdf()
else:
    if values['_formato_texto_']:
        despliega_resumen_txt()
    if values['_formato_PDF_']:
        despliega_resumen_pdf()

# --------------------------------
# import pickle
# pickle.dump(ws_relacion_ingresos, open('resumen_de_ingresos.p', 'wb'))
#
# import pickle
# ws_relacion_ingresos = pickle.load(open('resumen_de_ingresos.p', 'rb'))
# --------------------------------
