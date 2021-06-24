# GyG APORTE PARA VIGILANTES
#
# A partir del 1ro. de abril 2021, los montos superiores al pago de la cuota de vigilancia serán
# cancelados a los vigilantes como un aporte especial de los vecinos del sector, por lo que los
# pagos estarán clasificados en 'Vigilancia' y 'Aporte Vigilantes'

"""
    PENDIENTES POR HACER
      - 

    NOTAS
      - Añadir opción para seleccionar la categoría a analizar; no solicitarla en caso de haber
        sólo una categoría posible.
         -> Revisar título del reporte: "GyG PAGOS ADICIONALES PARA VIGILANTES, Mayo/2021"

    HISTORICO
      √ ¿Se incluye el campo 'Dirección', existente en la versión original?
         -> Se omite la dirección cuando se muestran dos o más meses (22-06-2021)
      √ Agregar parámetro '--meses=n' con la cantidad de meses a mostrar por defecto. En caso de
        no estar presente, tomar '1' como opción por defecto. Igualmente se agregó la opción
        '--mes_actual' para tomar el mes actual como referencia y no el mes anterior (22/06/2021)
      - Se reescribe el código para permitir el despliegue de múltiples columnas de aportes,
        facilitando la comparación con meses anteriores (16/06/2021)
      - Se añade la opción para generar las tablas de detalle en fuente monoespaciada para ser
        enviadas por WhatsApp (19/05/2921)
      - Un vecino puede haber hecho más de un aporte en el mes. Totalizar para mostrar un solo
        registro
         -> Corregido (04/05/2021)
      - Versión inicial (02/05/2021)
    
"""

print('Cargando librerías...')
import GyG_constantes
from GyG_utilitarios import *
from pandas import read_excel, pivot_table, isnull, notnull
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys, os

import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

nMeses = 1
toma_opciones_por_defecto = False
selecciona_mes_actual = False

for idx in range(1, len(sys.argv)):
    if sys.argv[idx].startswith("--meses="):
        try:
            nMeses = int(sys.argv[idx].replace("--meses=", ""))
        except:
            nMeses = 1
    elif sys.argv[idx] == '--mes_actual':
        selecciona_mes_actual = True
    else:
        toma_opciones_por_defecto = sys.argv[idx] == '--toma_opciones_por_defecto'


#
# DEFINE CONSTANTES
#

nombre_análisis   = GyG_constantes.txt_aporte_vigilantes     # "GyG Aporte Vigilantes {:%Y-%m (%b)}.txt"
attach_path       = GyG_constantes.ruta_analisis_de_pagos    # "./GyG Recibos/Análisis de Pago"

excel_workbook    = GyG_constantes.pagos_wb_estandar         # '1.1. GyG Recibos.xlsm'
excel_worksheet   = GyG_constantes.pagos_ws_vigilancia       # 'Vigilancia'
fmt_fecha         = "%b-%Y"

strTotal          = 'Total'
muestra_dirección = True

LONG_BENEFICIARIO = 18
LONG_DIRECCIÓN    = 14
LONG_MONTOS       = 11
LONG_TOTAL        = 13

# CATEGORÍAS DE PAGO DEFINIDAS ACTUALMENTE (1.1. GyG Recibos.xlsm > Vigilancia)
#     'Vigilancia',        'Aporte Vigilantes',
#     'Reparación Portón', 'Caseta Vigilancia',   'Reja CC',          'Reparación Portón 2',
#     'Aporte caseta',     'Teléfono Vigilancia', 'Cesta de Navidad', 'Comida Vigilantes',
#     'Control',           'Logo',                'Ingresos por Ventas',
#     'ANULADO',           'SOLVENTE',            'ANTICIPO',         'DONACION',
#     'REVERSADO',         'DEPOSITO'

posibles_categorías = [GyG_constantes.CATEGORIA_APORTE_VIGILANTES,
                       GyG_constantes.CATEGORIA_VIGILANCIA,
                      ]


#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def esAporte(x):
    return x == strCategoría


def genera_resumen():

    def linea_detalle(r):
        detalle = list()
        if r['Beneficiario'] == strTotal:
            es = 'es' if nMeses + muestra_dirección > 1 else ''
            detalle.append(alinea_texto(f'Total{es}', anchura=LONG_BENEFICIARIO, alineación=">"))
        else:
            detalle.append(alinea_texto(
                                trunca_texto(
                                    edita_beneficiario(r['Beneficiario']),
                                    max_width=LONG_BENEFICIARIO),
                                anchura=LONG_BENEFICIARIO, alineación="<")
                            )
        if muestra_dirección:
            if r['Beneficiario'] == strTotal:
                s = '' if df_aportes.shape[0] - 1 == 1 else 's'
                familia = 'pago' if nMeses == 1 else 'familia'
                detalle.append(alinea_texto(f"{df_aportes.shape[0] - 1} {familia}{s}",
                                anchura=LONG_DIRECCIÓN, alineación=">"))
            else:
                detalle.append(alinea_texto(edita_dirección(r['Dirección']),
                                anchura=LONG_DIRECCIÓN, alineación="<"))
        for mes in últimos_meses:
            try:
                monto = r[mes]
                if isnull(monto):
                    monto = 0.0
            except:
                monto = 0.0
            if monto == 0.0:
                detalle.append(alinea_texto('-', anchura=LONG_MONTOS))
            else:
                detalle.append(alinea_texto(edita_número(monto, num_decimals=0), anchura=LONG_MONTOS))
        if nMeses > 1:
            detalle.append(alinea_texto(edita_número(r[strTotal], num_decimals=0), anchura=LONG_TOTAL))

        return ' | '.join(detalle) + '\n'

    am_pm = 'pm' if hoy.hour > 12 else 'm' if hoy.hour == 12 else 'am'
    encabezado = ''.join([
                        wa_bold, "GyG PAGOS ADICIONALES PARA VIGILANTES, ",
                        f"{fecha_referencia.strftime('%B/%Y').capitalize()}", wa_bold, "\n",
                        f"al {hoy.strftime('%d/%m/%Y %I:%M')} {am_pm}\n",
                        wa_table,
                        "\n"
                    ])
    tabla_encabezado = ''.join([
                            alinea_texto('Vecinos', anchura=LONG_BENEFICIARIO, alineación='^'),
                            ' | ',
                            alinea_texto('Dirección', anchura=LONG_DIRECCIÓN, alineación='^') if muestra_dirección else '',
                            ' | ' if muestra_dirección else '',
                            ' | '.join(
                                [alinea_texto(mes.capitalize(), anchura=LONG_MONTOS, alineación='^') for mes in últimos_meses]),
                            ' | ' if nMeses > 1 else '',
                            alinea_texto(strTotal, anchura=LONG_TOTAL - 2, alineación='>') if nMeses > 1 else '',
                            '\n'
                        ])
    long_separador = (LONG_BENEFICIARIO + 3 + (LONG_MONTOS + 3) * nMeses + LONG_TOTAL + 1)
    long_separador += (LONG_DIRECCIÓN + 3) if muestra_dirección else 0
    long_separador -= (LONG_TOTAL + 3) if nMeses <= 1 else 0
    tabla_separador  = "-" * long_separador + "\n"

    análisis = encabezado + tabla_encabezado + tabla_separador

    for _, r in df_aportes.iterrows():
        if r['Beneficiario'] == strTotal:
            análisis += tabla_separador
        análisis += linea_detalle(r)
    análisis += wa_table + '\n'

    return análisis

def selecciona_categoría() -> str:
    size = len(max(posibles_categorías, key=len))
    nDigits = len(str(len(posibles_categorías)))
    nColumns = 3
    if nColumns * (3 + nDigits + 2 + size + 1) > 80:
        nColumns -= 1

    print("\nCategorías:", end="")
    for idx, categoría in enumerate(posibles_categorías):
        newline = '\n' if idx % nColumns == 0 else ''
        # print(f"{newline}   [{alinea_texto(str(idx + 1), anchura=nDigits, alineación=">")}]", end="")
        print(f"{newline}   [{(idx + 1):<{nDigits}}] {alinea_texto(categoría, anchura=size, alineación='<')}", end="")
    print('\n')
    while True:
        idx_categoría = input_valor('Indique la categoría a analizar', 1, toma_opciones_por_defecto)
        if not (1 <= idx_categoría <= len(posibles_categorías)):
            print(f"   -> ERROR: Categoría erronea ({idx_categoría})")
        else:
            break
    return posibles_categorías[idx_categoría - 1]


#
# PROCESO
#

# Determina el mes actual, a fin de utilizarlo como opción por defecto
hoy = datetime.now()
fecha_análisis = datetime(hoy.year, hoy.month, 1)
mes_actual = (fecha_análisis - relativedelta(days=0 if selecciona_mes_actual else 1)).strftime('%m-%Y')
print()

# Selecciona el mes y año a procesar
mes_año = input_mes_y_año('Indique el mes y año a analizar', mes_actual, toma_opciones_por_defecto)

# Selecciona la categoría a analizar
strCategoría = selecciona_categoría() if len(posibles_categorías) > 1 else posibles_categorías[0]

# Selecciona el número de meses a desplegar
while True:
    nMeses = input_valor('Indique la cantidad de meses a mostrar', nMeses, toma_opciones_por_defecto)
    if nMeses <= 0:
        print(f"   -> ERROR: cantidad de meses inválido ({nMeses})")
    else:
        break

# Selecciona si se ordenan alfabéticamente los vecinos
ordenado = input_si_no('Vecinos ordenados alfabéticamente', 'sí', toma_opciones_por_defecto)

# Selecciona si se colocarán marcas adicionales para ser interpretadas por WhatsApp
whatsapp = input_si_no("Para ser enviado por WhatsApp", 'no', toma_opciones_por_defecto)

# Prepara los caracteres de negrita, italizado y monoespaciado para los textos, en caso de WhatsApp
wa_bold, wa_italic, wa_table = ('*', '_', '```') if whatsapp else ('', '', '')

# La dirección sólo se mostrará cuando se presente un mes. Esto para facilitar el despliegue
# de la tabla en WhatsApp
muestra_dirección = nMeses == 1


año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = datetime(año, mes, 1)
# str_fecha_referencia = fecha_referencia.strftime(fmt_fecha)

print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))

# Lee la hoja de cálculo con el detalle de los pagos
df_pagos = read_excel(excel_workbook, sheet_name=excel_worksheet)

# Elimina los registros que no no corresponden a aportes para vigilancia
df_pagos = df_pagos.loc[esAporte(df_pagos['Categoría'])]

# Selecciona los meses a desplegar
últimos_meses = [fecha_referencia - relativedelta(months=offset) for offset in reversed(range(nMeses))]
df_pagos = df_pagos[df_pagos['Mes'].isin(últimos_meses)]
últimos_meses = [col.strftime(fmt_fecha) for col in últimos_meses]

if not df_pagos.shape[0]:
    print()
    print(f'*** Error: No hay montos registrados para {strCategoría.upper()}')
    print(f'{espacios(10)} durante el período de {reagrupa_meses(", ".join(últimos_meses), mes_completo=True)}')
    print()
    sys.exit()

# Genera la tabla pivote
df_aportes = pivot_table(df_pagos, index=['Beneficiario', 'Dirección'], columns='Mes', values='Monto',
                                   margins=True, margins_name=strTotal, aggfunc='sum').reset_index()

df_aportes.columns = [(col.strftime(fmt_fecha) if isinstance(col, datetime) else col) for col in df_aportes.columns]

# Ordena los beneficiarios en orden alfabético
if ordenado:
    r = df_aportes[df_aportes['Beneficiario'] == strTotal]
    df_aportes = df_aportes[df_aportes['Beneficiario'] != strTotal]
    df_aportes['_benef_sort_'] = df_aportes['Beneficiario'] \
        .apply(lambda benef: trunca_texto(edita_beneficiario(remueve_acentos(benef)), LONG_BENEFICIARIO))
    df_aportes.sort_values(by='_benef_sort_', inplace=True)
    df_aportes = df_aportes.append(r, ignore_index=True)

# Crea el archivo con el análisis
print(f"Creando análisis '{nombre_análisis.format(fecha_referencia)}'...")
print()

análisis = genera_resumen()

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(análisis)

filename = os.path.join(attach_path, 'Windows', nombre_análisis.format(fecha_referencia))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(análisis)
