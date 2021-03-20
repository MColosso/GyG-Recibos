# GyG CARTELERA VIRTUAL
#
# Elabora la Cartelera Virtual a partir de la hoja de cálculo de pagos

"""
    PENDIENTE POR HACER
    -   Después de ajustar todos los programas para usar esta nueva clase, cambiar el título
        de la columna "2016" en la hoja de Excel a "=fecha(2016;1;1)" y eliminar el cambio
        de nombre en la rutina principal
    -   En ocasiones es necesario que el 'resumen_de_cuotas()' muestre también el monto de la
        cuota del mes en curso, y en otros no. ¿Cómo se determina cuando sí y cuándo no? o,
        ¿cómo se le indica...?
    -   Revisar cálculo del saldo pendiente:
        El Colegio El Trigal quedó debiendo Bs. 20.000 en el mes de febrero 2021 (Bs. 5.220.000)
        y canceló por completo la cuota del mes de marzo 2021 (Bs. 5.500.000).
        Al ejecutar el reporte el 1ro de marzo, saldo pendiente mostrado es de Bs. 300.000.
        ¿El saldo pendiente a mostrar serían Bs. 20.000 (el monto pendiente por cancelar en febrero)
        o Bs. 300.000 (el monto pendiente por cancelar para completar la cuota vigente)
         -> El saldo es correcto y acorde con la política: "las cuotas atrasadas serán canceladas
            según la cuota vigente a la fecha de pago, indistintamente si se hubiera realizado un
            anticipo para dicho mes o no"
             -> Hay un pago completo para el mes siguiente (marzo). ¿Debe reflejarse esto de alguna
                forma...?

    
    HISTORICO
    -   Corregir:
         - Scaletta Briceño, Av. "L" anexo de Los Tronquitos. Tiene cuotas pendientes desde 2016. Saldo
            actual Bs. 41.800.000
        La fecha de cuotas pendientes y el saldo actual no concuerdan...
         -> Si el vecino no tiene pagos registrados, automáticamente se colocaba '2016' como fecha de
            último pago, indistintamente de su inicio en el sector (19/03/2021)
    -   Agregar opción para mostrar el saldo deudor y la fecha de última cancelación en la cartelera
        de los que no participan. (19/03/2021)
    -   Mostrar los saldos pendientes ajustados por inflación, así como se hace en el Resumen de
        Saldos (resumen_saldos.py) (16/03/2021)
    -   Ajustado el texto final de la cartelera 1: "Si usted paga el 100%+ de la cuota y no aparece
        en este listado..." a "Si usted paga la totalidad de la cuota o más, y no aparece en este
        listado..." para facilitar su lectura (10/10/2020)
    -   Si el mes seleccionado corresponde al mes actual, colocar la fecha de hoy como fecha de
        referencia (17/05/2020)
    -   Corregir inconsistencia en 'no_participa_desde()': Al seleccionar una fecha de referencia
        anterior a la última colaboración efectuada, se muestra esta última, y no la última desde
        la fehcha indicada:
             Ej. Cartelera Virtual al 31 de agosto de 2019                            <-- ago. 2019
                  - Rivero, Güigüe 88-40. Tiene colaboraciones pendientes desde octubre 2018. Último
                    pago: 04 dic 2019, Bs. 20.000, correspondiente a Noviembre 2019   <-- nov. 2019
          -> Corregido: Al abrir la hoja de cálculo de pagos, se ignoran los pagos posteriores a
             la fecha de referencia (24/02/2020)
    -   En la cartelera de COLABORADORES, colocar el monto y último mes cancelado
          -> Cambia "Muestra los saldos de colaboradores" a "Muestra el último pago de
             colaboradores"
             Si la respuesta es afirmativa, cambia el texto  a desplegar de " - Castillo Rodríguez,
             Güigüe 88-91. Tiene pendiente diciembre 2019" a " - Castillo Rodríguez, Güigüe 88-91.
             Tiene pendiente diciembre 2019. Última colaboración: 26 nov. 2019, Bs. 81.750, mes de
             Noviembre 2019"
             (27/01/2020)
    -   Cambiar la forma para la lectura de opciones desde el standard input, estandarizando su uso e
        implementando la opción '--toma_opciones_por_defecto' en la linea de comandos (08/12/2019)
    -   Incluir en el total de la deuda cualquier saldo pendiente del último mes pagado (tomar Resumen
        de Saldos a la fecha -resumen_saldos.py- como referencia) (26/11/2019)
    -   Ajustar la codificación de caracteres para que el archivo de saldos sea legible en Apple y
        Windows
          -> Se generan archivos con codificaciones 'UTF-8' (Apple) y 'cp1252' (Windows) en carpetas
             separadas (29/10/2019)
    -   Se cambiaron las ubicaciones de los archivos resultantes a la carpeta GyG Recibos dentro de la
        carpeta actual para compatibilidad entre Windows y macOS (21/10/2019)
    -   Incluir opción para mostrar saldos en Colaboradores o no (14/10/2019)
    -   Cambiar el manejo de cuotas para usar las rutinas en la clase Cuota (GyG_cuotas) (28/09/2019)
    -   "Buenos días vecinos, en reunión de Junta Directiva se estableció la cuota para el pago de la
        vigilancia en un dolar en base a la tasa de cierre semanal del Banco Central, actualmente en
        Bs. 23.400
        Las cuotas pendientes por cancelar, hasta Agosto 2019, quedarán en los montos fijos ya
        establecidos. A partir del mes de Septiembre, toda cuota atrasada se cancelará con la tasa de
        la semana en curso"
          -> Ajustado (10/09/2019)
          -> Corregir: Se muestran vecinos con saldo '0' en la cartelera (1): "Familias que
             pagan el 100% de la cuota y tienen retraso en el pago" (17/09/2019)
    -   Corregir: No se está sumando el monto de la cuota del mes de referencia cuando el último mes
        cancelado es el inmediato anterior al mes de referencia (corregido: 26/06/2019)
    -   Mostrar opcionalmente el monto adeudado por los vecinos
          -> Agregar el parámetro opcional '--saldos' para controlar el despliegue del saldo adeudado a
             la fecha (17/06/2019)
    -   Revisar: Anahiz Seijas no apareció en las cartelera de mayo/2019. Ella está como 'Cuota completa'
        y 'Comida' hasta mayo/2019
          -> Se corrigió un error en la comparación de la fecha de referencia
             (fecha_referencia) con las fechas de inicio y finalización (F.Desde y F.Hasta)
             (07/06/2019)
    -   Ordenar las carteleras virtuales por tipo (1, 2, 3, 4 y 5) - Evaluar si generarlas en una lista
        en memoria y luego imprimirlas
          -> Sólo las carteleras 4 y 5 se generan en memoria para imprimirlas al final. Se forzó la lista
             de categorías para obtener el ordenamiento deseado (10/04/2019)
    -   Generar una cartelera con las familias que colaboran con la comida para los vigilantes.
        Omitir de la cartelera con las familias que no participan aquellos que aparecen en la lista de
        colaboración de comida para los vigilantes (10/04/2019)
          -> Esta nueva cartelera está formada por las familias marcadas con una tilde (ü) en la columna
             "Comida" en la pestaña "R.VIGILANCIA (reordenado)"
          -> Para que no se desplieguen adicionalmente en la cartelera de las familias que
             no colaboran, cambiar la categoría "No participa" a "No participa_"
    -   Cambiar encabezado de los vecinos que "No participan", de "Familias que no participan en el
        mantenimiento del servicio de vigilancia" a "Familias que se benefician del servicio de vigilancia
        gracias al aporte del resto de los vecinos que sí cancelan su cuota" o similar (09/04/2019)
    -   Se muestra en la cartelera correspondiente un vecino en fecha posterior a "F.Hasta"
        y "Categoría" no nula (Ej. En la cartelera de enero 2019, la Familia Morales Sesar
        se muestran en cartelera Colaboración a pesar de que están hasta septiembre 2018)
        -- Corregido 02/02/2019
    -   Reducir las referencias redundantes al año en el resumen de cuotas de vigilancia
        y eliminar los decimales iguales a cero
    -   Generar automáticamente el resumen de cuotas de vigilancia de los últimos <nMeses>
        a partir de la misma hoja de cálculo
    -   Nuevos propietarios se muestran con saldos pendientes desde antes de haber asumido
        el pago del inmueble (Ej. Familia Acevedo, compra en Julio/2018 y se muestra con
        saldos pendientes desde el 2016; la familia Zampini Scaccia, a quien habían com-
        prado, no habían cancelado desde dicha fecha. De igual forma, Wilson Escamilla)
        Establecer fechas de inicio y final para evaluar dentro de ese rango. Una fecha en
        blanco indica que ese extremo está abierto
    -   Omitir la impresión de una cartelera cuando 'seña' sea igual a cero (0): OCULTO,
        etc. (!= 'Colaboración', 'Cuota completa' o 'No participa')
    -   Generar una cuarta cartelera con aquellos vecinos que se encuentran al día o que tienen cuotas
        pagadas por adelantado (4)
    -   Destacar los registros de aquellos vecinos que pagan cuota completa y el monto cancelado en el mes
        de análisis sea inferior a la cuota establecida
    -   Indicar '(1)' para los que pagan cuota completa, '(2)' para los que colaboran y (3) para los que no
        participan (08/feb/2018)
    
"""

# Selecciona las librerías a utilizar
print('Cargando librerías...')
import GyG_constantes
from GyG_cuotas import *
from GyG_utilitarios import *
from pandas import read_excel, isnull, notnull
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import sys
import os
import numbers
import re
import locale

# Define textos
nombre_cartelera   = GyG_constantes.txt_cartelera_virtual         # "GyG Cartelera Virtual {:%Y-%m (%B)}.txt"
attach_path        = GyG_constantes.ruta_cartelera_virtual        # "./GyG Recibos/Cartelera Virtual"

excel_workbook     = GyG_constantes.pagos_wb_estandar             # '1.1. GyG Recibos.xlsm'
excel_worksheet    = GyG_constantes.pagos_ws_resumen_reordenado   # 'R.VIGILANCIA (reordenado)'
excel_cuotas       = GyG_constantes.pagos_ws_cuotas               # 'CUOTA'
excel_pagos        = GyG_constantes.pagos_ws_vigilancia           # 'Vigilancia'

encabezado = 'CARTELERA VIRTUAL\n   CUADRA SEGURA GyG ({})\n' + \
             'al {:%d de %B de %Y}{}\n\n'
encabezado_colaboración   = 'Familias que colaboran con un monto inferior a la cuota establecida y tienen ' + \
                            'pendiente el pago.\n'
encabezado_cuota_completa = 'Familias que pagan el 100% de la cuota y tienen retrasos en el pago.\n'
# encabezado_no_participa   = 'Familias que se benefician del servicio de vigilancia gracias al aporte del resto ' + \
#                             'de los vecinos que cancelan su cuota para la seguridad de la zona.\n'
encabezado_no_participa   = 'Familias que no contribuyen con el pago del servicio de vigilancia.\n'
encabezado_al_día         = 'Familias que pagan el 100% de la cuota o colaboran con dicho pago y se encuentran ' + \
                            'al día en el mantenimiento del servicio de vigilancia.\n'
encabezado_comida_vigilantes = 'Familias que colaboran con almuerzos, cenas o agua para los vigilantes.\n'
pie_cuota_completa = '<< Generado automáticamente >>'
pie_no_participa   = 'Los invitamos a que se sumen a este esfuerzo. Ello nos ayudará a mantener la seguridad ' + \
                     'que hemos alcanzado en el sector. Tu contribución es importante.\n\n'
pie_de_página      = 'Equipo de cobranza, Cuadra Segura GyG\n\n' + \
                     'Cualquier comentario por Trivialidades.\n({:%d/%m/%Y %I:%M} {})\n'

nMeses             =  5   # Se muestran los <nMeses> últimos meses en el resumen de cuotas

dummy = locale.setlocale(locale.LC_ALL, 'es_es')

toma_opciones_por_defecto = False
if len(sys.argv) > 1:
    toma_opciones_por_defecto = sys.argv[1] == '--toma_opciones_por_defecto'

muestra_ultimo_pago_colaboracion = False
aplica_IPC = False

#
# DEFINE ALGUNAS RUTINAS UTILITARIAS
#

def get_filename(filename):
    return os.path.basename(filename)

#def edita_número(number, num_decimals=2):
#    return locale.format_string(f'%.{num_decimals}f', number, grouping=True, monetary=True)

def is_numeric(valor):
    return isinstance(valor, numbers.Number)

def list_and(l1, l2): return [a and b     for a, b in zip(l1, l2)]
def list_or(l1, l2):  return [a or  b     for a, b in zip(l1, l2)]
def list_nor(l1, l2): return [not(a or b) for a, b in zip(l1, l2)]
def list_not(l1):     return [not a       for a    in     l1     ]

def seleccionaRegistro(beneficiarios, categorías, montos):

    def list_lt_cuota(l1):
        # Para cada beneficiario,
        #    busca la fecha del último pago para el mes 'last_col'
        #    busca la cuota establecida para el momento del último pago
        #    agrega a la lista de resultados la comparación 'monto < cuota'
        l2 = list()
        for beneficiario, monto in zip(beneficiarios, montos):
            f_ultimo_pago = fecha_ultimo_pago(beneficiario, columns[last_col].strftime('%m-%Y'), fecha_real=False)
            if f_ultimo_pago == None:
                l2.append(False)
            else:
                cuota = cuotas_obj.cuota_vigente(beneficiario, f_ultimo_pago)
                l2.append(monto < cuota)
        return l2

    # Vecinos que pagan cuota completa y no han cancelado la totalidad de la cuota
    list_1 = list_or(
                        list_and(categorías == 'Cuota completa', list_lt_cuota(montos)),
                        list_and(categorías == 'Cuota completa', isnull(montos))
                    )
    # Vecinos que colaboran con el pago y no han colaborado en el mes de análisis
    list_2 = list_and(categorías == 'Colaboración', isnull(montos))
    # Vecinos que no participan
    list_3 = categorías == 'No participa'

    return list_or(list_1, list_or(list_2, list_3))

def fecha_ultimo_pago(beneficiario, str_mes_año, fecha_real=True):
    try:
        df_fecha = df_pagos[(df_pagos['Beneficiario'] == beneficiario) & (df_pagos['Meses'].str.contains(str_mes_año))]
    except:
        print(f"ERROR: fecha_ultimo_pago({beneficiario}, {str_mes_año}): {str(sys.exc_info()[1])}")
        return None
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

# a_evaluar = ['Familia Scaletta Briceño', 'Familia Rodríguez Chiappeta', ]

def no_participa_desde(r):
    """
        Busca a partir de qué fecha no se han recibido pagos
        (evalúa desde el mes y año indicado, hasta el 2016)
    """

    # if r['Beneficiario'] in a_evaluar: print(f"\nBeneficiario: {r['Beneficiario']}\n{'-'*(len('Beneficiario: ')+len(r['Beneficiario']))}")
    # if r['Beneficiario'] in a_evaluar: print(f"  (0): f_desde: {columns.index(r['F.Desde'])} ({columns[columns.index(r['F.Desde'])]}), last_col: {last_col} ({columns[last_col]})")

    x = last_col       # <-------- 'x' es la primera columna vacía
    ultimo_mes_con_pagos = None
    saldo_ultimo_mes = 0.00
    cuotas_pendientes = 0.00
    saldo_pendiente = False
    f_desde = columns.index(r['F.Desde'])

    for idx in reversed(range(f_desde, last_col+1)):
        # if r['Beneficiario'] in a_evaluar: print(f"  (1): idx: {idx}, cuota: {cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx])}, r.iloc[idx]: {r.iloc[idx]}")
        # if r['Beneficiario'] in a_evaluar: print(f"       columns[{idx}]: {columns[idx]}")
        if notnull(r.iloc[idx]):
            ultimo_mes_con_pagos = idx
            break
        x = idx
        cuotas_pendientes += cuotas_obj.cuota_actual(r['Beneficiario'], columns[idx], aplica_IPC=aplica_IPC) # cuotas[idx]
    if ultimo_mes_con_pagos == None:
        ultimo_mes_con_pagos = this_col = f_desde
        fecha_de_inicio = columns[columns.index(r['F.Desde'])]
        # fecha_txt = '2016'
        fecha_txt = '2016' if fecha_de_inicio == datetime(2016, 1, 1) else fecha_de_inicio.strftime('%B %Y')
    else:
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
            cuota_actual = cuotas_obj.cuota_vigente(r['Beneficiario'], f_ultimo_pago)
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

    info_deuda = ''
    if saldo_ultimo_mes == 0.00 and ultimo_mes_con_pagos < last_col and fecha_txt != '2016':
        ultimo_mes_con_pagos += 1
        this_col = columns[ultimo_mes_con_pagos] if isnull(r.iloc[last_col]) else columns[ultimo_mes_con_pagos]   # <<<=== anteriormente 'x+1' en lugar de 'x'
        fecha_txt = '2016' if this_col == datetime(2016, 1, 1) else this_col.strftime('%B %Y')
    # if r['Beneficiario'] in a_evaluar: print(f"  (9): x: {ultimo_mes_con_pagos}, last_col: {last_col}, r[{fecha_referencia}]: {r[fecha_referencia]}, deuda: {saldo_ultimo_mes}")

    if deuda_actual != 0.00:
        if saldo_ultimo_mes != 0.00:
            info_deuda = 'Tiene una diferencia pendiente en ' + fecha_txt
            if ultimo_mes_con_pagos != last_col:
                info_deuda += ' y cuotas de meses subsiguientes'
        else:
            if ultimo_mes_con_pagos == last_col:
                info_deuda = 'Tiene pendiente ' + fecha_txt
            else:
                info_deuda = f"Tiene {'colaboraciones' if r['Categoría'] == 'Colaboración' else 'cuotas'} " + \
                          f"pendientes desde {fecha_txt}"

    if muestra_saldos and (deuda_actual != 0.00) and \
        (((r['Categoría'] == 'Colaboración') and muestra_ultimo_pago_colaboracion) or \
         ((r['Categoría'] == 'No participa') and muestra_saldo_no_participa) or \
         (r['Categoría'] == 'Cuota completa')):
        # if r['Beneficiario'] in a_evaluar: print(f" (9a): Complemento de mensaje sobre deuda")
        if saldo_pendiente:
            sep1, sep2 = (' por ', '')
            if x == last_col:
                sep1 = ' de '
        elif r['Categoría'] in ['Colaboración', 'No participa']:    # , 'No participa'
            ultimo_pago = df_pagos[df_pagos['Beneficiario'] == r['Beneficiario']].tail(1).squeeze()
            #
            #    <-- Verificar, adicionalmente, si df_pagos['Fecha'] < "Fecha de referencia"
            if len(ultimo_pago) > 0:
                u_fecha = ultimo_pago['Fecha'].strftime('%d/%m/%Y')
                # u_monto = edita_número(ultimo_pago.Monto, num_decimals=2).replace(',00', '')
                u_concepto = re.search(r'.*([,:] )(.*)$', str(ultimo_pago.Concepto)).group(2)   # Busca ', ' (como en 'Cancelación Vigilancia, ') o
                                                                                                # ': ' (como en 'Vigilancia: ') y toma el resto del
                                                                                                # string
                # if r['Beneficiario'] in a_evaluar: print(f"  (9c): Concepto: {ultimo_pago.Concepto}, concepto editado: {u_concepto}")

                u_concepto = re.sub('mes(es)* de', 'correspondiente a', u_concepto)             # Cambia 'mes[es] de' por 'correspondiente a'
                sep1 = f". Último pago: {u_fecha}, "
                sep2 = f", {u_concepto}"
                if (r['Categoría'] == 'No participa') and muestra_saldo_no_participa:
                    sep2 += f". Saldo Bs. {edita_número(deuda_actual, num_decimals=2).replace(',00', '')}"
                deuda_actual = ultimo_pago['Monto']
            else:
                sep1, sep2 = '. Saldo actual ', ''
        else:
            sep1, sep2 = (' (', ')')
        # if r['Beneficiario'] in a_evaluar: print(f' (9b): Complemento: sep1 = "{sep1}", sep2 = "{sep2}", {deuda_actual = }')
        info_deuda += f"{sep1}Bs. {edita_número(deuda_actual, num_decimals=2).replace(',00', '')}{sep2}"

    # if r['Beneficiario'] in a_evaluar: print(f" (10): info_deuda: {info_deuda}")

    return info_deuda

def ajusta_nombre_y_dirección(str):
    """
        Remueve los textos "Familia " y "Calle " del texto recibido
    """
    return ' - ' + str.replace('Familia ', '').replace('Calle ', '').replace('Nro. ', '').replace('Nros. ', '')


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

def separa_meses(mensaje, as_string=False):
    tokens_validos = meses + meses_abrev + conectores
    mensaje = re.sub("\([^()]*\)", "", mensaje)
    mensaje = re.sub(r"\W ", " ", mensaje.lower().replace('-', ' a ').replace('/', ' ')).split()
    mensaje_ed = [x for x in mensaje if (x in tokens_validos) or x.isdigit()]
    last_year = None
    last_month = None
    mensaje_final = list()
    maneja_conector = False
    for x in reversed(mensaje_ed):
        token = meses[meses_abrev.index(x)] if x in meses_abrev else x
        if token.isdigit():
            last_year = token
            last_month = None
            continue
        elif token in meses:
            if maneja_conector:
                try:
                    n_last_month = meses.index(last_month)
                except:
                    continue    # ignora los mensajes que contienen textos del tipo:
                                # "(saldo a favor: Bs. 69.862,95)"
                n_token = meses.index(token)
                for t in reversed(range(n_token + 1, n_last_month)):
#                    mensaje_final = [f"{meses_abrev[t]}.{last_year}"] + mensaje_final
                    mensaje_final = [f"{t:02d}-{last_year}"] + mensaje_final
                maneja_conector = False
            last_month = token
#            mensaje_final = [f"{meses_abrev[meses.index(last_month)]}.{last_year}"] + mensaje_final
            mensaje_final = [f"{meses.index(last_month)+1:02d}-{last_year}"] + mensaje_final
        elif x in conectores:
            maneja_conector = True

    if as_string:
        mensaje_final = '|'.join(mensaje_final)

    return mensaje_final


def cartelera_vecinos_al_día():

    def list_ge_cuota(l1, l2):
        l = [a if is_numeric(a) else cuotas_mensuales[last_col] for a in l2]
        # L3: monto cancelado mayor o igual a la cuota del mes
        l3 = [a >= cuotas_mensuales[last_col] for a in l]
        # L4: colaboración con algún monto pagado en el mes
        l4 = [notnull(a) and b == 'Colaboración' for a, b in zip(l, l1)]

        return list_or(l3, l4)

    texto_cartelera = ''

    # Selecciona aquellos vecinos con pagos en el mes mes objeto de análisis mayor o
    # igual a la cuota del mes
    vad = df.loc[df['Pendiente'] == '']

    # Genera el encabezado
    texto_cartelera += encabezado.format(4, f_ref_último_día, '')
    texto_cartelera += encabezado_al_día

    # Genera el detalle
    inicio_dirección = ''
    for index, row in vad.iterrows():
        if inicio_dirección != row['Dirección'][: 3]:
            inicio_dirección = row['Dirección'][: 3]
            texto_cartelera += '\n'
        texto_cartelera += ajusta_nombre_y_dirección(row['Beneficiario'] + ', ' + row['Dirección']) + '\n'
        if inicio_dirección != row['Dirección'][: 3]:
            inicio_dirección = row['Dirección'][: 3]
            texto_cartelera += '\n'
        
    # Genera el pie de página
    texto_cartelera += '\n'
    texto_cartelera += pie_cuota_completa
    texto_cartelera += pie_de_página.format(hoy, am_pm)
    
    # Genera una linea de separación entre categorías
    texto_cartelera += '\n' + '-' * 25 + '\n\n'

    return texto_cartelera


def cartelera_comida_vigilantes(df_comida):

    texto_cartelera = ''

    # Genera el encabezado
    texto_cartelera += encabezado.format(5, f_ref_último_día, '')
    texto_cartelera += encabezado_comida_vigilantes

    # Genera el detalle
    inicio_dirección = ''
    for index, row in df_comida.iterrows():
        if inicio_dirección != row['Dirección'][: 3]:
            inicio_dirección = row['Dirección'][: 3]
            texto_cartelera += '\n'
        texto_cartelera += ajusta_nombre_y_dirección(row['Beneficiario'] + ', ' + row['Dirección']) + '\n'
        if inicio_dirección != row['Dirección'][: 3]:
            inicio_dirección = row['Dirección'][: 3]
            texto_cartelera += '\n'
        
    # Genera el pie de página
    texto_cartelera += '\n'
    texto_cartelera += pie_de_página.format(hoy, am_pm)
    
    # Genera una linea de separación entre categorías
    #cartelera += '\n' + '-' * 25 + '\n\n'

    return texto_cartelera


#
# PROCESO
#

# Determina la fecha de elaboración del informe
am_pm = 'pm' if datetime.now().hour > 12 else 'm' if datetime.now().hour == 12 else 'am'
hoy = datetime.now()

# Determina el período anterior al actual, a fin de utilizarlo como opción por defecto
período_anterior = (date.today() + relativedelta(months=-1)).strftime('%m-%Y')
print()

# Selecciona el mes y año a procesar
mes_año = input_mes_y_año('Indique el mes y año de la cartelera virtual', período_anterior, toma_opciones_por_defecto)

# Selecciona si se muestran los saldos deudores o no
muestra_saldos = input_si_no('Muestra los saldos pendientes', 'sí', toma_opciones_por_defecto)

# Selecciona si se muestran los saldos deudores de los colaboradores o no
muestra_ultimo_pago_colaboracion = False
if muestra_saldos:
    muestra_ultimo_pago_colaboracion = input_si_no('Muestra el último pago de los colaboradores', 'si', toma_opciones_por_defecto)
    muestra_saldo_no_participa = input_si_no('Muestra el saldo de los que no participan', 'no', toma_opciones_por_defecto)

    # Selecciona si se aplica el ajuste por inflación (IPC - Indice de Precios al Consumidor)
    aplica_IPC = input_si_no("Aplica ajuste por inflación (IPC)", 'sí', toma_opciones_por_defecto)

año = int(mes_año[3:7])
mes = int(mes_año[0:2])
fecha_referencia = datetime(año, mes, 1)
f_ref_último_día = fecha_referencia + relativedelta(months=1) + relativedelta(days=-1)
# Si estamos en el mismo mes, mostrar la fecha de hoy
if (fecha_referencia.year == hoy.year) and (fecha_referencia.month == hoy.month):
    f_ref_último_día = hoy

# Abre la hoja de cálculo de Recibos de Pago
print()
print('Cargando hoja de cálculo "{filename}"...'.format(filename=excel_workbook))
df = read_excel(excel_workbook, sheet_name=excel_worksheet)

# Elimina los registros que no tienen una categoría definida
df.dropna(subset=['Categoría'], inplace=True)

# Ajusta las columnas F.Desde y F.Hasta en aquellos en los que estén vacíos:
# 01/01/2016 y 01/mes/año+1
df.loc[df[isnull(df['F.Desde'])].index, 'F.Desde'] = datetime(2016, 1, 1)
df.loc[df[isnull(df['F.Hasta'])].index, 'F.Hasta'] = datetime(año + 1, mes, 1)

# Elimina aquellos vecinos que compraron (o se iniciaron) en fecha posterior
# a la fecha de análisis
df = df[df['F.Desde'] < f_ref_último_día]

# Elimina aquellos vecinos que vendieron (o cambiaron su razón social) en fecha anterior
# a la fecha de análisis
df = df[df['F.Hasta'] >= f_ref_último_día]

# Cambia el nombre de la columna 2016 a datetimme(2016, 1, 1)
df.rename(columns={2016:datetime(2016, 1, 1)}, inplace=True)

# Define algunas variables necesarias
columns = df.columns.values.tolist()
last_col = columns.index(fecha_referencia)
categorías = list(set(df['Categoría']))
categorías = list(dict.fromkeys(['Cuota completa', 'Colaboración', 'No participa'] + categorías))

# Inicializa el handler para el manejo de las cuotas
cuotas_obj = Cuota(excel_workbook)
pie_cuota_completa = cuotas_obj.resumen_de_cuotas('Vecino genérico', fecha_referencia)

# Lee la hoja de cálculo de pagos
df_pagos = read_excel(excel_workbook, sheet_name=excel_pagos)
df_pagos.drop(df_pagos.index[df_pagos['Categoría'] != 'Vigilancia'], inplace=True)
df_pagos = df_pagos[['Beneficiario', 'Fecha', 'Monto', 'Concepto']]

df_pagos = df_pagos[df_pagos['Fecha'] <= f_ref_último_día]          # Ignora los pagos posteriores
                                                                    # a la fecha de referencia

meses_cancelados = df_pagos['Concepto'].apply(lambda x: separa_meses(x, as_string=True))
df_pagos.insert(column='Meses', value=meses_cancelados, loc=df_pagos.shape[1])

# Inserta la columna 'Pendiente' con la información de desde cuándo está pendiente
# el pago de cada vecino
pendientes_desde = df.apply(no_participa_desde, axis=1)
df.insert(column='Pendiente', value=pendientes_desde, loc=df.shape[1])

# Crea la cartelera virtual
print(f"Creando Cartelera Virtual '{nombre_cartelera.format(f_ref_último_día)}'...")
cartelera = ""

# Genera la cartelera virtual con los vecinos al día
cartelera_4 = cartelera_vecinos_al_día()

# Genera la cartelera con los vecinos que colaboran con la comida de los vigilantes
# (selecciona aquellos registros donde la columna 'Comida' tenga UN espacio en blanco
#  o esté vacía, en lugar de seleccionar aquellos que contengan específicamente una
#  tilde (ü))
df_comida = df.loc[list_nor(df['Comida'] == ' ', isnull(df['Comida']))]
cartelera_5 = cartelera_comida_vigilantes(df_comida)

#
# Genera el resto de las carteleras
#

# Elimina los registros con pago en el mes seleccionado
#df = df.loc[isnull(df[mes_año])]
df = df.loc[seleccionaRegistro(df['Beneficiario'], df['Categoría'], df[fecha_referencia])]

# Revisa cada pago de los vecinos y los clasifica adecuadamente
for categoría in categorías:
    # Genera el encabezado en la cartelera virtual
    if categoría   == 'Cuota completa':
        seña = 1
    elif categoría == 'Colaboración':
        seña = 2
    elif categoría == 'No participa':
        seña = 3
    else:
        seña = 0      # OCULTO, etc.

    if seña > 0:
        muestra_encabezado_IPC = aplica_IPC and (seña == 1 or (seña == 3 and muestra_saldo_no_participa))
        ajuste_por_inflación = '\n(Saldos ajustados por inflación -IPC-)' if muestra_encabezado_IPC else ''
        cartelera += encabezado.format(seña, f_ref_último_día, ajuste_por_inflación)
        if categoría   == 'Cuota completa':
            cartelera += encabezado_cuota_completa
        elif categoría == 'Colaboración':
            cartelera += encabezado_colaboración
        elif categoría == 'No participa':
            cartelera += encabezado_no_participa
        else:
            pass

        # Selecciona los vecinos de la categoría actual
        df_subset = df.loc[df['Categoría'] == categoría]
        df_subset = df_subset.loc[df_subset['Pendiente'] != '']

        # Para cada registro en la categoría actual,
        inicio_dirección = ''
        for index, row in df_subset.iterrows():
            if inicio_dirección != row['Dirección'][: 3]:
                inicio_dirección = row['Dirección'][: 3]
                cartelera += '\n'
            cartelera += ajusta_nombre_y_dirección(row['Beneficiario'] + ', ' + row['Dirección'])
            if categoría in ['Colaboración', 'Cuota completa', 'No participa']:
                cartelera += '. ' + str(row['Pendiente'])
#            if (categoría == 'Colaboración' and muestra_ultimo_pago_colaboracion):
#                cartelera += '. ' + str(row['Pendiente'])
            cartelera += '\n'
            if inicio_dirección != row['Dirección'][: 3]:
                inicio_dirección = row['Dirección'][: 3]
                cartelera += '\n'

        # Genera el pie de página
        cartelera += '\n'
        if categoría == 'Colaboración':
            pass
        elif categoría == 'Cuota completa':
            cartelera += pie_cuota_completa
        elif categoría == 'No participa':
            cartelera += pie_no_participa
        else:
            pass

        if seña in [1, 2]:
            cartelera += ' '.join([
                'Si usted',
                'paga la totalidad de la cuota o más,' if seña == 1 else 'colabora con un monto inferior a la cuota',
                'y no aparece en este listado, significa que se encuentra solvente a la',
                'fecha de este reporte.\n\n'
            ])

        cartelera += pie_de_página.format(hoy, am_pm)
        
        # Genera una linea de separación entre categorías
        cartelera += '\n' + '-' * 25 + '\n\n'

# Imprime las carteleras de pagos al 100% y colaboración de comidas
cartelera += cartelera_4
cartelera += cartelera_5

# Graba los archivos de análisis (encoding para Windows y para macOS X)
filename = os.path.join(attach_path, 'Apple', nombre_cartelera.format(f_ref_último_día))
with open(filename, 'w', encoding=GyG_constantes.Apple_encoding) as output:
    output.write(cartelera)

filename = os.path.join(attach_path, 'Windows', nombre_cartelera.format(f_ref_último_día))
with open(filename, 'w', encoding=GyG_constantes.Windows_encoding) as output:
    output.write(cartelera)

print()

