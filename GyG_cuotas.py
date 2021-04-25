# GyG CUOTAS
#
# Módulo para el manejo de cuotas:
#   - cuota_vigente()       cuota a una fecha determinada
#   - cuota_actual()        cuota vigente si la fecha es anterior al 1° Septiembre 2019, o la última
#                           cuota registrada
#   - tasa_actual()         tasa a ser utilizada para el cálculo de la cuota
#   - resumen_de_cuotas()   resumen de las cuotas establecidas para los últimos meses

"""
    POR HACER
    -   

    HISTORICO
    -   Se agregó el parámetro 'aplica_IPC' a las rutinas 'cuota_vigente()' y 'cuota_actual()'
        para aplicar o no el ajuste por inflación en el cálculo de los saldos pendientes
        (26/01/2021)
    -   Corregir: En el resumen de cuotas, no se muestra el año final cuando es diferente al
        año inicial (23/02/2020)
    -   Corregir: Se muestra el monto de la cuota en lugar de la tasa en el párrafo iniciado
        con "A partir del 1° de Septiembre 2019[...]" (21/10/2019)
    -   Ajustar el resumen_de_cuotas() para mostrar las cuotas en VEB o USD (15/10/2019)
    -   Ajustar el resumen_de_cuotas() para emitirlo en formato PDF (11/10/2019)
    -   

"""

import GyG_constantes
from pandas import read_excel, notnull
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import locale
dummy = locale.setlocale(locale.LC_ALL, 'es_es')

# Constantes
fmtText = 0     # Resumen de cuotas en formato texto
fmtHtml = 1     #    "    "    "    "     "    html
fmtPdf  = 2     #    "    "    "    "     "    pdf

_nMeses = 5

class Cuota:
    def __init__(self, workbook=os.path.join(GyG_constantes.ruta_gyg_recibos, GyG_constantes.pagos_wb_estandar),
                       worksheet=GyG_constantes.pagos_ws_cuotas):
        self.workbook = workbook
        self.worksheet = worksheet
        try:
            self.df_cuotas = read_excel(workbook, sheet_name=worksheet, skiprows=[0, 1])
        except:
            raise 

        self.df_cuotas.dropna(inplace=True, subset=['CUOTA'])
        self.df_cuotas = self.df_cuotas[self.df_cuotas['Fecha'] <= datetime.today()]
        columnas = [str.replace('CUOTA ', '') for str in self.df_cuotas.columns.to_list()]
        self.df_cuotas.columns = columnas

    def cuota_vigente(self, beneficiario: str, fecha: datetime, aplica_IPC: bool=False):
        """
        Cuota vigente para el vecino indicado en la fecha establecida.
        Empleado para determinar si la cuota fue totalmente cancelada para la fecha del
        último pago.
        """
        column_name = beneficiario if beneficiario in self.df_cuotas.columns.to_list() else 'CUOTA'
        if (column_name == 'CUOTA') and aplica_IPC:
            column_name = 'Reexpr. por inflación'
        return self.df_cuotas[self.df_cuotas['Fecha'] <= fecha].iloc[-1][column_name]

    def cuota_actual(self, beneficiario: str, fecha: datetime, aplica_IPC: bool=False):
        """
        Cuota a ser utilizada para el cálculo del saldo deudor.
        A partir del 1° de Septiembre 2019, se tomará como cuota actual la última registrada.
        Si aplica_IPC es verdadero, los valores de las cuotas se reexpresarán en base al Indice de Precios
        al Consumidor (df_cuotas['Reexpr. por inflación'])
        """
        return self.cuota_vigente(beneficiario, fecha if fecha < GyG_constantes.fecha_de_corte else datetime.today(),
                                  aplica_IPC=aplica_IPC)

    def tasa_actual(self, fecha: datetime, redondeada: bool=True):
        """
        Tasa a ser utilizada para el cálculo de la cuota.
        """
        return self.df_cuotas[self.df_cuotas['Fecha'] <= fecha].iloc[-1]['Tasa redondeada' if redondeada else 'Tasa Bs./US$']


    def resumen_de_cuotas(self, beneficiario: str, fecha_final, 
                                fecha_inicial=None, formato: str=fmtText, lista: bool=False):
        """
        Retorna el resumen de las últimas <n-meses> -5- cuotas mensuales, hasta el mes de
        Agosto 2019, y el mensaje indicativo de los cambios vigentes a partir del 1° de
        Septiembre y la última cuota registrada.
        El texto puede ser de forma de texto lineal (por defecto) o en forma de lista, con
        atributos HTML o no (por defecto)
        """

        def spaces(n: int) -> str:
            return (' ' if formato == fmtText else '&nbsp;') * n

        def rango_fechas(idx_inicial, idx_final, cuota, año_anterior, rango_final=False, más_de_un_elemento=False):
            global fechas, cuotas

    #        final = ', sujeto a revisiones mensuales.' if rango_final and not lista else ''
            inicio_linea = '<li>'  if formato == fmtHtml and lista else spaces(5) + '•' + spaces(2) if lista else ''
            final_linea  = '</li>' if formato == fmtHtml and lista else EOL if lista else ''
            final = '.' if rango_final and not lista else ''
            texto = ', y ' if rango_final and not lista and más_de_un_elemento else ''
#            año_final = '' if fechas[idx_inicial].year == año_anterior else ' %Y'
            año_final = '' if fechas[idx_final].year == año_anterior else ' %Y'
            if idx_inicial == idx_final:
                texto += f"{inicio_linea}" + \
                         f"{cuota}" + \
                         f" en {fechas[idx_inicial]:%B{año_final}}{final}" + \
                         f"{final_linea}"
            else:
                a_y = 'y' if idx_final - idx_inicial == 1 else 'a'
                en_de = 'de' if idx_final - idx_inicial != 1 else 'en'
                año_inicio = '' if fechas[idx_inicial].year == fechas[idx_final].year else ' %Y'
                texto += f"{inicio_linea}" + \
                         f"{cuota}" + \
                         f" {en_de} {fechas[idx_inicial]:%B{año_inicio}}" + \
                         f" {a_y} {fechas[idx_final]:%B{año_final}}{final}" + \
                         f"{final_linea}"
            texto += ", " if not rango_final and not lista else ""

            return texto

        def genera_resumen():
            global fechas, cuotas, EOL

            column_name = beneficiario if beneficiario in self.df_cuotas.columns.to_list() else 'CUOTA'
            monto_usd   = 'USD '+beneficiario if beneficiario in self.df_cuotas.columns.to_list() else 'Cantidad'
            df_subset = self.df_cuotas.copy()
            df_subset = df_subset[(df_subset['Fecha'] >= dt_inicial) & \
                                  (df_subset['Fecha'] <= dt_final)]
            mes_anterior = None
            fechas = list()
            cuotas = list()
            for index, r in df_subset.iterrows():
                if datetime(r['Fecha'].year, r['Fecha'].month, 1) == mes_anterior:
                    pass
                else:
                    mes_anterior = datetime(r['Fecha'].year, r['Fecha'].month, 1)
                    fechas.append(mes_anterior)
                    if r['Moneda'] == 'VEB':
                        cuotas.append(f"Bs. {locale.format_string(f'%.2f', r[column_name], grouping=True, monetary=True).replace(',00', '')}")
                    else:
                        cuotas.append(f"US$ {locale.format_string(f'%.2f', r[monto_usd], grouping=True, monetary=True).replace(',00', '')}")

            txtCuotasMensuales = ''

            EOL = '<br/>' if formato in [fmtHtml, fmtPdf] else '\n'
            UL = '<UL>' if formato == fmtHtml and lista else EOL if lista else ' '
            END_UL = '</UL>' if formato == fmtHtml and lista else EOL if lista else ''

            if len(fechas) > 0:
                idx_inicial = 0
                if fechas[idx_inicial] < GyG_constantes.fecha_de_corte:
                    cuota = cuotas[idx_inicial]
                    last_notnull = idx_inicial
                    año_anterior = 0
                    txtCuotasMensuales = f'Las cuotas para el pago de la vigilancia en los últimos ' + \
                                         f'meses han sido las siguientes:{UL}'
                    más_de_un_elemento = False
                    for idx in range(1, len(fechas)):
                        if notnull(cuotas[idx]):
                            last_notnull = idx
                            if cuotas[idx] != cuota:
                                txtCuotasMensuales += rango_fechas(idx_inicial, idx - 1, cuota, año_anterior, rango_final=False)
                                más_de_un_elemento = True
                                idx_inicial = idx
                                cuota = cuotas[idx_inicial]
                                año_anterior = fechas[idx - 1].year
                    txtCuotasMensuales = (txtCuotasMensuales[:-2] if más_de_un_elemento and not lista else txtCuotasMensuales) + \
                                         rango_fechas(idx_inicial, last_notnull, cuota, año_anterior, 
                                                      rango_final=True, más_de_un_elemento=más_de_un_elemento) + \
                                         END_UL

            if dt_final >= GyG_constantes.fecha_de_corte:
                tasa = self.tasa_actual(datetime.today())
                str_tasa = locale.format_string(f'%.2f', tasa, grouping=True, monetary=True).replace(',00', '')
                if (len(fechas) > 0) and (fechas[0] >= GyG_constantes.fecha_de_corte):
                    separador = ""
                else:
                    separador = "" if (len(fechas) == 0) or lista else EOL * 2
                cuota_a_mostrar = self.cuota_vigente(beneficiario, datetime.today())
                cuota_a_mostrar = locale.format_string(f'%.2f', cuota_a_mostrar, grouping=True, monetary=True).replace(',00', '')
                txtCuotasMensuales += separador + \
                              "Las cuotas pendientes por pagar hasta Agosto 2019 se cancelarán según los montos fijos " + \
                              "ya establecidos. A partir del 1° de Septiembre de dicho año, toda cuota atrasada " + \
                              "se cancelará en base a la cuota vigente " + \
                             f"(Bs. {cuota_a_mostrar} al {datetime.today().strftime('%d/%m/%Y')})"
                             #  "A partir del 1° de Septiembre 2019, las cuotas mensuales han sido fijadas " + \
                             #  "en dólares, pagaderas en divisas o en bolívares a la tasa de cambio semanal publicada " + \
                             # f"por el Banco Central. A la fecha, la misma es de Bs. {str_tasa} por dólar." + EOL + \

                             #  "Las cuotas pendientes por cancelar hasta Agosto 2019 quedarán en los montos fijos " + \
                             #  "ya establecidos. A partir del mes de Septiembre, toda cuota atrasada se cancelará " + \
                             #  "con la tasa de la semana en curso."

            txtCuotasMensuales += EOL * 2

            return txtCuotasMensuales

        dt_inicial = fecha_inicial if fecha_inicial != None else fecha_final - relativedelta(months=_nMeses - 1)
        dt_final = fecha_final

        return genera_resumen()
