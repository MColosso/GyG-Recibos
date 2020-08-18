# GyG VALIDA CODIGO DE SEGURIDAD

# Verifica si el código de seguridad impreso en el Recibo de Pago corresponde
# con los datos del mismo

"""
    POR HACER
    -   

    HISTORICO
    -   Versión inicial (18/08/2020)
    
"""

import GyG_constantes
from GyG_utilitarios import input_fecha, valida_codigo_seguridad, _unobscure
from datetime import datetime


#
# RUTINAS DE UTILIDAD
#

def input_entero(mensaje):
    while True:
        valor_actual = input(f"*** {mensaje}: ")
        if len(valor_actual) == 0:
            valor_actual = None
            break
        if valor_actual.isnumeric():
            valor_actual = int(valor_actual)
            break
        else:
            print("  > Ingrese un valor válido")
    return valor_actual

def input_código(mensaje):
    while True:
        valor_actual = input(f"*** {mensaje}: ")
        if len(valor_actual) > 0:
            break
        print("  > Ingrese el código a validar")
    return valor_actual


#
# PROCESO
#

print()

while True:
    recibo  = input_entero("Indique el número de recibo a validar (en blanco para terminar)")
    if recibo is None:
        break
    fecha   = input_fecha("Fecha de emisión")
    código  = input_código("Código de validación")
    print()

    # if valida_codigo_seguridad(recibo, fecha, código):
    #     print('OK. Recibo válido.')
    # else:
    #     print('El código de seguridad NO corresponde con los datos del recibo.')

    try:
        decodificado = str(_unobscure(bytes(código, 'UTF-8')), 'UTF-8')
    except:
        print('El código de validación es inválido.\n')
        continue
    if valida_codigo_seguridad(recibo, fecha, código):
        print('OK. Recibo válido.')
    else:
        s_recibo, s_fecha = decodificado.split(' ')
        s_fecha = datetime.strptime(s_fecha, '%d/%m/%Y')
        print(f"Recibo inválido: El código de validación corresponde al " + \
              f"recibo {s_recibo} del {s_fecha.strftime('%d de %B de %Y')}")
    print()

print()
