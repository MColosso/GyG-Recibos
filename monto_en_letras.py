def MontoEnLetras(número, mostrar_céntimos=True, céntimos_en_letras=False):
    #
    # Constantes
    Moneda = "Bolívar"       # Nombre de Moneda (Singular)
    Monedas = "Bolívares"    # Nombre de Moneda (Plural)
    Céntimo = "Céntimo"      # Nombre de Céntimos (Singular)
    Céntimos = "Céntimos"    # Nombre de Céntimos (Plural)
    Preposición = "con"      # Preposición entre Moneda y Céntimos
    Máximo = 1999999999.99   # Máximo valor admisible

    def _número_recursivo(número):
        UNIDADES = ("", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez",
                    "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho",
                    "Diecinueve", "Veinte", "Veintiun", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco",
                    "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve")
        DECENAS  = ("", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta",
                    "Noventa", "Cien")
        CENTENAS = ("", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos",
                    "Setecientos", "Ochocientos", "Novecientos")
#        print(f"DEBUG: _número_recursivo(número={número})")
        if número == 0:
            resultado = 'cero'
        elif número <= 29:
            resultado = UNIDADES[número]
        elif número <= 100:
            resultado = DECENAS[número // 10]
            if número % 10 != 0: resultado += ' y ' + _número_recursivo(número % 10)
        elif número <= 999:
            resultado = CENTENAS[número // 100]
            if número % 100 != 0: resultado += ' ' + _número_recursivo(número % 100)
        elif número <= 1999:
            resultado = 'Mil'
            if número % 1000 != 0: resultado += ' ' + _número_recursivo(número % 1000)
        elif número <= 999999:
            resultado = _número_recursivo(número // 1000) + ' Mil'
            if número % 1000 != 0: resultado += ' ' + _número_recursivo(número % 1000)
        elif número <= 1999999:
            resultado = 'Un Millón'
            if número % 1000000 != 0: resultado += ' ' + _número_recursivo(número % 1000000)
        elif número <= 1999999999:
#            resultado = CENTENAS(número // 1000000) + ' Millones'
            resultado = _número_recursivo(número // 1000000) + ' Millones'
            if número % 1000000 != 0: resultado += ' ' + _número_recursivo(número % 1000000)
        else:
            resultado = format(número, ',.0f').replace(',', '.')

        return resultado

    if 0 <= número <= Máximo:
        # convertir la parte entera del número en letras
        letra = _número_recursivo(int(número))

        # Agregar la descripción de la moneda
        letra += ' ' + (Moneda if int(número) == 1 else Monedas)

        # Obtener los céntimos del Numero
        num_céntimos = int(round((número - int(número)) * 100, 0))

        if mostrar_céntimos:
            if céntimos_en_letras:
                letra += ' ' + Preposición + ' ' + _número_recursivo(num_céntimos)
                letra += ' ' + (Céntimo if num_céntimos == 1 else Céntimos)
            else:
                letra += ' ' + Preposición + (' 0' if num_céntimos < 10 else ' ') + str(num_céntimos) + '/100'

        return letra
    else:
        return 'ERROR: El número excede los límites.'
