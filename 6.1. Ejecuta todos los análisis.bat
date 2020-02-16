:: 5.1 Ejecuta todos los Análisis
::
:: Ejecuta el conjunto de análisis a disposición: Cartelera Virtual,
:: Análisis de Pagos, Saldos Pendientes y Resumen de Saldos usando
:: sus opciones por defecto.


:: GyG Cartelera Virtual --------------------------------------------

@echo 2.1 GyG Cartelera Virtual
@echo -------------------------

python ./cartelera_virtual.py --toma_opciones_por_defecto


:: GyG Analisis de Pagos --------------------------------------------

@echo 3.1 GyG Analisis de Pagos
@echo -------------------------

python ./analisis_de_pagos.py --toma_opciones_por_defecto


:: GyG Saldos pendientes --------------------------------------------

@echo 3.3 GyG Saldos pendientes
@echo -------------------------

python ./saldos_pendientes.py --toma_opciones_por_defecto


:: GyG Resumen de Saldos a la fecha ---------------------------------

@echo 3.4 GyG Resumen de Saldos a la fecha
@echo ------------------------------------

python ./resumen_saldos.py --toma_opciones_por_defecto

@pause
