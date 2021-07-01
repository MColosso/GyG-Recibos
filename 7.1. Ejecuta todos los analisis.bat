:: 7.1 Ejecuta todos los Análisis
::
:: Ejecuta el conjunto de análisis a disposición: Cartelera Virtual,
:: Análisis de Pagos, Saldos Pendientes y Resumen de Saldos usando
:: sus opciones por defecto.


:: GyG Cartelera Virtual --------------------------------------------

@echo 2.1 GyG Cartelera Virtual
@echo -------------------------

python ./cartelera_virtual.py --toma_opciones_por_defecto


:: GyG Saldos pendientes --------------------------------------------

@echo 2.2 GyG Saldos pendientes
@echo -------------------------

python ./saldos_pendientes.py --toma_opciones_por_defecto


:: GyG Resumen de Saldos a la fecha ---------------------------------

@echo 2.3 GyG Resumen de Saldos a la fecha
@echo ------------------------------------

python ./resumen_saldos.py --toma_opciones_por_defecto


:: GyG Distribución de Pagos ----------------------------------------

@echo 2.5 GyG Distribución de Pagos
@echo -----------------------------

python ./distribucion_pagos.py --toma_opciones_por_defecto


:: GyG Resumen de Ingresos ------------------------------------------

@echo 3.1 GyG Resumen de Ingresos
@echo ---------------------------

python ./resumen_de_ingresos_GUI.py --toma_opciones_por_defecto


:: GyG Aporte para Vigilantes ---------------------------------------

@echo 3.2 GyG Aporte para Vigilantes
@echo ------------------------------

python ./aporte_para_vigilantes.py --toma_opciones_por_defecto


:: GyG Propuestas de cambio de categoría ----------------------------

@echo 5.3 GyG Propuestas de cambio de categoría
@echo -----------------------------------------

python ./cambios_de_categoria.py --toma_opciones_por_defecto


:: GyG Mantenimiento de espacio en disco ----------------------------

@echo 5.2 GyG Mantenimiento de espacio en disco
@echo -----------------------------------------

python ./mantenimiento_GUI.py --toma_opciones_por_defecto

@pause
