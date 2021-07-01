:: GyG Aporte para Vigilantes
::
:: Genera un archivo con el resumen de pagos en los mes indicados para la categoría
:: seleccionada
::
:: Resultados:
::   Carpeta:  ../GyG Archivos/Análisis de Pago
::   Archivo:  GyG Pagos Adicionales {«año»-«mes» («abreviatura mes»)}.txt
::
:: Parámetros:
::   --meses=n                    -  Despliega por defecto 'n' meses
::   --mes_actual                 -  Finalizando en el mes actual
::   --toma_opciones_por_defecto  -  No interactúa solicitando parámetros


@echo 2.6 GyG Aporte para Vigilantes
@echo ------------------------------

python aporte_para_vigilantes.py  --meses=3 --mes_actual

@pause
