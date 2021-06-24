# GyG Distribución de Pagos
#
# Genera un archivo con el resumen de pagos en el mes, indicando aquellos inferiores,
# iguales o superiores a la cuota


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.5 GyG Distribución de Pagos"
echo "-----------------------------"

python3 ./distribucion_pagos.py

echo ""
echo "Proceso terminado . . . "
