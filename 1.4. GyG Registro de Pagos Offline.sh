# GyG Recibos - Registro de Pagos Off-Line
#
# Actualiza la hoja de cálculo con el registro de nuevos pagos; especialmente útil en
# ocasiones en las cuales no se dispone de conexión a Internet
#
# Para reemplazar la hoja de cálculo original con la versión off-line, primeramente
# se deba recalcular la hoja off-line con Excel.


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "1.4 GyG Recibos - Registro de Pagos Off-Line"
echo "---------------------------------------------"

python3 ./registro_de_pagos_offline.py

echo ""
echo "Proceso terminado . . . "
echo ""
