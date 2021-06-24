# GyG Analisis de Pagos
#
# Determina los pagos no realizados en el mes actual a fin de contactar
# a sus beneficiarios


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.1 GyG Analisis de Pagos"
echo "-------------------------"

python3 ./analisis_de_pagos.py

echo ""
echo "Proceso terminado . . . "
