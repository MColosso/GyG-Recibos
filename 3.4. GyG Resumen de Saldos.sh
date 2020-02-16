# GyG Resumen de Saldos a la fecha
#
# Genera los totales de cuotas pendientes por cada vecino en el sector
#


cd $HOME/Dropbox/"GyG Recibos"/ > nul
echo ""
echo "3.4 GyG Resumen de Saldos a la fecha"
echo "------------------------------------"

python3 ./resumen_saldos.py

echo ""
echo "Proceso terminado . . . "
