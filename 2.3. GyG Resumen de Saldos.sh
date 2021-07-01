# GyG Resumen de Saldos a la fecha
#
# Genera los totales de cuotas pendientes por cada vecino en el sector
#


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.3 GyG Resumen de Saldos a la fecha"
echo "------------------------------------"

python3 ./resumen_saldos.py

echo ""
echo "Proceso terminado . . . "
