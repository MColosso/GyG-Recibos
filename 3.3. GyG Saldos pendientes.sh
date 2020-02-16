# GyG Saldos pendientes
#
# Genera los totales de cuotas pendientes por cada vecino en el sector
#
# La opción '--solo_deudores' permite filtrar sólo a aquellos vecinos con un
# saldo pendiente
#    Ejemplo: "python saldos_pendientes.py --solo_deudores"


cd $HOME/Dropbox/"GyG Recibos"/ > nul
echo ""
echo "3.3 GyG Saldos pendientes"
echo "-------------------------"

python3 ./saldos_pendientes.py

echo ""
echo "Proceso terminado . . . "
