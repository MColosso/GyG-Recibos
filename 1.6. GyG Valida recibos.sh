# GyG Recibos - Valida recibos
#
# Valida que haya correspondencia entre el número de recibo y fecha y el
# código de validación impreso en el recibo de pago


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "1.9 GyG Recibos - Valida recibos"
echo "--------------------------------"

python3 ./valida_recibos_de_pago.py

echo ""
echo "Proceso terminado . . . "
echo ""
