# GyG Reconversión Monetaria
#
# Toma el archivo actual de pagos y genera uno nuevo al cual se le ha
# aplicado el factor de reconversión especificado


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "5.1 GyG Reconversión Monetaria"
echo "------------------------------"

python3 ./reconversion_monetaria_GUI.py

echo ""
echo "Proceso terminado . . . "
