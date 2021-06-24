# GyG Aporte para Vigilantes
#
# Genera un archivo con el resumen de pagos en el mes, indicando aquellos inferiores,
# iguales o superiores a la cuota


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.2 GyG Aporte para Vigilantes"
echo "------------------------------"

python3 ./aporte_para_vigilantes.py

echo ""
echo "Proceso terminado . . . "
