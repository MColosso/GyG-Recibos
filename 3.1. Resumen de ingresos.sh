# GyG Resumen de Ingresos
#
# Genera un resumen de todos los ingresos del mes, agrupados por
# categoría para el período seleccionado


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.1 GyG Resumen de ingresos"
echo "---------------------------"

python3 ./resumen_de_ingresos_GUI.py

echo ""
echo "Proceso terminado . . . "
