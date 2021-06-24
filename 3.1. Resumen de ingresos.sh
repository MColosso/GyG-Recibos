# GyG Gráficas
#
# Genera las gráficas de Gestión de Cobranzas y Pagos 100% equivalentes
# en base a los valores de la pestaña 'Cobranzas' de la hoja de cálculo
# '1.1. GyG Recibos.xlsm'


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "7.2 GyG Resumen de ingresos"
echo "---------------------------"

python3 ./resumen_de_ingresos_GUI.py

echo ""
echo "Proceso terminado . . . "
