# 6.3 Resumen de Ingresos
#
# Genera el Resumen de Ingresos de los últimos 12 meses previos al mes
# actual en formato PDF


# GyG Gráficas -----------------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "6.3 Resumen de Ingresos"
echo "-----------------------"

python3 ./resumen_de_ingresos_GUI.py --toma_opciones_por_defecto

echo ""
echo "Proceso terminado . . . "
