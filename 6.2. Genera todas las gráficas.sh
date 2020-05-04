# 6.2 Genera todas las gráficas
#
# Genera las gráficas de Gestión de Cobranzas y Pagos 100% equivalentes
# a la fecha


# GyG Gráficas -----------------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "7.1 GyG Gráficas"
echo "----------------"

python3 ./graficas_GUI.py --toma_opciones_por_defecto

echo ""
echo "Proceso terminado . . . "
