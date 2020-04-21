# GyG Mantenimiento
#
# Borra los archivos anteriores a la fecha indicada, los cuales pueden ser
# reproducidos nuevamente, a fin de mantener espacio libre en disco


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "7.3 GyG Mantenimiento de espacio en disco"
echo "-----------------------------------------"

python3 ./mantenimiento_GUI.py

echo ""
echo "Proceso terminado . . . "
