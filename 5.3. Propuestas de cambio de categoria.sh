# GyG Mantenimiento
#
# Borra los archivos anteriores a la fecha indicada, los cuales pueden ser
# reproducidos nuevamente, a fin de mantener espacio libre en disco


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "5.3 GyG Propuestas de cambio de categoría"
echo "-----------------------------------------"

python3 ./cambios_de_categoria.py

echo ""
echo "Proceso terminado . . . "
