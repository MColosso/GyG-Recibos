# GyG Recibos - Recibos de Pago como imágenes
#
# Genera los Recibos de Pago pendientes por imprimir, como archivos con
# formato .png
#
# Un parámetro opcional permite indicar la ruta en la cual serán almacenados
# los Recibos de Pago generados.
#    Ejemplo: "python genera_recibos.py Imágenes" genera los Recibos de Pago
#              en la carpeta "Imágenes" en el directorio actual.


cd $HOME/Dropbox/"GyG Recibos"/ > nul
echo ""
echo "1.2 GyG Recibos - Genera recibos"
echo "--------------------------------"

python3 ./genera_recibos.py

echo ""
echo "Proceso terminado . . . "
