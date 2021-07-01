# GyG Recibos - Copia Recibos de Pago
#
# Copia los Recibos de Pago indicados para facilitar su envío a través de
# WhatsApp
#
# Un parámetro opcional permite indicar la ruta en la cual serán almacenadas
# las imágenes resultantes.
#    Ejemplo: "python copia_recibos.py Imágenes" copia los Recibos de Pago
#              en la carpeta "Imágenes" en el directorio actual.


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "6.3 GyG Recibos - Copia Recibos de Pago"
echo "---------------------------------------"

python3 ./copia_recibos.py

echo ""
echo "Proceso terminado . . . "
echo ""
