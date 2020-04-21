# GyG Cartelera Virtual
#
# Genera los resúmenes de cuotas pendientes por categoría, para poder
# enviarlos posteriormente a las listas de difusión de WhatsApp


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.1 GyG Cartelera Virtual"
echo "-------------------------"

python3 ./cartelera_virtual.py

echo ""
echo "Proceso terminado . . . "
