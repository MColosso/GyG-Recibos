# GyG Recibos - Envía recibos
#
# Envía los recibos generados en el paso anterior a sus destinatarios vía
# correo electrónico o WhatsApp


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "1.3 GyG Recibos - Envia recibos"
echo "--------------------------------"

python3 ./envia_recibos.py

echo ""
echo "Proceso terminado . . . "
echo ""
