# GyG - Recibos a Solicitud
#
# Genera los recibos indicados a partir de la hoja de cálculo estándar o
# de la hoja de cálculo con datos históricos antes de la reconversión.
# Ello permite que los recibos generados puedan ser borrados en cualquier
# momento y regenerados a voluntad.


cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "1.7 GyG - Genera los recibos solicitados"
echo "----------------------------------------"
echo ""
echo "Ejemplos de recibos individuales o rangos a generar:"
echo "   12     Recibo  00012"
echo "   1-5    Recibos 00001 hasta 00005"
echo "   -8     Recibos 00001 hasta 00008"
echo "   12-    Recibos 00012 hasta último recibo existente"
echo "   -5, 8, 12-14   Se pueden indicar diferentes rangos de recibos a generar"
echo "                  colocándolos en una linea y separándolos por comas"
echo ""

python3 ./genera_recibos.py --seleccion_manual

echo ""
echo "Proceso terminado . . . "
