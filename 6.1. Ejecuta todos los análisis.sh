# 6.1 Ejecuta todos los Análisis
#
# Ejecuta el conjunto de análisis a disposición: Cartelera Virtual,
# Análisis de Pagos, Saldos Pendientes y Resumen de Saldos usando
# sus opciones por defecto.


# GyG Cartelera Virtual --------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.1 GyG Cartelera Virtual"
echo "-------------------------"

python3 ./cartelera_virtual.py --toma_opciones_por_defecto


# GyG Analisis de Pagos --------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.1 GyG Analisis de Pagos"
echo "-------------------------"

python3 ./analisis_de_pagos.py --toma_opciones_por_defecto


# GyG Estadística de Pagos --------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.2 GyG Estadística de Pagos"
echo "----------------------------"

python3 ./estadistica_pagos.py --toma_opciones_por_defecto


# GyG Saldos pendientes --------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.3 GyG Saldos pendientes"
echo "-------------------------"

python3 ./saldos_pendientes.py --toma_opciones_por_defecto


# GyG Resumen de Saldos a la fecha ---------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.4 GyG Resumen de Saldos a la fecha"
echo "------------------------------------"

python3 ./resumen_saldos.py --toma_opciones_por_defecto


# GyG Propuestas de cambio de categoría ----------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "7.4 GyG Propuestas de cambio de categoría"
echo "-----------------------------------------"

python3 ./cambios_de_categoria.py --toma_opciones_por_defecto

echo ""
echo "Proceso terminado . . . "
