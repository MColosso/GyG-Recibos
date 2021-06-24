# 7.1 Ejecuta todos los Análisis
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


# GyG Saldos pendientes --------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.2 GyG Saldos pendientes"
echo "-------------------------"

python3 ./saldos_pendientes.py --toma_opciones_por_defecto


# GyG Resumen de Saldos a la fecha ---------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.3 GyG Resumen de Saldos a la fecha"
echo "------------------------------------"

python3 ./resumen_saldos.py --toma_opciones_por_defecto


# GyG Distribución de Pagos ----------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "2.5 GyG Distribución de Pagos"
echo "-----------------------------"

python3 ./distribucion_pagos.py --toma_opciones_por_defecto


# GyG Resumen de Ingresos ------------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.1 GyG Resumen de Ingresos"
echo "---------------------------"

python3 ./resumen_de_ingresos_GUI.py --toma_opciones_por_defecto


# GyG Aporte para Vigilantes ---------------------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "3.2 GyG Aporte para Vigilantes"
echo "------------------------------"

python3 ./aporte_para_vigilantes.py --toma_opciones_por_defecto


# GyG Propuestas de cambio de categoría ----------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "5.3 GyG Propuestas de cambio de categoría"
echo "-----------------------------------------"

python3 ./cambios_de_categoria.py --toma_opciones_por_defecto


# GyG Mantenimiento de espacio en disco ----------------------------

cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
echo ""
echo "5.2 GyG Mantenimiento de espacio en disco"
echo "-----------------------------------------"

python3 ./mantenimiento_GUI.py --toma_opciones_por_defecto


# # GyG Analisis de Pagos --------------------------------------------

# cd $HOME/Dropbox/"GyG Recibos"/ > /dev/null
# echo ""
# echo "6.4 GyG Analisis de Pagos"
# echo "-------------------------"

# python3 ./analisis_de_pagos.py --toma_opciones_por_defecto


echo ""
echo "Proceso terminado . . . "
