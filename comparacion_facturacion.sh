#!/bin/bash


# # Establecer el directorio actual como punto de partida
# BASE_DIR=$(pwd)

# # Exportar PYTHONPATH para incluir el directorio src
# export PYTHONPATH="$BASE_DIR/src"
PYTHONPATH=$(pwd)/src python src/sunchemical_precios/sun_prices.py


# Activar el entorno virtual
source ./solumaster_venv/bin/activate
# Ejecutar el script edit_excel.py
python src/sunchemical_comp/sun_comp.py

# # Ejecutar el script comparacion de precios
# python src/sunchemical_precios/sun_prices.py
