import pandas as pd
import glob
import os
from datetime import datetime
# Solo NO haz skiprows de 2 en el primer archivo
excel_files = glob.glob(os.path.join("./docs_extractos/", "*.xlsx"))
df_list = [pd.read_excel(f) for f in excel_files]
df = pd.concat(df_list, ignore_index=True)

# Asegurarse de que la columna "Fecha" esté en formato datetime y quitar la hora
df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce').dt.date
df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
df['Valor'] = df['Valor'].apply(lambda x: f"${x:,.2f}")


# Ordenar el DataFrame por la columna "Fecha"
df = df.sort_values(by='Fecha', ascending=True)

print("DataFrame concatenado y ordenado por 'Fecha':")
print(df)

# Obtener la fecha actual en formato ddmmyyyy
fecha_actual = datetime.now().strftime("%d%m%Y")

# Crear el nombre del archivo con la fecha
output_file = f"./movimientos_bancolombia_{fecha_actual}.xlsx"

# Guardar el DataFrame en el archivo
df.to_excel(output_file, index=False)
print(f"DataFrame guardado en {output_file}")

# Eliminar todos los archivos Excel excepto el archivo 'movimientos_bancolombia.xlsx'
excel_files = glob.glob(os.path.join("./docs_extractos/", "*.xlsx"))
file_to_keep = output_file  # Este es el archivo que no se eliminará

files_to_delete = [f for f in excel_files if f != file_to_keep]
print(file_to_keep)
print(files_to_delete)
# Eliminar los archivos que no sean 'movimientos_bancolombia.xlsx'
for file in files_to_delete:
    os.remove(file)
    print(f"Archivo eliminado: {file}")

print("Eliminación de archivos completada.")
