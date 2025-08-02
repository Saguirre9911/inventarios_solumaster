import pandas as pd

# Dirección del archivo en el computador
archivo_inventarios = "Inventario_Solumaster_enviar.xlsx"
archivo_lotes = "docs/Excel_enviar.xlsx"
# lectura del archivo y creación del dataframe
df_inventarios = pd.read_excel(
    archivo_inventarios, sheet_name="Sheet1"
)  # ,skipfooter=1)
df_lotes = pd.read_excel(archivo_lotes, sheet_name="Excel_enviar")
print(df_inventarios)

df_lotes_sin_duplicados = df_lotes.drop_duplicates("N° de ítem        ")
rango_for = df_lotes_sin_duplicados["N° de ítem        "]
suma_lote = 0
df_inventarios["Total_lotes"] = 0

df_lotes["Total disponible    "] = (
    df_lotes["Total disponible    "]
      .astype(str)                     # Garantiza str
      .str.strip()                     # Quita espacios
      .str.replace(',', '', regex=False)  # Elimina comas
      .astype(float)                   # Convierte a float
)

for x in rango_for:
    dataframe_pequeño = df_lotes[df_lotes["N° de ítem        "] == x]
    suma_lote = 0
    import pandas as pd




    # print(dataframe_pequeño)
    if len(dataframe_pequeño) > 1:
        for y in range(len(dataframe_pequeño)):
            cantidad_en_lote = dataframe_pequeño.iloc[y, 16]

            cantidad_en_lote=float(cantidad_en_lote)
            print("esto es sumalote",suma_lote)
            print("esto es cantidad_en_lote",cantidad_en_lote)
            suma_lote = suma_lote + cantidad_en_lote
            suma_lote = float(suma_lote)
            # print(dataframe_pequeño['Total disponible    '])
            # print(suma_lote)
    else:
        # print(dataframe_pequeño['Total disponible    '])
        suma_lote = float(dataframe_pequeño["Total disponible    "])

        # print(suma_lote/1000)

    df_inventarios["Total_lotes"][df_inventarios["N° de ítem"] == x] = suma_lote
    # print(df_inventarios)

df_inventarios["Cant.uso"] = df_inventarios["Cant.uso"].replace(
    "                    ", 0
)

df_inventarios["Diferencia"] = (
    df_inventarios["Total_lotes"] - df_inventarios["Cant.uso"]
)
df_inventarios = df_inventarios.iloc[:, 1:]  # Elimina la primera columna
df_inventarios.set_index(
    "N° de ítem", inplace=True
)  # Establece el índice como "N° de ítem"
print("df inventarios:", df_inventarios)


with pd.ExcelWriter("Inventario_Solumaster_papa.xlsx") as writer:
    df_inventarios.to_excel(writer)
# No need to call writer.save()


# HACER UN FOR POR CADA UNO DE LOS CODIGOS Y DENTRO DE ESE FOR SUMAR CADA UNO DE LOS REGISTROS CON UNA TABLA INDIVIDUAL
