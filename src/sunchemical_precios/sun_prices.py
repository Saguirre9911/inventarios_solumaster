import pandas as pd

from sunchemical_comp.sun_comp import Comparador


def leer_informe_costos():
    costos_path = "docs/sun_chemical_costos_dolar/PRECIOS_SISTEMA_SUNCHEMICAL.xlsx"
    sheet_name = input("Ingrese el nombre de la hoja del archivo de costos (DEBE SER EN EL FORMATO MES AÑO EJ DIC 2024): ")
    costos = pd.read_excel(costos_path, sheet_name,skiprows=1)
    columns_to_keep = ["Material", "P.COP"]
    costos = costos[columns_to_keep]
    costos["Material"] = costos["Material"].astype(str)
    return costos

if __name__ == "__main__":
    # Ejemplo de uso
    comparador = Comparador()
    corte, mes = comparador.leer_opcion()
    comparador.ruta_zip = f"docs/sun_chemical_xml/{corte}"
    print(comparador.ruta_zip)
    
    df_agrupado, exact_descriptions = comparador.procesar()
    print(df_agrupado.head())
    df_costos = leer_informe_costos()
    print(df_costos.head())
    print(df_costos.columns)
    df_precios = df_agrupado.join(df_costos.set_index("Material"), on="id", how="left")
    df_precios["precio"] = df_precios["precio"].str.replace(",", "").astype(float)
    df_precios["diff"] = df_precios["precio"] - df_precios["P.COP"]
    # df_precios["diff"] = df_precios["diff"].abs()
    intervalo = input("Ingrese el intervalo de precios a buscar: ")
    df_precios = df_precios[df_precios["diff"] > int(intervalo)]
    print(df_precios.head())
    df_precios.to_excel("src/sunchemical_precios/diferencias_precios.xlsx", index=False)
