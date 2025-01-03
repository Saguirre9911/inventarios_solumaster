import os
import xml.etree.ElementTree as ET
import zipfile
from io import TextIOWrapper

import pandas as pd


class Comparador:
    def __init__(self):
        
        self.namespaces = {
            "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
            "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
            "ext": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
            "sts": "dian:gov:co:facturaelectronica:Structures-2-1",
            "ds": "http://www.w3.org/2000/09/xmldsig#",
        }
        self.data = []
        self.df = pd.DataFrame()
        self.df_agrupado = pd.DataFrame()
        self.exact_descriptions = pd.DataFrame()
        self.inventario = pd.DataFrame()
        self.ruta_zip = ""
    def procesar(self):
        for file_name in os.listdir(self.ruta_zip):
            if file_name.endswith(".zip"):
                zip_path = os.path.join(self.ruta_zip, file_name)

                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    for member in zip_ref.namelist():
                        if member.lower().endswith(".xml"):
                            with zip_ref.open(member, "r") as xml_file:
                                with TextIOWrapper(xml_file, encoding="utf-8") as text_wrapper:
                                    contenido = text_wrapper.read()

                                    # Parsear el XML principal
                                    root = ET.fromstring(contenido)

                                    # Buscar el cbc:Description con el CDATA interno
                                    description_elem = root.find(".//cbc:Description", namespaces=self.namespaces)
                                    if description_elem is not None and description_elem.text is not None:
                                        cdata_content = description_elem.text.strip()

                                        # Parsear el XML embebido en el CDATA
                                        try:
                                            embedded_root = ET.fromstring(cdata_content)
                                        except ET.ParseError:
                                            # Si no se puede parsear, continuar con el siguiente archivo
                                            continue

                                        # Buscar las cac:InvoiceLine en el XML embebido
                                        for invoice_line in embedded_root.findall(".//cac:InvoiceLine", namespaces=self.namespaces):
                                            item = invoice_line.find("cac:Item", namespaces=self.namespaces)
                                            if item is None:
                                                continue

                                            descripcion_elem = item.find("cbc:Description", namespaces=self.namespaces)
                                            descripcion = descripcion_elem.text if descripcion_elem is not None else None

                                            sellers_id_elem = item.find("cac:SellersItemIdentification/cbc:ID", namespaces=self.namespaces)
                                            producto_id = sellers_id_elem.text if sellers_id_elem is not None else None

                                            price = invoice_line.find("cac:Price", namespaces=self.namespaces)
                                            if price is None:
                                                continue

                                            price_amount_elem = price.find("cbc:PriceAmount", namespaces=self.namespaces)
                                            precio = price_amount_elem.text if price_amount_elem is not None else None

                                            base_quantity_elem = price.find("cbc:BaseQuantity", namespaces=self.namespaces)
                                            cantidad = base_quantity_elem.text if base_quantity_elem is not None else None
                                            unitcode = base_quantity_elem.get("unitCode") if base_quantity_elem is not None else None

                                            self.data.append({
                                                "id": producto_id,
                                                "descripcion": descripcion,
                                                "precio": precio,
                                                "unitcode": unitcode,
                                                "cantidad": cantidad,
                                            })

        # Crear el DataFrame con la información recopilada
        self.df = pd.DataFrame(self.data)

        # Convertir la columna 'cantidad' a numérica
        self.df["cantidad"] = pd.to_numeric(self.df["cantidad"], errors="coerce")

        # Agrupar por las columnas clave y sumar la cantidad
        self.df_agrupado = self.df.groupby(["id", "descripcion", "precio", "unitcode"], as_index=False).agg({"cantidad": "sum"})

        # Filtrar las descripciones que contengan "PSO"
        self.exact_descriptions = self.df_agrupado[self.df_agrupado["descripcion"].str.contains("PSO", na=False)]

        return self.df_agrupado, self.exact_descriptions

    def leer_informe_inventario(self):
        inventario_path = "Inventario_Solumaster_papa.xlsx"
        self.inventario = pd.read_excel(inventario_path)
        return self.inventario
    
    def leer_opcion(self):
        # Primer input: corte del mes (1 o 2)
        while True:
            opcion_corte = input("Escoja el corte del mes (1 o 2): ")
            if opcion_corte in ["1", "2"]:
                break
            else:
                print("Entrada inválida. Por favor, ingrese solo '1' o '2'.")

        # Segundo input: mes (1 al 12, por ejemplo)
        while True:
            opcion_mes = input("Escoja el número del mes (1 al 12): ")
            # Verificamos que sea un dígito y que esté dentro del rango permitido
            if opcion_mes.isdigit() and 1 <= int(opcion_mes) <= 12:
                break
            else:
                print("Entrada inválida. Por favor, ingrese un número entre 1 y 12.")

        return opcion_corte, opcion_mes


if __name__ == "__main__":
    # Ejemplo de uso
    comparador = Comparador()
    corte, mes = comparador.leer_opcion()
    comparador.ruta_zip = f"docs/sun_chemical_xml/{corte}"
    print(comparador.ruta_zip)
    
    df_agrupado, exact_descriptions = comparador.procesar()
    inventario = comparador.leer_informe_inventario()

    # Convertir la columna 'id' de df_agrupado a string
    df_agrupado['id'] = df_agrupado['id'].astype(str)
    print("antes de drop", inventario.shape)
    # Eliminar filas con NaN en la columna 'N° de ítem'
    inventario = inventario.dropna(subset=['N° de ítem'])
    print("despues de drop", inventario.shape)
    # Convertir 'N° de ítem' a entero y luego a string
    inventario['N° de ítem'] = inventario['N° de ítem'].astype(int).astype(str)

    # Asegurar que 'Diferencia' en inventario sea numérica
    inventario['Diferencia'] = pd.to_numeric(inventario['Diferencia'], errors='coerce')

    # Realizar el merge
    df_combinado = pd.merge(df_agrupado, inventario, left_on='id', right_on='N° de ítem', how='left')

    # Asegurarnos de que ambas columnas usadas para el cálculo sean float
    df_combinado['cantidad'] = df_combinado['cantidad'].astype(float)
    df_combinado['Diferencia'] = df_combinado['Diferencia'].astype(float)

    # Calcular la nueva columna 'diferencia'
    df_combinado['diferencia_facturacion'] = df_combinado['Diferencia'] - df_combinado['cantidad']

    print("DataFrame Agrupado:")
    print(df_agrupado)
    print("\nDescripciones Exactas:")
    print(exact_descriptions)
    print("\nInventario:")
    print(inventario)
    print("\nDataFrame Combinado con Diferencia:")
    print(df_combinado[['id', 'descripcion', 'cantidad', 'Diferencia', 'diferencia_facturacion']])
    df_combinado = df_combinado[['id', 'descripcion', 'cantidad', 'Diferencia', 'diferencia_facturacion']]
    df_combinado = df_combinado.rename(columns={'Diferencia': 'inventario', 'cantidad': 'facturacion'})
    # Guardar el DataFrame combinado en un archivo Excel
    df_combinado.to_excel(f"src/sunchemical_comp/diferencia_facturacion_corte{corte}_mes{mes}.xlsx", index=False)
