import subprocess  # Importar el módulo subprocess
import threading
import tkinter as tk
from tkinter import messagebox

import pandas as pd
from pynput import keyboard as kb
import pandas as pd
import sys
import io
##SE REQUIERE GUARDAR EL ARCHIVO EN LA CARPETA CON EL NOMBRE Excel enviar.xlsx

# declaración globales
producto_actual = "Hola mundo"
ultimo_escrito_codigo = ""
ultimo_escrito_cantidad = ""
# PROPUESTA SIGUIENTE, CREAR EL BOT QUE ACTUALICE ESTE CODIGO FUENTE AUTOMATICAMENTE PONIENDO LOS ARCHIVOS EN SU LUGAR.
# Cambia la codificación de la salida estándar
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Se requiere ajustar a la dirección del archivo en el pc de mamá
# Ruta al archivo SIEMPRE SE DEBE EDITAR SI CAMBIA EL NOMBRE DEL ARCHIVO
archivo_inventarios = "docs/Excel_enviar.xlsx"

# Se lee el archivo
# SIEMPRE SE DEBERÁ EDITAR EL NOMBRE DE LA HOJA SI CAMBIA
df_exportar = pd.read_excel(archivo_inventarios, sheet_name="Excel_enviar")
# print(df_exportar.head())
# Se eliminan los duplicados de la columna N° de Item
# este es el dataframe que se exportará a excel al final, y es en el cual se pondran los datos que se vayan ingresando
df_exportar = df_exportar.drop_duplicates("N° de ítem        ")
df_exportar.reset_index(drop=True, inplace=True)

# Se lee nuevamente el dataframe para sacar los datos del nombre y código de las tintas
# SIEMPRE SE DEBERÁ EDITAR EL NOMBRE DE LA HOJA
df = pd.read_excel(archivo_inventarios, sheet_name="Excel_enviar")  # ,skipfooter=1)
# print(df.head())
df_creacion_vector_cantidades = df[["N° de ítem        ", "Cant. uso           "]]
# se transforman a enteros para que coincidan con la lectura del codigo de barras.

df_creacion_vector_cantidades["N° de ítem        "] = df_creacion_vector_cantidades[
    "N° de ítem        "
].astype(int)
df_creacion_vector_cantidades.columns = ["Numero", "Cantidad"]
df_creacion_vector_cantidades = df_creacion_vector_cantidades.drop_duplicates()
lista = df_creacion_vector_cantidades.values.tolist()
# print(lista)

# Se filtran los datos que se requieren. para la determinación del producto
df = df[["N° de ítem        ", "Descrip. art.                 "]]
# se transforman a enteros para que coincidan con la lectura del codigo de barras.
df["N° de ítem        "] = df["N° de ítem        "].astype(int)
# Se eliminan los duplicados para tener la lista de los codigos y su nombre
df = df.drop_duplicates()
# se cambian los nombres de las columnas para que sea más facil su adquisición
df.columns = ["Numero", "Nombre"]


# Creación de ventana
ventana = tk.Tk()
ventana.geometry("700x350")

# label codigo leido
label_codigo = tk.Label(text="Código de la tinta N°", font="Helvetica 25")
label_codigo.pack()
# Creación de cuadro de texto para lectura del codigo con el lector
codigo_tinta = tk.Entry(ventana, font="helvetica 25")
codigo_tinta.pack()


# FUNCIONES


# Función para ejecutar el archivo externo
def ejecutar_script_externo():
    subprocess.run(["python", "consolidado_inventarios.py"], check=True)


# Frame para los botones
frame_botones = tk.Frame(ventana)
frame_botones.pack(side=tk.BOTTOM, pady=10)


def confirmar_ejecucion():
    respuesta = messagebox.askyesno(
        "Confirmar Ejecución",
        "¿Está seguro que desea ejecutar el consolidado? Ya debió haber revisado que el informe de inventario 'Inventario_Solumaster_enviar.xlsx' esté de manera correcta.",
    )
    if respuesta:
        ejecutar_script_externo()


# Botón para ejecutar el script externo
boton_ejecutar_script = tk.Button(
    frame_botones,
    text="Ejecutar Consolidado de Inventarios",
    command=confirmar_ejecucion,  # Llama a la función de confirmación en lugar de ejecutar directamente
)
boton_ejecutar_script.pack(side=tk.LEFT, padx=10)


def crear_excel():
    df_excel_append = pd.DataFrame(lista)
    df_exportar["Cant. uso           "] = df_excel_append[1]
    df_exportar_final = pd.DataFrame()
    df_exportar_final["N° de ítem"] = df_exportar["N° de ítem        "]
    df_exportar_final["Descrip. art."] = df_exportar["Descrip. art.                 "]
    df_exportar_final["Cant.uso"] = df_exportar["Cant. uso           "]
    print(df_exportar_final)
    # Crear el objeto ExcelWriter con un contexto para manejar automáticamente el guardado
    with pd.ExcelWriter("Inventario_Solumaster_enviar.xlsx") as writer:
        df_exportar_final.to_excel(writer)

    print("El DataFrame se ha escrito con éxito en el archivo de Excel.")


# Botón para crear el Excel
boton_excel = tk.Button(frame_botones, text="Crear Excel", command=crear_excel)
boton_excel.pack(side=tk.LEFT, padx=10)

bandera_enter = 0


def pulsa(tecla):
    global bandera_enter

    #    #Función para determinar el producto actual leído del codigo de barras
    def nombre_producto_actual(codigo_tinta_actual):
        producto_actual = df[df.Numero == codigo_tinta_actual]
        producto_actual = producto_actual["Nombre"]
        producto_actual = producto_actual.to_string()
        print(producto_actual.encode('utf-8', errors='ignore').decode('utf-8'))
        return producto_actual

    def habilitar_cantidad():
        global bandera_enter
        # Codigo para agregar al excel cada uno de los codigos
        valor_cantidad = cantidad.get()
        # print(codigo_tinta_actual)
        for i in range(len(lista)):
            for j in range(len(lista[i])):
                # print(lista[i][j])
                if lista[i][j] == codigo_tinta_actual:
                    # print(lista[i][j])
                    if lista[i][j + 1] == "                    ":
                        print("Nuevo codigo")
                        lista[i][j + 1] = "=" + valor_cantidad
                        print(valor_cantidad)

                    else:
                        print("codigo ya ingresado anteriormente")
                        lista[i][j + 1] = lista[i][j + 1] + "+" + valor_cantidad
                        print(valor_cantidad)

        # Código para llenar el excel

        # Se eliminan las credenciales ya enviadas
        codigo_tinta.delete("0", "end")
        cantidad.pack_forget()
        boton_ingreso.pack_forget()
        label_producto.pack_forget()
        label_cantidad.pack_forget()
        # print('Se ha pulsado la tecla ' + str(tecla))
        codigo_tinta.focus()
        bandera_enter = 0

    if tecla == kb.Key.enter:
        if bandera_enter == 0:
            bandera_enter = 1
            codigo_tinta_actual = codigo_tinta.get()
            print(codigo_tinta_actual)
            try:
                int(codigo_tinta_actual)
                it_is = True
            except ValueError:
                it_is = False
            if it_is == True:
                codigo_tinta_actual = int(codigo_tinta_actual)
                producto_actual = nombre_producto_actual(codigo_tinta_actual)
                label_producto = tk.Label(text=producto_actual, font="Helvetica 20")
                label_producto.pack()

                # label codigo leido
                label_cantidad = tk.Label(
                    text="Cantidad en 'KG' de la tinta", font="Helvetica 25"
                )
                label_cantidad.pack()
                cantidad = tk.Entry(ventana, font="helvetica 25")
                cantidad.pack()
                boton_ingreso = tk.Button(
                    ventana, text="Agregar al excel", command=habilitar_cantidad
                )
                boton_ingreso.pack()
                cantidad.focus()
            else:
                bandera_enter = 0


codigo_tinta.focus()


# funcion para correr lectura de codigo de barras para activacion del teclado
def runleccod():
    # Correr el escuchador de teclas
    kb.Listener(pulsa).run()


lecturacodigo = threading.Thread(target=runleccod)
lecturacodigo.start()
# Correr la applicación
ventana.mainloop()
# Correr el escuchador de teclas
# kb.Listener(pulsa).run()
