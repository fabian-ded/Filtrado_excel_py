import streamlit as st
import pandas as pd
import io
from openpyxl.utils.dataframe import dataframe_to_rows

# Función para procesar el archivo Excel y generar la tabla dinámica con filtros automáticos
def generar_tabla_dinamica_con_filtros(df):
    # elimina espacios en blanco al principio y al final de los nombres de las columnas
    df.columns = df.columns.str.strip()

    # Verificar que las columnas necesarias existen en el DataFrame
    columnas_necesarias = ["Número de Documento", "Nombre", "Apellidos", "Estado", "Competencia", "Juicio de Evaluación"]
    columnas_faltantes = [col for col in columnas_necesarias if col not in df.columns]
    if columnas_faltantes:
        st.error(f"Faltan columnas en el archivo: {', '.join(columnas_faltantes)}")
        return None

    # Filtrar los datos según las condiciones requeridas
    df_filtrado = df[
        (df["Juicio de Evaluación"].isin(["NO APROBADO", "POR EVALUAR"])) &  # Filtrar "Juicio de Evaluación"
        (df["Estado"] == "EN FORMACION")  # Filtrar "Estado"
    ]

    # Generar la tabla dinámica con los datos filtrados
    tabla_dinamica = pd.pivot_table(
        df_filtrado,
        values="Juicio de Evaluación",  # Columna que se cuenta
        index=["Número de Documento", "Nombre", "Apellidos"],  # Índices principales
        columns=["Estado", "Competencia"],  # Columnas dinámicas
        aggfunc='count',  # Contar valores de la columna
        fill_value=0  # Rellenar valores vacíos con 0
    )

    # Limpieza del MultiIndex para que las columnas sean más comprensibles
    tabla_dinamica.columns = [' '.join(col).strip() for col in tabla_dinamica.columns.values]

    # Agregar una columna "Total General" que sume todas las competencias para cada persona
    tabla_dinamica["Total General"] = tabla_dinamica.sum(axis=1)

    return tabla_dinamica

# Función para guardar la tabla dinámica en un archivo Excel
def guardar_excel_con_tabla_dinamica(df, tabla_dinamica, file_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Escribir datos originales en la hoja "Hoja"
        df.to_excel(writer, sheet_name="Hoja", index=False)

        # Escribir la tabla dinámica en la hoja "Tabla_Dinamica"
        tabla_dinamica.to_excel(writer, sheet_name="Tabla_Dinamica")

    output.seek(0)
    return output

# Interfaz de usuario con Streamlit
st.title('Generador de Tabla Dinámica con Filtros Automáticos')

# Cargar archivo Excel
uploaded_file = st.file_uploader("Sube un archivo Excel (.xls o .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # Obtener el nombre del archivo subido
        nombre_archivo = uploaded_file.name

        # Leer el archivo Excel, saltando las primeras 12 filas para los primeros datos
        df = pd.read_excel(uploaded_file, sheet_name="Hoja", skiprows=12)  # Saltar las primeras 12 filas

        # Mostrar los primeros registros (de la fila 2 hasta la 12)
        df_preliminar = pd.read_excel(uploaded_file, sheet_name="Hoja", skiprows=1, nrows=10)  # Fila 2 hasta la 12
        st.write("Datos cargados de la hoja:"(uploaded_file.name))
        st.write(df_preliminar)

        # Procesar y generar la tabla dinámica
        if st.button('Generar Tabla Dinámica'):
            tabla_dinamica = generar_tabla_dinamica_con_filtros(df)
            if tabla_dinamica is not None:
                # Mostrar la tabla dinámica de forma interactiva
                st.write("Visualiza la tabla dinámica:")
                st.dataframe(tabla_dinamica)

                # Crear un archivo Excel en memoria con la tabla dinámica
                archivo_excel = guardar_excel_con_tabla_dinamica(df, tabla_dinamica, nombre_archivo)

                # Colocar el botón de descarga debajo de la tabla dinámica
                st.download_button(
                    label="Descargar Excel",
                    data=archivo_excel,
                    file_name=f"tabla_dinamica_{nombre_archivo}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
