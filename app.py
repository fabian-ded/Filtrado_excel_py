import streamlit as st
import pandas as pd
import io

# Función para procesar el archivo Excel y generar la tabla dinámica con la columna de Total General
def generar_tabla_dinamica_con_filtros(df):
    # Limpiar los nombres de las columnas
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

# Interfaz de usuario con Streamlit
st.title('Generador de Tabla Dinámica con Filtros Automáticos y Total General')

# Cargar archivo Excel
uploaded_file = st.file_uploader("Sube un archivo Excel (.xls o .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # Leer el archivo Excel, tomando la hoja "Hoja"
        df = pd.read_excel(uploaded_file, sheet_name="Hoja", skiprows=12)  # Saltar las primeras 12 filas

        # Mostrar los primeros registros
        st.write("Datos cargados de la hoja 'Hoja':")
        st.write(df.head())

        # Procesar y generar la tabla dinámica
        if st.button('Generar Tabla Dinámica'):
            tabla_dinamica = generar_tabla_dinamica_con_filtros(df)

            if tabla_dinamica is not None:
                # Crear un archivo Excel en memoria con dos hojas
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Escribir datos originales en Hoja
                    df.to_excel(writer, sheet_name="Hoja", index=False)

                    # Escribir la tabla dinámica en Hoja2
                    tabla_dinamica.to_excel(writer, sheet_name="Hoja2")

                output.seek(0)

                # Descargar el archivo con la tabla dinámica filtrada en Hoja2
                st.download_button(
                    label="Descargar Excel con Tabla Dinámica Filtrada en Hoja2",
                    data=output,
                    file_name="tabla_dinamica_filtrada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Mostrar la tabla dinámica de forma interactiva
                st.write("Visualiza la tabla dinámica (Hoja2):")
                st.dataframe(tabla_dinamica)
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")