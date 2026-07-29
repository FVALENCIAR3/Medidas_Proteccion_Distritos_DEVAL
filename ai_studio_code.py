# Instalar librería para quitar tildes fácilmente y manejar excel
!pip install unidecode xlsxwriter -q

import pandas as pd
import io
from google.colab import files
from unidecode import unidecode
import warnings
import numpy as np # Import numpy for NaN handling
warnings.filterwarnings('ignore')

# 1. Definir el diccionario de distritos y sus municipios (Actualizado con Cascajal)
distritos = {
    'DISTRITO 1 BUGA': ['CALIMA', 'DARIEN', 'GINEBRA', 'GUACARI', 'BUGA', 'RESTREPO', 'SAN PEDRO', 'YOTOCO'],
    'DISTRITO 2 TULUA': ['ANDALUCIA', 'BUGALAGRANDE', 'RIO FRIO', 'RIOFRIO', 'TRUJILLO', 'TULUA'],
    'DISTRITO 3 SEVILLA': ['CAICEDONIA', 'SEVILLA'],
    'DISTRITO 4 ROLDANILLO': ['BOLIVAR', 'EL DOVIO', 'LA UNION', 'LA VICTORIA', 'ROLDANILLO', 'TORO', 'VERSALLES', 'ZARZAL'],
    'DISTRITO 5 DAGUA': ['DAGUA'],
    'DISTRITO 6 PALMIRA': ['PALMIRA', 'PRADERA', 'CERRITO', 'FLORIDA'],
    'DISTRITO 7 CARTAGO': ['ALCALA', 'ANSERMANUEVO', 'ARGELIA', 'CARTAGO', 'EL AGUILA', 'EL CAIRO', 'OBANDO', 'SAN JOSE DEL PALMAR', 'ULLOA'],
    'DISTRITO 8 BUENAVENTURA': ['BUENAVENTURA', 'CASCAJAL']
}

# Función para mapear el texto al distrito
def asignar_distrito(texto):
    if pd.isna(texto):
        return 'SIN CLASIFICAR'

    # Convertir a mayúsculas y quitar tildes para evitar errores
    texto_limpio = unidecode(str(texto)).upper()

    for distrito, municipios in distritos.items():
        for municipio in municipios:
            if municipio in texto_limpio:
                return distrito
    return 'SIN CLASIFICAR'

# 2. Cargar el archivo
print("Por favor, sube tu archivo Excel (.xlsx, .xls) o CSV (.csv)")
uploaded = files.upload()
nombre_archivo = list(uploaded.keys())[0]

print(f"\nProcesando archivo: {nombre_archivo}...")

# Leer el archivo dependiendo de su extensión
if nombre_archivo.endswith('.csv'):
    try:
        df = pd.read_csv(io.BytesIO(uploaded[nombre_archivo]), encoding='utf-8', sep=';')
    except:
        df = pd.read_csv(io.BytesIO(uploaded[nombre_archivo]), encoding='latin1', sep=';')
else:
    df = pd.read_excel(io.BytesIO(uploaded[nombre_archivo]))

# 3. Identificar columnas clave (Estación, Identificación, Unidad)
col_estacion = next((c for c in df.columns if 'ESTACI' in c.upper()), None)
col_id = next((c for c in df.columns if 'IDENTIF' in c.upper()), None)
col_unidad = next((c for c in df.columns if 'UNIDAD' in c.upper()), None)

# 4. Asignar Distrito a toda la base
if col_estacion:
    df['DISTRITO_ASIGNADO'] = df[col_estacion].apply(asignar_distrito)
    print(f"Distritos asignados basados en la columna: '{col_estacion}'")
    # Ordenar la base principal por Distrito y luego por Estación
    df = df.sort_values(by=['DISTRITO_ASIGNADO', col_estacion])
else:
    print("⚠️ ADVERTENCIA: No se encontró columna de 'Estación' o similar. Revisa los nombres de tus columnas.")
    df['DISTRITO_ASIGNADO'] = 'SIN CLASIFICAR'
    # Ordenar la base principal por Distrito
    df = df.sort_values(by='DISTRITO_ASIGNADO')

# 5. Clasificar Duplicados por Unidad e Identificación (Y organizarlos por Distrito)
if col_id and col_unidad:
    # Marca como True TODOS los registros repetidos
    df['REGISTRO_DUPLICADO'] = df.duplicated(subset=[col_id, col_unidad], keep=False)

    # Extraer duplicados y ORDENAR primero por Distrito, luego por Estación, luego por ID y luego por Unidad
    if col_estacion:
        df_duplicados = df[df['REGISTRO_DUPLICADO'] == True].sort_values(by=['DISTRITO_ASIGNADO', col_estacion, col_id, col_unidad])
    else:
        df_duplicados = df[df['REGISTRO_DUPLICADO'] == True].sort_values(by=['DISTRITO_ASIGNADO', col_id, col_unidad])
    print(f"Se evaluaron duplicados usando columnas: '{col_id}' y '{col_unidad}' (Discriminados por Distrito)")
else:
    print("⚠️ ADVERTENCIA: No se encontraron columnas de Identificación o Unidad. Se buscarán filas idénticas completas.")
    df['REGISTRO_DUPLICADO'] = df.duplicated(keep=False)
    df_duplicados = df[df['REGISTRO_DUPLICADO'] == True].sort_values(by=['DISTRITO_ASIGNADO'])

# 6. Exportar los resultados
nombre_salida = 'Datos_Clasificados_Por_Distritos.xlsx'

with pd.ExcelWriter(nombre_salida, engine='xlsxwriter') as writer:
    # Pestaña 1: Base de datos completa ordenada
    df.to_excel(writer, sheet_name='Base_Completa', index=False)

    # Pestaña 2: Registros repetidos ordenados por distrito con títulos
    if not df_duplicados.empty:
        worksheet_dups = writer.book.add_worksheet('Duplicados_Por_Distrito')
        title_format = writer.book.add_format({'bold': True, 'font_size': 14, 'bg_color': '#DCE6F1'})
        current_row = 0

        # Write header once at the top of the sheet
        header = df_duplicados.columns.tolist()
        for col_idx, col_name in enumerate(header):
            worksheet_dups.write(current_row, col_idx, col_name)
        current_row += 2 # Move to the row after the header, leaving a blank row

        # Iterate through districts, df_duplicados is already sorted by 'DISTRITO_ASIGNADO'
        for dist_name, group in df_duplicados.groupby('DISTRITO_ASIGNADO'):
            # Write district title
            worksheet_dups.write(current_row, 0, f"DISTRITO: {dist_name}", title_format)
            current_row += 1 # Move to the next row for data

            # Write data for the current district group
            # Convert group DataFrame to list of lists (rows) for writing
            for r_idx, row_data in enumerate(group.values):
                for c_idx, cell_value in enumerate(row_data):
                    # Convert NaN to empty string to avoid TypeError with xlsxwriter
                    if pd.isna(cell_value):
                        cell_value = ''
                    worksheet_dups.write(current_row + r_idx, c_idx, cell_value)
            current_row += len(group) + 2 # Move current_row past this group's data and add two blank rows for separation

    else:
        # If no duplicates, create a simple sheet with a message
        pd.DataFrame({'Mensaje': ['No se encontraron duplicados']}).to_excel(writer, sheet_name='Duplicados_Por_Distrito', index=False)

    # Pestañas por cada Distrito
    distritos_encontrados = sorted(df['DISTRITO_ASIGNADO'].unique())
    for dist in distritos_encontrados:
        if dist != 'SIN CLASIFICAR':
            df_temp = df[df['DISTRITO_ASIGNADO'] == dist].copy() # Use .copy() to avoid SettingWithCopyWarning
            # Limitar el nombre de la hoja a 31 caracteres (regla de Excel)
            nombre_hoja = dist[:31]

            # Create a new worksheet for the district
            worksheet_dist = writer.book.add_worksheet(nombre_hoja)
            title_format_station = writer.book.add_format({'bold': True, 'font_size': 12, 'bg_color': '#DCE6F1'})
            current_row_dist = 0

            # Write header once at the top of the sheet
            header = df_temp.columns.tolist()
            for col_idx, col_name in enumerate(header):
                worksheet_dist.write(current_row_dist, col_idx, col_name)
            current_row_dist += 2 # Move to the row after the header, leaving a blank row

            # Iterate through stations within the district if col_estacion is valid
            if col_estacion and col_estacion in df_temp.columns:
                df_temp = df_temp.sort_values(by=col_estacion) # Sort by station for consistent grouping
                for station_name, group_station in df_temp.groupby(col_estacion):
                    # Write station title
                    worksheet_dist.write(current_row_dist, 0, f"ESTACIÓN: {station_name}", title_format_station)
                    current_row_dist += 1 # Move to the next row for data

                    # Write data for the current station group
                    for r_idx, row_data in enumerate(group_station.values):
                        for c_idx, cell_value in enumerate(row_data):
                            if pd.isna(cell_value):
                                cell_value = ''
                            worksheet_dist.write(current_row_dist + r_idx, c_idx, cell_value)
                    current_row_dist += len(group_station) + 2 # Move current_row_dist past this group's data and add two blank rows for separation
            else:
                # If no station column found, just write the entire district data without station separation
                for r_idx, row_data in enumerate(df_temp.values):
                    for c_idx, cell_value in enumerate(row_data):
                        if pd.isna(cell_value):
                            cell_value = ''
                        worksheet_dist.write(current_row_dist + r_idx, c_idx, cell_value)
                current_row_dist += len(df_temp) + 2

    # Pestaña para los que no coincidieron con ningún municipio
    df_sin_clasificar = df[df['DISTRITO_ASIGNADO'] == 'SIN CLASIFICAR']
    if not df_sin_clasificar.empty:
        df_sin_clasificar.to_excel(writer, sheet_name='Sin_Clasificar', index=False)

print("\n✅ Proceso terminado con éxito. Descargando archivo...")
files.download(nombre_salida)
