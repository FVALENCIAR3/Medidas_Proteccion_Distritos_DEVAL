import streamlit as st
import pandas as pd
import io
from unidecode import unidecode
import warnings
warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(page_title="Clasificador por Distritos", page_icon="🏢")

st.title("📊 Clasificador de Registros por Distritos")
st.write("Sube tu archivo de Excel o CSV para organizar la información por distritos y clasificar los registros duplicados automáticamente.")

# Definir el diccionario de distritos y sus municipios
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

def asignar_distrito(texto):
    if pd.isna(texto):
        return 'SIN CLASIFICAR'
    texto_limpio = unidecode(str(texto)).upper()
    for distrito, municipios in distritos.items():
        for municipio in municipios:
            if municipio in texto_limpio:
                return distrito
    return 'SIN CLASIFICAR'

# Botón para subir archivo
uploaded_file = st.file_uploader("📥 Sube tu archivo Excel (.xlsx, .xls) o CSV (.csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        with st.spinner('Procesando archivo...'):
            # Leer archivo
            if uploaded_file.name.endswith('.csv'):
                try:
                    df = pd.read_csv(uploaded_file, encoding='utf-8')
                except:
                    df = pd.read_csv(uploaded_file, encoding='latin1')
            else:
                df = pd.read_excel(uploaded_file)
            
            # Identificar columnas
            col_estacion = next((c for c in df.columns if 'ESTACI' in c.upper()), None)
            col_id = next((c for c in df.columns if 'IDENTIF' in c.upper()), None)
            col_unidad = next((c for c in df.columns if 'UNIDAD' in c.upper()), None)

            # Asignar Distrito
            if col_estacion:
                df['DISTRITO_ASIGNADO'] = df[col_estacion].apply(asignar_distrito)
                st.info(f"✅ Distritos asignados basados en la columna: '{col_estacion}'")
            else:
                st.warning("⚠️ No se encontró columna de 'Estación' o similar.")
                df['DISTRITO_ASIGNADO'] = 'SIN CLASIFICAR'

            df = df.sort_values(by='DISTRITO_ASIGNADO')

            # Duplicados
            if col_id and col_unidad:
                df['REGISTRO_DUPLICADO'] = df.duplicated(subset=[col_id, col_unidad], keep=False)
                df_duplicados = df[df['REGISTRO_DUPLICADO'] == True].sort_values(by=['DISTRITO_ASIGNADO', col_id, col_unidad])
                st.info(f"✅ Se evaluaron duplicados usando columnas: '{col_id}' y '{col_unidad}' (Discriminados por Distrito)")
            else:
                st.warning("⚠️ No se encontraron columnas de Identificación o Unidad. Se buscaron filas idénticas completas.")
                df['REGISTRO_DUPLICADO'] = df.duplicated(keep=False)
                df_duplicados = df[df['REGISTRO_DUPLICADO'] == True].sort_values(by=['DISTRITO_ASIGNADO'])

            # Crear archivo Excel en memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Base_Completa', index=False)
                
                if not df_duplicados.empty:
                    df_duplicados.to_excel(writer, sheet_name='Duplicados_Por_Distrito', index=False)
                else:
                    pd.DataFrame({'Mensaje': ['No se encontraron duplicados']}).to_excel(writer, sheet_name='Duplicados_Por_Distrito', index=False)
                
                distritos_encontrados = sorted(df['DISTRITO_ASIGNADO'].unique())
                for dist in distritos_encontrados:
                    if dist != 'SIN CLASIFICAR':
                        df_temp = df[df['DISTRITO_ASIGNADO'] == dist]
                        nombre_hoja = dist[:31] 
                        df_temp.to_excel(writer, sheet_name=nombre_hoja, index=False)
                        
                df_sin_clasificar = df[df['DISTRITO_ASIGNADO'] == 'SIN CLASIFICAR']
                if not df_sin_clasificar.empty:
                    df_sin_clasificar.to_excel(writer, sheet_name='Sin_Clasificar', index=False)

            st.success("¡Procesamiento terminado con éxito! Puedes descargar tu archivo organizado.")
            
            # Botón de descarga
            st.download_button(
                label="⬇️ Descargar Archivo Organizado (.xlsx)",
                data=output.getvalue(),
                file_name="Datos_Clasificados_Por_Distritos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")