import streamlit as st
import pandas as pd
import base64
import io
import re
import calendar
from datetime import datetime
import numpy as np
import xlsxwriter

# Configuración de la página
st.set_page_config(
    page_title="Unificador de Reportes Netsuite-Salesforce",
    page_icon="📊",
    layout="wide"
)

# Título de la aplicación
st.title("Unificador de Reportes Netsuite-Salesforce")
st.write("Esta aplicación te permite combinar reportes de Netsuite y Salesforce en un único archivo XLSX o CSV.")

# Función para convertir fecha a formato unificado
def convert_date_format(date_str, source_format="netsuite"):
    try:
        if pd.isna(date_str) or date_str == "" or date_str is None:
            return None
        
        date_str = str(date_str).strip()
        
        # Diferentes patrones para detectar el formato de fecha
        patterns = {
            # DD/MM/YYYY
            "dd_mm_yyyy": r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$',
            # MM/DD/YYYY
            "mm_dd_yyyy": r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$',
            # YYYY/MM/DD
            "yyyy_mm_dd": r'^(\d{4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})$',
            # DD-MMM-YYYY o DD MMM YYYY
            "dd_mmm_yyyy": r'^(\d{1,2})[\s\-\.]+([A-Za-z]{3})[\s\-\.]+(\d{4})$',
        }
        
        for pattern_name, pattern in patterns.items():
            match = re.match(pattern, date_str)
            if match:
                if pattern_name == "dd_mm_yyyy":
                    day, month, year = match.groups()
                    # Verificar si el formato es realmente MM/DD/YYYY (común en EEUU)
                    if int(month) <= 12 and int(day) <= 12:
                        # Si source_format es "netsuite", asumimos que viene en formato europeo DD/MM/YYYY
                        if source_format == "netsuite":
                            return f"{int(day):02d}/{int(month):02d}/{year}"
                        # En caso de duda, seguir el formato MM/DD/YYYY (para EEUU)
                        else:
                            # Aquí invertimos día y mes si viene en formato americano
                            return f"{int(day):02d}/{int(month):02d}/{year}"
                    else:
                        # Si month > 12, entonces es claramente DD/MM/YYYY
                        return f"{int(day):02d}/{int(month):02d}/{year}"
                
                elif pattern_name == "mm_dd_yyyy":
                    month, day, year = match.groups()
                    return f"{int(day):02d}/{int(month):02d}/{year}"
                
                elif pattern_name == "yyyy_mm_dd":
                    year, month, day = match.groups()
                    return f"{int(day):02d}/{int(month):02d}/{year}"
                
                elif pattern_name == "dd_mmm_yyyy":
                    day, month_str, year = match.groups()
                    month_dict = {
                        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
                        'Ene': 1, 'Feb': 2, 'Mar': 3, 'Abr': 4, 'May': 5, 'Jun': 6,
                        'Jul': 7, 'Ago': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dic': 12
                    }
                    month = month_dict.get(month_str, 1)
                    return f"{int(day):02d}/{month:02d}/{year}"
        
        # Si no coincide con ningún patrón común, intentar con datetime
        try:
            # Intentar con datetime para detectar automáticamente el formato
            for fmt in ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y"]:
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt.strftime("%d/%m/%Y")
                except ValueError:
                    continue
        except:
            pass
        
        # Si no se pudo convertir, devolver el valor original
        return date_str
    except Exception as e:
        print(f"Error al convertir fecha '{date_str}': {e}")
        return date_str

# Función para formatear números (eliminar separadores de miles y mantener punto decimal)
def format_number(value):
    try:
        # Si es un número, convertir a string primero
        if isinstance(value, (int, float)):
            return str(value)
        
        # Si ya es string, procesar
        if isinstance(value, str):
            # Eliminar caracteres no numéricos excepto punto y coma
            value = ''.join(c for c in value if c.isdigit() or c in '.,')
            
            # Si hay puntos y comas, asumir que el último es el decimal
            if '.' in value and ',' in value:
                # Determinar cuál es el separador decimal (el último)
                last_dot_pos = value.rfind('.')
                last_comma_pos = value.rfind(',')
                
                if last_dot_pos > last_comma_pos:  # El punto es el separador decimal
                    # Eliminar todas las comas (separadores de miles)
                    value = value.replace(',', '')
                else:  # La coma es el separador decimal
                    # Eliminar todos los puntos (separadores de miles) y cambiar la última coma por punto
                    value = value.replace('.', '')
                    value = value[:last_comma_pos] + '.' + value[last_comma_pos+1:]
            elif ',' in value:
                # Si solo hay comas, la última es el separador decimal
                last_comma_pos = value.rfind(',')
                if last_comma_pos == len(value) - 3 or last_comma_pos == len(value) - 2:
                    # Parece ser un separador decimal, cambiar por punto
                    value = value.replace(',', '.')
                else:
                    # Probablemente son separadores de miles, eliminarlos
                    value = value.replace(',', '')
            
            # Intentar convertir a float y luego de nuevo a string para asegurar formato consistente
            try:
                return str(float(value))
            except ValueError:
                return value
        return value
    except Exception as e:
        # st.warning(f"Error al formatear número '{value}': {e}")
        return value

# Función para formatear números específicamente para Excel
def format_number_for_excel(value):
    try:
        if pd.isna(value):
            return ""
            
        # Si es un número o puede convertirse a uno
        try:
            # Primero limpiar el valor usando la función anterior
            cleaned_value = format_number(value)
            # Convertir a número
            num_value = float(cleaned_value)
            
            # Formatear el número con el formato español (coma como decimal)
            # Y asegurar que tenga comillas para que Excel no lo convierta
            if num_value == int(num_value):
                # Es un entero
                formatted = f'"{int(num_value)}"'
            else:
                # Es un decimal - usar 2 decimales y coma como separador
                formatted = f'"{str(num_value).replace(".", ",")}"'
                
            return formatted
        except:
            # Si no se puede convertir a número, devolver como string con comillas
            return f'"{str(value)}"'
    except Exception as e:
        # Si hay cualquier error, devolver el valor original
        return value

# Función para leer y mostrar datos
def read_and_display_data(file, title):
    if file is not None:
        try:
            df = pd.read_csv(file)
            st.write(f"**Vista previa de {title}:**")
            st.dataframe(df.head())
            return df
        except Exception as e:
            st.error(f"Error al cargar el archivo {title}: {e}")
            return None
    return None

# Función para generar link de descarga
def get_csv_download_link(df, filename="datos_unificados.csv"):
    # Reemplazar puntos por comas en las columnas numéricas antes de descargar
    df_download = df.copy()
    numeric_columns = ["Total", "Total USD", "Quantity", "FX Rate", "FX Rate Item", "Consolidated FX Rate"]
    
    for col in numeric_columns:
        if col in df_download.columns:
            df_download[col] = df_download[col].apply(
                lambda x: format_number_for_excel(x) if pd.notna(x) else x
            )
    
    # Usar punto y coma como separador para evitar conflictos con comas
    # Asegurar que los números no se conviertan a notación científica
    csv = df_download.to_csv(index=False, sep=';', float_format='%.10f')
    
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Descargar CSV unificado</a>'
    return href

# Mapeo predeterminado de columnas Salesforce a Netsuite
default_mapping = {
    "Probability (%)": "Quantity",
    "Client Leader": "_Client Leader AUX",  # Nombre exacto de la columna
    "Project Manager": "_PM",              # Nombre exacto de la columna
    "Amount Currency": "Proj. Currency",
    "Amount (converted)": "Total",
    "Account Name": "Customer Parent",
    "Opportunity Name": "Project(PLAN)",
    "Month": "Date"
}

# Cargar archivos CSV
st.header("1. Cargar archivos CSV")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Archivo de Netsuite")
    netsuite_file = st.file_uploader("Cargar CSV de Netsuite", type=["csv"])

with col2:
    st.subheader("Archivo de Salesforce")
    salesforce_file = st.file_uploader("Cargar CSV de Salesforce", type=["csv"])

# Procesar archivos si se han cargado
if netsuite_file and salesforce_file:
    st.header("2. Vista previa de los datos")
    
    # Leer y mostrar datos
    netsuite_df = read_and_display_data(netsuite_file, "Netsuite")
    salesforce_df = read_and_display_data(salesforce_file, "Salesforce")
    
    # Verificar y mostrar problemas con las columnas
    if netsuite_df is not None:
        st.write("**Columnas en Netsuite:**")
        st.write(", ".join(netsuite_df.columns.tolist()))
        if "_PM" not in netsuite_df.columns and "_Client Leader AUX" not in netsuite_df.columns:
            st.warning("⚠️ No se encontraron las columnas '_PM' o '_Client Leader AUX' en el archivo de Netsuite. Verifica que los nombres sean exactos.")
    
    if salesforce_df is not None:
        st.write("**Columnas en Salesforce:**")
        st.write(", ".join(salesforce_df.columns.tolist()))
    
    if netsuite_df is not None and salesforce_df is not None:
        st.header("3. Mapeo de columnas")
        st.write("El mapeo predeterminado de columnas ya está configurado según tus especificaciones:")
        
        # Crear mapeo entre las columnas de Salesforce y Netsuite
        salesforce_columns = salesforce_df.columns.tolist()
        mapping = {col: "No mapear" for col in salesforce_columns}
        
        # Aplicar mapeo predeterminado
        for sf_col, ns_col in default_mapping.items():
            if sf_col in mapping:
                mapping[sf_col] = ns_col
            # También revisar si hay alguna columna que contenga el nombre (para casos como "Amount (converted)")
            else:
                for col in salesforce_columns:
                    if sf_col in col:
                        mapping[col] = ns_col
                        break
        
        col1, col2 = st.columns(2)
        
        # Obtener las columnas de Netsuite y agregar "Estado" si no existe
        netsuite_columns_list = netsuite_df.columns.tolist()
        if "Estado" not in netsuite_columns_list:
            netsuite_columns_list.append("Estado")
            
        with col1:
            st.subheader("Columnas de Salesforce")
            for sf_col in salesforce_columns:
                options = ["No mapear"] + netsuite_columns_list
                default_index = options.index(mapping[sf_col]) if mapping[sf_col] in options else 0
                mapping[sf_col] = st.selectbox(
                    f"Mapear '{sf_col}' a:",
                    options=options,
                    index=default_index,
                    key=f"map_{sf_col}"
                )
        
        with col2:
            st.subheader("Vista previa del mapeo")
            mapping_df = pd.DataFrame({
                'Columna Salesforce': mapping.keys(),
                'Columna Netsuite': mapping.values()
            })
            mapping_df = mapping_df[mapping_df['Columna Netsuite'] != "No mapear"]
            st.dataframe(mapping_df)
        
        st.header("4. Unificar datos")
        
        st.info("Al hacer clic en 'Unificar datos', la información del CSV de Salesforce se incorporará al formato de Netsuite, generando un único archivo CSV con toda la información integrada.")
        
        if st.button("Unificar datos"):
            with st.spinner("Procesando e incorporando datos de Salesforce a Netsuite..."):
                try:
                    # Filtrar los mensajes de depuración para que no sobrecarguen la interfaz
                    with st.expander("Ver detalles de procesamiento"):
                        st.write("Datos procesados correctamente. Expande para ver los detalles de mapeo.")
                    
                    # Crear una copia de la función de información para capturar mensajes
                    original_info = st.info
                    original_warning = st.warning
                    
                    # Redefine temporalmente las funciones para capturar mensajes
                    debug_messages = []
                    def capture_info(message):
                        debug_messages.append(f"ℹ️ {message}")
                    
                    def capture_warning(message):
                        debug_messages.append(f"⚠️ {message}")
                    
                    # Reemplazar las funciones temporalmente
                    st.info = capture_info
                    st.warning = capture_warning
                    
                    # Mostrar mensaje de información
                    st.info("Procesando los datos... Por favor espera.")
                    
                    # Verificar si hay columnas duplicadas en Netsuite
                    if len(netsuite_df.columns) != len(set(netsuite_df.columns)):
                        st.warning("Se detectaron columnas duplicadas en el archivo de Netsuite. Se renombrarán automáticamente para evitar conflictos.")
                    
                    # Preparar DataFrame de Netsuite para recibir datos de Salesforce
                    result_df = netsuite_df.copy()
                    
                    # Añadir columna "Estado" a Netsuite con valor "CONFIRMADO"
                    if "Estado" not in result_df.columns:
                        result_df["Estado"] = "CONFIRMADO"
                    
                    # Unificar formato de fechas en el DataFrame de Netsuite
                    if "Date" in result_df.columns:
                        result_df["Date"] = result_df["Date"].apply(lambda x: convert_date_format(x, "netsuite"))
                    
                    # Crear lista para almacenar las filas de Salesforce transformadas
                    salesforce_rows = []
                    
                    # Añadir variables para depuración del mapeo
                    mappings_applied = []
                    debug_rows = []
                    
                    # Transferir datos de Salesforce según el mapeo
                    for idx, row in salesforce_df.iterrows():
                        # Crear diccionario para la nueva fila con todas las columnas de netsuite_df y "Estado"
                        new_row = {col: None for col in result_df.columns}
                        
                        row_mappings = []
                        
                        for sf_col, ns_col in mapping.items():
                            if ns_col != "No mapear" and ns_col in new_row:
                                row_mappings.append(f"{sf_col} → {ns_col}")
                                
                                # Procesamiento especial para Client Leader (cambiar formato de nombre)
                                if (sf_col == "Client Leader" or "Client Leader" in sf_col) and ("_Client Leader AUX" in ns_col or "_Client Leader" in ns_col):
                                    # Cambiar formato "Nombre Apellido" a "Apellido, Nombre"
                                    if pd.notna(row[sf_col]) and str(row[sf_col]).strip() != "":
                                        try:
                                            name_str = str(row[sf_col]).strip()
                                            parts = name_str.split(maxsplit=1)
                                            if len(parts) > 1:
                                                formatted_name = f"{parts[1]}, {parts[0]}"
                                                st.info(f"Mapeando Client Leader: '{name_str}' a '{formatted_name}' en columna '{ns_col}'")
                                                new_row[ns_col] = formatted_name
                                            else:
                                                new_row[ns_col] = name_str
                                        except Exception as e:
                                            st.warning(f"Error al formatear el nombre '{row[sf_col]}': {e}")
                                            new_row[ns_col] = row[sf_col]
                                # Procesamiento especial para Project Manager a PM
                                elif (sf_col == "Project Manager" or "Project Manager" in sf_col) and ("_PM" in ns_col or ns_col.endswith("PM")):
                                    if pd.notna(row[sf_col]):
                                        # Asegurarse de que el valor se transfiera correctamente
                                        try:
                                            pm_value = str(row[sf_col]).strip()
                                            st.info(f"Mapeando Project Manager: '{pm_value}' a '{ns_col}'")
                                            new_row[ns_col] = pm_value
                                        except Exception as e:
                                            st.warning(f"Error al procesar Project Manager '{row[sf_col]}': {e}")
                                            new_row[ns_col] = str(row[sf_col])
                                # Procesamiento especial para Month a Date (Mmm.YYYY a DD/MM/YYYY)
                                elif (sf_col == "Month" or "Month" in sf_col) and "Date" in ns_col:
                                    if pd.notna(row[sf_col]) and str(row[sf_col]) != "":
                                        try:
                                            # Extraer mes y año del formato Mmm.YYYY (por ejemplo, Feb.2025)
                                            month_str = str(row[sf_col])
                                            st.info(f"Procesando fecha: '{month_str}'")
                                            
                                            # Diferentes patrones posibles para extraer mes y año
                                            patterns = [
                                                r'([A-Za-z]+)\.(\d{4})',  # Mmm.YYYY
                                                r'([A-Za-z]+)(\d{4})',    # MmmYYYY
                                                r'([A-Za-z]+)[^0-9]+(\d{4})',  # Cualquier separador
                                                r'(\d{1,2})[/\-](\d{4})'  # MM/YYYY o MM-YYYY
                                            ]
                                            
                                            month_num = None
                                            year_num = None
                                            
                                            # Probar cada patrón
                                            for pattern in patterns:
                                                match = re.match(pattern, month_str)
                                                if match:
                                                    part1, part2 = match.groups()
                                                    
                                                    # Ver si la primera parte es un mes en texto
                                                    month_dict = {
                                                        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                                                        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
                                                        'Ene': 1, 'Feb': 2, 'Mar': 3, 'Abr': 4, 'May': 5, 'Jun': 6,
                                                        'Jul': 7, 'Ago': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dic': 12
                                                    }
                                                    
                                                    # Intentar extraer el mes y año
                                                    if part1 in month_dict:
                                                        month_num = month_dict[part1]
                                                        year_num = int(part2)
                                                    elif part1.isdigit() and int(part1) >= 1 and int(part1) <= 12:
                                                        month_num = int(part1)
                                                        year_num = int(part2)
                                                    
                                                    break
                                            
                                            # Si no se pudo extraer con los patrones, intentar buscar partes numéricas
                                            if month_num is None or year_num is None:
                                                # Buscar 4 dígitos consecutivos para el año
                                                year_match = re.search(r'(\d{4})', month_str)
                                                if year_match:
                                                    year_num = int(year_match.group(1))
                                                    
                                                    # Buscar 1-2 dígitos para el mes
                                                    month_match = re.search(r'(?<!\d)([1-9]|1[0-2])(?!\d)', month_str)
                                                    if month_match:
                                                        month_num = int(month_match.group(1))
                                            
                                            # Si se encontró mes y año, formatear la fecha
                                            if month_num is not None and year_num is not None:
                                                # Obtener último día del mes
                                                last_day = calendar.monthrange(year_num, month_num)[1]
                                                
                                                # Formatear como DD/MM/YYYY (formato universal para ordenar)
                                                formatted_date = f"{last_day:02d}/{month_num:02d}/{year_num}"
                                                st.info(f"Fecha convertida: '{month_str}' → '{formatted_date}'")
                                                new_row[ns_col] = formatted_date
                                            else:
                                                st.warning(f"No se pudo extraer mes y año de '{month_str}'")
                                                new_row[ns_col] = month_str
                                        except Exception as e:
                                            st.warning(f"Error al convertir la fecha '{row[sf_col]}': {e}")
                                            new_row[ns_col] = row[sf_col]
                                # Procesamiento especial para Amount (converted) a Total
                                elif ("Amount" in sf_col and "converted" in sf_col) and ("Total" in ns_col):
                                    if pd.notna(row[sf_col]):
                                        try:
                                            # Intentar convertir a número y limpiar formato
                                            value_str = str(row[sf_col])
                                            value_clean = value_str.replace(',', '').replace('$', '').strip()
                                            new_row[ns_col] = float(value_clean) if value_clean else None
                                        except Exception as e:
                                            st.warning(f"Error al convertir el monto '{row[sf_col]}': {e}")
                                            new_row[ns_col] = row[sf_col]
                                else:
                                    # Transferir el valor de la columna de Salesforce a Netsuite
                                    new_row[ns_col] = row[sf_col]
                        
                        # Calcular TOTAL USD = TOTAL * (Probability / 100)
                        if "Total" in new_row and new_row["Total"] is not None and "Quantity" in new_row and new_row["Quantity"] is not None:
                            try:
                                # Asegurar que ambos valores sean numéricos
                                total_value = new_row["Total"]
                                qty_value = new_row["Quantity"]
                                
                                # Limpiar y convertir valores si son strings
                                if isinstance(total_value, str):
                                    total_value = total_value.replace(',', '').replace('$', '').strip()
                                    total_value = float(total_value) if total_value else 0
                                
                                if isinstance(qty_value, str):
                                    qty_value = qty_value.replace('%', '').replace(',', '').strip()
                                    qty_value = float(qty_value) if qty_value else 0
                                
                                # Realizar el cálculo
                                if isinstance(total_value, (int, float)) and isinstance(qty_value, (int, float)):
                                    new_row["Total USD"] = float(total_value) * (float(qty_value) / 100)
                                    
                                    # Determinar el valor de "Estado" basado en Probability (Quantity)
                                    if qty_value == 100:
                                        new_row["Estado"] = "CONFIRMADO"
                                    elif qty_value in [50, 70]:
                                        new_row["Estado"] = "PIPELINE"
                                    else:
                                        new_row["Estado"] = "NO INCLUIR"
                            except Exception as e:
                                st.warning(f"No se pudo calcular TOTAL USD: {e}")
                                new_row["Total USD"] = None
                                new_row["Estado"] = "NO INCLUIR"  # Valor por defecto si hay un error
                        else:
                            new_row["Estado"] = "NO INCLUIR"  # Si no hay datos para el cálculo
                        
                        # Agregar la fila a la lista y la información de depuración
                        salesforce_rows.append(new_row)
                        mappings_applied.append(row_mappings)
                        debug_rows.append(dict(row))
                    
                    # Restaurar las funciones originales
                    st.info = original_info
                    st.warning = original_warning
                    
                    # Mostrar los mensajes de depuración capturados en el expander
                    with st.expander("Ver detalles de procesamiento"):
                        for msg in debug_messages:
                            st.write(msg)
                    
                    # Mostrar un resumen de las columnas mapeadas
                    st.subheader("Resumen del mapeo aplicado:")
                    st.markdown("**Columnas mapeadas para la primera fila:**")
                    if mappings_applied:
                        for mapping_info in mappings_applied[0]:
                            st.write(f"- {mapping_info}")
                    
                    # Mostrar ejemplos de los valores mapeados
                    if debug_rows and salesforce_rows:
                        st.subheader("Ejemplos de valores mapeados (primera fila):")
                        important_cols = ["_PM", "_Client Leader AUX", "Date", "Total", "Total USD", "Estado"]
                        for col in important_cols:
                            if col in salesforce_rows[0]:
                                source_col = next((sf for sf, ns in mapping.items() if ns == col), "Desconocido")
                                source_value = debug_rows[0].get(source_col, "N/A") if source_col != "Desconocido" else "N/A"
                                mapped_value = salesforce_rows[0].get(col, "No mapeado")
                                st.markdown(f"**{col}**: `{source_value}` → `{mapped_value}`")
                    
                    # Verificar y mostrar advertencias importantes sobre columnas mapeadas
                    column_warnings = []
                    important_columns = ["_PM", "_Client Leader AUX", "Date", "Total", "Quantity"]
                    
                    # Comprobar si los datos de Salesforce se mapearon correctamente
                    if salesforce_rows:
                        mapped_columns = set(salesforce_rows[0].keys())
                        for col in important_columns:
                            if col in mapped_columns and salesforce_rows[0][col] is None:
                                column_warnings.append(f"⚠️ La columna '{col}' no parece tener datos mapeados correctamente.")
                    
                    if column_warnings:
                        st.warning("Se detectaron problemas en el mapeo:")
                        for warning in column_warnings:
                            st.write(warning)
                    
                    # Crear DataFrame con las filas de Salesforce
                    if salesforce_rows:
                        # Convertir la lista de filas de Salesforce a DataFrame
                        temp_salesforce = pd.DataFrame(salesforce_rows)
                        
                        # Combinar los DataFrames de manera segura, asegurando que tengan las mismas columnas
                        combined_df = pd.concat([result_df, temp_salesforce], axis=0, ignore_index=True)
                        
                        # Reemplazar 'nan' string con valores nulos reales
                        combined_df = combined_df.replace('nan', np.nan)
                        combined_df = combined_df.replace('None', np.nan)
                        
                        # Formatear valores numéricos (eliminar separadores de miles y formatear decimales)
                        numeric_columns = ["Total", "Total USD", "Quantity", "FX Rate", "FX Rate Item", "Consolidated FX Rate"]
                        for col in numeric_columns:
                            if col in combined_df.columns:
                                combined_df[col] = combined_df[col].apply(lambda x: format_number(x) if pd.notna(x) else x)
                        
                        # Unificar formato de todas las fechas para que sean ordenables
                        if "Date" in combined_df.columns:
                            combined_df["Date"] = combined_df["Date"].apply(lambda x: convert_date_format(x))
                    else:
                        combined_df = result_df.copy()
                    
                    # Mostrar resultado
                    st.subheader("Vista previa del resultado:")
                    st.dataframe(combined_df.head(10))
                    
                    # Información sobre el formato de descarga
                    st.info("""
                    El archivo CSV de descarga ha sido optimizado para Excel:
                    - Usa punto y coma (;) como separador de columnas
                    - Los valores numéricos usan coma (,) como separador decimal
                    - Los números tienen formato óptimo para evitar conversiones automáticas
                    - No aparecerá la advertencia de conversión a notación científica
                    
                    Al abrir el archivo en Excel, simplemente haz clic en "Aceptar" si aparece algún diálogo.
                    """)
                    
                    # Generar opción para descargar en formato Excel directamente
                    try:
                        # Crear un archivo Excel en memoria
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            combined_df.to_excel(writer, index=False, sheet_name='Datos Unificados')
                            # Configurar formato de números para las columnas numéricas
                            workbook = writer.book
                            worksheet = writer.sheets['Datos Unificados']
                            num_format = workbook.add_format({'num_format': '#,##0.00'})
                            
                            # Aplicar formato a columnas numéricas
                            numeric_columns = ["Total", "Total USD", "Quantity", "FX Rate", "FX Rate Item", "Consolidated FX Rate"]
                            for col in numeric_columns:
                                if col in combined_df.columns:
                                    col_idx = combined_df.columns.get_loc(col) + 1  # +1 porque en Excel las columnas comienzan en 1
                                    worksheet.set_column(col_idx, col_idx, None, num_format)
                            
                            # Aplicar formato para la columna Estado (resaltar visualmente)
                            if "Estado" in combined_df.columns:
                                estado_col_idx = combined_df.columns.get_loc("Estado") + 1
                                
                                # Crear formatos para cada tipo de estado
                                confirmado_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})  # Verde claro
                                pipeline_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})    # Amarillo
                                no_incluir_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})  # Rojo claro
                                
                                # Aplicar formato condicional
                                worksheet.conditional_format(1, estado_col_idx, len(combined_df)+1, estado_col_idx, {
                                    'type': 'cell',
                                    'criteria': 'equal to',
                                    'value': '"CONFIRMADO"',
                                    'format': confirmado_format
                                })
                                
                                worksheet.conditional_format(1, estado_col_idx, len(combined_df)+1, estado_col_idx, {
                                    'type': 'cell',
                                    'criteria': 'equal to',
                                    'value': '"PIPELINE"',
                                    'format': pipeline_format
                                })
                                
                                worksheet.conditional_format(1, estado_col_idx, len(combined_df)+1, estado_col_idx, {
                                    'type': 'cell',
                                    'criteria': 'equal to',
                                    'value': '"NO INCLUIR"',
                                    'format': no_incluir_format
                                })
                        
                        # Botón para descargar como Excel
                        excel_data = output.getvalue()
                        st.download_button(
                            label="⬇️ Descargar como Excel (.xlsx)",
                            data=excel_data,
                            file_name="datos_unificados.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.warning(f"No se pudo crear el archivo Excel: {e}")
                    
                    # Generar link de descarga CSV
                    st.markdown("<h4>O descarga como CSV:</h4>", unsafe_allow_html=True)
                    st.markdown(get_csv_download_link(combined_df), unsafe_allow_html=True)
                    
                    # Guardar mapeo para futuros usos
                    mapping_json = pd.Series(mapping).to_json()
                    st.download_button(
                        label="Guardar configuración de mapeo",
                        data=mapping_json,
                        file_name="mapeo_columnas.json",
                        mime="application/json"
                    )
                    
                except Exception as e:
                    st.error(f"Error al unificar los datos: {e}")
else:
    st.info("Por favor, carga ambos archivos CSV para continuar.")

# Información adicional
st.sidebar.header("Instrucciones")
st.sidebar.write("""
Previamente debes descargar los siguientes reportes:
Netsuite: "Delivery Tracking - Consolidated View"
Salesforce: "Copy of Pipeline by Month (All Accounts)"

1. Carga los archivos CSV de Netsuite y Salesforce.
2. Revisa la vista previa de los datos.
3. Verifica el mapeo predefinido de columnas de Salesforce a Netsuite.
4. Haz clic en 'Unificar datos' para incorporar la información de Salesforce al formato de Netsuite.
5. Descarga el XLSX o CSV unificado con toda la información integrada.

**Importante**: Esta aplicación incorpora la información de Salesforce al CSV de Netsuite, respetando la estructura de columnas de Netsuite. El resultado es un único archivo XLSX o CSV que contiene tanto los datos originales de Netsuite como los datos de Salesforce mapeados al formato de Netsuite.
""")

st.sidebar.header("Mapeo predeterminado")
st.sidebar.write("""
El mapeo predeterminado configurado es:
- Probability (%) → Quantity
- Client Leader → _Client Leader AUX (con formato "Apellido, Nombre")
- Project Manager → _PM
- Amount Currency → Proj. Currency
- Amount (converted) → Total
- Account Name → Customer Parent
- Opportunity Name → Project(PLAN)
- Month → Date (con formato convertido de "Mmm.YYYY" a "DD/MM/YYYY")

Además, se añaden automáticamente:
- Total USD = Total * (Probability / 100)
- Estado = CONFIRMADO, PIPELINE o NO INCLUIR según procedencia y Probability
""")