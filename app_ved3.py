import streamlit as st
import pandas as pd
import numpy as np
import os
import tempfile
import re
import math
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from collections import defaultdict
from io import BytesIO


# Configuración de la página
st.set_page_config(
    page_title="Herramienta de Análisis VED",
    page_icon="📊",
    layout="wide"
)

# ===================== CONSTANTES Y MAPEOS =====================

MAPEO_CENTROS = {
    "CENTRE PENITENCIARI LLEDONERS": "CP Lledoners",
    "CENTRE PENITENCIARI DE DONES": "CP Dones",
    "CENTRE PENITENCIARI DE DONES BARCELONA": "CP Dones",
    "CENTRE PENITENCIARI PUIG DE LES BASSES": "CP Puig de les Basses",
    "CENTRE PENITENCIARI MAS D'ENRIC": "CP Mas D'Enric",
    "CENTRE PENITENCIARI PONENT": "CP Ponent",
    "CENTRE PENITENCIARI BRIANS 1": "CP Brians 1",
    "CENTRE PENITENCIARI BRIANS 2": "CP Brians 2",
    "CENTRE PENITENCIARI DE JOVES": "CP de Joves",
    "CENTRE PENITENCIARI OBERT GIRONA": "CP Obert Girona",
    "CENTRE PENITENCIARI OBERT DE GIRONA": "CP Obert Girona",
    "CENTRE PENITENCIARI OBERT LLEIDA": "CP Obert Lleida",
    "CENTRE PENITENCIARI OBERT DE LLEIDA": "CP Obert Lleida",
    "CENTRE PENITENCIARI OBERT BARCELONA": "CP Obert BCN",
    "CENTRE PENITENCIARI OBERT DE BARCELONA": "CP Obert BCN",
    "CENTRE PENITENCIARI OBERT TARRAGONA": "CP Obert Tarragona",
    "CENTRE PENITENCIARI OBERT DE TARRAGONA": "CP Obert Tarragona",
    "CENTRE PENITENCIARI TERRASSA": "CP Terrassa",
    "CENTRE PENITENCIARI QUATRE CAMINS": "CP Quatre Camins",
    "CENTRE EDUCATIU CAN LLUPIÀ": "CE Can Llupià",
    "CENTRE EDUCATIU L'ALZINA": "CE L'Alzina",
    "CENTRE EDUCATIU MONTILIVI": "CE Montilivi",
    "CENTRE EDUCATIU TIL.LERS": "CE Til.lers",
    "CENTRE EDUCATIU ELS TIL·LERS": "CE Til.lers",
    "UNITAT TERAPÈUTICA TIL.LERS": "UT Til.lers",
    "UNITAT TERAPÈUTICA CENTRE EDUCATIU TIL·LERS": "UT Til.lers",
    "CENTRE EDUCATIU ORIOL BADIA": "CE Oriol Badia",
    "CENTRE EDUCATIU FOLCH I TORRES": "CE Folch i Torres",
    "CENTRE EDUCATIU EL SEGRE": "CE El Segre"
}

ORDEN_CENTROS = [
    "CP Lledoners (Pilot M1)",
    "CP Lledoners",
    "CP Dones",
    "CP Puig de les Basses",
    "CP Mas D'Enric",
    "CP Ponent",
    "CP Brians 1",
    "CP de Joves",
    "CP Obert Girona",
    "CP Obert Lleida",
    "CP Obert BCN",
    "CP Obert Tarragona",
    "CP Quatre Camins",
    "CP Brians 2",  
    "CE Can Llupià",
    "CE L'Alzina",
    "CE Montilivi",
    "CE Til.lers",
    "UT Til.lers",
    "CE Oriol Badia",
    "CE Folch i Torres",
    "CE El Segre"
]

def extraer_tabla_acumulada_anterior(file):
    """Extrae la tabla acumulada del mes más reciente del archivo Excel"""
    if not file:
        return None
    
    try:
        df = pd.read_excel(file, header=None)
        
        # Buscar todas las fechas en formato dd/mm/yyyy
        fechas_encontradas = []
        for idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                if pd.notna(cell):
                    cell_str = str(cell)
                    # Buscar patrón de fecha
                    match = re.search(r'(\d{2}/\d{2}/\d{4})', cell_str)
                    if match:
                        fechas_encontradas.append({
                            'fecha': match.group(1),
                            'fila': idx,
                            'col': col_idx
                        })
        
        if not fechas_encontradas:
            st.error("No se encontraron fechas en el archivo")
            return None
        
        # Ordenar por fila para obtener la última tabla
        fechas_encontradas.sort(key=lambda x: x['fila'], reverse=True)
        ultima_fecha_info = fechas_encontradas[0]
        
        # La tabla empieza 2 filas después de la fecha
        inicio_tabla = ultima_fecha_info['fila'] + 2
        
        # Leer la tabla desde esa posición
        df_tabla = pd.read_excel(file, header=inicio_tabla, nrows=21)  # 20 centros + 1 TOTAL
        
        # Verificar que tiene las columnas esperadas
        if df_tabla.shape[1] < 20:
            st.error("La tabla no tiene el formato esperado (faltan columnas)")
            return None
        
        # Renombrar columnas para que coincidan con el formato esperado
        # La estructura es: Interns, Inici VeD, [5 cols Trucades], [5 cols Videotrucades], 
        # [5 cols Videovisites], [2 cols Consultes], [1 col Reserves], [2 cols Alta Digital]
        
        columnas_nuevas = [
            'Interns', 'Inici VeD', 
            'N_Trucades', '%Incre_Trucades', 'Minuts_Trucades', '%Incre_Minuts_T', 'Interns_Trucades',
            'N_Videotrucades', '%Incre_Video', 'Minuts_Videotrucades', '%Incre_Minuts_V', 'Interns_Videotrucades',
            'N_Videovisites', '%Incre_Visites', 'Minuts_Videovisites', '%Incre_Minuts_Vi', 'Interns_Videovisites',
            'N_Consultes', 'Interns_Consultes',
            'N_Reserves', 'Interns_Reserves',
            'Alta_Digital', 'N_Alta'
        ]
        
        # Ajustar si hay menos columnas
        if df_tabla.shape[1] < len(columnas_nuevas):
            columnas_nuevas = columnas_nuevas[:df_tabla.shape[1]]
        
        df_tabla.columns = columnas_nuevas[:df_tabla.shape[1]]
        
        # Añadir columna Centro si no existe
        if 'Centro' not in df_tabla.columns:
            # Buscar la columna con nombres de centros (probablemente la primera con texto)
            for col in df_tabla.columns:
                if df_tabla[col].dtype == 'object':
                    df_tabla['Centro'] = df_tabla[col]
                    break
        
        st.success(f"✅ Tabla acumulada anterior cargada: {ultima_fecha_info['fecha']}")
        return df_tabla
        
    except Exception as e:
        st.error(f"Error procesando tabla acumulada: {e}")
        return None

def calcular_tabla_acumulada(df_actual, df_anterior):
    """Calcula la tabla acumulada sumando datos actuales + anteriores y calculando incrementos"""
    if df_anterior is None:
        return df_actual
    
    #print("df_actual: ",df_actual)
    #print("df_anterior: ",df_anterior)

    df_acumulada = df_actual.copy()

    # Copiar columnas no acumulables (Interns, Inici VeD, Alta_Digital) desde df_anterior
    columnas_copiar = ['Interns', 'Inici VeD', 'Alta_Digital']
    
    for idx, row in df_acumulada.iterrows():
        centro = row['Centro']
        
        if centro in df_anterior['Centro'].values:
            fila_anterior = df_anterior[df_anterior['Centro'] == centro].iloc[0]
            
            for col in columnas_copiar:
                if col in df_anterior.columns and col in df_acumulada.columns:
                    df_acumulada.at[idx, col] = fila_anterior[col]
    
    # Columnas numéricas a sumar
    columnas_sumar = [col for col in df_actual.columns 
                      if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_')]
    columnas_sumar = [col for col in columnas_sumar if col != 'N_Alta']  # Excluir N_Alta
    
    # Para cada centro
    for idx, row in df_acumulada.iterrows():
        centro = row['Centro']
        
        # Buscar el centro en la tabla anterior
        if centro in df_anterior['Centro'].values:
            fila_anterior = df_anterior[df_anterior['Centro'] == centro].iloc[0]
            
            # Sumar valores
            for col in columnas_sumar:
                if col in df_anterior.columns and col in df_acumulada.columns:
                    valor_anterior = fila_anterior[col] if pd.notna(fila_anterior[col]) else 0
                    valor_actual = row[col] if pd.notna(row[col]) else 0
                    
                    # Convertir a numérico si es necesario
                    if isinstance(valor_anterior, str):
                        valor_anterior = 0
                    if isinstance(valor_actual, str):
                        valor_actual = 0
                    
                    df_acumulada.at[idx, col] = valor_anterior + valor_actual
    
    # Calcular incrementos porcentuales
    # El incremento debe ser: (valor_actual_mes / valor_anterior_acumulado) * 100
    for idx, row in df_acumulada.iterrows():
        if row['Centro'] == 'TOTAL':
            continue
            
        centro = row['Centro']
        if centro not in df_anterior['Centro'].values:
            continue
        
        fila_anterior = df_anterior[df_anterior['Centro'] == centro].iloc[0]
        
        # Calcular incrementos para cada servicio
        servicios = [
            ('Trucades', 'N_Trucades', '%Incre_Trucades'),
            ('Trucades', 'Minuts_Trucades', '%Incre_Minuts_T'),
            ('Videotrucades', 'N_Videotrucades', '%Incre_Video'),
            ('Videotrucades', 'Minuts_Videotrucades', '%Incre_Minuts_V'),
            ('Videovisites', 'N_Videovisites', '%Incre_Visites'),
            ('Videovisites', 'Minuts_Videovisites', '%Incre_Minuts_Vi')
        ]
        
        for servicio, col_valor, col_incre in servicios:
            if col_valor in df_acumulada.columns and col_valor in df_anterior.columns:
                # Valor anterior acumulado
                valor_anterior_acum = fila_anterior[col_valor] if pd.notna(fila_anterior[col_valor]) else 0
                if isinstance(valor_anterior_acum, str):
                    valor_anterior_acum = 0
                
                # Valor actual acumulado (ya sumado anteriormente)
                valor_actual_acum = row[col_valor] if pd.notna(row[col_valor]) else 0
                if isinstance(valor_actual_acum, str):
                    valor_actual_acum = 0
                
                # Calcular incremento: ((actual_acum - anterior_acum) / actual_acum) * 100
                if valor_actual_acum > 0:
                    incremento = ((valor_actual_acum - valor_anterior_acum) / valor_actual_acum) * 100
                    df_acumulada.at[idx, col_incre] = f"{incremento:.1f}%"
                elif valor_anterior_acum > 0:
                    df_acumulada.at[idx, col_incre] = "-100.0%"
                else:
                    df_acumulada.at[idx, col_incre] = "0.0%"

    return df_acumulada

def formatear_numeros_df(df):
    """Elimina formato de comas en números del DataFrame"""
    df_formatted = df.copy()
    for col in df_formatted.columns:
        if df_formatted[col].dtype in ['int64', 'float64']:
            df_formatted[col] = df_formatted[col].astype(int)
    return df_formatted

# ===================== FUNCIONES DE PROCESAMIENTO =====================

def read_excel_or_csv(file):
    """Lee archivo Excel o CSV desde un objeto de archivo de Streamlit"""
    file_name = file.name
    _, extension = os.path.splitext(file_name)
    extension = extension.lower()
    
    if extension in ['.xlsx', '.xls']:
        df = pd.read_excel(file)
    elif extension == '.csv':
        file.seek(0)  # Resetear el puntero del archivo
        try:
            df = pd.read_csv(file, encoding='utf-8')
        except UnicodeDecodeError:
            file.seek(0)
            try:
                df = pd.read_csv(file, encoding='latin-1')
            except UnicodeDecodeError:
                file.seek(0)
                df = pd.read_csv(file, encoding='iso-8859-1')
    else:
        raise ValueError(f"Formato no soportado: {extension}")
    
    return df

def parsear_duracion(duracion_str):
    """Convierte string de duración a minutos totales"""
    if pd.isna(duracion_str) or duracion_str == '':
        return 0
    
    try:
        duracion_str = str(duracion_str).lower()
        minutos = 0
        segundos = 0
        
        min_match = re.search(r'(\d+)\s*min', duracion_str)
        if min_match:
            minutos = int(min_match.group(1))
        
        seg_match = re.search(r'(\d+)\s*seg', duracion_str)
        if seg_match:
            segundos = int(seg_match.group(1))
        
        minutos_totales = minutos + (segundos / 60)
        return minutos_totales
    except:
        return 0

def procesar_trucades_video(files, tipo="Trucades"):
    """Procesa archivos de trucades, videotrucades o videovisites"""
    if not files:
        return None
    
    # Acumular todos los DataFrames raw
    dfs_raw = []
    
    for file in files:
        try:
            df = read_excel_or_csv(file)
            
            # Buscar columnas flexiblemente
            col_centre = None
            col_nis = None
            col_duracio = None
            
            for col in df.columns:
                if 'CENTRE' in col.upper():
                    col_centre = col
                if 'NIS' in col.upper() or 'EXPEDIENT' in col.upper():
                    col_nis = col
                if 'DURAC' in col.upper():
                    col_duracio = col
            
            if not col_centre or not col_nis:
                st.warning(f"⚠️ {file.name}: Columnas necesarias no encontradas")
                continue
            
            dfs_raw.append(df)
            
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")
    
    if not dfs_raw:
        return None
    
    # Concatenar todos los DataFrames
    df_combined = pd.concat(dfs_raw, ignore_index=True)
    
    # Buscar columnas en el DataFrame combinado
    col_centre = None
    col_nis = None
    col_duracio = None
    
    for col in df_combined.columns:
        if 'CENTRE' in col.upper():
            col_centre = col
        if 'NIS' in col.upper() or 'EXPEDIENT' in col.upper():
            col_nis = col
        if 'DURAC' in col.upper():
            col_duracio = col
    
    if not col_centre or not col_nis:
        return None
    
    # Filtrar NaN
    df_combined = df_combined.dropna(subset=[col_centre, col_nis])
    
    # Filtrar por duración si existe
    if col_duracio:
        df_combined = df_combined[df_combined[col_duracio] != '-']
        df_combined = df_combined[df_combined[col_duracio] != ' - ']
        df_combined = df_combined[~df_combined[col_duracio].astype(str).str.strip().eq('-')]
        df_combined['minutos_totales'] = df_combined[col_duracio].apply(parsear_duracion)
        df_combined = df_combined[df_combined['minutos_totales'] > (5/60)]
    
    # Procesar por centro
    resultados = []
    for centro_original in df_combined[col_centre].unique():
        df_centro = df_combined[df_combined[col_centre] == centro_original]
        centro = MAPEO_CENTROS.get(str(centro_original).upper(), centro_original)
        
        n_llamadas = len(df_centro)
        n_internos = df_centro[col_nis].nunique()  # Elimina duplicados correctamente
        
        minutos_enteros = 0
        if col_duracio:
            minutos_totales = sum(df_centro[col_duracio].apply(parsear_duracion))
            minutos_enteros = math.ceil(minutos_totales)
        
        resultados.append({
            'Centro': centro,
            'N': n_llamadas,
            'Minuts': minutos_enteros,
            'Interns': n_internos
        })
    
    # Crear DataFrame base con todos los centros
    df_base = pd.DataFrame({
        'Centro': ORDEN_CENTROS,
        f'N_{tipo}': 0,
        f'Minuts_{tipo}': 0,
        f'Interns_{tipo}': 0
    })
    
    df_resultados = pd.DataFrame(resultados)
    
    # Actualizar valores donde existan
    for idx, centro in enumerate(df_base['Centro']):
        centro_buscar = centro
        
        if centro_buscar in df_resultados['Centro'].values:
            fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
            df_base.at[idx, f'N_{tipo}'] = fila['N']
            df_base.at[idx, f'Minuts_{tipo}'] = fila['Minuts'] if 'Minuts' in fila else 0
            df_base.at[idx, f'Interns_{tipo}'] = fila['Interns']
    
    return df_base

def procesar_consultes(files):
    """Procesa archivos de consultes con múltiples hojas"""
    if not files:
        return None
    
    acumulados = defaultdict(lambda: {'llamadas': 0, 'internos_unicos': set()})
    
    for file in files:
        try:
            excel_file = pd.ExcelFile(file)
        
            for hoja in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file, sheet_name=hoja)
                    
                    if 'CENTRE' not in df.columns or 'ID_USUARI' not in df.columns:
                        continue
                    
                    for centro_original in df['CENTRE'].unique():
                        if pd.isna(centro_original):
                            continue
                            
                        df_centro = df[df['CENTRE'] == centro_original]
                        centro = MAPEO_CENTROS.get(centro_original.upper(), centro_original)
                        
                        n_llamadas = len(df_centro)
                        ids_unicos = set(df_centro['ID_USUARI'].dropna().unique())
                        
                        acumulados[centro]['llamadas'] += n_llamadas
                        acumulados[centro]['internos_unicos'].update(ids_unicos)
                        
                except Exception as e:
                    st.warning(f"Error en hoja {hoja}: {e}")
                    continue
            
            # Crear DataFrame con resultados
            df_base = pd.DataFrame({
                'Centro': ORDEN_CENTROS,
                'N_Consultes': 0,
                'Interns_Consultes': 0
            })
            
            for idx, centro in enumerate(df_base['Centro']):
                centro_buscar = centro
                
                if centro_buscar in acumulados:
                    df_base.at[idx, 'N_Consultes'] = acumulados[centro_buscar]['llamadas']
                    df_base.at[idx, 'Interns_Consultes'] = len(acumulados[centro_buscar]['internos_unicos'])
        except Exception as e:
            st.error(f"Error procesando consultes: {e}")

    return df_base

def procesar_reserves(files):
    """Procesa archivos de reserves"""
    if not files:
        return None
    
    # Acumular todos los DataFrames raw para unificar y eliminar duplicados correctamente
    dfs_raw = []
    
    for file in files:
        try:
            df = read_excel_or_csv(file)
            
            # Buscar columnas necesarias
            col_centre = 'Centre' if 'Centre' in df.columns else None
            col_nis = 'NIS/Exp.' if 'NIS/Exp.' in df.columns else None
            
            if not col_centre or not col_nis:
                st.warning(f"⚠️ {file.name}: Columnas necesarias no encontradas")
                continue
            
            if df.empty:
                continue
            
            dfs_raw.append(df)
            
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")
    
    if not dfs_raw:
        return None
    
    # Concatenar todos los DataFrames
    df_combined = pd.concat(dfs_raw, ignore_index=True)
    
    # Buscar columnas
    col_centre = 'Centre' if 'Centre' in df_combined.columns else None
    col_nis = 'NIS/Exp.' if 'NIS/Exp.' in df_combined.columns else None
    
    if not col_centre or not col_nis:
        return None
    
    # Procesar por centro
    resultados = []
    for centro_original in df_combined[col_centre].unique():
        if pd.isna(centro_original):
            continue
            
        df_centro = df_combined[df_combined[col_centre] == centro_original]
        centro = MAPEO_CENTROS.get(str(centro_original).upper(), centro_original)
        
        n_reserves = len(df_centro)
        n_internos = df_centro[col_nis].nunique()  # Elimina duplicados correctamente
        
        resultados.append({
            'Centro': centro,
            'N': n_reserves,
            'Interns': n_internos
        })
    
    # Crear DataFrame base
    df_base = pd.DataFrame({
        'Centro': ORDEN_CENTROS,
        'N_Reserves': 0,
        'Interns_Reserves': 0
    })
    
    df_resultados = pd.DataFrame(resultados)
    
    for idx, centro in enumerate(df_base['Centro']):
        centro_buscar = centro
        
        if centro_buscar in df_resultados['Centro'].values:
            fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
            df_base.at[idx, 'N_Reserves'] = fila['N']
            df_base.at[idx, 'Interns_Reserves'] = fila['Interns']
    
    return df_base

def procesar_video_combinado(files):
    """Procesa archivos combinados con hojas VIDEOTRUCADES y VIDEOVISITES"""
    if not files:
        return None, None
    
    # Acumular DataFrames RAW de todos los archivos para poder eliminar duplicados correctamente
    dfs_videotrucades_raw = []
    dfs_videovisites_raw = []
    
    for file in files:
        try:
            excel_file = pd.ExcelFile(file)
            
            # Procesar hoja VIDEOTRUCADES
            if 'VIDEOTRUCADES' in excel_file.sheet_names:
                df_temp = pd.read_excel(file, sheet_name='VIDEOTRUCADES')
                temp_file = BytesIO()
                df_temp.to_excel(temp_file, index=False)
                temp_file.seek(0)
                temp_file.name = f"videotrucades_{file.name}"
                dfs_videotrucades_raw.append(df_temp)
            
            # Procesar hoja VIDEOVISITES
            if 'VIDEOVISITES' in excel_file.sheet_names:
                df_temp = pd.read_excel(file, sheet_name='VIDEOVISITES')
                temp_file = BytesIO()
                df_temp.to_excel(temp_file, index=False)
                temp_file.seek(0)
                temp_file.name = f"videovisites_{file.name}"
                dfs_videovisites_raw.append(df_temp)
                
        except Exception as e:
            st.error(f"Error procesando archivo combinado {file.name}: {e}")
            continue
    
    # Procesar VIDEOTRUCADES unificando datos
    df_videotrucades_final = None
    if dfs_videotrucades_raw:
        # Concatenar todos los DataFrames raw
        df_combined = pd.concat(dfs_videotrucades_raw, ignore_index=True)
        
        # Buscar columnas necesarias
        col_centre = None
        col_nis = None
        col_duracio = None
        
        for col in df_combined.columns:
            if 'CENTRE' in col.upper():
                col_centre = col
            if 'NIS' in col.upper() or 'EXPEDIENT' in col.upper():
                col_nis = col
            if 'DURAC' in col.upper():
                col_duracio = col
        
        if col_centre and col_nis:
            # Filtrar NaN
            df_combined = df_combined.dropna(subset=[col_centre, col_nis])
            
            # Filtrar por duración si existe
            if col_duracio:
                df_combined = df_combined[df_combined[col_duracio] != '-']
                df_combined = df_combined[df_combined[col_duracio] != ' - ']
                df_combined = df_combined[~df_combined[col_duracio].astype(str).str.strip().eq('-')]
                df_combined['minutos_totales'] = df_combined[col_duracio].apply(parsear_duracion)
                df_combined = df_combined[df_combined['minutos_totales'] > (5/60)]
            
            # Procesar por centro
            resultados = []
            for centro_original in df_combined[col_centre].unique():
                df_centro = df_combined[df_combined[col_centre] == centro_original]
                centro = MAPEO_CENTROS.get(str(centro_original).upper(), centro_original)
                
                n_llamadas = len(df_centro)
                n_internos = df_centro[col_nis].nunique()  # AQUÍ elimina duplicados correctamente
                
                minutos_enteros = 0
                if col_duracio:
                    minutos_totales = sum(df_centro[col_duracio].apply(parsear_duracion))
                    minutos_enteros = math.ceil(minutos_totales)
                
                resultados.append({
                    'Centro': centro,
                    'N': n_llamadas,
                    'Minuts': minutos_enteros,
                    'Interns': n_internos
                })
            
            # Crear DataFrame base
            df_base = pd.DataFrame({
                'Centro': ORDEN_CENTROS,
                'N_Videotrucades': 0,
                'Minuts_Videotrucades': 0,
                'Interns_Videotrucades': 0
            })
            
            df_resultados = pd.DataFrame(resultados)
            
            for idx, centro in enumerate(df_base['Centro']):
                centro_buscar = centro
                
                if centro_buscar in df_resultados['Centro'].values:
                    fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
                    df_base.at[idx, 'N_Videotrucades'] = fila['N']
                    df_base.at[idx, 'Minuts_Videotrucades'] = fila['Minuts'] if 'Minuts' in fila else 0
                    df_base.at[idx, 'Interns_Videotrucades'] = fila['Interns']
            
            df_videotrucades_final = df_base
    
    # Procesar VIDEOVISITES unificando datos (mismo proceso)
    df_videovisites_final = None
    if dfs_videovisites_raw:
        df_combined = pd.concat(dfs_videovisites_raw, ignore_index=True)
        
        col_centre = None
        col_nis = None
        col_duracio = None
        
        for col in df_combined.columns:
            if 'CENTRE' in col.upper():
                col_centre = col
            if 'NIS' in col.upper() or 'EXPEDIENT' in col.upper():
                col_nis = col
            if 'DURAC' in col.upper():
                col_duracio = col
        
        if col_centre and col_nis:
            df_combined = df_combined.dropna(subset=[col_centre, col_nis])
            
            if col_duracio:
                df_combined = df_combined[df_combined[col_duracio] != '-']
                df_combined = df_combined[df_combined[col_duracio] != ' - ']
                df_combined = df_combined[~df_combined[col_duracio].astype(str).str.strip().eq('-')]
                df_combined['minutos_totales'] = df_combined[col_duracio].apply(parsear_duracion)
                df_combined = df_combined[df_combined['minutos_totales'] > (5/60)]
            
            resultados = []
            for centro_original in df_combined[col_centre].unique():
                df_centro = df_combined[df_combined[col_centre] == centro_original]
                centro = MAPEO_CENTROS.get(str(centro_original).upper(), centro_original)
                
                n_llamadas = len(df_centro)
                n_internos = df_centro[col_nis].nunique()
                
                minutos_enteros = 0
                if col_duracio:
                    minutos_totales = sum(df_centro[col_duracio].apply(parsear_duracion))
                    minutos_enteros = math.ceil(minutos_totales)
                
                resultados.append({
                    'Centro': centro,
                    'N': n_llamadas,
                    'Minuts': minutos_enteros,
                    'Interns': n_internos
                })
            
            df_base = pd.DataFrame({
                'Centro': ORDEN_CENTROS,
                'N_Videovisites': 0,
                'Minuts_Videovisites': 0,
                'Interns_Videovisites': 0
            })
            
            df_resultados = pd.DataFrame(resultados)
            
            for idx, centro in enumerate(df_base['Centro']):
                centro_buscar = centro
                
                if centro_buscar in df_resultados['Centro'].values:
                    fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
                    df_base.at[idx, 'N_Videovisites'] = fila['N']
                    df_base.at[idx, 'Minuts_Videovisites'] = fila['Minuts'] if 'Minuts' in fila else 0
                    df_base.at[idx, 'Interns_Videovisites'] = fila['Interns']
            
            df_videovisites_final = df_base
    
    return df_videotrucades_final, df_videovisites_final

def combinar_resultados(dfs_dict):
    """Combina todos los DataFrames de resultados en una tabla final"""
    
    # Empezar con el DataFrame base
    df_final = pd.DataFrame({
        'Centro': ORDEN_CENTROS,
        'Interns': [''] * len(ORDEN_CENTROS),
        'Inici VeD': [''] * len(ORDEN_CENTROS)
    })
    
    # Definir el orden de las columnas según la tabla objetivo
    columnas_orden = []
    
    # Trucades de veu
    if 'Trucades' in dfs_dict:
        df = dfs_dict['Trucades']
        df_final['N_Trucades'] = df['N_Trucades']
        df_final['%Incre_Trucades'] = ''
        df_final['Minuts_Trucades'] = df['Minuts_Trucades']
        df_final['%Incre_Minuts_T'] = ''
        df_final['Interns_Trucades'] = df['Interns_Trucades']
    
    # Videotrucades
    if 'Videotrucades' in dfs_dict:
        df = dfs_dict['Videotrucades']
        df_final['N_Videotrucades'] = df['N_Videotrucades']
        df_final['%Incre_Video'] = ''
        df_final['Minuts_Videotrucades'] = df['Minuts_Videotrucades']
        df_final['%Incre_Minuts_V'] = ''
        df_final['Interns_Videotrucades'] = df['Interns_Videotrucades']
    
    # Videovisites
    if 'Videovisites' in dfs_dict:
        df = dfs_dict['Videovisites']
        df_final['N_Videovisites'] = df['N_Videovisites']
        df_final['%Incre_Visites'] = ''
        df_final['Minuts_Videovisites'] = df['Minuts_Videovisites']
        df_final['%Incre_Minuts_Vi'] = ''
        df_final['Interns_Videovisites'] = df['Interns_Videovisites']
    
    # Consultes Autoservei
    if 'Consultes' in dfs_dict:
        df = dfs_dict['Consultes']
        df_final['N_Consultes'] = df['N_Consultes']

    # Reserves
    if 'Reserves' in dfs_dict:
        df = dfs_dict['Reserves']
        df_final['N_Reserves'] = df['N_Reserves']

    # Alta Digital
    df_final['Alta_Digital'] = ''
    
    # Añadir fila de totales
    totales = {'Centro': 'TOTAL', 'Interns': '', 'Inici VeD': '', 'Alta_Digital': '', 'N_Alta': ''}
    
    for col in df_final.columns:
        if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_'):
            totales[col] = df_final[col].apply(lambda x: x if isinstance(x, (int, float)) else 0).sum()
        elif col.startswith('%'):
            totales[col] = ''
    
    df_totales = pd.DataFrame([totales])
    df_final = pd.concat([df_final, df_totales], ignore_index=True)
    
    return df_final

# ===================== INTERFAZ STREAMLIT =====================

def main():
    st.title("🎯 Herramienta de Análisis VED")
    
    st.markdown("""
    ### 📊 Càlcul de Volumetries per Mes
    
    Aquesta eina permet tractar les dades de **Trucades**, **Videotrucades**, **Videovisites**, 
    **Consultes Autoservei** i **Reserves** corresponents al servei de VED.
    
    L'eina processa automàticament els arxius proporcionats, un cop pujats, i clicant en -Processar Dades- genera un informe consolidat amb les 
    estadístiques per cada centre penitenciari i educatiu.
    """)
    
    # --- INICIALIZACIÓN DE SESSION_STATE (NUEVO) ---
    if 'processing_complete' not in st.session_state:
        st.session_state['processing_complete'] = False
    if 'resultados_ved' not in st.session_state:
        st.session_state['resultados_ved'] = {}
    if 'df_final_ved' not in st.session_state:
        st.session_state['df_final_ved'] = None
    if 'df_acumulada_ved' not in st.session_state:
        st.session_state['df_acumulada_ved'] = None
    if 'df_acumulada_manual_confirmada' not in st.session_state:
        st.session_state['df_acumulada_manual_confirmada'] = None
    if 'df_actual_manual_confirmada' not in st.session_state:
        st.session_state['df_actual_manual_confirmada'] = None
    if 'df_acum_manual_modo2_confirmada' not in st.session_state:
        st.session_state['df_acum_manual_modo2_confirmada'] = None
    # -----------------------------------------------
    
    
    st.markdown("### ⚙️ Mode d'entrada de dades")
    modo_manual = st.toggle("Activar mode manual (sense pujar arxius)", value=False, key="modo_manual")

    st.divider()
    
    #MODO CON DOS INSERTAR----------------- (se puede hacer acortar reutilizando lo de la acumulada)
    if not modo_manual:
        # Crear columnas para los file uploaders
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📞 Trucades de veu")
            st.caption("Formats acceptats: .xlsx, .csv")
            trucades_files = st.file_uploader(
                "Selecciona arxius de Trucades",
                type=['xlsx', 'xls', 'csv'],
                accept_multiple_files=True,
                key="trucades"
            )
            
            st.markdown("#### 🎥 Videotrucades i Videovisites")
            st.caption("Format acceptat: .xlsx amb 2 fulles (VIDEOTRUCADES i VIDEOVISITES)")
            video_combined_files = st.file_uploader(
                "Selecciona l'arxiu combinat de Video",
                type=['xlsx', 'xls'],
                accept_multiple_files=True,
                key="video_combined"
            )
        
        with col2:
            st.markdown("#### 💻 Consultes Autoservei")
            st.caption("Format acceptat: .xlsx (amb múltiples fulles)")
            consultes_files = st.file_uploader(
                "Selecciona l'arxiu de Consultes",
                type=['xlsx', 'xls'],
                accept_multiple_files=True,
                key="consultes"
            )
            
            st.markdown("#### 📅 Reserves")
            st.caption("Formats acceptats: .xlsx, .csv")
            reserves_files = st.file_uploader(
                "Selecciona arxius de Reserves",
                type=['xlsx', 'xls', 'csv'],
                accept_multiple_files=True,
                key="reserves"
            )

            st.divider()

            # Sección para tabla acumulada (opcional)
            st.markdown("#### 📊 Taula Acumulada Anterior (Opcional)")
            st.caption("Tria com vols proporcionar les dades acumulades del mes anterior:")

            # Pestañas para elegir método
            tab_acum1, tab_acum2 = st.tabs(["📋 Enganxar des d'Excel", "📁 Pujar arxiu Excel"])

            with tab_acum1:
                st.markdown("**Copia les dades des d'Excel i enganxa-les directament a la taula:**")
                st.caption("Pots copiar des d'Excel (Ctrl+C) i enganxar aquí (Ctrl+V)")
                
                # Crear DataFrame template vacío con las columnas correctas
                # Crear DataFrame template vacío con las columnas correctas
                columnas_template = [
                    'Centro', 'Interns', 'Inici VeD',
                    'N_Trucades', '%Incre_Trucades', 'Minuts_Trucades', '%Incre_Minuts_T', 'Interns_Trucades',
                    'N_Videotrucades', '%Incre_Video', 'Minuts_Videotrucades', '%Incre_Minuts_V', 'Interns_Videotrucades',
                    'N_Videovisites', '%Incre_Visites', 'Minuts_Videovisites', '%Incre_Minuts_Vi', 'Interns_Videovisites',
                    'N_Consultes',
                    'N_Reserves',
                    'Alta_Digital'
                ]
                
                # DataFrame vacío con 21 filas (20 centros + TOTAL)
                num_centros = len(ORDEN_CENTROS)
                df_template = pd.DataFrame('', index=range(num_centros + 1), columns=columnas_template)

                # Pre-llenar la columna Centro con los nombres
                for idx, centro in enumerate(ORDEN_CENTROS):
                    df_template.at[idx, 'Centro'] = centro

                # Añadir fila TOTAL al final
                df_template.at[num_centros, 'Centro'] = 'TOTAL'
                
                # Data editor editable
                df_acumulada_manual = st.data_editor(
                    df_template,
                    use_container_width=True,
                    height=400,
                    key="acumulada_manual",
                    num_rows="fixed"
                )
                
                # Guardar en session_state si hay datos
                if st.button("✅ Confirmar dades enganxades", key="confirmar_manual"):
                    # Verificar que hay datos (al menos una celda numérica rellena)
                    columnas_numericas = [col for col in df_acumulada_manual.columns 
                                        if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_')]
                    hay_datos = False
                    for col in columnas_numericas:
                        if df_acumulada_manual[col].astype(str).str.strip().ne('').any():
                            hay_datos = True
                            break
                    
                    if hay_datos:
                        st.session_state['df_acumulada_manual_confirmada'] = df_acumulada_manual.copy()
                        st.success("✅ Dades de la taula acumulada confirmades!")
                    else:
                        st.warning("⚠️ No s'han detectat dades. Enganxa les dades des d'Excel.")

            with tab_acum2:

                acumulada_file = st.file_uploader(
                    "Selecciona l'arxiu de Taula Acumulada",
                    type=['xlsx', 'xls'],
                    accept_multiple_files=False,
                    key="acumulada"
                )

    else:
    # MODO MANUAL: Insertar tabla actual y acumulada
        st.markdown("### 📋 Mode Manual: Enganxa les taules directament")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 Taula Actual (Dades del mes)")
            st.caption("Enganxa les dades actuals des d'Excel")
            
            # Crear DataFrame template para tabla actual
            columnas_actual = [
                'Centro', 'Interns', 'Inici VeD',
                'N_Trucades', '%Incre_Trucades', 'Minuts_Trucades', '%Incre_Minuts_T', 'Interns_Trucades',
                'N_Videotrucades', '%Incre_Video', 'Minuts_Videotrucades', '%Incre_Minuts_V', 'Interns_Videotrucades',
                'N_Videovisites', '%Incre_Visites', 'Minuts_Videovisites', '%Incre_Minuts_Vi', 'Interns_Videovisites',
                'N_Consultes',
                'N_Reserves',
                'Alta_Digital'
            ]
            
            num_centros = len(ORDEN_CENTROS)
            df_template_actual = pd.DataFrame('', index=range(num_centros + 1), columns=columnas_actual)
            
            for idx, centro in enumerate(ORDEN_CENTROS):
                df_template_actual.at[idx, 'Centro'] = centro
            df_template_actual.at[num_centros, 'Centro'] = 'TOTAL'
            
            df_actual_manual = st.data_editor(
                df_template_actual,
                use_container_width=True,
                height=500,
                key="actual_manual",
                num_rows="fixed"
            )
            
            if st.button("✅ Confirmar Taula Actual", key="confirmar_actual_manual"):
                # Limpiar y convertir datos
                columnas_numericas = [col for col in df_actual_manual.columns 
                                    if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_')]
                
                df_limpio = df_actual_manual.copy()
                for col in columnas_numericas:
                    df_limpio[col] = df_limpio[col].astype(str).str.replace(' ', '').str.strip()
                    df_limpio[col] = pd.to_numeric(df_limpio[col], errors='coerce').fillna(0)
                
                st.session_state['df_actual_manual_confirmada'] = df_limpio
                st.success("✅ Taula Actual confirmada!")
        
        with col2:
            st.markdown("#### 📊 Taula Acumulada Anterior")
            st.caption("Enganxa les dades acumulades del mes anterior")
            
            num_centros = len(ORDEN_CENTROS)
            df_template_acum_manual = pd.DataFrame('', index=range(num_centros + 1), columns=columnas_actual)
            
            for idx, centro in enumerate(ORDEN_CENTROS):
                df_template_acum_manual.at[idx, 'Centro'] = centro
            df_template_acum_manual.at[num_centros, 'Centro'] = 'TOTAL'
            
            df_acum_manual_modo2 = st.data_editor(
                df_template_acum_manual,
                use_container_width=True,
                height=500,
                key="acum_manual_modo2",
                num_rows="fixed"
            )
            
            if st.button("✅ Confirmar Taula Acumulada", key="confirmar_acum_manual_modo2"):
                # Limpiar y convertir datos
                columnas_numericas = [col for col in df_acum_manual_modo2.columns 
                                    if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_')]
                
                df_limpio = df_acum_manual_modo2.copy()
                for col in columnas_numericas:
                    # Eliminar espacios y reemplazar . y , por nada (asumir que son separadores de miles)
                    df_limpio[col] = df_limpio[col].astype(str).str.replace(' ', '').str.replace('.', '').str.replace(',', '').str.strip()
                    df_limpio[col] = pd.to_numeric(df_limpio[col], errors='coerce').fillna(0)
                
                st.session_state['df_acum_manual_modo2_confirmada'] = df_limpio
                st.success("✅ Taula Acumulada confirmada!")
        
    st.divider()
    
    # Botón de procesar
    if st.button("🚀 Processar Dades", type="primary", use_container_width=True):
        
        # MODO MANUAL
        if modo_manual:
            # Verificar que hay tablas confirmadas
            if st.session_state.get('df_actual_manual_confirmada') is None:
                st.error("⚠️ Si us plau, confirma la Taula Actual.")
                st.session_state['processing_complete'] = False
                return
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Processant taules manuals...")
            
            df_actual = st.session_state['df_actual_manual_confirmada'].copy()
            df_anterior = st.session_state.get('df_acum_manual_modo2_confirmada')
            
            st.session_state['resultados_ved'] = {}
            st.session_state['df_final_ved'] = df_actual
            
            # Calcular acumulada si existe tabla anterior
            if df_anterior is not None:
                df_acumulada = calcular_tabla_acumulada(df_actual, df_anterior)
                st.session_state['df_acumulada_ved'] = df_acumulada
            else:
                st.session_state['df_acumulada_ved'] = None
            
            progress_bar.progress(1.0)
            st.session_state['processing_complete'] = True
            st.success("✅ Processament completat amb èxit!")
            st.rerun()
        
        else:
            # Verificar que hay al menos un archivo
            if not any([trucades_files, video_combined_files, consultes_files, reserves_files]):
                st.error("⚠️ Si us plau, carrega almenys un arxiu per processar.")
                # Borrar estado si estaba en True
                st.session_state['processing_complete'] = False 
                st.session_state['resultados_ved'] = {}
                st.session_state['df_final_ved'] = None
                return
            
            # Crear barra de progreso
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            resultados = {}
            progress = 0
            total_steps = 5
            
            # Procesar Trucades
            if trucades_files:
                status_text.text("Processant Trucades de veu...")
                df_trucades = procesar_trucades_video(trucades_files, "Trucades")
                if df_trucades is not None:
                    resultados['Trucades'] = df_trucades
                progress += 1
                progress_bar.progress(progress / total_steps)
            
            # Procesar Videotrucades
            # Procesar Videotrucades y Videovisites desde archivo combinado
            # Procesar Videotrucades y Videovisites desde archivos combinados
            if video_combined_files:
                status_text.text("Processant Videotrucades i Videovisites...")
                df_videotrucades, df_videovisites = procesar_video_combinado(video_combined_files)
                
                if df_videotrucades is not None:
                    resultados['Videotrucades'] = df_videotrucades
                if df_videovisites is not None:
                    resultados['Videovisites'] = df_videovisites
                
                progress += 1
                progress_bar.progress(progress / total_steps)
            
            # Procesar Consultes
            if consultes_files:
                status_text.text("Processant Consultes Autoservei...")
                df_consultes = procesar_consultes(consultes_files)
                if df_consultes is not None:
                    resultados['Consultes'] = df_consultes
                progress += 1
                progress_bar.progress(progress / total_steps)
            
            # Procesar Reserves
            if reserves_files:
                status_text.text("Processant Reserves...")
                df_reserves = procesar_reserves(reserves_files)
                if df_reserves is not None:
                    resultados['Reserves'] = df_reserves
                progress += 1
                progress_bar.progress(progress / total_steps)
            
            progress_bar.progress(1.0)
            status_text.text("Generant informe final...")
            
            # --- GUARDAR RESULTADOS EN SESSION_STATE (MODIFICADO) ---
            if resultados:
                df_actual = combinar_resultados(resultados)
                st.session_state['resultados_ved'] = resultados
                st.session_state['df_final_ved'] = df_actual
                
                # Procesar tabla acumulada si existe
                # Procesar tabla acumulada (desde Excel o manual)
                df_anterior = None

                # Prioridad 1: Desde archivo Excel
                if acumulada_file:
                    status_text.text("Processant taula acumulada des d'arxiu...")
                    df_anterior = extraer_tabla_acumulada_anterior(acumulada_file)

                # Prioridad 2: Desde datos manuales pegados
                elif st.session_state.get('df_acumulada_manual_confirmada') is not None:
                    status_text.text("Processant taula acumulada enganxada...")
                    df_anterior = st.session_state['df_acumulada_manual_confirmada'].copy()
                    
                    # Limpiar datos pegados (eliminar espacios y convertir a números)
                    columnas_numericas = [col for col in df_anterior.columns 
                                        if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_')]
                    
                    for col in columnas_numericas:
                        # Eliminar espacios en blanco de los números antes de convertir
                        df_anterior[col] = df_anterior[col].astype(str).str.replace(' ', '').str.strip()
                        df_anterior[col] = pd.to_numeric(df_anterior[col], errors='coerce').fillna(0)
                    
                    st.success("✅ Taula acumulada carregada des de dades enganxades")

                # Calcular acumulada si hay datos anteriores
                if df_anterior is not None:
                    df_acumulada = calcular_tabla_acumulada(df_actual, df_anterior)
                    st.session_state['df_acumulada_ved'] = df_acumulada
                else:
                    st.session_state['df_acumulada_ved'] = None
                
                st.session_state['processing_complete'] = True
                
                st.success("✅ Processament completat amb èxit!")
                st.rerun()
            else:
                st.error("⚠️ No s'ha pogut processar cap dada. Assegura't que els arxius són correctes.")
                st.session_state['processing_complete'] = False
            # ------------------------------------------------------
        
    
    # --- MOSTRAR PESTAÑAS SÓLO SI EL PROCESAMIENTO ESTÁ COMPLETO (NUEVO SCOPE) ---
    if st.session_state['processing_complete']:
        
        resultados = st.session_state['resultados_ved']
        df_final = st.session_state['df_final_ved']
        
        # Crear tabs para Tablas y Visualizar
        tab1, tab2 = st.tabs(["📊 Taules", "📈 Visualitzar"])
        
        with tab1:
            st.markdown("### Selecciona la taula a visualitzar")
            
            col1, col2 = st.columns([1, 3])
            
            with col1:
                # Selector de tabla
                opciones_tabla = []
                dataframes_disponibles = {}
                
                # Añadir las tablas individuales disponibles
                if 'Trucades' in resultados:
                    opciones_tabla.append("Trucades de veu")
                    dataframes_disponibles["Trucades de veu"] = resultados['Trucades']
                if 'Videotrucades' in resultados:
                    opciones_tabla.append("Videotrucades")
                    dataframes_disponibles["Videotrucades"] = resultados['Videotrucades']
                if 'Videovisites' in resultados:
                    opciones_tabla.append("Videovisites")
                    dataframes_disponibles["Videovisites"] = resultados['Videovisites']
                if 'Consultes' in resultados:
                    opciones_tabla.append("Consultes Autoservei")
                    dataframes_disponibles["Consultes Autoservei"] = resultados['Consultes']
                if 'Reserves' in resultados:
                    opciones_tabla.append("Reserves")
                    dataframes_disponibles["Reserves"] = resultados['Reserves']
                
                # Añadir tabla general
                opciones_tabla.append("Taula General (Actual)")
                dataframes_disponibles["Taula General (Actual)"] = df_final
                
                # Añadir tabla acumulada si existe
                if st.session_state.get('df_acumulada_ved') is not None:
                    opciones_tabla.append("Taula General (Acumulada)")
                    dataframes_disponibles["Taula General (Acumulada)"] = st.session_state['df_acumulada_ved']
                
                # El selectbox ahora está fuera del bloque if del botón y mantiene el estado.
                tabla_seleccionada = st.selectbox(
                    "Tipus de taula:",
                    opciones_tabla,
                    index=len(opciones_tabla)-1  # Por defecto la General
                )
            
            with col2:
                # Mostrar la tabla seleccionada
                df_mostrar = dataframes_disponibles[tabla_seleccionada].copy()
                
                # Añadir fila de totales si no es la tabla general (que ya la tiene)
                # Añadir fila de totales si no es la tabla general (que ya la tiene)
                if tabla_seleccionada != "Taula General (Actual)" and tabla_seleccionada != "Taula General (Acumulada)":
                    totales = {}
                    for col in df_mostrar.columns:
                        if col == 'Centro':
                            totales[col] = 'TOTAL'
                        elif col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_'):
                            totales[col] = df_mostrar[col].apply(lambda x: x if isinstance(x, (int, float)) else 0).sum()
                        else:
                            totales[col] = ''
                    df_mostrar = pd.concat([df_mostrar, pd.DataFrame([totales])], ignore_index=True)
                
                df_mostrar_formatted = formatear_numeros_df(df_mostrar)
                st.dataframe(df_mostrar_formatted, use_container_width=True, height=600)
                
                # Botón de descarga para la tabla seleccionada
                output = BytesIO()
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_mostrar_formatted.to_excel(writer, sheet_name=tabla_seleccionada, index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets[tabla_seleccionada]
                    
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'center',
                        'align': 'center',
                        'border': 1,
                        'bg_color': '#D7E4BD'
                    })
                    
                    for col_num, value in enumerate(df_mostrar.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    worksheet.set_column('A:Z', 12)
                    if 'Centro' in df_mostrar.columns:
                        worksheet.set_column('A:A', 20)
                
                st.download_button(
                    label=f"⬇️ Descarregar {tabla_seleccionada}",
                    data=output.getvalue(),
                    file_name=f"{tabla_seleccionada.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with tab2:
            st.markdown("### 📊 Visualitzacions de les dades")
            
            # Preparar datos para visualizaciones
            datos_viz = df_final[df_final['Centro'] != 'TOTAL'].copy()
            
            # 1. Gráfico comparativo por tipo de servicio
            st.markdown("#### 1. Comparació de volum per tipus de servei")
            
            col1, col2 = st.columns([3, 1])
            with col2:
                metrica_comparar = st.radio(
                    "Mètrica a comparar:",
                    ["Número de serveis (N)", "Minuts", "Interns únics"],
                    key="metrica1"
                )
            
            with col1:
                # Recopilar datos según la métrica seleccionada
                data_comparison = []
                
                if metrica_comparar == "Número de serveis (N)":
                    columnas = [col for col in datos_viz.columns if col.startswith('N_') and col != 'N_Alta']
                elif metrica_comparar == "Minuts":
                    columnas = [col for col in datos_viz.columns if col.startswith('Minuts_')]
                else:
                    columnas = [col for col in datos_viz.columns if col.startswith('Interns_')]
                
                for col in columnas:
                    servicio = col.split('_')[1]
                    for idx, row in datos_viz.iterrows():
                        if row[col] > 0:  # Solo incluir valores > 0
                            data_comparison.append({
                                'Centro': row['Centro'],
                                'Servei': servicio,
                                'Valor': row[col]
                            })
                
                if data_comparison:
                    df_comparison = pd.DataFrame(data_comparison)
                    fig = px.bar(df_comparison, 
                                x='Centro', 
                                y='Valor', 
                                color='Servei',
                                title=f"{metrica_comparar} per centre i servei",
                                labels={'Valor': metrica_comparar, 'Centro': 'Centre'},
                                color_discrete_map={
                                    'Trucades': '#1f77b4',
                                    'Videotrucades': '#ff7f0e', 
                                    'Videovisites': '#2ca02c',
                                    'Consultes': '#d62728',
                                    'Reserves': '#9467bd'
                                })
                    fig.update_xaxes(tickangle=45) # <--- CORRECCIÓN DE update_xaxis a update_xaxes
                    fig.update_layout(height=500)
                    st.plotly_chart(fig, use_container_width=True)
            
            # 2. Top 10 Centros por actividad
            st.markdown("#### 2. Top 10 centres amb més activitat")
            
            col1, col2 = st.columns([3, 1])
            with col2:
                servicio_top = st.selectbox(
                    "Servei:",
                    [s for s in ['Trucades', 'Videotrucades', 'Videovisites', 'Consultes', 'Reserves'] 
                     if s in resultados],
                    key="servicio_top"
                )
            
            with col1:
                col_n = f'N_{servicio_top}'
                col_interns = f'Interns_{servicio_top}'
                
                if col_n in datos_viz.columns:
                    # Preparar datos para el top 10
                    top_data = datos_viz[['Centro', col_n, col_interns]].copy()
                    top_data = top_data.sort_values(col_n, ascending=False).head(10)
                    
                    # Crear gráfico de barras doble
                    fig = make_subplots(
                        rows=1, cols=2,
                        subplot_titles=(f"Número de {servicio_top}", f"Interns únics - {servicio_top}")
                    )
                    
                    fig.add_trace(
                        go.Bar(x=top_data['Centro'], y=top_data[col_n], 
                              name=f"N {servicio_top}",
                              marker_color='lightblue'),
                        row=1, col=1
                    )
                    
                    fig.add_trace(
                        go.Bar(x=top_data['Centro'], y=top_data[col_interns], 
                              name="Interns",
                              marker_color='lightcoral'),
                        row=1, col=2
                    )
                    
                    fig.update_xaxes(tickangle=45)
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            
            # 3. Distribución de minutos (para servicios con minutos)
            st.markdown("#### 3. Distribució de minuts per centre")
            
            servicios_con_minutos = []
            if 'Trucades' in resultados and 'Minuts_Trucades' in datos_viz.columns:
                servicios_con_minutos.append('Trucades')
            if 'Videotrucades' in resultados and 'Minuts_Videotrucades' in datos_viz.columns:
                servicios_con_minutos.append('Videotrucades')
            if 'Videovisites' in resultados and 'Minuts_Videovisites' in datos_viz.columns:
                servicios_con_minutos.append('Videovisites')
            
            if servicios_con_minutos:
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    tipo_grafico = st.radio(
                        "Tipus de gràfic:",
                        ["Gràfic circular (Pie)", "Gràfic de barres apilades"],
                        key="tipo_minutos"
                    )
                    
                    incluir_zeros = st.checkbox("Incloure centres sense activitat", value=False)
                
                with col2:
                    # Preparar datos de minutos
                    minutos_data = []
                    for servicio in servicios_con_minutos:
                        col_min = f'Minuts_{servicio}'
                        for idx, row in datos_viz.iterrows():
                            if incluir_zeros or row[col_min] > 0:
                                minutos_data.append({
                                    'Centro': row['Centro'],
                                    'Servei': servicio,
                                    'Minuts': row[col_min]
                                })
                    
                    df_minutos = pd.DataFrame(minutos_data)
                    
                    if tipo_grafico == "Gràfic circular (Pie)":
                        # Agrupar por servicio
                        df_pie = df_minutos.groupby('Servei')['Minuts'].sum().reset_index()
                        fig = px.pie(df_pie, values='Minuts', names='Servei',
                                   title="Distribució total de minuts per servei")
                    else:
                        fig = px.bar(df_minutos, x='Centro', y='Minuts', color='Servei',
                                   title="Minuts per centre i servei",
                                   labels={'Minuts': 'Minuts totals', 'Centro': 'Centre'})
                        fig.update_xaxes(tickangle=45) # <--- CORRECCIÓN DE update_xaxis a update_xaxes
                    
                    fig.update_layout(height=500)
                    st.plotly_chart(fig, use_container_width=True)
            
            # 4. Análisis de eficiencia
            st.markdown("#### 4. Anàlisi d'eficiència (Serveis per intern)")
            
            # Calcular ratio de servicios por interno
            eficiencia_data = []
            for idx, row in datos_viz.iterrows():
                for servicio in ['Trucades', 'Videotrucades', 'Videovisites', 'Consultes', 'Reserves']:
                    col_n = f'N_{servicio}'
                    col_interns = f'Interns_{servicio}'
                    if col_n in row and col_interns in row and row[col_interns] > 0:
                        ratio = row[col_n] / row[col_interns]
                        if ratio > 0:
                            eficiencia_data.append({
                                'Centro': row['Centro'],
                                'Servei': servicio,
                                'Ratio': ratio
                            })
            
            if eficiencia_data:
                df_eficiencia = pd.DataFrame(eficiencia_data)
                
                # Heatmap de eficiencia
                pivot_eficiencia = df_eficiencia.pivot_table(
                    values='Ratio', 
                    index='Centro', 
                    columns='Servei', 
                    fill_value=0
                )
                
                fig = px.imshow(pivot_eficiencia.T,
                              labels=dict(x="Centre", y="Servei", color="Serveis/Intern"),
                              title="Mapa de calor: Serveis per intern",
                              color_continuous_scale='YlOrRd',
                              aspect="auto")
                fig.update_xaxes(tickangle=45) # <--- CORRECCIÓN DE update_xaxis a update_xaxes
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            
            # Métricas resumen
            st.markdown("---")
            st.markdown("### 📊 Resum estadístic")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_centres = len(datos_viz)
                st.metric("Centres actius", total_centres)
            
            with col2:
                total_serveis = sum([datos_viz[col].sum() for col in datos_viz.columns 
                                   if col.startswith('N_') and col != 'N_Alta'])
                st.metric("Total serveis", f"{int(total_serveis):,}")
            
            with col3:
                if servicios_con_minutos:
                    total_minuts = sum([datos_viz[f'Minuts_{s}'].sum() for s in servicios_con_minutos])
                    st.metric("Total minuts", f"{int(total_minuts):,}")
            
            with col4:
                # Calcular internos únicos aproximados (el máximo de cada centro)
                interns_cols = [col for col in datos_viz.columns if col.startswith('Interns_')]
                if interns_cols:
                    # Nota: Esto es una aproximación, no suma de internos únicos reales entre servicios
                    max_interns = datos_viz[interns_cols].max(axis=1).sum()
                    st.metric("Interns (aprox.)", f"{int(max_interns):,}")


if __name__ == "__main__":
    main()