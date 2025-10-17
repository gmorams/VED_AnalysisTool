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
    "CP Terrassa",
    "CE Can Llupià",
    "CE L'Alzina",
    "CE Montilivi",
    "CE Til.lers",
    "UT Til.lers",
    "CE Oriol Badia",
    "CE Folch i Torres",
    "CE El Segre"
]

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
    
    resultados = []
    
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
            
            # Filtrar NaN
            df = df.dropna(subset=[col_centre, col_nis])
            
            # Filtrar por duración si existe
            if col_duracio:
                # Primero eliminar filas con "-" en duración
                df = df[df[col_duracio] != '-']
                df = df[df[col_duracio] != ' - ']  # Por si tiene espacios
                df = df[~df[col_duracio].astype(str).str.strip().eq('-')]  # Más robusto
                
                # Luego aplicar el filtro de 5 segundos
                df['minutos_totales'] = df[col_duracio].apply(parsear_duracion)
                df = df[df['minutos_totales'] > (5/60)]
            
            # Procesar por centro
            for centro_original in df[col_centre].unique():
                df_centro = df[df[col_centre] == centro_original]
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
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")
    
    if not resultados:
        return None
    
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
        # Ajustar para CP Lledoners (Pilot M1)
        centro_buscar = "CP Lledoners" if centro == "CP Lledoners (Pilot M1)" else centro
        
        if centro_buscar in df_resultados['Centro'].values:
            fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
            df_base.at[idx, f'N_{tipo}'] = fila['N']
            df_base.at[idx, f'Minuts_{tipo}'] = fila['Minuts'] if 'Minuts' in fila else 0
            df_base.at[idx, f'Interns_{tipo}'] = fila['Interns']
    
    return df_base

def procesar_consultes(file):
    """Procesa archivo de consultes con múltiples hojas"""
    if not file:
        return None
    
    try:
        excel_file = pd.ExcelFile(file)
        acumulados = defaultdict(lambda: {'llamadas': 0, 'internos_unicos': set()})
        
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
            centro_buscar = "CP Lledoners" if centro == "CP Lledoners (Pilot M1)" else centro
            
            if centro_buscar in acumulados:
                df_base.at[idx, 'N_Consultes'] = acumulados[centro_buscar]['llamadas']
                df_base.at[idx, 'Interns_Consultes'] = len(acumulados[centro_buscar]['internos_unicos'])
        
        return df_base
        
    except Exception as e:
        st.error(f"Error procesando consultes: {e}")
        return None

def procesar_reserves(files):
    """Procesa archivos de reserves"""
    if not files:
        return None
    
    resultados = []
    
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
                
            centro_original = df[col_centre].iloc[0]
            centro = MAPEO_CENTROS.get(centro_original.upper(), centro_original)
            
            n_reserves = len(df)
            n_internos = df[col_nis].nunique()
            
            resultados.append({
                'Centro': centro,
                'N': n_reserves,
                'Interns': n_internos
            })
            
        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")
    
    if not resultados:
        return None
    
    # Crear DataFrame base
    df_base = pd.DataFrame({
        'Centro': ORDEN_CENTROS,
        'N_Reserves': 0,
        'Interns_Reserves': 0
    })
    
    df_resultados = pd.DataFrame(resultados)
    
    for idx, centro in enumerate(df_base['Centro']):
        centro_buscar = "CP Lledoners" if centro == "CP Lledoners (Pilot M1)" else centro
        
        if centro_buscar in df_resultados['Centro'].values:
            fila = df_resultados[df_resultados['Centro'] == centro_buscar].iloc[0]
            df_base.at[idx, 'N_Reserves'] = fila['N']
            df_base.at[idx, 'Interns_Reserves'] = fila['Interns']
    
    return df_base

def procesar_video_combinado(file):
    """Procesa archivo combinado con hojas VIDEOTRUCADES y VIDEOVISITES"""
    if not file:
        return None, None
    
    try:
        excel_file = pd.ExcelFile(file)
        
        df_videotrucades = None
        df_videovisites = None
        
        # Procesar hoja VIDEOTRUCADES
        if 'VIDEOTRUCADES' in excel_file.sheet_names:
            df_temp = pd.read_excel(file, sheet_name='VIDEOTRUCADES')
            # Convertir a lista de un solo archivo para reutilizar función existente
            temp_file = BytesIO()
            df_temp.to_excel(temp_file, index=False)
            temp_file.seek(0)
            temp_file.name = "videotrucades_temp.xlsx"
            
            df_videotrucades = procesar_trucades_video([temp_file], "Videotrucades")
        
        # Procesar hoja VIDEOVISITES
        if 'VIDEOVISITES' in excel_file.sheet_names:
            df_temp = pd.read_excel(file, sheet_name='VIDEOVISITES')
            temp_file = BytesIO()
            df_temp.to_excel(temp_file, index=False)
            temp_file.seek(0)
            temp_file.name = "videovisites_temp.xlsx"
            
            df_videovisites = procesar_trucades_video([temp_file], "Videovisites")
        
        return df_videotrucades, df_videovisites
        
    except Exception as e:
        st.error(f"Error procesando archivo combinado: {e}")
        return None, None

def combinar_resultados(dfs_dict):
    """Combina todos los DataFrames de resultados en una tabla final"""
    
    # Empezar con el DataFrame base
    df_final = pd.DataFrame({
        'Interns': [''] * len(ORDEN_CENTROS),
        'Inici VeD': [''] * len(ORDEN_CENTROS),
        'Centro': ORDEN_CENTROS
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
        df_final['Interns_Consultes'] = df['Interns_Consultes']
    
    # Reserves
    if 'Reserves' in dfs_dict:
        df = dfs_dict['Reserves']
        df_final['N_Reserves'] = df['N_Reserves']
        df_final['Interns_Reserves'] = df['Interns_Reserves']
    
    # Alta Digital
    df_final['Alta_Digital'] = ''
    df_final['N_Alta'] = ''
    
    # Añadir fila de totales
    totales = {'Interns': '', 'Inici VeD': '', 'Centro': 'TOTAL', 'Alta_Digital': '', 'N_Alta': ''}
    
    for col in df_final.columns:
        if col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_'):
            if col not in ['N_Alta']:
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
    # -----------------------------------------------
    
    st.divider()
    
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
        video_combined_file = st.file_uploader(
            "Selecciona l'arxiu combinat de Video",
            type=['xlsx', 'xls'],
            accept_multiple_files=False,
            key="video_combined"
        )
    
    with col2:
        st.markdown("#### 💻 Consultes Autoservei")
        st.caption("Format acceptat: .xlsx (amb múltiples fulles)")
        consultes_file = st.file_uploader(
            "Selecciona l'arxiu de Consultes",
            type=['xlsx', 'xls'],
            accept_multiple_files=False,
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
    
    # Botón de procesar
    if st.button("🚀 Processar Dades", type="primary", use_container_width=True):
        
        # Verificar que hay al menos un archivo
        if not any([trucades_files, video_combined_file, consultes_file, reserves_files]):
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
        if video_combined_file:
            status_text.text("Processant Videotrucades i Videovisites...")
            df_videotrucades, df_videovisites = procesar_video_combinado(video_combined_file)
            
            if df_videotrucades is not None:
                resultados['Videotrucades'] = df_videotrucades
            if df_videovisites is not None:
                resultados['Videovisites'] = df_videovisites
            
            progress += 1
            progress_bar.progress(progress / total_steps)
        
        # Procesar Consultes
        if consultes_file:
            status_text.text("Processant Consultes Autoservei...")
            df_consultes = procesar_consultes(consultes_file)
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
            st.session_state['resultados_ved'] = resultados
            st.session_state['df_final_ved'] = combinar_resultados(resultados)
            st.session_state['processing_complete'] = True
            
            st.success("✅ Processament completat amb èxit!")
            st.rerun() # Fuerza la re-ejecución para salir de este bloque y mostrar las pestañas
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
                opciones_tabla.append("Taula General")
                dataframes_disponibles["Taula General"] = df_final
                
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
                if tabla_seleccionada != "Taula General":
                    totales = {}
                    for col in df_mostrar.columns:
                        if col == 'Centro':
                            totales[col] = 'TOTAL'
                        elif col.startswith('N_') or col.startswith('Minuts_') or col.startswith('Interns_'):
                            totales[col] = df_mostrar[col].sum()
                        else:
                            totales[col] = 0
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