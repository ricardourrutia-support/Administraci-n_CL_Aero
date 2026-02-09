import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import re

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto (XLSX)", layout="wide")

st.title("✈️ Gestor de Turnos y Cobertura (Versión Excel)")
st.markdown("""
Esta aplicación procesa tus archivos **Excel (.xlsx)** originales. 
Selecciona el mes a analizar y la aplicación intentará adivinar qué hoja del Excel corresponde.
""")

# --- FUNCIONES AUXILIARES ---

def parse_time_range(time_str):
    """Convierte strings como '09:00 - 20:00' en una lista de horas (0-23)."""
    if pd.isna(time_str) or str(time_str).lower().strip() in ['libre', 'nan', 'dia libre', 'dias libres', 'l', 'x']:
        return []
    
    try:
        # Limpieza de texto basura
        clean_str = str(time_str).lower()
        clean_str = clean_str.replace(" diurno", "").replace(" nocturno", "").strip()
        
        # Manejo de formatos variados
        parts = clean_str.split('-')
        if len(parts) != 2:
            return []
        
        start_str = parts[0].strip()
        end_str = parts[1].strip()
        
        formats = ["%H:%M:%S", "%H:%M", "%H"]
        
        start_dt = None
        end_dt = None
        
        for fmt in formats:
            try:
                if not start_dt: start_dt = datetime.strptime(start_str, fmt)
            except: pass
            try:
                if not end_dt: end_dt = datetime.strptime(end_str, fmt)
            except: pass
            
        if not start_dt or not end_dt:
            return []
            
        start_h = start_dt.hour
        end_h = end_dt.hour
        
        # Lógica de rango (Ej: 9 a 20)
        if end_h > start_h:
            hours_covered = list(range(start_h, end_h)) 
        # Lógica turno noche (Ej: 20 a 07)
        elif end_h < start_h:
            hours_covered = list(range(start_h, 24)) + list(range(0, end_h))
        else:
            hours_covered = [start_h] # Caso raro mismo inicio y fin
            
        return hours_covered
        
    except Exception as e:
        return []

def find_header_row(df, keywords=["nombre", "colaborador", "supervisor", "cargo"]):
    """Busca en las primeras 10 filas dónde empieza realmente la tabla."""
    for i in range(min(15, len(df))):
        row_values = df.iloc[i].astype(str).str.lower().tolist()
        if any(key in " ".join(row_values) for key in keywords):
            return i
    return 0

# --- LECTURA DE EXCEL ---

def load_excel_sheet(file, sheet_name, role_type, analysis_month_num, year):
    """
    Lee una hoja específica y normaliza los datos.
    role_type: 'Ejecutivo', 'Coordinador', 'Anfitrion', 'Supervisor'
    """
    data_extracted = []
    
    try:
        # 1. Leer sin header para inspeccionar
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # 2. Encontrar la fila de cabecera
        header_idx = 0
        if role_type == 'Supervisor':
            # Supervisores tienen una estructura muy distinta (Días en fila X, Nombres abajo)
            # Buscamos donde dice "Supervisor"
            header_idx = find_header_row(df_raw, keywords=["supervisor"])
        else:
            # Ejecutivos/Coordinadores/Anfitriones: Buscamos "Nombre" o Fechas
            header_idx = find_header_row(df_raw, keywords=["nombre", "cargo"])
            
        # 3. Releer con la cabecera correcta
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        # 4. Procesamiento específico por tipo
        
        if role_type == 'Supervisor':
            # Lógica especial Supervisores (días son columnas numéricas 1, 2, 3...)
            # Asumimos que la fila anterior al header tenía el Mes, pero usaremos el mes seleccionado por el usuario
            
            # Identificar columnas que son números (días del mes)
            day_cols = []
            for col in df.columns:
                if str(col).isdigit():
                    if 1 <= int(col) <= 31:
                        day_cols.append(col)
            
            # Iterar filas
            for idx, row in df.iterrows():
                name = row.iloc[0] # Asumimos columna 0 es Nombre
                if pd.isna(name) or str(name).lower() in ['supervisor', 'nan']: continue
                
                for day in day_cols:
                    try:
                        date_str = f"{year}-{analysis_month_num:02d}-{int(day):02d}"
                        shift = row[day]
                        data_extracted.append({
                            'Fecha': date_str,
                            'Nombre': name,
                            'Rol': role_type,
                            'Turno_Raw': shift
                        })
                    except: pass
                    
        else:
            # Lógica Ejecutivos, Coordinadores, Anfitriones
            # Buscan columnas que sean fechas (datetime)
            
            cols = df.columns
            # Identificar columna nombre
            name_col = cols[0] 
            for c in cols:
                if "nombre" in str(c).lower():
                    name_col = c
                    break
            
            date_cols = []
            for c in cols:
                # Intentar ver si la columna es una fecha
                if isinstance(c, (datetime, pd.Timestamp)):
                     date_cols.append(c)
                else:
                    # Intentar parsear string a fecha
                    try:
                        pd.to_datetime(c)
                        date_cols.append(c)
                    except:
                        pass

            for idx, row in df.iterrows():
                name = row[name_col]
                if pd.isna(name) or "nombre" in str(name).lower(): continue
                
                for date_col in date_cols:
                    shift = row[date_col]
                    # Convertir fecha de columna a string estandar
                    try:
                        d_val = pd.to_datetime(date_col)
                        # Filtrar solo si coincide con el mes de análisis (para evitar errores de hojas mixtas)
                        if d_val.month == analysis_month_num:
                            data_extracted.append({
                                'Fecha': d_val.strftime("%Y-%m-%d"),
                                'Nombre': name,
                                'Rol': role_type,
                                'Turno_Raw': shift
                            })
                    except: pass

    except Exception as e:
        st.error(f"Error procesando hoja '{sheet_name}' para {role_type}: {str(e)}")
        
    return pd.DataFrame(data_extracted)

# --- MOTOR DE REGLAS ---

def run_assignment_algorithm(df_raw):
    """Aplica las reglas de negocio (Counters, Colaciones, Tareas)."""
    
    # Expandir turnos a horas
    expanded_rows = []
    for _, row in df_raw.iterrows():
        hours = parse_time_range(row['Turno_Raw'])
        for h in hours:
            expanded_rows.append({
                'Fecha': row['Fecha'],
                'Hora': h,
                'Nombre': row['Nombre'],
                'Rol': row['Rol']
            })
            
    if not expanded_rows:
        return pd.DataFrame()
        
    df_hourly = pd.DataFrame(expanded_rows)
    
    # Preparar columnas de salida
    df_hourly['Tarea'] = "Disponible"
    df_hourly['Ubicacion'] = "-" # Counter o Zona
    
    # Procesar grupo por grupo (Fecha y Hora)
    grouped = df_hourly.groupby(['Fecha', 'Hora'])
    
    results = []
    counters_pool = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (date, hour), group in grouped:
        
        # Sub-dataframes por rol
        execs = group[group['Rol'] == 'Ejecutivo']
        coords = group[group['Rol'] == 'Coordinador']
        hosts = group[group['Rol'] == 'Anfitrion']
        sups = group[group['Rol'] == 'Supervisor']
        
        # 1. SUPERVISORES (Regla: No hacer nada, solo marcar)
        for idx in sups.index:
            results.append((idx, "Supervisión", "General"))
            
        # 2. EJECUTIVOS (Regla: Counters y Colaciones)
        active_execs_indices = []
        
        for idx, row in execs.iterrows():
            is_colacion = False
            # Regla Colación: Diurnos (entran AM) colación 12-14. Nocturnos (entran PM) colación 2-4
            # Usamos hash del nombre para distribuir aleatoriamente la colación en esos rangos
            # Hora 13:00 - 15:00 (Rango 12-14 real cubre horas 12 y 13, o 13 y 14)
            if hour in [13, 14]: 
                if hash(row['Nombre']) % 2 == (hour % 2): is_colacion = True
            elif hour in [2, 3]: # Madrugada
                if hash(row['Nombre']) % 2 == (hour % 2): is_colacion = True
                
            if is_colacion:
                results.append((idx, "Colación", "-"))
            else:
                active_execs_indices.append(idx)
        
        # Asignar Counters
        demand = 4
        supply = len(active_execs_indices)
        
        for i, idx in enumerate(active_execs_indices):
            if i < 4:
                # Cubrir T1A, T1T, T2A, T2T
                results.append((idx, "Atención", counters_pool[i]))
            else:
                # Exceso -> Aire
                extra = "T1 AIRE" if i % 2 == 0 else "T2 AIRE"
                results.append((idx, "Refuerzo", extra))
                
        uncovered_counters = max(0, 4 - supply)
        # Lista de counters que quedaron vacíos
        counters_needing_cover = counters_pool[supply:] if supply < 4 else []
        
        # 3. COORDINADORES (Regla: Tarea 1, Tarea 2 Admin, Tarea 4 Quiebre)
        active_coords = []
        
        for idx, row in coords.iterrows():
            # Regla Tarea 2 (Admin): Max 2 horas. Horarios tipicos 10-11, 14-16, 5-8
            is_tarea2 = False
            if uncovered_counters == 0: # Solo si no hay quiebre
                if hour in [10, 11, 15, 16, 5, 6]:
                    is_tarea2 = True
            
            if is_tarea2:
                results.append((idx, "Tarea 2 (Admin)", "Oficina"))
            else:
                active_coords.append(idx)
                
        # Asignar disponibles a Tarea 4 (si hay quiebre) o Tarea 1
        for idx in active_coords:
            if uncovered_counters > 0:
                cnt = counters_needing_cover.pop(0)
                results.append((idx, "Tarea 4 (Cobertura)", cnt))
                uncovered_counters -= 1
            else:
                results.append((idx, "Tarea 1 (Coord)", "Piso"))
                
        # 4. ANFITRIONES (Regla: Zona Int/Nac, Tarea 4 si Coord no cubre)
        for idx, row in hosts.iterrows():
            if uncovered_counters > 0:
                # Quiebre critico
                cnt = counters_needing_cover.pop(0) if counters_needing_cover else "Cualquiera"
                results.append((idx, "Tarea 4 (Apoyo Extremo)", cnt))
                uncovered_counters -= 1
            else:
                # Zona
                zona = "Zona Internacional" if hash(row['Nombre']) % 2 == 0 else "Zona Nacional"
                results.append((idx, "Tarea 1 (Anfitrion)", zona))

    # Escribir resultados
    for idx, task, loc in results:
        df_hourly.at[idx, 'Tarea'] = task
        df_hourly.at[idx, 'Ubicacion'] = loc
        
    return df_hourly

# --- INTERFAZ GRÁFICA ---

st.sidebar.header("1. Configuración del Análisis")
month_names = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
               "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
selected_month_name = st.sidebar.selectbox("Mes a Analizar", month_names, index=1) # Default Febrero
selected_year = st.sidebar.number_input("Año", value=2026, step=1)

# Calcular número de mes para filtrar
month_map = {name: i+1 for i, name in enumerate(month_names)}
month_num = month_map[selected_month_name]

st.sidebar.markdown("---")
st.sidebar.header("2. Carga de Excel (.xlsx)")

def smart_select_sheet(excel_file, label):
    """Helper para UI: Muestra dropdown con sugerencia inteligente."""
    if excel_file is not None:
        try:
            xl = pd.ExcelFile(excel_file)
            sheets = xl.sheet_names
            
            # Buscar sugerencia (insensitive case)
            default_ix = 0
            for i, s in enumerate(sheets):
                if selected_month_name.lower() in s.lower():
                    default_ix = i
                    break
            
            selected = st.sidebar.selectbox(f"Hoja para {label}", sheets, index=default_ix, key=label)
            return selected
        except Exception as e:
            st.sidebar.error(f"Error leyendo archivo: {e}")
            return None
    return None

# Uploaders
u_exec = st.sidebar.file_uploader("Ejecutivos", type=["xlsx"])
s_exec = smart_select_sheet(u_exec, "Ejecutivos")

u_host = st.sidebar.file_uploader("Anfitriones", type=["xlsx"])
s_host = smart_select_sheet(u_host, "Anfitriones")

u_coord = st.sidebar.file_uploader("Coordinadores", type=["xlsx"])
s_coord = smart_select_sheet(u_coord, "Coordinadores")

u_sup = st.sidebar.file_uploader("Supervisores", type=["xlsx"])
s_sup = smart_select_sheet(u_sup, "Supervisores")

# --- PROCESAMIENTO ---

if st.sidebar.button("🚀 Generar Planificación"):
    if all([u_exec, s_exec, u_host, s_host, u_coord, s_coord, u_sup, s_sup]):
        with st.spinner(f"Analizando turnos para {selected_month_name}..."):
            
            # 1. Cargar Datos
            df_exec = load_excel_sheet(u_exec, s_exec, 'Ejecutivo', month_num, selected_year)
            df_host = load_excel_sheet(u_host, s_host, 'Anfitrion', month_num, selected_year)
            df_coord = load_excel_sheet(u_coord, s_coord, 'Coordinador', month_num, selected_year)
            df_sup = load_excel_sheet(u_sup, s_sup, 'Supervisor', month_num, selected_year)
            
            # Unir todo
            df_master = pd.concat([df_exec, df_host, df_coord, df_sup], ignore_index=True)
            
            if not df_master.empty:
                st.success(f"Datos extraídos correctamente. Total registros brutos: {len(df_master)}")
                
                # 2. Ejecutar Algoritmo
                df_final = run_assignment_algorithm(df_master)
                
                if df_final.empty:
                    st.warning("No se generaron turnos. Verifica que las fechas del Excel coincidan con el mes seleccionado.")
                else:
                    # 3. Mostrar Resultados
                    st.header(f"Planificación: {selected_month_name} {selected_year}")
                    
                    # KPIs
                    kpi1, kpi2, kpi3 = st.columns(3)
                    kpi1.metric("Horas Totales Cubiertas", len(df_final))
                    quiebres = df_final[df_final['Tarea'].str.contains("Tarea 4", na=False)]
                    kpi2.metric("Apoyos por Quiebre (Tarea 4)", len(quiebres))
                    colaciones = df_final[df_final['Tarea'] == "Colación"]
                    kpi3.metric("Colaciones Asignadas", len(colaciones))
                    
                    # Visualización Matriz
                    st.subheader("Disponibilidad por Hora")
                    heatmap_data = df_final.groupby(['Fecha', 'Hora'])['Nombre'].count().reset_index()
                    st.vega_lite_chart(heatmap_data, {
                        'mark': 'rect',
                        'encoding': {
                            'x': {'field': 'Hora', 'type': 'ordinal'},
                            'y': {'field': 'Fecha', 'type': 'ordinal', 'timeUnit': 'yearmonthdate'},
                            'color': {'field': 'Nombre', 'aggregate': 'count', 'type': 'quantitative', 'scale': {'scheme': 'yellowgreenblue'}}
                        }
                    }, use_container_width=True)
                    
                    # Tabla Detalle
                    st.dataframe(df_final, use_container_width=True)
                    
                    # 4. Descarga Excel
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, sheet_name='Detalle_Completo', index=False)
                        
                        # Resumen de Cobertura
                        piv_cob = df_final.pivot_table(index='Fecha', columns='Hora', values='Nombre', aggfunc='count')
                        piv_cob.to_excel(writer, sheet_name='Mapa_Calor')
                        
                        # Resumen por Persona
                        piv_per = df_final.pivot_table(index=['Rol','Nombre'], columns='Tarea', values='Hora', aggfunc='count', fill_value=0)
                        piv_per.to_excel(writer, sheet_name='Resumen_Persona')
                        
                    st.download_button(
                        label="📥 Descargar Excel Final",
                        data=buffer.getvalue(),
                        file_name=f"Planificacion_{selected_month_name}_{selected_year}.xlsx",
                        mime="application/vnd.ms-excel"
                    )
            else:
                st.error("No se encontraron datos. Revisa que las hojas seleccionadas tengan el formato correcto.")
    else:
        st.warning("Por favor, carga todos los archivos y asegúrate de seleccionar una hoja válida para cada uno.")
