import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")

st.title("✈️ Gestor de Asignación de Turnos y Tareas")
st.markdown("""
Esta aplicación procesa los archivos de turnos mensuales, aplica las reglas de negocio 
(cobertura de counters, colaciones, tareas especiales) y genera un reporte descargable.
""")

# --- FUNCIONES DE PARSEO (LECTURA DE ARCHIVOS) ---

def parse_time_range(time_str):
    """Convierte strings como '09:00 - 20:00' en una lista de horas (0-23) cubiertas."""
    if pd.isna(time_str) or str(time_str).lower().strip() in ['libre', 'nan', 'dia libre', 'dias libres']:
        return []
    
    try:
        # Limpieza básica
        time_str = str(time_str).replace(" Diurno", "").replace(" Nocturno", "").strip()
        parts = time_str.split('-')
        if len(parts) != 2:
            return []
        
        start_str = parts[0].strip()
        end_str = parts[1].strip()
        
        # Formatos posibles: H:M:S o H:M
        fmt_long = "%H:%M:%S"
        fmt_short = "%H:%M"
        
        try:
            start_dt = datetime.strptime(start_str, fmt_long)
        except:
            start_dt = datetime.strptime(start_str, fmt_short)
            
        try:
            end_dt = datetime.strptime(end_str, fmt_long)
        except:
            end_dt = datetime.strptime(end_str, fmt_short)
            
        start_h = start_dt.hour
        end_h = end_dt.hour
        
        hours_covered = []
        
        # Caso turno normal (ej 9 a 20)
        if end_h > start_h:
            hours_covered = list(range(start_h, end_h)) # No incluimos la hora final exacta
        # Caso turno nocturno (ej 20 a 07)
        elif end_h < start_h:
            hours_covered = list(range(start_h, 24)) + list(range(0, end_h))
        else:
            # Caso raro inicio == fin
            hours_covered = [start_h]
            
        return hours_covered
        
    except Exception as e:
        # Si falla el parseo, asumimos libre (o se podría loguear error)
        return []

def load_data(file_execs, file_coords, file_hosts, file_sups, year=2026):
    """Carga y estandariza los 4 tipos de inputs."""
    data_master = []

    # 1. EJECUTIVOS (Busca fila de encabezado '202X-XX-XX')
    try:
        df_e = pd.read_csv(file_execs, header=None)
        # Buscar la fila que tiene fechas. Asumimos que fila 3 (indice) tiene fechas
        header_idx = 3
        df_e_clean = pd.read_csv(file_execs, header=header_idx)
        # Filtrar filas vacías o basura
        df_e_clean = df_e_clean.dropna(subset=[df_e_clean.columns[0]]) # Nombre/Cargo no nulo
        
        # Iterar columnas de fecha
        cols = df_e_clean.columns
        for idx, row in df_e_clean.iterrows():
            name = row[cols[0]] # Asumimos col 0 es nombre/cargo
            if "Nombre" in str(name) or "Cargo" in str(name): continue

            for col in cols[1:]:
                date_val = col
                shift_val = row[col]
                # Validar que col es una fecha
                try:
                    pd.to_datetime(date_val)
                    data_master.append({
                        'Fecha': date_val,
                        'Nombre': name,
                        'Rol': 'Ejecutivo',
                        'Turno_Raw': shift_val
                    })
                except:
                    pass
    except Exception as e:
        st.error(f"Error leyendo Ejecutivos: {e}")

    # 2. ANFITRIONES (Similar a ejecutivos)
    try:
        df_h = pd.read_csv(file_hosts, header=3) # Asumimos misma estructura
        df_h = df_h.dropna(subset=[df_h.columns[0]])
        cols = df_h.columns
        for idx, row in df_h.iterrows():
            name = row[cols[0]]
            if "Nombre" in str(name): continue
            for col in cols[1:]:
                try:
                    pd.to_datetime(col)
                    data_master.append({
                        'Fecha': col,
                        'Nombre': name,
                        'Rol': 'Anfitrion',
                        'Turno_Raw': row[col]
                    })
                except:
                    pass
    except Exception as e:
        st.error(f"Error leyendo Anfitriones: {e}")

    # 3. COORDINADORES (Header en fila 4 generalmente)
    try:
        df_c = pd.read_csv(file_coords, header=4)
        df_c = df_c.dropna(subset=[df_c.columns[0]])
        cols = df_c.columns
        for idx, row in df_c.iterrows():
            name = row[cols[0]]
            if "Nombre" in str(name): continue
            for col in cols[1:]:
                try:
                    pd.to_datetime(col)
                    data_master.append({
                        'Fecha': col,
                        'Nombre': name,
                        'Rol': 'Coordinador',
                        'Turno_Raw': row[col]
                    })
                except:
                    pass
    except Exception as e:
        st.error(f"Error leyendo Coordinadores: {e}")

    # 4. SUPERVISORES (Estructura compleja: Dia numero en fila 1, Mes en fila 0)
    try:
        # Leemos raw para reconstruir fechas
        df_s_raw = pd.read_csv(file_sups, header=None)
        month_str = str(df_s_raw.iloc[0, 1]).strip().upper() # "FEBRERO"
        days_row = df_s_raw.iloc[1] # 2, 3, 4...
        
        # Mapeo de mes a numero
        months = {'ENERO':1, 'FEBRERO':2, 'MARZO':3, 'ABRIL':4, 'MAYO':5, 'JUNIO':6, 
                  'JULIO':7, 'AGOSTO':8, 'SEPTIEMBRE':9, 'OCTUBRE':10, 'NOVIEMBRE':11, 'DICIEMBRE':12}
        month_num = months.get(month_str, 2) # Default feb
        
        # Data real empieza en fila 3
        df_s_data = pd.read_csv(file_sups, header=2)
        
        # Iterar
        for idx, row in df_s_data.iterrows():
            name = row[0] # "Leo", "Hugo"
            if str(name) == "Supervisor" or pd.isna(name): continue
            
            # Las columnas desde la 1 son los días.
            # Necesitamos matchear el indice de columna con el día en days_row
            for c_idx in range(1, len(df_s_data.columns)):
                try:
                    day_num = int(days_row[c_idx])
                    # Construir fecha YYYY-MM-DD
                    date_str = f"{year}-{month_num:02d}-{day_num:02d}"
                    shift_val = row.iloc[c_idx]
                    
                    data_master.append({
                        'Fecha': date_str,
                        'Nombre': name,
                        'Rol': 'Supervisor',
                        'Turno_Raw': shift_val
                    })
                except:
                    pass # Columna sin dia valido o fin de mes
    except Exception as e:
        st.warning(f"Advertencia leyendo Supervisores (formato complejo): {e}")

    return pd.DataFrame(data_master)

# --- MOTOR DE REGLAS ---

def run_assignment_algorithm(df_raw, start_date, end_date):
    """
    Expande los turnos diarios a una matriz horaria y asigna tareas.
    """
    # 1. Filtrar por fecha
    df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'])
    mask = (df_raw['Fecha'] >= pd.to_datetime(start_date)) & (df_raw['Fecha'] <= pd.to_datetime(end_date))
    df = df_raw.loc[mask].copy()
    
    # 2. Expansión Horaria
    expanded_rows = []
    
    for _, row in df.iterrows():
        hours = parse_time_range(row['Turno_Raw'])
        for h in hours:
            expanded_rows.append({
                'Fecha': row['Fecha'],
                'Hora': h,
                'Nombre': row['Nombre'],
                'Rol': row['Rol'],
                'Turno_Origen': row['Turno_Raw']
            })
            
    df_hourly = pd.DataFrame(expanded_rows)
    
    if df_hourly.empty:
        return pd.DataFrame()

    # 3. Asignación Lógica (Iterar por Fecha -> Hora)
    df_hourly['Tarea'] = "Disponible"
    df_hourly['Counter'] = "-"
    
    # Ordenamos para procesar cronológicamente
    grouped = df_hourly.groupby(['Fecha', 'Hora'])
    
    results = []
    
    counters_pool = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (date, hour), group in grouped:
        # Separar grupos
        execs = group[group['Rol'] == 'Ejecutivo'].copy()
        coords = group[group['Rol'] == 'Coordinador'].copy()
        hosts = group[group['Rol'] == 'Anfitrion'].copy()
        sups = group[group['Rol'] == 'Supervisor'].copy()
        
        # --- REGLA: SUPERVISORES ---
        # "Simplemente marcar con 1" -> Tarea: Supervisión
        for i, idx in enumerate(sups.index):
             results.append((idx, "Supervisión", "General"))
             
        # --- REGLA: EJECUTIVOS Y COLACIÓN ---
        # Asignar colación según horario
        # Diurno (entra 8-10): Colación 12-14. Nocturno (entra 20-22): Colación 2-4.
        # Simplificación heurística basada en la hora actual
        
        active_execs_indices = []
        
        for idx, row in execs.iterrows():
            # Determinar si está en horario de colación
            is_colacion = False
            
            # Heurística simple: Si es ejecutivo y son las 13:00 o 14:00 (día) o 02:00/03:00 (noche)
            # Para mayor precisión necesitaríamos la hora de inicio del turno, aquí usamos probabilidad
            if hour in [13, 14]: # Colación Diurna (asumida para simplificar)
                # Dividir grupo: mitad a las 13, mitad a las 14
                # Usamos hash del nombre para ser deterministas
                if hash(row['Nombre']) % 2 == (hour % 2): 
                    is_colacion = True
            elif hour in [2, 3]: # Colación Nocturna
                 if hash(row['Nombre']) % 2 == (hour % 2):
                    is_colacion = True
            
            if is_colacion:
                results.append((idx, "Colación", "-"))
            else:
                active_execs_indices.append(idx)

        # --- REGLA: EJECUTIVOS A COUNTERS ---
        # Llenar los 4 counters
        # Si sobran, a counters de Aire
        
        demand = 4 # Necesitamos 1 en cada uno de los 4 counters
        supply = len(active_execs_indices)
        
        # Asignación Cíclica
        for i, idx in enumerate(active_execs_indices):
            if i < 4:
                assigned_counter = counters_pool[i] # Llenar T1A, T1T, T2A, T2T
                results.append((idx, "Atención Counter", assigned_counter))
            else:
                # Exceso -> Aire (T1 Aire o T2 Aire)
                extra_counter = "T1 AIRE" if i % 2 == 0 else "T2 AIRE"
                results.append((idx, "Refuerzo Aire", extra_counter))
        
        uncovered_counters = max(0, 4 - supply)
        counters_needing_cover = counters_pool[supply:] if supply < 4 else []
        
        # --- REGLA: COORDINADORES ---
        # Tarea 2: Max 2 horas. Horarios 10-11, 14-16, 5-8(noche).
        # Tarea 4: Cubrir quiebre de ejecutivos (uncovered_counters)
        
        active_coords_indices = []
        for idx, row in coords.iterrows():
            # Lógica Tarea 2 (Administrativa)
            is_tarea2 = False
            # Heurística horarios Tarea 2
            if hour in [10, 15] or (hour in [5, 6, 7]): 
                # Solo asignar si no se necesita cubrir counter (Tarea 4 tiene prioridad en quiebre?)
                # El usuario dijo: "Tarea 4 es cuando ocurre quiebre... si hay coordinador en Tarea 2"
                # Asumiremos Tarea 2 por defecto en esos horarios salvo emergencia
                is_tarea2 = True
            
            if is_tarea2 and uncovered_counters == 0:
                 results.append((idx, "Tarea 2 (Admin)", "Oficina"))
            else:
                active_coords_indices.append(idx)
        
        # Asignar Tarea 4 (Cobertura) o Tarea 1 (Supervisión Pista)
        for i, idx in enumerate(active_coords_indices):
            if uncovered_counters > 0:
                # Cubrir el counter faltante
                cnt_to_cover = counters_needing_cover[0]
                results.append((idx, "Tarea 4 (Cobertura)", cnt_to_cover))
                counters_needing_cover.pop(0)
                uncovered_counters -= 1
            else:
                results.append((idx, "Tarea 1 (Coord)", "Piso"))

        # --- REGLA: ANFITRIONES ---
        # 4 por franja. Tarea 1. Si quiebre persiste y coord no puede, Tarea 4.
        
        for idx, row in hosts.iterrows():
            if uncovered_counters > 0:
                # Emergencia extrema
                cnt_to_cover = counters_needing_cover[0]
                results.append((idx, "Tarea 4 (Apoyo Counter)", cnt_to_cover))
                counters_needing_cover.pop(0)
                uncovered_counters -= 1
            else:
                # Asignar zona Int/Nac
                zone = "Zona Internacional" if hash(row['Nombre']) % 2 == 0 else "Zona Nacional"
                results.append((idx, "Tarea 1 (Anfitrión)", zone))

    # Aplicar resultados al DF original
    for idx, task, cnt in results:
        df_hourly.at[idx, 'Tarea'] = task
        df_hourly.at[idx, 'Counter'] = cnt
        
    return df_hourly

# --- UI PRINCIPAL ---

st.sidebar.header("1. Carga de Archivos")

f_execs = st.sidebar.file_uploader("Turnos Ejecutivos (.csv)", type="csv")
f_hosts = st.sidebar.file_uploader("Turnos Anfitriones (.csv)", type="csv")
f_coords = st.sidebar.file_uploader("Turnos Coordinadores (.csv)", type="csv")
f_sups = st.sidebar.file_uploader("Turnos Supervisores (.csv)", type="csv")

st.sidebar.header("2. Parámetros")
start_d = st.sidebar.date_input("Fecha Inicio", datetime(2026, 2, 1))
end_d = st.sidebar.date_input("Fecha Fin", datetime(2026, 2, 28))

if st.sidebar.button("Generar Planificación"):
    if f_execs and f_hosts and f_coords and f_sups:
        with st.spinner("Procesando reglas de negocio..."):
            # 1. Cargar
            df_raw = load_data(f_execs, f_coords, f_hosts, f_sups)
            
            if not df_raw.empty:
                st.success(f"Datos cargados: {len(df_raw)} registros de turnos brutos.")
                
                # 2. Procesar
                df_final = run_assignment_algorithm(df_raw, start_d, end_d)
                
                # 3. Mostrar Resultados
                st.header("Visualización de Turnos Asignados")
                
                # Matriz visual (Pivot)
                st.subheader("Mapa de Calor de Cobertura (Ejecutivos)")
                heatmap_data = df_final[df_final['Rol']=='Ejecutivo'].groupby(['Fecha', 'Hora'])['Nombre'].count().reset_index()
                st.vega_lite_chart(heatmap_data, {
                    'mark': 'rect',
                    'encoding': {
                        'x': {'field': 'Hora', 'type': 'ordinal'},
                        'y': {'field': 'Fecha', 'type': 'ordinal', 'timeUnit': 'yearmonthdate'},
                        'color': {'field': 'Nombre', 'type': 'quantitative', 'title': 'Cant. Personal'}
                    }
                }, use_container_width=True)

                st.subheader("Detalle Hora a Hora")
                st.dataframe(df_final[['Fecha', 'Hora', 'Nombre', 'Rol', 'Tarea', 'Counter']], use_container_width=True)
                
                # 4. Estadísticas
                st.header("Estadísticas")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Horas Asignadas", len(df_final))
                with col2:
                    quiebres = df_final[df_final['Tarea'].str.contains("Tarea 4")]
                    st.metric("Horas de Apoyo (Tarea 4)", len(quiebres))
                
                # Gráfico de uso de counters
                chart_data = df_final[df_final['Counter'] != '-']['Counter'].value_counts()
                st.bar_chart(chart_data)

                # 5. Descarga
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Detalle_Horario')
                    
                    # Hoja Resumen por Persona
                    pivot_person = df_final.pivot_table(index='Nombre', columns='Tarea', values='Hora', aggfunc='count', fill_value=0)
                    pivot_person.to_excel(writer, sheet_name='Resumen_Persona')
                    
                
                st.download_button(
                    label="📥 Descargar Excel con Planificación",
                    data=buffer.getvalue(),
                    file_name="Planificacion_Turnos_Procesada.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            else:
                st.error("No se pudieron extraer datos válidos. Revisa el formato de los archivos.")
    else:
        st.warning("Por favor carga los 4 archivos CSV para comenzar.")

else:
    st.info("Sube los archivos en el menú lateral y presiona 'Generar Planificación'.")
