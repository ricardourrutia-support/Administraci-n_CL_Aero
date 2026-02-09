import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import xlsxwriter

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto (Admin)", layout="wide")

st.title("✈️ Gestor Avanzado de Turnos y Cobertura")
st.markdown("""
Esta aplicación genera la **Planificación Mensual (Formato Admin)** y permite gestionar **Ausencias**.
El sistema recalcula automáticamente las coberturas y detecta necesidades de horas extra (Tarea 4).
""")

# --- ESTADO DE SESIÓN (PARA AUSENCIAS) ---
if 'absences' not in st.session_state:
    st.session_state.absences = []

def add_absence(name, date, reason):
    st.session_state.absences.append({'Nombre': name, 'Fecha': date, 'Motivo': reason})

def remove_absence(idx):
    if 0 <= idx < len(st.session_state.absences):
        st.session_state.absences.pop(idx)

# --- FUNCIONES DE PARSEO ---

def parse_time_range(time_str):
    """Convierte '09:00 - 20:00' a lista de horas [9, 10, ..., 19]."""
    if pd.isna(time_str): return []
    s = str(time_str).lower().strip()
    if s in ['libre', 'nan', 'dia libre', 'dias libres', 'l', 'x', 'vacaciones', 'licencia']:
        return []
    
    # Limpieza
    s = s.replace(" diurno", "").replace(" nocturno", "").strip()
    
    try:
        parts = s.split('-')
        if len(parts) != 2: return []
        
        start_str, end_str = parts[0].strip(), parts[1].strip()
        
        # Intentar varios formatos
        fmts = ["%H:%M:%S", "%H:%M", "%H"]
        start_dt, end_dt = None, None
        
        for fmt in fmts:
            if not start_dt:
                try: start_dt = datetime.strptime(start_str, fmt)
                except: pass
            if not end_dt:
                try: end_dt = datetime.strptime(end_str, fmt)
                except: pass
                
        if not start_dt or not end_dt: return []
        
        start_h = start_dt.hour
        end_h = end_dt.hour
        
        if end_h > start_h:
            return list(range(start_h, end_h))
        elif end_h < start_h: # Turno noche (ej 22 a 06)
            return list(range(start_h, 24)) + list(range(0, end_h))
        else:
            return [start_h]
            
    except:
        return []

def get_shift_type(time_str):
    """Determina si es Diurno o Nocturno basado en hora inicio."""
    hours = parse_time_range(time_str)
    if not hours: return "Libre"
    start_h = hours[0]
    # Criterio simple: Si entra entre 05:00 y 14:00 es Diurno
    if 5 <= start_h < 14:
        return "Diurno"
    else:
        return "Nocturno"

def find_header_row(df, keywords=["nombre", "colaborador", "supervisor", "cargo"]):
    for i in range(min(20, len(df))):
        row_str = " ".join(df.iloc[i].astype(str).str.lower())
        if any(k in row_str for k in keywords):
            return i
    return 0

# --- CARGA DE DATOS ---

def load_excel_sheet(file, sheet_name, role_type, month_num, year):
    data = []
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        header_idx = find_header_row(df_raw, ["nombre", "supervisor", "colaborador"])
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        # Identificar columnas
        cols = df.columns
        name_col = next((c for c in cols if "nombre" in str(c).lower() or "supervisor" in str(c).lower()), cols[0])
        
        # Identificar fechas
        date_map = {} # {col_name: date_obj}
        
        if role_type == 'Supervisor':
            # Supervisores: Días numéricos (1, 2, 3...)
            for c in cols:
                if str(c).isdigit():
                    d = int(c)
                    if 1 <= d <= 31:
                        try:
                            date_obj = datetime(year, month_num, d)
                            date_map[c] = date_obj
                        except: pass
        else:
            # Otros: Fechas datetime o strings
            for c in cols:
                if isinstance(c, (datetime, pd.Timestamp)):
                    if c.month == month_num:
                        date_map[c] = c
                else:
                    try:
                        d_obj = pd.to_datetime(c)
                        if d_obj.month == month_num and d_obj.year == year:
                            date_map[c] = d_obj
                    except: pass
        
        # Extraer datos
        for idx, row in df.iterrows():
            name = row[name_col]
            if pd.isna(name) or str(name).lower() in ['nombre', 'supervisor', 'cargo', 'nan']: continue
            
            for col, date_val in date_map.items():
                shift = row[col]
                data.append({
                    'Fecha': date_val,
                    'Nombre': name,
                    'Rol': role_type,
                    'Turno_Raw': shift
                })
                
    except Exception as e:
        st.error(f"Error en {sheet_name}: {e}")
        
    return pd.DataFrame(data)

# --- ALGORITMO DE ASIGNACIÓN ---

def run_assignment(df_raw, absences_list):
    """
    Genera la matriz horaria aplicando reglas y ausentismo.
    """
    # 1. Marcar Ausencias
    df_raw['Estado'] = 'Presente'
    df_raw['Fecha'] = pd.to_datetime(df_raw['Fecha'])
    
    for abs_rec in absences_list:
        mask = (df_raw['Nombre'] == abs_rec['Nombre']) & (df_raw['Fecha'] == pd.to_datetime(abs_rec['Fecha']))
        if mask.any():
            df_raw.loc[mask, 'Turno_Raw'] = "AUSENTE" # Sobrescribir turno
            df_raw.loc[mask, 'Estado'] = 'Ausente'

    # 2. Expandir a Horas
    expanded = []
    for _, row in df_raw.iterrows():
        hours = parse_time_range(row['Turno_Raw'])
        
        # Determinar tipo turno para ordenamiento
        shift_type = get_shift_type(row['Turno_Raw'])
        
        if not hours:
            # Registrar al menos una fila para el día libre/ausente
            expanded.append({
                'Fecha': row['Fecha'], 'Hora': -1, 'Nombre': row['Nombre'],
                'Rol': row['Rol'], 'Turno_Raw': row['Turno_Raw'], 
                'Shift_Type': shift_type, 'Estado': row['Estado']
            })
        else:
            for h in hours:
                expanded.append({
                    'Fecha': row['Fecha'], 'Hora': h, 'Nombre': row['Nombre'],
                    'Rol': row['Rol'], 'Turno_Raw': row['Turno_Raw'], 
                    'Shift_Type': shift_type, 'Estado': row['Estado']
                })
                
    df_h = pd.DataFrame(expanded)
    if df_h.empty: return pd.DataFrame()
    
    df_h['Tarea'] = "-"
    df_h['Counter'] = "-"
    
    # 3. Procesar por Bloque Horario
    grouped = df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora'])
    
    counters_pool = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (date, hour), group in grouped:
        # Separar roles
        execs = group[(group['Rol']=='Ejecutivo') & (group['Estado']=='Presente')]
        coords = group[(group['Rol']=='Coordinador') & (group['Estado']=='Presente')]
        hosts = group[(group['Rol']=='Anfitrion') & (group['Estado']=='Presente')]
        sups = group[(group['Rol']=='Supervisor') & (group['Estado']=='Presente')]
        
        # --- LÓGICA EJECUTIVOS ---
        active_execs = []
        
        # Colaciones (Regla Probabilística para simplificar)
        for idx, row in execs.iterrows():
            is_colacion = False
            # Diurno: Colación 13-15 approx. Nocturno: 2-4
            if row['Shift_Type'] == 'Diurno' and hour in [13, 14]:
                if hash(row['Nombre'] + str(date)) % 2 == (hour % 2): is_colacion = True
            elif row['Shift_Type'] == 'Nocturno' and hour in [2, 3]:
                if hash(row['Nombre'] + str(date)) % 2 == (hour % 2): is_colacion = True
                
            if is_colacion:
                df_h.at[idx, 'Tarea'] = "C" # Colación
            else:
                active_execs.append(idx)
        
        # Asignar Counters
        supply = len(active_execs)
        uncovered = max(0, 4 - supply)
        need_cover_list = counters_pool[supply:] if supply < 4 else []
        
        for i, idx in enumerate(active_execs):
            if i < 4:
                # Tarea 1: Counter asignado
                cnt = counters_pool[i]
                df_h.at[idx, 'Tarea'] = "1"
                df_h.at[idx, 'Counter'] = cnt
            else:
                # Exceso: Refuerzo Aire
                cnt = "T1 AIRE" if i%2==0 else "T2 AIRE"
                df_h.at[idx, 'Tarea'] = "1" # Sigue siendo tarea 1 (atención)
                df_h.at[idx, 'Counter'] = cnt
                
        # --- LÓGICA COORDINADORES ---
        active_coords = []
        for idx, row in coords.iterrows():
            # Tarea 2 (Admin) vs Disponibilidad
            is_admin = False
            if uncovered == 0:
                if row['Shift_Type'] == 'Diurno' and hour in [10, 11, 15, 16]: is_admin = True
                if row['Shift_Type'] == 'Nocturno' and hour in [5, 6]: is_admin = True
            
            if is_admin:
                df_h.at[idx, 'Tarea'] = "2"
                df_h.at[idx, 'Counter'] = "Oficina"
            else:
                active_coords.append(idx)
                
        # Cobertura de Quiebres (Tarea 4)
        for idx in active_coords:
            if uncovered > 0:
                cnt = need_cover_list.pop(0)
                df_h.at[idx, 'Tarea'] = "4" # Cobertura
                df_h.at[idx, 'Counter'] = cnt
                uncovered -= 1
            else:
                df_h.at[idx, 'Tarea'] = "1" # Supervisión piso
                df_h.at[idx, 'Counter'] = "Piso"
                
        # --- LÓGICA ANFITRIONES ---
        for idx, row in hosts.iterrows():
            if uncovered > 0:
                cnt = need_cover_list.pop(0) if need_cover_list else "Apoyo"
                df_h.at[idx, 'Tarea'] = "4"
                df_h.at[idx, 'Counter'] = cnt
                uncovered -= 1
            else:
                df_h.at[idx, 'Tarea'] = "1"
                df_h.at[idx, 'Counter'] = "Zona Int" if i%2==0 else "Zona Nac"

        # --- LÓGICA SUPERVISORES ---
        for idx, row in sups.iterrows():
            df_h.at[idx, 'Tarea'] = "1"
            df_h.at[idx, 'Counter'] = "General"

    # 4. Post-Proceso: Detectar Tarea 3 (Cambio de Counter)
    # Calculamos el counter "Moda" (más frecuente) por persona por día
    # Si en una hora específica el counter != Moda, es Tarea 3 (Movimiento)
    
    # Solo para filas con hora válida
    valid_rows = df_h[df_h['Hora'] != -1]
    
    # Moda por día/persona
    mode_counters = valid_rows.groupby(['Fecha', 'Nombre'])['Counter'].agg(lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
    
    for idx, row in valid_rows.iterrows():
        if row['Tarea'] == "1" and row['Rol'] == 'Ejecutivo':
            try:
                main_cnt = mode_counters.loc[(row['Fecha'], row['Nombre'])]
                curr_cnt = row['Counter']
                if curr_cnt != "-" and main_cnt != "-" and curr_cnt != main_cnt:
                    df_h.at[idx, 'Tarea'] = "3" # Floating
            except: pass

    return df_h

# --- GENERADOR EXCEL (FORMATO ADMIN) ---

def generate_admin_excel(df_processed, month_name, year):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet(f"Admin {month_name}")
    
    # Estilos
    fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
    fmt_day = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#BDD7EE', 'border': 1})
    fmt_cell = workbook.add_format({'align': 'center', 'border': 1, 'font_size': 9})
    fmt_task4 = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Rojo
    fmt_colacion = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#E2EFDA'}) # Verde
    fmt_libre = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1})
    
    # Preparar datos
    # Necesitamos una lista de personas única ordenada
    people = df_processed[['Nombre', 'Rol', 'Shift_Type']].drop_duplicates().sort_values(['Rol', 'Shift_Type', 'Nombre'])
    
    # Fechas únicas
    dates = sorted(df_processed['Fecha'].unique())
    
    # --- ESCRIBIR CABECERAS ---
    # Fila 0: Días (Agrupados)
    # Fila 1: Columnas fijas por día (Turno, Term, Count, 0...23)
    
    row_day = 0
    row_sub = 1
    start_col = 0
    
    # Columna Nombre Fija
    worksheet.merge_range(0, 0, 1, 0, "Colaborador", fmt_header)
    worksheet.merge_range(0, 1, 1, 1, "Rol", fmt_header)
    start_col = 2
    
    date_col_map = {} # fecha -> indice columna inicio
    
    for d in dates:
        d_str = pd.to_datetime(d).strftime("%d-%b")
        # Ancho del bloque: Turno (1) + Terminal (1) + Counter (1) + 24 Horas = 27 cols
        block_width = 27 
        
        # Escribir Día
        worksheet.merge_range(row_day, start_col, row_day, start_col + block_width - 1, d_str, fmt_day)
        
        # Escribir Sub-cabeceras
        worksheet.write(row_sub, start_col, "Turno", fmt_header)
        worksheet.write(row_sub, start_col+1, "Term", fmt_header)
        worksheet.write(row_sub, start_col+2, "Count", fmt_header)
        
        for h in range(24):
            worksheet.write(row_sub, start_col + 3 + h, h, fmt_header)
            
        date_col_map[d] = start_col
        start_col += block_width + 1 # +1 para separador
        
    # --- ESCRIBIR DATOS ---
    curr_row = 2
    
    # Agrupar datos para acceso rápido
    # df_processed indexado por (Nombre, Fecha, Hora)
    df_idx = df_processed.set_index(['Nombre', 'Fecha', 'Hora'])
    df_day_idx = df_processed.set_index(['Nombre', 'Fecha']) # Para sacar turno raw y counter principal
    
    last_role = None
    
    for _, person in people.iterrows():
        name = person['Nombre']
        role = person['Rol']
        
        # Separador visual entre roles
        if last_role and role != last_role:
            curr_row += 1 
        last_role = role
        
        worksheet.write(curr_row, 0, name, fmt_cell)
        worksheet.write(curr_row, 1, role, fmt_cell)
        
        for d in dates:
            if d not in date_col_map: continue
            c_start = date_col_map[d]
            
            # Obtener datos del día para la persona
            # Info general del día (Turno, Counter Principal)
            try:
                # Buscar cualquier registro de ese día para sacar turno y counter moda
                day_records = df_processed[(df_processed['Nombre'] == name) & (df_processed['Fecha'] == d)]
                if day_records.empty:
                    worksheet.write(curr_row, c_start, "-", fmt_libre)
                    continue
                
                # Turno Raw
                turno_str = day_records.iloc[0]['Turno_Raw']
                worksheet.write(curr_row, c_start, turno_str, fmt_cell)
                
                # Counter Principal (Moda)
                cnt_list = day_records[day_records['Counter'] != '-']['Counter']
                main_cnt = cnt_list.mode().iloc[0] if not cnt_list.empty else "-"
                
                # Separar Term y Count
                term = "T1" if "T1" in main_cnt else ("T2" if "T2" in main_cnt else "-")
                count_type = "AIRE" if "AIRE" in main_cnt else ("TIERRA" if "TIERRA" in main_cnt else main_cnt)
                
                worksheet.write(curr_row, c_start+1, term, fmt_cell)
                worksheet.write(curr_row, c_start+2, count_type, fmt_cell)
                
                # Escribir Horas
                for h in range(24):
                    try:
                        val = df_idx.loc[(name, d, h), 'Tarea']
                        # Si hay duplicados (raro), toma el primero
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        
                        fmt = fmt_cell
                        if val == "4": fmt = fmt_task4
                        if val == "C": fmt = fmt_colacion
                        
                        worksheet.write(curr_row, c_start + 3 + h, val, fmt)
                    except KeyError:
                        # No hay tarea asignada para esa hora (fuera de turno)
                        worksheet.write(curr_row, c_start + 3 + h, "", fmt_libre)
                        
            except Exception as e:
                pass
                
        curr_row += 1
        
    workbook.close()
    return output

# --- UI PRINCIPAL ---

st.sidebar.header("1. Carga de Archivos (.xlsx)")

# Uploaders
files = {}
roles = ["Ejecutivo", "Anfitrion", "Coordinador", "Supervisor"]
for r in roles:
    files[r] = st.sidebar.file_uploader(f"Turnos {r}", type=["xlsx"], key=r)

# Configuración
st.sidebar.markdown("---")
st.sidebar.header("2. Configuración")
months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
          "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
sel_month = st.sidebar.selectbox("Mes", months, index=1)
sel_year = st.sidebar.number_input("Año", 2025, 2030, 2026)
month_num = months.index(sel_month) + 1

# --- GESTOR DE AUSENCIAS ---
st.sidebar.markdown("---")
st.sidebar.header("3. Registro de Ausencias")

# Necesitamos cargar nombres primero para llenar esto
loaded_names = []
if st.session_state.get('df_master_names') is not None:
    loaded_names = st.session_state['df_master_names']

with st.sidebar.expander("Gestionar Ausencias", expanded=False):
    if loaded_names:
        a_name = st.selectbox("Colaborador", loaded_names)
        a_date = st.date_input("Fecha Ausencia", datetime(sel_year, month_num, 1))
        a_reason = st.text_input("Motivo", "Licencia")
        if st.button("Agregar Ausencia"):
            add_absence(a_name, a_date.strftime("%Y-%m-%d"), a_reason)
            st.success("Agregado")
            
    # Listar
    if st.session_state.absences:
        st.write("Ausencias registradas:")
        for i, abs_rec in enumerate(st.session_state.absences):
            c1, c2 = st.columns([4, 1])
            c1.caption(f"{abs_rec['Nombre']} - {abs_rec['Fecha']}")
            if c2.button("X", key=f"del_{i}"):
                remove_absence(i)
                st.rerun()

# --- PROCESO PRINCIPAL ---

if st.button("Generar Planificación Admin"):
    # Validar cargas
    if not all(files.values()):
        st.error("Por favor carga los 4 archivos Excel.")
    else:
        with st.spinner("Leyendo Excels y procesando reglas..."):
            # 1. Cargar Datos Raw
            dfs = []
            for r in roles:
                # Intentar adivinar hoja
                xl = pd.ExcelFile(files[r])
                sheet = next((s for s in xl.sheet_names if sel_month.lower() in s.lower()), xl.sheet_names[0])
                df = load_excel_sheet(files[r], sheet, r, month_num, sel_year)
                dfs.append(df)
            
            df_master = pd.concat(dfs, ignore_index=True)
            
            if df_master.empty:
                st.error("No se encontraron datos válidos para el mes seleccionado.")
            else:
                # Guardar nombres para el selector de ausencias
                st.session_state['df_master_names'] = sorted(df_master['Nombre'].unique().tolist())
                
                # 2. Ejecutar Algoritmo
                df_results = run_assignment(df_master, st.session_state.absences)
                
                if not df_results.empty:
                    # 3. Mostrar KPIs
                    st.success("Planificación Generada Exitosamente")
                    
                    c1, c2, c3 = st.columns(3)
                    total_hours = len(df_results[df_results['Tarea'].isin(['1','2','3','4'])])
                    overtime_hours = len(df_results[df_results['Tarea'] == '4'])
                    colaciones = len(df_results[df_results['Tarea'] == 'C'])
                    
                    c1.metric("Horas Totales", total_hours)
                    c2.metric("Horas Extra (Tarea 4)", overtime_hours, delta_color="inverse")
                    c3.metric("Colaciones", colaciones)
                    
                    # 4. Generar Excel Admin
                    excel_data = generate_admin_excel(df_results, sel_month, sel_year)
                    
                    st.download_button(
                        label="📥 Descargar Sábana Mensual (Formato Admin)",
                        data=excel_data.getvalue(),
                        file_name=f"Admin_{sel_month}_{sel_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Preview
                    st.subheader("Vista Previa (Primeros 100 registros)")
                    st.dataframe(df_results.head(100))
                    
                else:
                    st.warning("No se generaron turnos. Revisa las fechas.")
