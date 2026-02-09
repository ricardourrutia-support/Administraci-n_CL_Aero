import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Planificador Maestro de Turnos y Cobertura")

# --- ESTADO DE SESIÓN ---
if 'absences' not in st.session_state:
    st.session_state.absences = []
if 'collaborators_list' not in st.session_state:
    st.session_state.collaborators_list = []

# --- FUNCIONES DE PARSEO INTELIGENTE ---

def parse_shift_time(shift_str):
    """
    Convierte textos de turnos (ej: '09:00 - 20:00', '22:00 - 07:00') 
    en una lista de horas cubiertas [9, 10...].
    Maneja cruces de medianoche y formatos sucios.
    """
    if pd.isna(shift_str): return []
    s = str(shift_str).lower().strip()
    
    # Palabras clave de "Libre"
    if any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla']):
        return []
    
    # Limpieza
    s = s.replace(" diurno", "").replace(" nocturno", "").replace("hrs", "").strip()
    
    try:
        # Separar inicio y fin
        parts = re.split(r'\s*-\s*|\s*a\s*', s) # Separa por "-" o " a "
        if len(parts) < 2: return []
        
        start_str, end_str = parts[0].strip(), parts[1].strip()
        
        # Formatos de hora soportados
        formats = ["%H:%M:%S", "%H:%M", "%H"]
        
        def parse_h(h_str):
            for fmt in formats:
                try: return datetime.strptime(h_str, fmt).hour
                except: pass
            return None
            
        start_h = parse_h(start_str)
        end_h = parse_h(end_str)
        
        if start_h is None or end_h is None: return []
        
        # Generar rango
        if start_h < end_h:
            return list(range(start_h, end_h))
        elif start_h > end_h: # Cruce de medianoche (ej 22 a 07)
            return list(range(start_h, 24)) + list(range(0, end_h))
        else:
            return [start_h] # Caso 1 hora (raro)
            
    except:
        return []

def find_header_and_data(df, role_type):
    """
    Busca la fila de cabecera basándose en palabras clave.
    Retorna: (indice_header, dataframe_cortado)
    """
    # Palabras clave para buscar la fila de nombres
    keywords = ["nombre", "colaborador", "funcionario", "personal"]
    if role_type == "Supervisor":
        keywords = ["supervisor"]
        
    for i in range(min(30, len(df))):
        row_str = " ".join(df.iloc[i].astype(str).str.lower())
        if any(k in row_str for k in keywords):
            # Encontrado header
            df_new = df.iloc[i+1:].reset_index(drop=True)
            df_new.columns = df.iloc[i]
            return i, df_new
            
    return 0, df

# --- CARGA DE DATOS ---

def process_file_sheet(file, sheet_name, role, target_month, target_year):
    """
    Lee una hoja específica y extrae (Nombre, Fecha, Turno).
    Adaptado para Ejecutivos, Coordinadores y Supervisores.
    """
    extracted_data = []
    
    try:
        # Leer hoja completa sin header para inspección
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Detectar estructura
        header_idx, df = find_header_and_data(df_raw, role)
        
        # Identificar columna de Nombres
        cols = df.columns
        name_col = None
        for c in cols:
            if isinstance(c, str) and any(x in c.lower() for x in ["nombre", "colaborador", "supervisor"]):
                name_col = c
                break
        if name_col is None: name_col = cols[0] # Fallback a primera columna
        
        # Identificar columnas de Fechas
        date_map = {} # {col_name: datetime_obj}
        
        if role == "Supervisor":
            # LOGICA SUPERVISORES: Los días suelen ser números (1, 2, 3...) en la fila de header
            # O en la fila superior. Asumiremos que el header encontrado tiene los números.
            for c in cols:
                try:
                    d_num = int(float(str(c))) # Maneja "1.0"
                    if 1 <= d_num <= 31:
                        # Construir fecha con el mes/año seleccionado
                        try:
                            dt = datetime(target_year, target_month, d_num)
                            date_map[c] = dt
                        except: pass # Fecha inválida (ej 30 febrero)
                except:
                    pass
        else:
            # LOGICA OTROS: Las columnas son fechas datetime o strings tipo '2025-02-01'
            for c in cols:
                if isinstance(c, datetime):
                    if c.month == target_month:
                        date_map[c] = c
                else:
                    try:
                        dt = pd.to_datetime(c)
                        if dt.month == target_month and dt.year == target_year:
                            date_map[c] = dt
                    except: pass
        
        # Extraer filas
        for idx, row in df.iterrows():
            name = row[name_col]
            # Validar nombre
            if pd.isna(name) or str(name).strip() == "" or str(name).lower() in ["nombre", "supervisor", "turno"]:
                continue
            
            for col_key, date_val in date_map.items():
                shift = row[col_key]
                extracted_data.append({
                    'Nombre': str(name).strip().title(),
                    'Rol': role,
                    'Fecha': date_val,
                    'Turno_Raw': shift
                })
                
    except Exception as e:
        st.error(f"Error leyendo {role} en hoja {sheet_name}: {e}")
        
    return pd.DataFrame(extracted_data)

# --- UI LATERAL ---

st.sidebar.header("1. Configuración de Fecha")
months = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
    "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
}
sel_month_name = st.sidebar.selectbox("Mes", list(months.keys()), index=1)
sel_year = st.sidebar.number_input("Año", value=2026)
target_month = months[sel_month_name]

st.sidebar.markdown("---")
st.sidebar.header("2. Archivos y Hojas")

files_config = [
    {"label": "Ejecutivos", "key": "exec"},
    {"label": "Anfitriones", "key": "host"},
    {"label": "Coordinadores", "key": "coord"},
    {"label": "Supervisores", "key": "sup"}
]

uploaded_sheets = {} # Almacena {role: (file, sheet_name)}

for fconf in files_config:
    role = fconf["label"]
    file = st.sidebar.file_uploader(f"Excel {role}", type=["xlsx"], key=fconf["key"])
    
    if file:
        # Leer nombres de hojas
        try:
            xl = pd.ExcelFile(file)
            sheet_names = xl.sheet_names
            # Sugerir hoja
            default_ix = 0
            for i, s in enumerate(sheet_names):
                if sel_month_name.lower() in s.lower():
                    default_ix = i
                    break
            
            selected_sheet = st.sidebar.selectbox(
                f"Hoja para {role}", 
                sheet_names, 
                index=default_ix, 
                key=f"{fconf['key']}_sheet"
            )
            uploaded_sheets[role] = (file, selected_sheet)
        except:
            st.sidebar.error("Error leyendo archivo Excel.")

# --- BOTÓN PARA CARGAR NOMBRES (SOLUCIÓN AUSENCIAS) ---
st.sidebar.markdown("---")
st.sidebar.header("3. Gestión de Ausencias")

if st.sidebar.button("🔄 1° Cargar Nombres de Archivos"):
    if not uploaded_sheets:
        st.sidebar.error("Sube archivos primero.")
    else:
        all_names = set()
        with st.spinner("Escaneando nombres en las hojas seleccionadas..."):
            for role, (u_file, u_sheet) in uploaded_sheets.items():
                # Leemos rápido solo para sacar nombres
                try:
                    df_temp = process_file_sheet(u_file, u_sheet, role, target_month, sel_year)
                    if not df_temp.empty:
                        all_names.update(df_temp['Nombre'].unique())
                except: pass
        
        st.session_state.collaborators_list = sorted(list(all_names))
        st.sidebar.success(f"Se encontraron {len(all_names)} colaboradores.")

# Selector de Ausencias
if st.session_state.collaborators_list:
    with st.sidebar.expander("Registrar Ausencia", expanded=True):
        a_name = st.selectbox("Colaborador", st.session_state.collaborators_list)
        a_dates = st.date_input(
            "Rango de Fechas", 
            [datetime(sel_year, target_month, 1), datetime(sel_year, target_month, 1)]
        )
        if st.button("Añadir Ausencia"):
            if isinstance(a_dates, list) and len(a_dates) == 2:
                # Expandir rango
                start, end = a_dates[0], a_dates[1]
                curr = start
                while curr <= end:
                    st.session_state.absences.append({
                        "Nombre": a_name,
                        "Fecha": curr
                    })
                    curr += timedelta(days=1)
                st.success(f"Ausencia registrada para {a_name}")
            else:
                st.session_state.absences.append({
                    "Nombre": a_name,
                    "Fecha": a_dates[0] if isinstance(a_dates, list) else a_dates
                })
                st.success("Ausencia registrada")

# Listar ausencias
if st.session_state.absences:
    st.sidebar.write("---")
    st.sidebar.write("**Lista de Ausencias:**")
    abs_df = pd.DataFrame(st.session_state.absences)
    # Mostrar resumen simple
    st.sidebar.dataframe(abs_df, use_container_width=True)
    if st.sidebar.button("Borrar Todas las Ausencias"):
        st.session_state.absences = []
        st.rerun()

# --- LÓGICA PRINCIPAL ---

def run_logic(df_all):
    # 1. Aplicar Ausencias
    for abs_rec in st.session_state.absences:
        mask = (df_all['Nombre'] == abs_rec['Nombre']) & (df_all['Fecha'] == pd.to_datetime(abs_rec['Fecha']))
        df_all.loc[mask, 'Turno_Raw'] = 'AUSENTE'
        
    # 2. Expandir
    rows = []
    for _, row in df_all.iterrows():
        hours = parse_shift_time(row['Turno_Raw'])
        
        # Determinar si es Diurno/Nocturno (simple check hora inicio)
        start_h = hours[0] if hours else 0
        tipo_turno = "Diurno" if 5 <= start_h < 15 else "Nocturno"
        
        if not hours:
            # Fila vacía para mostrar libres/ausentes
            rows.append({**row, 'Hora': -1, 'Tarea': row['Turno_Raw'], 'Counter': '-', 'Tipo': tipo_turno})
        else:
            for h in hours:
                rows.append({**row, 'Hora': h, 'Tarea': '1', 'Counter': '?', 'Tipo': tipo_turno})
                
    df = pd.DataFrame(rows)
    if df.empty: return df
    
    # 3. Asignación Hora a Hora
    # Pools
    counters = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (date, hour), group in df[df['Hora'] != -1].groupby(['Fecha', 'Hora']):
        # Separar
        execs = group[(group['Rol']=='Ejecutivo')].index.tolist()
        coords = group[(group['Rol']=='Coordinador')].index.tolist()
        hosts = group[(group['Rol']=='Anfitriones')].index.tolist()
        
        # --- EJECUTIVOS ---
        # Colación
        active_execs = []
        for idx in execs:
            # Simular colación
            shift_type = df.at[idx, 'Tipo']
            is_col = False
            if shift_type == "Diurno" and hour in [13, 14]:
                if hash(df.at[idx, 'Nombre']) % 2 == (hour % 2): is_col = True
            elif shift_type == "Nocturno" and hour in [2, 3]:
                if hash(df.at[idx, 'Nombre']) % 2 == (hour % 2): is_col = True
            
            if is_col:
                df.at[idx, 'Tarea'] = 'C'
                df.at[idx, 'Counter'] = 'Casino'
            else:
                active_execs.append(idx)
        
        # Asignar Counters
        n_active = len(active_execs)
        uncovered = max(0, 4 - n_active)
        
        for i, idx in enumerate(active_execs):
            if i < 4:
                df.at[idx, 'Counter'] = counters[i]
                df.at[idx, 'Tarea'] = '1'
            else:
                # Sobra gente -> Refuerzo
                df.at[idx, 'Counter'] = "T1 AIRE" if i%2==0 else "T2 AIRE"
                df.at[idx, 'Tarea'] = '1'
                
        # --- COORDINADORES ---
        active_coords = []
        for idx in coords:
            # Tarea 2 (Admin) si no hay quiebre
            if uncovered == 0 and hour in [10, 11, 15, 16, 5, 6]:
                df.at[idx, 'Tarea'] = '2'
                df.at[idx, 'Counter'] = 'Oficina'
            else:
                active_coords.append(idx)
        
        # Cubrir Quiebre (Tarea 4)
        for idx in active_coords:
            if uncovered > 0:
                df.at[idx, 'Tarea'] = '4'
                df.at[idx, 'Counter'] = 'Cobertura'
                uncovered -= 1
            else:
                df.at[idx, 'Tarea'] = '1'
                df.at[idx, 'Counter'] = 'Piso'
                
        # --- ANFITRIONES ---
        for idx in hosts:
            if uncovered > 0:
                df.at[idx, 'Tarea'] = '4'
                df.at[idx, 'Counter'] = 'Apoyo'
                uncovered -= 1
            else:
                df.at[idx, 'Tarea'] = '1'
                df.at[idx, 'Counter'] = 'Zona ' + ('Int' if idx%2==0 else 'Nac')

    return df

def generate_excel_admin(df):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws = wb.add_worksheet("Admin Turnos")
    
    # Formatos
    f_header = wb.add_format({'bold': True, 'bg_color': '#DDDDDD', 'border': 1, 'align': 'center'})
    f_date = wb.add_format({'bold': True, 'bg_color': '#9FC5E8', 'border': 1, 'align': 'center'})
    f_cell = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    f_task4 = wb.add_format({'bg_color': '#FF9999', 'border': 1, 'align': 'center'})
    f_col = wb.add_format({'bg_color': '#93C47D', 'border': 1, 'align': 'center'})
    
    # Ejes
    dates = sorted(df['Fecha'].unique())
    names = df[['Nombre', 'Rol', 'Tipo']].drop_duplicates().sort_values(['Rol', 'Tipo', 'Nombre'])
    
    # Cabecera
    ws.write(1, 0, "Nombre", f_header)
    ws.write(1, 1, "Rol", f_header)
    
    col_idx = 2
    date_cols = {}
    
    for d in dates:
        d_str = pd.to_datetime(d).strftime("%d-%b")
        ws.merge_range(0, col_idx, 0, col_idx+25, d_str, f_date)
        
        ws.write(1, col_idx, "Turno", f_header)
        ws.write(1, col_idx+1, "Lugar", f_header)
        for h in range(24):
            ws.write(1, col_idx+2+h, h, f_header)
            
        date_cols[d] = col_idx
        col_idx += 26
        
    # Datos
    row_idx = 2
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_raw_idx = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    for _, p in names.iterrows():
        n = p['Nombre']
        ws.write(row_idx, 0, n, f_cell)
        ws.write(row_idx, 1, p['Rol'], f_cell)
        
        for d in dates:
            c = date_cols.get(d)
            if not c: continue
            
            # Datos base del día
            try:
                raw_shift = df_raw_idx.loc[(n, d), 'Turno_Raw']
                ws.write(row_idx, c, raw_shift, f_cell)
            except:
                ws.write(row_idx, c, "-", f_cell)
            
            # Horas
            for h in range(24):
                try:
                    res = df_idx.loc[(n, d, h)]
                    if isinstance(res, pd.DataFrame): res = res.iloc[0]
                    
                    task = res['Tarea']
                    cnt = res['Counter']
                    
                    fmt = f_cell
                    if task == '4': fmt = f_task4
                    if task == 'C': fmt = f_col
                    
                    ws.write(row_idx, c+2+h, task, fmt)
                    
                    # Escribir Counter principal en la columna Lugar (solo 1 vez)
                    if h == 12: # Al mediodía escribimos el counter aproximado
                         ws.write(row_idx, c+1, cnt if cnt != '?' else '-', f_cell)
                         
                except:
                    pass
        row_idx += 1
        
    wb.close()
    return output

# --- EJECUCIÓN ---

if st.button("🚀 2° Generar Planificación Final"):
    if len(uploaded_sheets) < 1:
        st.error("Debes cargar y seleccionar hojas de al menos un archivo.")
    else:
        with st.spinner("Procesando datos..."):
            all_dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                df_part = process_file_sheet(uf, us, role, target_month, sel_year)
                all_dfs.append(df_part)
            
            full_df = pd.concat(all_dfs)
            
            if full_df.empty:
                st.error("No se extrajeron datos. Revisa que las hojas seleccionadas correspondan al mes.")
            else:
                # Correr lógica
                final_df = run_logic(full_df)
                
                # Resumen
                st.success(f"Procesados {len(final_df['Nombre'].unique())} colaboradores.")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Horas Totales", len(final_df[final_df['Hora']!=-1]))
                c2.metric("Horas Extra (Tarea 4)", len(final_df[final_df['Tarea']=='4']))
                c3.metric("Ausencias Aplicadas", len(st.session_state.absences))
                
                # Excel
                excel_file = generate_excel_admin(final_df)
                
                st.download_button(
                    "📥 Descargar Excel Admin",
                    data=excel_file.getvalue(),
                    file_name=f"Planificacion_Admin_{sel_month_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
