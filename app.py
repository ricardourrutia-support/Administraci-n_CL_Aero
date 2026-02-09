import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Planificador Maestro de Turnos (Todos los Ejecutivos)")

# --- ESTADO DE SESIÓN ---
if 'absences' not in st.session_state:
    st.session_state.absences = []
if 'collaborators_list' not in st.session_state:
    st.session_state.collaborators_list = []

# --- PARSEO DE TURNOS ---

def parse_shift_time(shift_str):
    """Convierte strings de turnos a listas de horas."""
    if pd.isna(shift_str): return []
    s = str(shift_str).lower().strip()
    
    # Filtros de "No turno"
    if any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla', 'domingos libres', 'festivo']):
        return []
    
    # Limpieza
    s = s.replace(" diurno", "").replace(" nocturno", "").replace("hrs", "").strip()
    
    try:
        parts = re.split(r'\s*-\s*|\s*a\s*', s)
        if len(parts) < 2: return []
        
        start_str, end_str = parts[0].strip(), parts[1].strip()
        
        formats = ["%H:%M:%S", "%H:%M", "%H"]
        
        def parse_h(h_str):
            for fmt in formats:
                try: return datetime.strptime(h_str, fmt).hour
                except: pass
            return None
            
        start_h = parse_h(start_str)
        end_h = parse_h(end_str)
        
        if start_h is None or end_h is None: return []
        
        if start_h < end_h:
            return list(range(start_h, end_h))
        elif start_h > end_h: # Cruce de medianoche (ej 22 a 07)
            return list(range(start_h, 24)) + list(range(0, end_h))
        else:
            return [start_h]
    except:
        return []

# --- LECTURA ROBUSTA (SALTA FILAS BASURA) ---

def find_date_header_row(df):
    """Busca la fila que contiene las fechas."""
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        date_count = 0
        number_count = 0
        
        for val in row:
            if isinstance(val, (datetime, pd.Timestamp)):
                date_count += 1
            elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val):
                date_count += 1
            elif isinstance(val, (int, float)):
                try:
                    if 1 <= int(val) <= 31: number_count += 1
                except: pass
                
        if date_count > 3: return i, 'date'
        if number_count > 15: return i, 'number'
            
    return None, None

def process_file_sheet(file, sheet_name, role, target_month, target_year):
    extracted_data = []
    try:
        # 1. Leer hoja completa (sin headers al principio)
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # 2. Buscar fila de fechas
        header_idx, header_type = find_date_header_row(df_raw)
        
        if header_idx is None:
            st.warning(f"⚠️ No se detectaron fechas en '{sheet_name}' ({role}).")
            return pd.DataFrame()
            
        # 3. Recargar usando esa fila como header
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        # 4. Encontrar columna de Nombre
        name_col = df.columns[0]
        for col in df.columns:
            c_str = str(col).lower()
            if "nombre" in c_str or "cargo" in c_str or "supervisor" in c_str:
                name_col = col
                break
        
        # 5. Mapear Columnas de Fecha
        date_map = {}
        for col in df.columns:
            col_date = None
            if header_type == 'date':
                if isinstance(col, (datetime, pd.Timestamp)): col_date = col
                elif isinstance(col, str):
                    try: col_date = pd.to_datetime(col)
                    except: pass
            elif header_type == 'number':
                try:
                    d = int(float(col))
                    col_date = datetime(target_year, target_month, d)
                except: pass
            
            if col_date and col_date.month == target_month and col_date.year == target_year:
                date_map[col] = col_date

        # 6. ITERAR TODAS LAS FILAS (SALTANDO BASURA)
        for idx, row in df.iterrows():
            name_val = row[name_col]
            
            # --- FILTROS DE LIMPIEZA ---
            # 1. Si es NaN o vacío
            if pd.isna(name_val) or str(name_val).strip() == "":
                continue
            
            s_name = str(name_val).strip()
            
            # 2. Si parece un encabezado repetido
            if s_name.lower() in ["nombre", "cargo", "supervisor", "turno", "fecha"]:
                continue
                
            # 3. Si es un número (Totales de horas, filas de resumen)
            # Esto elimina la fila de "números" que mencionaste
            if isinstance(name_val, (int, float)) or s_name.replace('.', '', 1).isdigit():
                continue
                
            # 4. Si contiene palabras clave de resumen
            if any(k in s_name.lower() for k in ["total", "suma", "horas", "hh", "resumen"]):
                continue

            # Si pasa los filtros, es una persona real
            clean_name = s_name.title()
            
            for col_name, date_obj in date_map.items():
                shift_val = row[col_name]
                if pd.notna(shift_val):
                    extracted_data.append({
                        'Nombre': clean_name,
                        'Rol': role,
                        'Fecha': date_obj,
                        'Turno_Raw': shift_val
                    })
                    
    except Exception as e:
        st.error(f"Error procesando {role}: {e}")
        
    return pd.DataFrame(extracted_data)

# --- UI LATERAL ---

st.sidebar.header("1. Fecha y Archivos")
months = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
    "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
}
sel_month_name = st.sidebar.selectbox("Mes", list(months.keys()), index=1)
sel_year = st.sidebar.number_input("Año", value=2026)
target_month = months[sel_month_name]

files_config = [
    {"label": "Ejecutivos", "key": "exec"},
    {"label": "Anfitriones", "key": "host"},
    {"label": "Coordinadores", "key": "coord"},
    {"label": "Supervisores", "key": "sup"}
]

uploaded_sheets = {} 

for fconf in files_config:
    role = fconf["label"]
    file = st.sidebar.file_uploader(f"{role}", type=["xlsx"], key=fconf["key"])
    
    if file:
        try:
            xl = pd.ExcelFile(file)
            sheets = xl.sheet_names
            # Auto-sugerencia
            def_ix = 0
            for i, s in enumerate(sheets):
                if sel_month_name.lower() in s.lower():
                    def_ix = i
                    break
            
            sel_sheet = st.sidebar.selectbox(f"Hoja ({role})", sheets, index=def_ix, key=f"{role}_sheet")
            uploaded_sheets[role] = (file, sel_sheet)
        except:
            st.sidebar.error("Error leyendo archivo.")

# --- AUSENCIAS ---
st.sidebar.markdown("---")
st.sidebar.header("2. Ausencias")

if st.sidebar.button("🔄 Cargar Lista de Nombres"):
    if not uploaded_sheets:
        st.sidebar.error("Sube archivos primero.")
    else:
        all_names = set()
        with st.spinner("Escaneando..."):
            for role, (uf, us) in uploaded_sheets.items():
                try:
                    df_t = process_file_sheet(uf, us, role, target_month, sel_year)
                    if not df_t.empty:
                        all_names.update(df_t['Nombre'].unique())
                except: pass
        st.session_state.collaborators_list = sorted(list(all_names))
        st.sidebar.success(f"Encontrados: {len(all_names)}")

if st.session_state.collaborators_list:
    with st.sidebar.expander("Añadir Ausencia"):
        p = st.selectbox("Persona", st.session_state.collaborators_list)
        d_range = st.date_input("Fechas", [])
        if st.button("Guardar"):
            st.session_state.absences.append({"Nombre": p, "Rango": str(d_range)})
            st.success("Guardado (Recuerda regenerar)")

# --- LÓGICA DE NEGOCIO ---

def logic_engine(df):
    # Expandir
    rows = []
    for _, r in df.iterrows():
        hours = parse_shift_time(r['Turno_Raw'])
        # Tipo turno
        sh_type = "Diurno"
        if hours and (hours[0] >= 20 or hours[0] < 6): sh_type = "Nocturno"
        
        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 'Tipo': sh_type})
        else:
            for h in hours:
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 'Tipo': sh_type})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    # Asignación
    counters_list = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        # --- EJECUTIVOS ---
        execs = g[g['Rol']=='Ejecutivo'].index.tolist()
        active = []
        
        for idx in execs:
            # Colación
            is_col = False
            row_hash = hash(df_h.at[idx, 'Nombre'])
            if df_h.at[idx, 'Tipo'] == 'Diurno' and h in [13, 14]:
                if row_hash % 2 == (h % 2): is_col = True
            elif df_h.at[idx, 'Tipo'] == 'Nocturno' and h in [2, 3]:
                if row_hash % 2 == (h % 2): is_col = True
                
            if is_col:
                df_h.at[idx, 'Tarea'] = 'C'
                df_h.at[idx, 'Counter'] = 'Casino'
            else:
                active.append(idx)
        
        # Asignar a Counters
        for i, idx in enumerate(active):
            if i < 4:
                df_h.at[idx, 'Counter'] = counters_list[i]
            else:
                df_h.at[idx, 'Counter'] = "T1 AIRE" if i%2==0 else "T2 AIRE"
                
        # --- COORDINADORES ---
        coords = g[g['Rol']=='Coordinador'].index.tolist()
        uncovered = max(0, 4 - len(active))
        avail_coords = []
        
        for idx in coords:
            # Tarea 2 si horario lo permite y no hay quiebre
            is_t2 = False
            if uncovered == 0:
                if h in [10, 11, 15, 16, 5, 6]: is_t2 = True
            
            if is_t2:
                df_h.at[idx, 'Tarea'] = '2'
                df_h.at[idx, 'Counter'] = 'Oficina'
            else:
                avail_coords.append(idx)
                
        # Tarea 4 (Quiebre)
        for idx in avail_coords:
            if uncovered > 0:
                df_h.at[idx, 'Tarea'] = '4'
                df_h.at[idx, 'Counter'] = 'Cobertura'
                uncovered -= 1
            else:
                df_h.at[idx, 'Tarea'] = '1'
                df_h.at[idx, 'Counter'] = 'Piso'

        # --- ANFITRIONES ---
        hosts = g[g['Rol']=='Anfitriones'].index.tolist()
        for i, idx in enumerate(hosts):
            if uncovered > 0:
                df_h.at[idx, 'Tarea'] = '4'
                df_h.at[idx, 'Counter'] = 'Apoyo'
                uncovered -= 1
            else:
                df_h.at[idx, 'Tarea'] = '1'
                df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'
                
        # --- SUPERVISORES ---
        sups = g[g['Rol']=='Supervisor'].index.tolist()
        for idx in sups:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'

    return df_h

# --- EXCEL ---

def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana Mensual")
    
    fmt_h = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
    fmt_d = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#CFE2F3', 'align': 'center'})
    fmt_c = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    fmt_warn = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#F4CCCC', 'font_color': '#990000'})
    fmt_ok = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#D9EAD3'})

    ws.write(1, 0, "Colaborador", fmt_h)
    ws.write(1, 1, "Rol", fmt_h)
    
    dates = sorted(df['Fecha'].unique())
    col = 2
    d_map = {}
    
    for d in dates:
        d_lbl = pd.to_datetime(d).strftime("%d-%b")
        ws.merge_range(0, col, 0, col+25, d_lbl, fmt_d)
        ws.write(1, col, "Turno", fmt_h)
        ws.write(1, col+1, "Lugar", fmt_h)
        for h in range(24):
            ws.write(1, col+2+h, h, fmt_h)
        d_map[d] = col
        col += 26
        
    people = df[['Nombre', 'Rol']].drop_duplicates().sort_values(['Rol', 'Nombre'])
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_base = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    row = 2
    for _, p in people.iterrows():
        n, r = p['Nombre'], p['Rol']
        ws.write(row, 0, n, fmt_c)
        ws.write(row, 1, r, fmt_c)
        
        for d in dates:
            if d not in d_map: continue
            c = d_map[d]
            
            try:
                t_raw = df_base.loc[(n, d), 'Turno_Raw']
                ws.write(row, c, str(t_raw), fmt_c)
            except: ws.write(row, c, "-", fmt_c)
            
            for h in range(24):
                try:
                    val = df_idx.loc[(n, d, h)]
                    if isinstance(val, pd.DataFrame): val = val.iloc[0]
                    
                    task = val['Tarea']
                    cnt = val['Counter']
                    
                    if h == 12: ws.write(row, c+1, cnt if cnt!='?' else '-', fmt_c)
                    
                    style = fmt_c
                    if task == '4': style = fmt_warn
                    if task == 'C': style = fmt_ok
                    
                    ws.write(row, c+2+h, task, style)
                except: pass
        row += 1
        
    wb.close()
    return out

# --- MAIN ---

if st.button("🚀 Generar Planificación"):
    if not uploaded_sheets:
        st.error("Carga archivos primero.")
    else:
        dfs = []
        with st.spinner("Procesando todas las filas (ignorando basura)..."):
            for role, (uf, us) in uploaded_sheets.items():
                d = process_file_sheet(uf, us, role, target_month, sel_year)
                dfs.append(d)
        
        full = pd.concat(dfs)
        if full.empty:
            st.error("No se encontraron datos.")
        else:
            final = logic_engine(full)
            st.success(f"Analizados {len(final['Nombre'].unique())} colaboradores.")
            st.download_button("📥 Descargar Excel", make_excel(final), f"Plan_{sel_month_name}.xlsx")
