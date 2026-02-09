import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Planificador Maestro de Turnos (Multi-Formato)")

# --- ESTADO DE SESIÓN ---
if 'absences' not in st.session_state:
    st.session_state.absences = []
if 'collaborators_list' not in st.session_state:
    st.session_state.collaborators_list = []

# --- PARSEO DE TURNOS ---

def parse_shift_time(shift_str):
    """Convierte '09:00 - 20:00' o '21:00 - 08:00' a lista de horas."""
    if pd.isna(shift_str): return []
    s = str(shift_str).lower().strip()
    
    # Palabras clave de ausencia/libre
    if any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla', 'domingos libres']):
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
        elif start_h > end_h: # Turno noche (cruce de día)
            return list(range(start_h, 24)) + list(range(0, end_h))
        else:
            return [start_h]
    except:
        return []

# --- LECTURA INTELIGENTE DE EXCEL ---

def find_date_header_row(df):
    """
    Busca la fila que contiene fechas (datetime) o números de día (1-31).
    Retorna: (indice_fila, tipo_detectado ['date', 'number'])
    """
    # Revisar las primeras 20 filas
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        date_count = 0
        number_count = 0
        
        for val in row:
            # Chequear si es fecha
            if isinstance(val, (datetime, pd.Timestamp)):
                date_count += 1
            # Chequear si es string fecha YYYY-MM-DD
            elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val):
                date_count += 1
            # Chequear si es número de día (Supervisor)
            elif isinstance(val, (int, float)):
                try:
                    if 1 <= int(val) <= 31:
                        number_count += 1
                except: pass
                
        # Umbral para decidir si es la cabecera
        if date_count > 3: # Si hay más de 3 fechas, es la fila de fechas
            return i, 'date'
        if number_count > 15: # Si hay muchos números (1..30), es supervisores
            return i, 'number'
            
    return None, None

def process_file_sheet(file, sheet_name, role, target_month, target_year):
    extracted_data = []
    try:
        # 1. Leer hoja cruda
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # 2. Buscar dónde están las fechas
        header_idx, header_type = find_date_header_row(df_raw)
        
        if header_idx is None:
            st.warning(f"⚠️ No se detectaron fechas en la hoja '{sheet_name}' ({role}). Revisa el formato.")
            return pd.DataFrame()
            
        # 3. Recargar usando esa fila como header
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        # 4. Identificar Columna de Nombres (generalmente la primera columna de texto a la izquierda)
        name_col = df.columns[0] # Default
        for col in df.columns:
            # Buscar columna que NO sea fecha y parezca nombre
            if "nombre" in str(col).lower() or "cargo" in str(col).lower() or "supervisor" in str(col).lower():
                name_col = col
                break
        
        # 5. Mapear columnas a fechas reales
        date_map = {} # {nombre_columna: objeto_datetime}
        
        for col in df.columns:
            col_date = None
            
            if header_type == 'date':
                # Intentar parsear la columna si es fecha
                if isinstance(col, (datetime, pd.Timestamp)):
                    col_date = col
                elif isinstance(col, str):
                    try: col_date = pd.to_datetime(col)
                    except: pass
            
            elif header_type == 'number':
                # Supervisores: la columna es un número (día del mes)
                try:
                    d = int(float(col))
                    col_date = datetime(target_year, target_month, d)
                except: pass
            
            # Validar que la fecha corresponde al mes seleccionado
            if col_date and col_date.month == target_month and col_date.year == target_year:
                date_map[col] = col_date

        # 6. Extraer Datos
        for idx, row in df.iterrows():
            name_val = row[name_col]
            
            # Filtros de basura (filas vacías o subtítulos)
            if pd.isna(name_val) or str(name_val).strip() == "" or str(name_val).lower() in ["nombre", "cargo", "supervisor", "turno", "fecha"]:
                continue
                
            # Extraer turnos de las columnas mapeadas
            for col_name, date_obj in date_map.items():
                shift_val = row[col_name]
                if pd.notna(shift_val):
                    extracted_data.append({
                        'Nombre': str(name_val).strip().title(),
                        'Rol': role,
                        'Fecha': date_obj,
                        'Turno_Raw': shift_val
                    })
                    
    except Exception as e:
        st.error(f"Error procesando {role}: {e}")
        
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
st.sidebar.header("2. Archivos")

files_config = [
    {"label": "Ejecutivos", "key": "exec"},
    {"label": "Anfitriones", "key": "host"},
    {"label": "Coordinadores", "key": "coord"},
    {"label": "Supervisores", "key": "sup"}
]

uploaded_sheets = {} 

for fconf in files_config:
    role = fconf["label"]
    file = st.sidebar.file_uploader(f"Excel {role}", type=["xlsx"], key=fconf["key"])
    
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

# --- UI AUSENCIAS ---
st.sidebar.markdown("---")
st.sidebar.header("3. Ausencias")

if st.sidebar.button("🔄 1° Leer Nombres"):
    if not uploaded_sheets:
        st.sidebar.error("Sube archivos primero.")
    else:
        all_names = set()
        with st.spinner("Leyendo nombres..."):
            for role, (uf, us) in uploaded_sheets.items():
                try:
                    # Leemos datos
                    df_t = process_file_sheet(uf, us, role, target_month, sel_year)
                    if not df_t.empty:
                        all_names.update(df_t['Nombre'].unique())
                except: pass
        
        st.session_state.collaborators_list = sorted(list(all_names))
        st.sidebar.success(f"Nombres cargados: {len(all_names)}")

if st.session_state.collaborators_list:
    with st.sidebar.expander("Agregar Ausencia"):
        p = st.selectbox("Persona", st.session_state.collaborators_list)
        d = st.date_input("Día", datetime(sel_year, target_month, 1))
        if st.button("Agregar"):
            st.session_state.absences.append({"Nombre": p, "Fecha": d})
            st.success("Listo")

if st.session_state.absences:
    st.sidebar.dataframe(pd.DataFrame(st.session_state.absences))
    if st.sidebar.button("Limpiar Ausencias"):
        st.session_state.absences = []
        st.rerun()

# --- LÓGICA DE NEGOCIO ---

def logic_engine(df):
    # 1. Marcar ausencias
    for a in st.session_state.absences:
        # Convertir fecha de ausencia a datetime64 para comparar
        abs_date = pd.to_datetime(a['Fecha'])
        mask = (df['Nombre'] == a['Nombre']) & (df['Fecha'] == abs_date)
        df.loc[mask, 'Turno_Raw'] = 'AUSENTE'

    # 2. Expandir
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
    
    # 3. Asignación
    counters_list = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        # --- EJECUTIVOS ---
        execs = g[g['Rol']=='Ejecutivo'].index.tolist()
        active = []
        
        for idx in execs:
            # Colación
            is_col = False
            # Regla simple: Diurno colación 13 o 14. Nocturno 2 o 3.
            # Usamos hash para distribuir mitad y mitad
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
        uncovered = max(0, 4 - len(active))
        
        for i, idx in enumerate(active):
            if i < 4:
                df_h.at[idx, 'Counter'] = counters_list[i]
            else:
                df_h.at[idx, 'Counter'] = "T1 AIRE" if i%2==0 else "T2 AIRE"
                
        # --- COORDINADORES ---
        coords = g[g['Rol']=='Coordinador'].index.tolist()
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
    ws = wb.add_worksheet("Sábana")
    
    fmt_h = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
    fmt_d = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#CFE2F3', 'align': 'center'})
    fmt_c = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    fmt_warn = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#F4CCCC', 'font_color': '#990000'})
    fmt_ok = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#D9EAD3'})

    # Headers
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
        
    # Filas
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
            
            # Turno Texto
            try:
                t_raw = df_base.loc[(n, d), 'Turno_Raw']
                ws.write(row, c, str(t_raw), fmt_c)
            except:
                ws.write(row, c, "-", fmt_c)
            
            # Horas
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
                except:
                    pass
        row += 1
        
    wb.close()
    return out

# --- MAIN ---

if st.button("🚀 2° Generar"):
    if not uploaded_sheets:
        st.error("Sube archivos y selecciona hojas.")
    else:
        dfs = []
        with st.spinner("Procesando..."):
            for role, (uf, us) in uploaded_sheets.items():
                d = process_file_sheet(uf, us, role, target_month, sel_year)
                dfs.append(d)
        
        full = pd.concat(dfs)
        if full.empty:
            st.error("No se encontraron datos. Verifica que el mes seleccionado coincida con las fechas en el Excel.")
        else:
            final = logic_engine(full)
            st.success(f"Procesados {len(final['Nombre'].unique())} personas.")
            
            # Métricas
            st.metric("Total Horas Cubiertas", len(final[final['Hora'] != -1]))
            
            x = make_excel(final)
            st.download_button("📥 Descargar Excel", x, f"Plan_{sel_month_name}.xlsx")
