import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Generador de Sábana de Turnos (V3: Ordenada y con Colaciones)")
st.markdown("""
**Novedades V3:**
1. **Orden Prioritario:** Ejecutivos primero (Diurno/Nocturno), luego Coordinadores (por horario), Anfitriones y Supervisores.
2. **Colaciones Visibles:** Regla 50/50 (mitad sale a una hora, mitad a la siguiente).
3. **Separadores Visuales:** El Excel incluye filas divisorias para facilitar la lectura.
""")

# --- PARSEO ROBUSTO ---
def parse_shift_time(shift_str):
    if pd.isna(shift_str): return [], None
    s = str(shift_str).lower().strip()
    if any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla', 'domingos libres', 'festivo']):
        return [], None
    
    s = s.replace(" diurno", "").replace(" nocturno", "").replace("hrs", "").strip()
    try:
        parts = re.split(r'\s*-\s*|\s*a\s*', s)
        if len(parts) < 2: return [], None
        
        formats = ["%H:%M:%S", "%H:%M", "%H"]
        start_h = -1
        end_h = -1
        
        for fmt in formats:
            try: 
                if start_h == -1: start_h = datetime.strptime(parts[0].strip(), fmt).hour
            except: pass
            try: 
                if end_h == -1: end_h = datetime.strptime(parts[1].strip(), fmt).hour
            except: pass
        
        if start_h == -1 or end_h == -1: return [], None
        
        if start_h < end_h: hours = list(range(start_h, end_h))
        elif start_h > end_h: hours = list(range(start_h, 24)) + list(range(0, end_h))
        else: hours = [start_h]
        
        return hours, start_h
    except: return [], None

# --- LECTURA ---
def find_date_header_row(df):
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        date_count = 0
        number_count = 0
        for val in row:
            if isinstance(val, (datetime, pd.Timestamp)): date_count += 1
            elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val): date_count += 1
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
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        header_idx, header_type = find_date_header_row(df_raw)
        
        if header_idx is None:
            st.warning(f"⚠️ No se detectaron fechas en '{sheet_name}' ({role}).")
            return pd.DataFrame()
            
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        name_col = df.columns[0]
        for col in df.columns:
            if "nombre" in str(col).lower() or "cargo" in str(col).lower() or "supervisor" in str(col).lower():
                name_col = col
                break
        
        date_map = {}
        for col in df.columns:
            col_date = None
            if header_type == 'date':
                if isinstance(col, (datetime, pd.Timestamp)): col_date = col
                elif isinstance(col, str):
                    try: col_date = pd.to_datetime(col)
                    except: pass
            elif header_type == 'number':
                try: col_date = datetime(target_year, target_month, int(float(col)))
                except: pass
            
            if col_date and col_date.month == target_month and col_date.year == target_year:
                date_map[col] = col_date

        for idx, row in df.iterrows():
            name_val = row[name_col]
            if pd.isna(name_val) or str(name_val).strip() == "": continue
            s_name = str(name_val).strip()
            if s_name.lower() in ["nombre", "cargo", "supervisor", "turno", "fecha"]: continue
            if any(k in s_name.lower() for k in ["total", "suma", "horas", "hh", "resumen"]): continue
            if s_name.replace('.', '', 1).isdigit(): continue

            clean_name = s_name.title()
            for col_name, date_obj in date_map.items():
                shift_val = row[col_name]
                if pd.notna(shift_val):
                    extracted_data.append({
                        'Nombre': clean_name, 'Rol': role, 'Fecha': date_obj, 'Turno_Raw': shift_val
                    })
    except Exception as e: st.error(f"Error en {role}: {e}")
    return pd.DataFrame(extracted_data)

# --- UI LATERAL ---
st.sidebar.header("1. Configuración")
months = {"Enero":1, "Febrero":2, "Marzo":3, "Abril":4, "Mayo":5, "Junio":6, 
          "Julio":7, "Agosto":8, "Septiembre":9, "Octubre":10, "Noviembre":11, "Diciembre":12}
sel_month_name = st.sidebar.selectbox("Mes", list(months.keys()), index=1)
sel_year = st.sidebar.number_input("Año", value=2026)
target_month = months[sel_month_name]

files_config = [{"label": "Ejecutivos", "key": "exec"}, {"label": "Anfitriones", "key": "host"},
                {"label": "Coordinadores", "key": "coord"}, {"label": "Supervisores", "key": "sup"}]
uploaded_sheets = {} 
for fconf in files_config:
    role = fconf["label"]
    file = st.sidebar.file_uploader(f"{role}", type=["xlsx"], key=fconf["key"])
    if file:
        try:
            xl = pd.ExcelFile(file)
            sheets = xl.sheet_names
            def_ix = next((i for i, s in enumerate(sheets) if sel_month_name.lower() in s.lower()), 0)
            sel_sheet = st.sidebar.selectbox(f"Hoja ({role})", sheets, index=def_ix, key=f"{role}_sheet")
            uploaded_sheets[role] = (file, sel_sheet)
        except: pass

# --- MOTOR DE LÓGICA V3 ---
def logic_engine(df):
    # 1. Expandir y Categorizar para Ordenamiento
    rows = []
    
    # Mapa de Prioridad de Roles (Para Excel)
    role_priority = {'Ejecutivo': 1, 'Coordinador': 2, 'Anfitriones': 3, 'Supervisor': 4}
    
    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        
        # Clasificación Diurno/Nocturno
        sub_group = "General"
        sort_val = 0
        
        if r['Rol'] == 'Ejecutivo' or r['Rol'] == 'Anfitriones':
            if start_h is not None:
                if 5 <= start_h < 15: 
                    sub_group = "Diurno"
                    sort_val = 1
                else: 
                    sub_group = "Nocturno"
                    sort_val = 2
        
        elif r['Rol'] == 'Coordinador':
            if start_h is not None:
                sub_group = f"Ingreso {start_h:02d}:00"
                sort_val = start_h # Ordenar por hora de ingreso
        
        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 
                         'Sub_Group': sub_group, 'Role_Rank': role_priority.get(r['Rol'], 9),
                         'Sort_Val': sort_val, 'Start_H': -1})
        else:
            for h in hours:
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                             'Sub_Group': sub_group, 'Role_Rank': role_priority.get(r['Rol'], 9),
                             'Sort_Val': sort_val, 'Start_H': start_h})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    # 2. Asignación Hora por Hora
    main_counters = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        # Obtener índices por rol
        execs = g[g['Rol']=='Ejecutivo'].index.tolist()
        coords = g[g['Rol']=='Coordinador'].index.tolist()
        hosts = g[g['Rol']=='Anfitriones'].index.tolist()
        sups = g[g['Rol']=='Supervisor'].index.tolist()
        
        # --- A. COLACIONES (Regla del 50/50) ---
        # Definimos función local para asignar colación
        def assign_colacion(indices, start_range, break_slots):
            # indices: lista de indices de personas
            # start_range: tupla (min, max) hora ingreso
            # break_slots: lista de horas de colación [12, 13]
            candidates = []
            for idx in indices:
                st_h = df_h.at[idx, 'Start_H']
                if start_range[0] <= st_h <= start_range[1]:
                    candidates.append(idx)
            
            # Distribuir
            n = len(candidates)
            for i, idx in enumerate(candidates):
                # Si hay 2 slots, mitad al slot 1, mitad al slot 2
                slot_idx = i % len(break_slots)
                target_hour = break_slots[slot_idx]
                
                if h == target_hour:
                    df_h.at[idx, 'Tarea'] = 'C'
                    df_h.at[idx, 'Counter'] = 'Casino'
        
        # Aplicar a Ejecutivos y Anfitriones
        # Diurnos (8-10): Colación 12, 13 (o 13, 14)
        assign_colacion(execs + hosts, (5, 14), [12, 13])
        # Nocturnos (20-22): Colación 2, 3
        assign_colacion(execs + hosts, (15, 23), [2, 3])
        assign_colacion(execs + hosts, (0, 4), [2, 3]) # Madrugada

        # Aplicar a Coordinadores (Reglas específicas)
        # Ingreso 05:00 -> Colación ~12
        assign_colacion(coords, (4, 6), [12]) 
        # Ingreso 10:00 -> Colación ~15
        assign_colacion(coords, (9, 11), [15])
        # Ingreso 21:00 -> Colación ~02
        assign_colacion(coords, (20, 23), [2])

        # --- B. FILTRAR DISPONIBLES ---
        # Solo los que NO están en 'C' pueden trabajar
        active_execs = [i for i in execs if df_h.at[i, 'Tarea'] != 'C']
        active_coords = [i for i in coords if df_h.at[i, 'Tarea'] != 'C']
        active_hosts = [i for i in hosts if df_h.at[i, 'Tarea'] != 'C']
        
        # --- C. ASIGNAR COUNTERS (Ejecutivos) ---
        spare_execs = []
        counters_state = {c: False for c in main_counters}
        
        # Ordenamos active_execs para que Diurnos y Nocturnos se mezclen o prioricen
        # (El orden natural del Excel suele funcionar, pero aquí es FIFO)
        for i, idx in enumerate(active_execs):
            if i < 4:
                cnt = main_counters[i]
                df_h.at[idx, 'Counter'] = cnt
                df_h.at[idx, 'Tarea'] = '1'
                counters_state[cnt] = True
            else:
                spare_execs.append(idx)
        
        # --- D. CUBRIR QUIEBRES (Tarea 3 y 4) ---
        # Si un counter quedó False (vacío por colación o falta de personal)
        for cnt_name, covered in counters_state.items():
            if not covered:
                # 1. Usar Flotante (Tarea 3)
                if spare_execs:
                    idx = spare_execs.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '3'
                # 2. Usar Coordinador (Tarea 4)
                elif active_coords:
                    # Verificar que el coordinador no tenga Tarea 2 obligatoria
                    idx = active_coords[0] 
                    # Lógica simplificada: si está libre de colación, cubre.
                    # (Idealmente chequearíamos su horario T2, pero priorizamos cobertura)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'
                    active_coords.pop(0)
                # 3. Usar Anfitrión (Tarea 4) - Ultimo recurso
                elif active_hosts:
                    idx = active_hosts.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'

        # --- E. ASIGNAR RESTO DE TAREAS ---
        
        # Ejecutivos Sobrantes -> Refuerzo Aire
        for i, idx in enumerate(spare_execs):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = "T1 AIRE" if i % 2 == 0 else "T2 AIRE"
            
        # Coordinadores Restantes -> Tarea 2 o Piso
        for idx in active_coords:
            st_h = df_h.at[idx, 'Start_H']
            task = '1'
            loc = 'Piso'
            
            # Reglas Tarea 2 Admin
            if st_h == 10:
                if h == 10 or h == 14: task='2'; loc='Oficina'
            elif st_h == 5:
                if h == 6 or h == 7: task='2'; loc='Oficina'
            elif st_h >= 21:
                if h == 5 or h == 6: task='2'; loc='Oficina'
                
            df_h.at[idx, 'Tarea'] = task
            df_h.at[idx, 'Counter'] = loc
            
        # Anfitriones Restantes -> Zona
        for i, idx in enumerate(active_hosts):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'
            
        # Supervisores
        for idx in sups:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'

    return df_h

# --- GENERADOR EXCEL CON SEPARADORES ---
def make_excel_ordered(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana Turnos")
    
    # Formatos
    f_header = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#404040', 'font_color': 'white', 'align': 'center'})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'left', 'indent': 1})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#CFE2F3', 'align': 'center'})
    
    # Estilos Condicionales
    styles = {
        '2': wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'}),
        '3': wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#DDEBF7'}),
        '4': wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#F4CCCC', 'font_color': '#990000', 'bold': True}),
        'C': wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#E2EFDA', 'font_color': '#38761D', 'bold': True})
    }
    f_def = f_base

    # Encabezados
    ws.write(1, 0, "Colaborador", f_header)
    ws.write(1, 1, "Rol", f_header)
    ws.freeze_panes(2, 2)
    
    dates = sorted(df['Fecha'].unique())
    col = 2
    d_map = {}
    for d in dates:
        d_lbl = pd.to_datetime(d).strftime("%d-%b")
        ws.merge_range(0, col, 0, col+25, d_lbl, f_date)
        ws.write(1, col, "Turno", f_header)
        ws.write(1, col+1, "Lugar", f_header)
        for h in range(24):
            ws.write(1, col+2+h, h, f_header)
        d_map[d] = col
        col += 26
    
    # ORDENAMIENTO DE DATOS PARA EXCEL
    # Clave de orden: Role_Rank -> Sort_Val (Diurno/Nocturno/Hora) -> Nombre
    df_unique = df[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank', 'Sort_Val']].drop_duplicates()
    df_sorted = df_unique.sort_values(by=['Role_Rank', 'Sort_Val', 'Nombre'])
    
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_base = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    row = 2
    last_group = None
    
    for _, p in df_sorted.iterrows():
        n, r, grp = p['Nombre'], p['Rol'], p['Sub_Group']
        
        # INSERTAR SEPARADOR SI CAMBIA EL GRUPO
        group_label = f"{r.upper()} - {grp.upper()}"
        if group_label != last_group:
            ws.merge_range(row, 0, row, col-1, group_label, f_group)
            row += 1
            last_group = group_label
        
        ws.write(row, 0, n, f_base)
        ws.write(row, 1, r, f_base)
        
        for d in dates:
            if d not in d_map: continue
            c = d_map[d]
            try:
                t_raw = df_base.loc[(n, d), 'Turno_Raw']
                ws.write(row, c, str(t_raw), f_base)
            except: ws.write(row, c, "-", f_base)
            
            for h in range(24):
                try:
                    val = df_idx.loc[(n, d, h)]
                    if isinstance(val, pd.DataFrame): val = val.iloc[0] # Manejo error duplicados
                    
                    task = val['Tarea']
                    cnt = val['Counter']
                    
                    if h == 12: ws.write(row, c+1, cnt if cnt!='?' else '-', f_base)
                    
                    style = styles.get(task, f_def)
                    ws.write(row, c+2+h, task, style)
                except: pass
        row += 1
        
    wb.close()
    return out

# --- EJECUCIÓN ---
if st.button("🚀 Generar Sábana Maestra V3"):
    if not uploaded_sheets:
        st.error("Por favor carga los archivos.")
    else:
        with st.spinner("Procesando Reglas V3 (Colaciones + Orden)..."):
            dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                dfs.append(process_file_sheet(uf, us, role, target_month, sel_year))
            
            full = pd.concat(dfs)
            if full.empty: st.error("Sin datos.")
            else:
                final = logic_engine(full)
                excel_file = make_excel_ordered(final)
                st.success("¡Planificación Generada!")
                st.download_button("📥 Descargar Excel V3", excel_file, f"Planificacion_V3_{sel_month_name}.xlsx")
