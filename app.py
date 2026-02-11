import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Gestor de Turnos: Reglas V6 (Filtro Fechas + Tarea 3)")
st.markdown("""
**Nuevas Funcionalidades:**
1. **Rango de Fechas:** Selecciona días específicos del mes.
2. **Sin TICA:** Carga agentes y restringe su asignación a Tierra.
3. **Clasificación:** Agentes separados en Diurno (AM) y Nocturno (PM).
4. **Tarea 3:** Lógica específica para cubrir quiebres por colación.
""")

# --- PARSEO DE FECHAS Y HORAS ---
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

def process_file_sheet(file, sheet_name, role, start_date, end_date):
    """
    Lee el Excel y filtra SOLO las fechas dentro del rango seleccionado.
    """
    extracted_data = []
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        header_idx, header_type = find_date_header_row(df_raw)
        
        if header_idx is None:
            return pd.DataFrame()
            
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        
        name_col = df.columns[0]
        for col in df.columns:
            if "nombre" in str(col).lower() or "cargo" in str(col).lower() or "supervisor" in str(col).lower():
                name_col = col
                break
        
        # Mapear columnas a fechas y FILTRAR por rango
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
                    # Asumimos que el usuario selecciona el año/mes correcto en el filtro
                    # Usamos el mes/año del start_date seleccionado
                    d_num = int(float(col))
                    col_date = datetime(start_date.year, start_date.month, d_num)
                except: pass
            
            # FILTRO DE FECHAS
            if col_date:
                # Convertir a date para comparar
                c_dt = col_date.date() if isinstance(col_date, datetime) else col_date
                if start_date <= c_dt <= end_date:
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
st.sidebar.header("1. Configuración de Periodo")

# Selector de Rango de Fechas
today = datetime.now()
date_range = st.sidebar.date_input(
    "Selecciona Rango de Fechas",
    (today.replace(day=1), today.replace(day=15)), # Default
    format="DD/MM/YYYY"
)

if len(date_range) != 2:
    st.sidebar.warning("Selecciona fecha de inicio y fin.")
    start_d, end_d = None, None
else:
    start_d, end_d = date_range[0], date_range[1]

st.sidebar.markdown("---")
st.sidebar.header("2. Carga de Archivos")

files_config = [{"label": "Agente", "key": "exec"}, {"label": "Coordinador", "key": "coord"},
                {"label": "Anfitrion", "key": "host"}, {"label": "Supervisor", "key": "sup"}]
uploaded_sheets = {} 

for fconf in files_config:
    role = fconf["label"]
    file = st.sidebar.file_uploader(f"{role}", type=["xlsx"], key=fconf["key"])
    if file:
        try:
            xl = pd.ExcelFile(file)
            sheets = xl.sheet_names
            # Intentar adivinar la hoja por el mes seleccionado en el rango
            month_name_guess = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto"][start_d.month] if start_d else ""
            def_ix = next((i for i, s in enumerate(sheets) if month_name_guess.lower() in s.lower()), 0)
            sel_sheet = st.sidebar.selectbox(f"Hoja ({role})", sheets, index=def_ix, key=f"{role}_sheet")
            uploaded_sheets[role] = (file, sel_sheet)
        except: pass

# --- SELECTOR SIN TICA ---
st.sidebar.markdown("---")
st.sidebar.header("3. Excepciones (TICA)")
agents_no_tica = []

if 'exec' in uploaded_sheets and start_d and end_d:
    if st.sidebar.button("🔄 Cargar Nombres de Agentes"):
        with st.spinner("Leyendo archivo de Agentes..."):
            uf, us = uploaded_sheets['exec']
            df_temp = process_file_sheet(uf, us, "Agente", start_d, end_d)
            if not df_temp.empty:
                unique_names = sorted(df_temp['Nombre'].unique().tolist())
                st.session_state['agent_names_list'] = unique_names
                st.sidebar.success(f"¡{len(unique_names)} agentes encontrados!")
            else:
                st.sidebar.warning("No se encontraron nombres. Revisa la hoja o el rango de fechas.")

    if 'agent_names_list' in st.session_state:
        agents_no_tica = st.sidebar.multiselect(
            "Agentes SIN TICA (Solo Tierra)", 
            st.session_state['agent_names_list']
        )
else:
    st.sidebar.info("Sube Agentes y define fechas para configurar TICA.")

# --- MOTOR LÓGICO V6 ---
def logic_engine(df, no_tica_list):
    rows = []
    
    # 1. EXPANDIR Y CLASIFICAR (DIURNO/NOCTURNO)
    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        
        # Clasificación Diurno (AM < 12) / Nocturno (PM >= 12)
        sub_group = "General"
        role_rank = 99
        
        if r['Rol'] == 'Agente':
            if start_h is not None:
                if start_h < 12: # Ingreso AM
                    sub_group = "Diurno (AM)"
                    role_rank = 10
                else: # Ingreso PM
                    sub_group = "Nocturno (PM)"
                    role_rank = 11
            else:
                role_rank = 12 # Libre/Desconocido
                
        elif r['Rol'] == 'Coordinador':
            role_rank = 20
        elif r['Rol'] == 'Anfitrion':
            role_rank = 30
        elif r['Rol'] == 'Supervisor':
            role_rank = 40
            
        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 
                         'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': -1})
        else:
            for h in hours:
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                             'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': start_h})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    main_counters = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    # 2. PROCESAR POR FRANJA HORARIA
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        agentes = g[g['Rol']=='Agente'].index.tolist()
        coords = g[g['Rol']=='Coordinador'].index.tolist()
        anfitriones = g[g['Rol']=='Anfitrion'].index.tolist()
        sups = g[g['Rol']=='Supervisor'].index.tolist()
        
        # --- A. ASIGNAR COLACIONES ---
        def apply_break(indices_list, valid_start_range, break_slots):
            candidates = []
            for idx in indices_list:
                st_h = df_h.at[idx, 'Start_H']
                # valid_start_range es (min, max) inclusivo
                if valid_start_range[0] <= st_h <= valid_start_range[1]:
                    candidates.append(idx)
            
            for i, idx in enumerate(candidates):
                slot_idx = i % len(break_slots)
                if h == break_slots[slot_idx]:
                    df_h.at[idx, 'Tarea'] = 'C'
                    df_h.at[idx, 'Counter'] = 'Casino'

        # Agentes Diurnos (Ingreso < 12) -> Colación 12-15
        apply_break(agentes, (0, 11), [13, 14]) 
        # Agentes Nocturnos (Ingreso >= 12) -> Colación 02-04
        apply_break(agentes, (12, 23), [2, 3])
        
        # Anfitriones y Coordinadores (Reglas estándar)
        apply_break(anfitriones, (0, 11), [13, 14])
        apply_break(anfitriones, (12, 23), [2, 3])
        apply_break(coords, (0, 6), [12]) 
        apply_break(coords, (18, 23), [2]) 
        
        # --- B. FILTRAR DISPONIBLES ---
        active_agentes = [i for i in agentes if df_h.at[i, 'Tarea'] != 'C']
        active_coords = [i for i in coords if df_h.at[i, 'Tarea'] != 'C']
        active_anfitriones = [i for i in anfitriones if df_h.at[i, 'Tarea'] != 'C']
        
        # --- C. ASIGNAR AGENTES A COUNTERS ---
        with_tica = []
        no_tica = []
        for idx in active_agentes:
            if df_h.at[idx, 'Nombre'] in no_tica_list:
                no_tica.append(idx)
            else:
                with_tica.append(idx)
        
        counters_status = {c: False for c in main_counters}
        
        # 1. Llenar TIERRA (Prioridad Sin TICA)
        tierra_cnts = ["T1 TIERRA", "T2 TIERRA"]
        for t_cnt in tierra_cnts:
            if no_tica:
                idx = no_tica.pop(0)
                df_h.at[idx, 'Counter'] = t_cnt
                df_h.at[idx, 'Tarea'] = '1'
                counters_status[t_cnt] = True
            elif with_tica:
                idx = with_tica.pop(0)
                df_h.at[idx, 'Counter'] = t_cnt
                df_h.at[idx, 'Tarea'] = '1'
                counters_status[t_cnt] = True
        
        spare_no_tica = no_tica # Estos sobran de tierra

        # 2. Llenar AIRE (Solo Con TICA)
        aire_cnts = ["T1 AIRE", "T2 AIRE"]
        for a_cnt in aire_cnts:
            if with_tica:
                idx = with_tica.pop(0)
                df_h.at[idx, 'Counter'] = a_cnt
                df_h.at[idx, 'Tarea'] = '1'
                counters_status[a_cnt] = True
        
        spare_with_tica = with_tica # Estos sobran de aire (Flotantes potenciales)
        
        # --- D. TAREA 3: COBERTURA DE QUIEBRE (Solo si hay quiebre) ---
        for cnt_name, covered in counters_status.items():
            if not covered:
                # ¡ALERTA! Quiebre detectado (por Colación probablemente)
                filled = False
                
                # Buscar un Flotante (Spare) para Tarea 3
                # REGLA: Tarea 3 es cubrir quiebre.
                
                # Intentar cubrir Aire
                if "AIRE" in cnt_name and spare_with_tica:
                    idx = spare_with_tica.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '3' # Asignación de Cobertura
                    filled = True
                
                # Intentar cubrir Tierra
                elif "TIERRA" in cnt_name:
                    if spare_no_tica:
                        idx = spare_no_tica.pop(0)
                        df_h.at[idx, 'Counter'] = cnt_name
                        df_h.at[idx, 'Tarea'] = '3'
                        filled = True
                    elif spare_with_tica:
                        idx = spare_with_tica.pop(0)
                        df_h.at[idx, 'Counter'] = cnt_name
                        df_h.at[idx, 'Tarea'] = '3'
                        filled = True

                # Si no hay agentes flotantes, usar Coordinador/Anfitrión (Tarea 4)
                if not filled and active_coords:
                    idx = active_coords[0]
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'
                    active_coords.pop(0)
                    filled = True
                elif not filled and active_anfitriones:
                    idx = active_anfitriones.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'
                    active_anfitriones.pop(0)
                    filled = True

        # --- E. ASIGNAR SOBRANTES (RESTO DE TAREAS) ---
        
        # Agentes que sobraron y NO hicieron Tarea 3 -> Refuerzo (Tarea 1)
        for idx in spare_no_tica:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = "Refuerzo Tierra"
        for i, idx in enumerate(spare_with_tica):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = "T1 AIRE" if i%2==0 else "T2 AIRE"
            
        # Coordinadores restantes
        for idx in active_coords:
            st_h = df_h.at[idx, 'Start_H']
            task = '1'
            cnt = 'Piso'
            # Heurística Tarea 2
            if st_h == 10 and (h == 10 or h in [14, 15]): task = '2'; cnt = 'Oficina'
            elif st_h == 5 and (h in [6, 7]): task = '2'; cnt = 'Oficina'
            df_h.at[idx, 'Tarea'] = task
            df_h.at[idx, 'Counter'] = cnt
            
        # Anfitriones restantes
        for i, idx in enumerate(active_anfitriones):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'
            
        # Supervisores
        for idx in sups:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'

    return df_h

# --- GENERADOR EXCEL ---
def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana Turnos")
    
    f_head = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#44546A', 'font_color': 'white', 'align': 'center'})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#BFBFBF', 'align': 'left', 'indent': 1})
    
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        '3': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center'}),
        '4': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'bold': True}),
        'C': wb.add_format({'bg_color': '#C6E0B4', 'border': 1, 'align': 'center'}),
        '1': f_base
    }

    ws.write(1, 0, "Colaborador", f_head)
    ws.write(1, 1, "Rol", f_head)
    ws.freeze_panes(2, 2)
    
    dates = sorted(df['Fecha'].unique())
    col = 2
    d_map = {}
    for d in dates:
        d_str = pd.to_datetime(d).strftime("%d-%b")
        ws.merge_range(0, col, 0, col+25, d_str, f_date)
        ws.write(1, col, "Turno", f_head)
        ws.write(1, col+1, "Lugar", f_head)
        for h in range(24):
            ws.write(1, col+2+h, h, f_head)
        d_map[d] = col
        col += 26
        
    # ORDEN: Agente Diurno -> Agente Nocturno -> Coord -> Anfitrion -> Supervisor
    df_sorted = df[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    
    row = 2
    curr_group = ""
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_base = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    for _, p in df_sorted.iterrows():
        n, r, grp = p['Nombre'], p['Rol'], p['Sub_Group']
        
        # Etiqueta de Grupo
        grp_label = f"{r.upper()}"
        if r == "Agente": grp_label += f" - {grp}"
        
        if grp_label != curr_group:
            ws.merge_range(row, 0, row, col-1, grp_label, f_group)
            row += 1
            curr_group = grp_label
            
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
                    res = df_idx.loc[(n, d, h)]
                    if isinstance(res, pd.DataFrame): res = res.iloc[0]
                    task = res['Tarea']
                    cnt = res['Counter']
                    if h == 12: ws.write(row, c+1, cnt if cnt!='?' else '-', f_base)
                    ws.write(row, c+2+h, task, st_map.get(task, f_base))
                except: pass
        row += 1
        
    wb.close()
    return out

# --- EJECUCIÓN ---
if st.button("🚀 Generar Planificación V6"):
    if not uploaded_sheets:
        st.error("Carga los archivos.")
    elif not (start_d and end_d):
        st.error("Selecciona rango de fechas.")
    else:
        with st.spinner(f"Analizando del {start_d} al {end_d}..."):
            dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                dfs.append(process_file_sheet(uf, us, role, start_d, end_d))
            full = pd.concat(dfs)
            
            if full.empty: st.error("No hay datos en ese rango.")
            else:
                final = logic_engine(full, agents_no_tica)
                st.success("¡Planificación Creada!")
                st.download_button("📥 Descargar Excel", make_excel(final), f"Planificacion_V6.xlsx")
