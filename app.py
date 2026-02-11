import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Gestor de Turnos: Reglas V8 (Tarea 3 Explícita + Continuidad)")
st.markdown("""
**Novedades V8:**
1. **Continuidad:** El agente mantiene su counter asignado durante todo el turno.
2. **Tarea 3 Detallada:** Si un counter queda vacío por colación, se asigna cobertura explícita (ej: "3: Cubrir T1 Tierra").
3. **Visualización:** Solo Diurno/Nocturno para agentes. Coordinadores/Supervisores en "General".
4. **Sin TICA:** Input selectivo que fuerza asignación a Tierra.
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
                    d_num = int(float(col))
                    col_date = datetime(start_date.year, start_date.month, d_num)
                except: pass
            
            if col_date:
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
st.sidebar.header("1. Periodo")
today = datetime.now()
date_range = st.sidebar.date_input("Rango de Fechas", (today.replace(day=1), today.replace(day=15)), format="DD/MM/YYYY")

if len(date_range) != 2:
    st.sidebar.warning("Selecciona inicio y fin.")
    start_d, end_d = None, None
else:
    start_d, end_d = date_range[0], date_range[1]

st.sidebar.markdown("---")
st.sidebar.header("2. Archivos")

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
            month_guess = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"][start_d.month] if start_d else ""
            def_ix = next((i for i, s in enumerate(sheets) if month_guess.lower() in s.lower()), 0)
            sel_sheet = st.sidebar.selectbox(f"Hoja ({role})", sheets, index=def_ix, key=f"{role}_sheet")
            uploaded_sheets[role] = (file, sel_sheet)
        except: pass

# --- SELECTOR SIN TICA ---
st.sidebar.markdown("---")
st.sidebar.header("3. Configuración TICA")
agents_no_tica = []

if 'exec' in uploaded_sheets and start_d and end_d:
    if st.sidebar.button("🔄 Cargar Nombres Agentes"):
        with st.spinner("Cargando..."):
            uf, us = uploaded_sheets['exec']
            df_temp = process_file_sheet(uf, us, "Agente", start_d, end_d)
            if not df_temp.empty:
                unique_names = sorted(df_temp['Nombre'].unique().tolist())
                st.session_state['agent_names_list'] = unique_names
                st.sidebar.success(f"¡{len(unique_names)} encontrados!")
    
    if 'agent_names_list' in st.session_state:
        agents_no_tica = st.sidebar.multiselect("Agentes SIN TICA (Asignar solo Tierra)", st.session_state['agent_names_list'])

# --- MOTOR LÓGICO V8 ---
def logic_engine(df, no_tica_list):
    rows = []
    
    # 1. EXPANDIR Y CLASIFICAR (DIURNO/NOCTURNO)
    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        
        # Clasificación simplificada
        sub_group = "General"
        role_rank = 99
        
        if r['Rol'] == 'Agente':
            if start_h is not None:
                if start_h < 12: # AM
                    sub_group = "Diurno"
                    role_rank = 10
                else: # PM
                    sub_group = "Nocturno"
                    role_rank = 11
            else:
                role_rank = 12
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
    
    # 2. PRE-ASIGNACIÓN DE COUNTERS (CONTINUIDAD DIARIA)
    # Asignar un "Counter Base" a cada agente para cada día
    main_counters_aire = ["T1 AIRE", "T2 AIRE"]
    main_counters_tierra = ["T1 TIERRA", "T2 TIERRA"]
    
    # Iterar por Día y por Agente para asignar su counter del día
    unique_dates = df_h['Fecha'].unique()
    
    daily_assignments = {} # (Nombre, Fecha) -> Counter
    
    for d in unique_dates:
        # Filtrar agentes que trabajan ese día
        agents_on_day = df_h[(df_h['Fecha'] == d) & (df_h['Rol'] == 'Agente')]['Nombre'].unique()
        
        # Contadores de carga para balancear
        load_balance = {c: 0 for c in main_counters_aire + main_counters_tierra}
        
        for ag_name in agents_on_day:
            # Determinar si tiene TICA
            has_tica = ag_name not in no_tica_list
            
            chosen_counter = None
            
            if not has_tica:
                # Solo Tierra: Buscar el menos cargado
                candidates = sorted(main_counters_tierra, key=lambda c: load_balance[c])
                chosen_counter = candidates[0]
            else:
                # Con TICA: Preferencia Aire, luego Tierra si Aire está muy lleno (heurística simple: balanceo total)
                # Pero la regla dice "Exceso a Aire", así que priorizamos llenar Aire y Tierra equitativamente
                # Vamos a llenar round-robin entre los 4 para asegurar cobertura min 1
                candidates = sorted(main_counters_aire + main_counters_tierra, key=lambda c: load_balance[c])
                chosen_counter = candidates[0]
            
            load_balance[chosen_counter] += 1
            daily_assignments[(ag_name, d)] = chosen_counter

    # 3. PROCESAMIENTO HORA POR HORA
    # Aplicar counter base, colaciones y gestionar quiebres
    
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        # Índices
        idx_agentes = g[g['Rol']=='Agente'].index.tolist()
        idx_coords = g[g['Rol']=='Coordinador'].index.tolist()
        idx_anf = g[g['Rol']=='Anfitrion'].index.tolist()
        idx_sups = g[g['Rol']=='Supervisor'].index.tolist()
        
        # --- A. APLICAR COUNTER BASE ---
        for idx in idx_agentes:
            name = df_h.at[idx, 'Nombre']
            base_cnt = daily_assignments.get((name, d), "General")
            df_h.at[idx, 'Counter'] = base_cnt
            df_h.at[idx, 'Tarea'] = '1'

        # --- B. APLICAR COLACIONES ---
        def apply_break(indices, start_range, slots):
            candidates = [i for i in indices if start_range[0] <= df_h.at[i, 'Start_H'] <= start_range[1]]
            for i, idx in enumerate(candidates):
                if h == slots[i % len(slots)]:
                    df_h.at[idx, 'Tarea'] = 'C'
                    df_h.at[idx, 'Counter'] = 'Casino' # Sale del counter
        
        # Agentes
        apply_break(idx_agentes, (0, 11), [13, 14]) # Diurno
        apply_break(idx_agentes, (12, 23), [2, 3]) # Nocturno
        # Anfitriones
        apply_break(idx_anf, (0, 11), [13, 14, 15])
        apply_break(idx_anf, (12, 23), [2, 3])
        # Coordinadores
        apply_break(idx_coords, (0, 6), [12]) 
        apply_break(idx_coords, (18, 23), [2])
        
        # --- C. DETECTAR QUIEBRES (Counter con 0 personas) ---
        # Contar cuántos activos hay en cada counter
        active_counts = {c: 0 for c in main_counters_aire + main_counters_tierra}
        donor_candidates = [] # Agentes trabajando (Tarea 1) que podrían moverse
        
        for idx in idx_agentes:
            if df_h.at[idx, 'Tarea'] == '1':
                cnt = df_h.at[idx, 'Counter']
                if cnt in active_counts:
                    active_counts[cnt] += 1
                    donor_candidates.append(idx)
        
        # Identificar counters vacíos (que deberían tener gente)
        # Asumimos que si un counter tiene 0 es un quiebre POTENCIAL si es horario operativo
        # Para simplificar: si un counter tiene 0, intentamos cubrirlo.
        
        empty_counters = [c for c, count in active_counts.items() if count == 0]
        
        # --- D. TAREA 3: CUBRIR QUIEBRES ---
        for target_cnt in empty_counters:
            covered = False
            
            # Buscar un donante: Alguien en un counter con > 1 persona
            # Ojo: Respetar TICA (Sin Tica no puede ir a Aire)
            
            best_donor = None
            
            # Prioridad: Mover de Aire a Aire, o de Tierra a Tierra, luego cruzado.
            # Filtrar donantes válidos (que su counter origen tenga > 1 persona para no destapar uno por tapar otro)
            valid_donors = []
            for d_idx in donor_candidates:
                origin_cnt = df_h.at[d_idx, 'Counter']
                if active_counts.get(origin_cnt, 0) > 1:
                    # Chequear restricción TICA
                    d_name = df_h.at[d_idx, 'Nombre']
                    is_no_tica = d_name in no_tica_list
                    
                    if "AIRE" in target_cnt and is_no_tica:
                        continue # Este no puede cubrir Aire
                    
                    valid_donors.append(d_idx)
            
            if valid_donors:
                # Tomar el primero (simplificación)
                best_donor = valid_donors[0]
                
                # EJECUTAR CAMBIO (Tarea 3)
                donor_origin = df_h.at[best_donor, 'Counter']
                df_h.at[best_donor, 'Tarea'] = f"3: Cubrir {target_cnt}" # Tarea explicita
                df_h.at[best_donor, 'Counter'] = target_cnt # Se mueve visualmente o lógicamente
                
                # Actualizar conteos
                active_counts[donor_origin] -= 1
                active_counts[target_cnt] += 1
                donor_candidates.remove(best_donor) # Ya no es donante
                covered = True
            
            # Si no hay donantes agentes, usar Coordinador/Anfitrión (Tarea 4)
            if not covered:
                # Coordinador
                avail_coords = [i for i in idx_coords if df_h.at[i, 'Tarea'] != 'C']
                if avail_coords:
                    c_idx = avail_coords[0]
                    df_h.at[c_idx, 'Tarea'] = f"4: Cubrir {target_cnt}"
                    df_h.at[c_idx, 'Counter'] = target_cnt
                    covered = True
                
                # Anfitrión
                elif not covered: # Si coord no pudo
                    avail_anf = [i for i in idx_anf if df_h.at[i, 'Tarea'] != 'C']
                    if avail_anf:
                        a_idx = avail_anf[0]
                        df_h.at[a_idx, 'Tarea'] = f"4: Cubrir {target_cnt}"
                        df_h.at[a_idx, 'Counter'] = target_cnt
                        covered = True

        # --- E. RESTO DE PERSONAL ---
        # Coordinadores y Supervisores en General
        for idx in idx_coords:
            if df_h.at[idx, 'Tarea'] == '1': # Si no está comiendo ni cubriendo
                df_h.at[idx, 'Counter'] = 'General'
        
        for idx in idx_sups:
            df_h.at[idx, 'Counter'] = 'General'
            df_h.at[idx, 'Tarea'] = '1'
            
        # Anfitriones
        avail_anf = [i for i in idx_anf if df_h.at[i, 'Tarea'] == '1']
        for i, idx in enumerate(avail_anf):
             df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'

    return df_h

# --- EXCEL ---
def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana V8")
    
    f_head = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#44546A', 'font_color': 'white', 'align': 'center'})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'text_wrap': True}) # Text wrap para Tarea 3
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#BFBFBF', 'align': 'left', 'indent': 1})
    
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        '3': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center', 'font_size': 8, 'text_wrap': True}), # Azulito y texto ajustado
        '4': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'bold': True, 'font_size': 8, 'text_wrap': True}),
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
        
    df_sorted = df[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    
    row = 2
    curr_group = ""
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_base = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    for _, p in df_sorted.iterrows():
        n, r, grp = p['Nombre'], p['Rol'], p['Sub_Group']
        
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
                    task = str(res['Tarea'])
                    cnt = res['Counter']
                    
                    if h == 12: ws.write(row, c+1, cnt if cnt!='?' else '-', f_base)
                    
                    # Detectar estilo (si empieza con 3 o 4)
                    style = f_base
                    if task.startswith('3'): style = st_map['3']
                    elif task.startswith('4'): style = st_map['4']
                    elif task == 'C': style = st_map['C']
                    elif task == '2': style = st_map['2']
                    
                    # Escribir solo número si es 1, o texto completo si es 3/4
                    val_to_write = task
                    
                    ws.write(row, c+2+h, val_to_write, style)
                except: pass
        row += 1
        
    wb.close()
    return out

# --- EJECUCIÓN ---
if st.button("🚀 Generar Planificación V8"):
    if not uploaded_sheets:
        st.error("Carga los archivos.")
    elif not (start_d and end_d):
        st.error("Selecciona fechas.")
    else:
        with st.spinner("Procesando..."):
            dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                dfs.append(process_file_sheet(uf, us, role, start_d, end_d))
            full = pd.concat(dfs)
            if full.empty: st.error("No hay datos.")
            else:
                final = logic_engine(full, agents_no_tica)
                st.success("¡Listo!")
                st.download_button("📥 Descargar Excel", make_excel(final), f"Planificacion_V8.xlsx")
