import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto (V53)", layout="wide")
st.title("✈️ Gestor de Turnos: V53 (Lógica Coordinación Estricta)")
st.markdown("""
**Corrección V53:**
* **Regla de Cobertura Coordinador:** Solo reemplazan si están en Tarea 2 (Admin) o si hay más de un Coordinador en Tarea 1 (Supervisión) al mismo tiempo.
* **Ubicación:** El Coordinador siempre mantiene ubicación "General".
* **Visual:** Se mantienen todas las mejoras gráficas y la bitácora inteligente.
""")

# --- INICIALIZACIÓN ---
if 'incidencias' not in st.session_state:
    st.session_state.incidencias = []

today = datetime.now()
uploaded_sheets = {}
start_d = None
end_d = None
agents_no_tica = []

# --- PARSEO ---
def parse_shift_time(shift_str):
    if pd.isna(shift_str): return [], None
    s = str(shift_str).lower().strip()
    if s == "" or any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla', 'domingos libres', 'festivo', 'feriado']):
        return [], None
    
    s = s.replace(" diurno", "").replace(" nocturno", "").replace("hrs", "").replace("horas", "").replace("de", "").replace("a", "-").replace("–", "-").replace("to", "-")
    match = re.search(r'(\d{1,2})(?:[:.]\d+)*\s*-\s*(\d{1,2})(?:[:.]\d+)*', s)
    
    start_h = -1
    end_h = -1
    if match:
        try:
            start_h = int(match.group(1))
            end_h = int(match.group(2))
            if 0 <= start_h <= 24 and 0 <= end_h <= 24:
                if start_h < end_h:
                    return list(range(start_h, end_h)), start_h
                elif start_h > end_h:
                    return list(range(start_h, 24)) + list(range(0, end_h)), start_h
        except: pass
    return [], None

def find_date_header_row(df):
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        date_count = 0
        for val in row:
            if isinstance(val, (datetime, pd.Timestamp)): date_count += 1
            elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val): date_count += 1
        if date_count > 3: return i, 'date'
    return None, None

def process_file_sheet(file, sheet_name, role, start_date, end_date):
    extracted_data = []
    try:
        file.seek(0)
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
        header_idx, header_type = find_date_header_row(df_raw)
        if header_idx is None: return pd.DataFrame()
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
        name_col = df.columns[0]
        for col in df.columns:
            if isinstance(col, str) and ("nombre" in col.lower() or "cargo" in col.lower() or "supervisor" in col.lower()):
                name_col = col
                break
        date_map = {}
        if start_date:
            load_start = start_date - timedelta(days=1)
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
                        if d_num > 20 and start_date.day < 5:
                             col_date = col_date - timedelta(days=30)
                    except: pass
                if col_date:
                    c_dt = col_date.date() if isinstance(col_date, datetime) else col_date
                    if load_start <= c_dt <= end_date:
                        date_map[col] = col_date

            for idx, row in df.iterrows():
                name_val = row[name_col]
                if pd.isna(name_val): continue
                s_name = str(name_val).strip()
                if s_name == "" or len(s_name) < 3: continue
                if any(k in s_name.lower() for k in ["nombre", "cargo", "turno", "fecha", "total", "suma", "horas"]): continue
                if s_name.replace('.', '', 1).isdigit(): continue

                clean_name = s_name.title()
                for col_name, date_obj in date_map.items():
                    shift_val = row[col_name]
                    if pd.isna(shift_val): shift_val = ""
                    extracted_data.append({
                        'Nombre': clean_name, 'Rol': role, 'Fecha': date_obj, 'Turno_Raw': shift_val
                    })
    except Exception as e: st.error(f"Error en {role}: {e}")
    return pd.DataFrame(extracted_data)

# --- APLICAR INCIDENCIAS ---
def apply_incidents(df, incidents):
    df_real = df.copy()
    for inc in incidents:
        tipo = inc['tipo']
        nombre = inc['nombre']
        fecha_ini = inc['fecha_inicio']
        fecha_fin = inc['fecha_fin']
        mask_name = df_real['Nombre'] == nombre
        
        if tipo == 'Inasistencia':
            mask_date = (df_real['Fecha'].dt.date >= fecha_ini) & (df_real['Fecha'].dt.date <= fecha_fin)
            target_rows = df_real[mask_name & mask_date].index
            df_real.loc[target_rows, 'Hora'] = -1
            df_real.loc[target_rows, 'Tarea'] = 'Ausente'
            df_real.loc[target_rows, 'Turno_Raw'] = 'Falta'
            
        elif tipo == 'Atraso':
            hora_llegada = inc['hora_impacto']
            mask_date = df_real['Fecha'].dt.date == fecha_ini
            mask_time = df_real['Hora'] < hora_llegada
            target_rows = df_real[mask_name & mask_date & mask_time & (df_real['Hora'] != -1)].index
            df_real.loc[target_rows, 'Hora'] = -1
            df_real.loc[target_rows, 'Tarea'] = 'Atraso'
            
        elif tipo == 'Salida Anticipada':
            hora_salida = inc['hora_impacto']
            mask_date = df_real['Fecha'].dt.date == fecha_ini
            mask_time = df_real['Hora'] >= hora_salida
            target_rows = df_real[mask_name & mask_date & mask_time & (df_real['Hora'] != -1)].index
            df_real.loc[target_rows, 'Hora'] = -1
            df_real.loc[target_rows, 'Tarea'] = 'Salida Ant.'
    return df_real

# --- UI LATERAL ---
st.sidebar.header("1. Periodo")
date_range = st.sidebar.date_input("Rango", (today.replace(day=1), today.replace(day=15)), format="DD/MM/YYYY")
if len(date_range) == 2:
    start_d, end_d = date_range

st.sidebar.markdown("---")
st.sidebar.header("2. Archivos")
for label, key in [("Agente", "exec"), ("Coordinador", "coord"), ("Anfitrion", "host"), ("Supervisor", "sup")]:
    f = st.sidebar.file_uploader(f"{label}", type=["xlsx"], key=key)
    if f and start_d:
        try:
            xl = pd.ExcelFile(f)
            def_ix = 0
            sel_sheet = st.sidebar.selectbox(f"Hoja ({label})", xl.sheet_names, index=def_ix, key=f"{key}_sheet")
            uploaded_sheets[key] = (f, sel_sheet)
        except: pass

st.sidebar.markdown("---")
st.sidebar.header("3. TICA")
if 'exec' in uploaded_sheets and start_d:
    f_exec, s_exec = uploaded_sheets['exec']
    try:
        df_temp = process_file_sheet(f_exec, s_exec, "Agente", start_d, end_d)
        if not df_temp.empty:
            unique_names = sorted(df_temp['Nombre'].unique().tolist())
            agents_no_tica = st.sidebar.multiselect("Agentes SIN TICA", unique_names)
    except: pass

# --- MOTOR LÓGICO V53 (Regla Coordinación) ---
def logic_engine(df, no_tica_list):
    rows = []
    
    agent_class = {}
    df_agentes = df[df['Rol'] == 'Agente']
    for name, group in df_agentes.groupby('Nombre'):
        am = 0; pm = 0
        for _, r in group.iterrows():
            _, start_h = parse_shift_time(r['Turno_Raw'])
            if start_h is not None:
                if start_h < 12: am += 1
                else: pm += 1
        agent_class[name] = "Nocturno" if pm > am else "Diurno"

    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        sub_group = "General"
        role_rank = 99
        if r['Rol'] == 'Agente':
            cat = agent_class.get(r['Nombre'], "Diurno")
            sub_group = cat
            role_rank = 10 if cat == "Diurno" else 11
        elif r['Rol'] == 'Coordinador': role_rank = 20
        elif r['Rol'] == 'Anfitrion': role_rank = 30
        elif r['Rol'] == 'Supervisor': role_rank = 40
            
        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '', 'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': -1, 'Base_Diaria': ''})
        else:
            for h in hours:
                current_date = r['Fecha']
                if start_h >= 18 and h < 12: current_date = current_date + timedelta(days=1)
                rows.append({
                    'Nombre': r['Nombre'], 'Rol': r['Rol'], 'Turno_Raw': r['Turno_Raw'],
                    'Fecha': current_date, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                    'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': start_h, 'Base_Diaria': '?'
                })
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    if 'incidencias' in st.session_state and st.session_state.incidencias:
        df_h = apply_incidents(df_h, st.session_state.incidencias)

    main_counters_aire = ["T1 AIRE", "T2 AIRE"]
    main_counters_tierra = ["T1 TIERRA", "T2 TIERRA"]
    daily_assignments = {} 
    anf_base_assignments = {}
    last_ag_counter = {}
    last_anf_zone = {}
    sorted_dates = sorted(df_h['Fecha'].unique())
    
    for d in sorted_dates:
        # AGENTES
        df_d_ag = df_h[(df_h['Fecha'] == d) & (df_h['Rol'] == 'Agente') & (df_h['Hora'] != -1)]
        active_ag_names = df_d_ag['Nombre'].unique()
        ag_start_times = {}
        for nm in active_ag_names:
            day_data = df_d_ag[df_d_ag['Nombre'] == nm]
            ag_start_times[nm] = day_data.iloc[0]['Start_H'] if not day_data.empty else 99

        continuers_ag = []; starters_ag = []
        for name in active_ag_names:
            works_midnight = not df_d_ag[(df_d_ag['Nombre'] == name) & (df_d_ag['Hora'] == 0)].empty
            if works_midnight and name in last_ag_counter: continuers_ag.append(name)
            else: starters_ag.append((name, ag_start_times[name]))
        
        load_ag = {c: 0 for c in main_counters_aire + main_counters_tierra}
        
        for name in continuers_ag:
            prev = last_ag_counter[name]; daily_assignments[(name, d)] = prev; load_ag[prev] += 1
            
        # Priority AM Assignment
        starters_ag.sort(key=lambda x: x[1])
        am_fixed_slots = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
        am_slot_idx = 0
        
        for name, s_h in starters_ag:
            has_tica = name not in no_tica_list
            if s_h < 12 and am_slot_idx < 4:
                chosen = am_fixed_slots[am_slot_idx]; am_slot_idx += 1
            else:
                if has_tica: options = main_counters_aire + main_counters_tierra
                else: options = main_counters_tierra
                chosen = sorted(options, key=lambda c: load_ag[c])[0]
            daily_assignments[(name, d)] = chosen; load_ag[chosen] += 1
            
        for name in active_ag_names: last_ag_counter[name] = daily_assignments[(name, d)]
            
        # ANFITRIONES
        df_d_anf = df_h[(df_h['Fecha'] == d) & (df_h['Rol'] == 'Anfitrion') & (df_h['Hora'] != -1)]
        active_anf = df_d_anf['Nombre'].unique()
        zone_load = {'Zona Int': 0, 'Zona Nac': 0}
        continuers = []; starters = []
        for name in active_anf:
            works_midnight = not df_d_anf[(df_d_anf['Nombre'] == name) & (df_d_anf['Hora'] == 0)].empty
            if works_midnight and name in last_anf_zone: continuers.append(name)
            else: starters.append(name)
        for name in continuers:
            z = last_anf_zone[name]; anf_base_assignments[(name, d)] = z; zone_load[z] += 1
        starters.sort()
        for name in starters:
            z = 'Zona Int' if zone_load['Zona Int'] <= zone_load['Zona Nac'] else 'Zona Nac'
            anf_base_assignments[(name, d)] = z; zone_load[z] += 1
        for name in active_anf: last_anf_zone[name] = anf_base_assignments[(name, d)]

    # Coordinadores (Default)
    coords_active = df_h[(df_h['Rol'] == 'Coordinador') & (df_h['Hora'] != -1)]
    for idx, row in coords_active.iterrows():
        st_h = row['Start_H']; h = row['Hora']; nm = row['Nombre']; is_odd = hash(nm) % 2 != 0
        if st_h == 10:
            if h == 10: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
            elif h in [14, 15]:
                if (h == 14 and is_odd) or (h == 15 and not is_odd): df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
        elif st_h == 5:
            if h in [11, 12, 13]:
                if h == 12: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
        elif st_h == 21:
            if h in [5, 6, 7]:
                if h == 6: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'

    hhee_counters = []; hhee_coord = []; hhee_anf = []

    # 5. Proceso Hora
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        idx_ag = g[g['Rol']=='Agente'].index.tolist()
        idx_co = g[g['Rol']=='Coordinador'].index.tolist()
        idx_an = g[g['Rol']=='Anfitrion'].index.tolist()
        idx_su = g[g['Rol']=='Supervisor'].index.tolist()
        
        # Asignar bases
        for idx in idx_ag:
            df_h.at[idx, 'Counter'] = daily_assignments.get((df_h.at[idx, 'Nombre'], d), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_an:
            df_h.at[idx, 'Counter'] = anf_base_assignments.get((df_h.at[idx, 'Nombre'], d), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_su:
            df_h.at[idx, 'Counter'] = "General"; df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_co:
            if df_h.at[idx, 'Tarea'] not in ['2', 'C']:
                df_h.at[idx, 'Counter'] = "General"
                df_h.at[idx, 'Tarea'] = '1'

        # Breaks
        def apply_break(indices, start_range, slots):
            cands = [i for i in indices if start_range[0] <= df_h.at[i, 'Start_H'] <= start_range[1]]
            for i, idx in enumerate(cands):
                if h == slots[i % len(slots)]: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
        
        apply_break(idx_ag, (0, 11), [12, 13, 14, 15]); apply_break(idx_ag, (12, 23), [2, 3])
        apply_break(idx_an, (0, 11), [13, 14, 15]); apply_break(idx_an, (12, 23), [2, 3, 4])

        # Quiebres Agentes
        active_counts = {c: 0 for c in main_counters_aire + main_counters_tierra}
        donors = [] 
        for idx in idx_ag:
            if df_h.at[idx, 'Tarea'] == '1':
                c = df_h.at[idx, 'Counter']
                if c in active_counts: active_counts[c] += 1; donors.append(idx)
        
        empty = [c for c, count in active_counts.items() if count == 0]
        
        # Función auxiliar para encontrar Coordinador Disponible (V53)
        def find_coord_cover():
            # Regla: 
            # 1. Coord en Tarea 2 -> Libre para ir
            # 2. Coord en Tarea 1 -> Solo si hay >1 Coord en Tarea 1
            
            c_t2 = [i for i in idx_co if df_h.at[i, 'Tarea'] == '2']
            c_t1 = [i for i in idx_co if df_h.at[i, 'Tarea'] == '1']
            
            if c_t2: return c_t2[0] # Preferencia T2
            if len(c_t1) > 1: return c_t1[0] # Si sobra uno en supervisión
            return None

        for target_cnt in empty:
            covered = False
            possible = []
            for d_idx in donors:
                orig = df_h.at[d_idx, 'Counter']
                if orig == target_cnt: continue
                if active_counts.get(orig, 0) > 1:
                    d_nm = df_h.at[d_idx, 'Nombre']
                    if "AIRE" in target_cnt and (d_nm in no_tica_list): continue
                    possible.append(d_idx)
            if possible:
                best = possible[0]
                df_h.at[best, 'Tarea'] = f"3: Cubrir {target_cnt}"; df_h.at[best, 'Counter'] = target_cnt
                active_counts[daily_assignments[(df_h.at[best, 'Nombre'], d)]] -= 1
                active_counts[target_cnt] += 1; donors.remove(best); covered = True
            
            if not covered:
                cand = find_coord_cover() # NEW LOGIC
                if cand is not None: 
                    df_h.at[cand, 'Tarea'] = f"4: Cubrir {target_cnt}"; df_h.at[cand, 'Counter'] = target_cnt; covered = True
            
            if not covered and idx_an:
                avail_a = [i for i in idx_an if df_h.at[i, 'Tarea'] != 'C']
                if len(avail_a) > 2:
                    ix = avail_a[0]; df_h.at[ix, 'Tarea'] = f"4: Cubrir {target_cnt}"; df_h.at[ix, 'Counter'] = target_cnt; covered = True
            
            if not covered: hhee_counters.append({'Fecha': d, 'Hora': h, 'Counter': target_cnt})

        # Cobertura Anfitriones
        anf_avail = [idx for idx in idx_an if df_h.at[idx, 'Tarea'] in ['?', '1']]
        zones_assigned = {'Nac': [], 'Int': []}
        for i, idx in enumerate(anf_avail):
            z = df_h.at[idx, 'Counter']
            z_key = 'Nac' if 'Nac' in z else 'Int'
            zones_assigned[z_key].append(idx)
            
        count_nac = len(zones_assigned['Nac'])
        count_int = len(zones_assigned['Int'])
        target_zone_name = None
        if count_nac == 0 and count_int > 0: target_zone_name = "Nac"
        elif count_int == 0 and count_nac > 0: target_zone_name = "Int"
        
        if target_zone_name:
            filled = False
            other = 'Int' if target_zone_name == 'Nac' else 'Nac'
            if len(zones_assigned[other]) > 1:
                cand = zones_assigned[other][0]
                df_h.at[cand, 'Tarea'] = f"3: Cubrir {target_zone_name}"; df_h.at[cand, 'Counter'] = f"Zona {target_zone_name}"; filled = True
            
            if not filled:
                cand = find_coord_cover() # NEW LOGIC
                if cand is not None:
                    df_h.at[cand, 'Tarea'] = f"Cubrir {target_zone_name}"; filled = True
            
            if not filled:
                for d_idx in donors:
                    orig = daily_assignments.get((df_h.at[d_idx, 'Nombre'], d))
                    if active_counts.get(orig, 0) > 1:
                        df_h.at[d_idx, 'Tarea'] = f"Cubrir {target_zone_name}"; active_counts[orig] -= 1; donors.remove(d_idx); filled = True; break
            
            if not filled: hhee_anf.append({'Fecha': d, 'Hora': h, 'Counter': f"Zona {target_zone_name}"})

    final_bases = {}
    grouped = df_h[df_h['Hora'] != -1].groupby(['Nombre', 'Fecha'])
    for (name, date), group in grouped:
        valid_counts = group[~group['Counter'].isin(['Casino', 'Oficina', 'General', ''])]['Counter'].value_counts()
        real_base = valid_counts.index[0] if not valid_counts.empty else group['Counter'].mode()[0]
        final_bases[(name, date)] = real_base
        for idx in group.index:
            task = str(df_h.at[idx, 'Tarea']); cnt = str(df_h.at[idx, 'Counter'])
            if "Cubrir" in task and (real_base in task or real_base in cnt): df_h.at[idx, 'Tarea'] = '1'

    for idx, row in df_h.iterrows():
        if row['Hora'] != -1: 
            if row['Rol'] == 'Coordinador': df_h.at[idx, 'Base_Diaria'] = "General"
            else: df_h.at[idx, 'Base_Diaria'] = final_bases.get((row['Nombre'], row['Fecha']), "")

    unique_dates = sorted(df_h['Fecha'].unique())
    hhee_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_counters)
    
    for cnt in ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]:
        base_row = {'Nombre': f"HHEE Agente - {cnt}", 'Rol': 'HHEE', 'Sub_Group': 'Counter', 'Role_Rank': 900, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': cnt}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h, cnt) in hhee_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)
            
    # HHEE Coord (Solo si se agregó externamente)
    hhee_c_set = set((x['Fecha'], x['Hora']) for x in hhee_coord)
    if hhee_c_set:
        base_row = {'Nombre': "HHEE Coordinación", 'Rol': 'HHEE', 'Sub_Group': 'Supervisión', 'Role_Rank': 901, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': 'General'}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h) in hhee_c_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)
        
    hhee_a_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_anf)
    for z in ['Zona Int', 'Zona Nac']:
        base_row = {'Nombre': f"HHEE Anfitriones - {z.upper()}", 'Rol': 'HHEE', 'Sub_Group': 'Losa', 'Role_Rank': 902, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': z}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h, z) in hhee_a_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)

    return df_h

# --- EXCEL GENERATOR (V53) ---
def make_excel(df, start_d, end_d):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    
    f_cabify = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#7145D6', 'font_color': 'white', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'text_wrap': True})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#F3F3F3', 'align': 'center'})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#EFEFEF', 'align': 'left', 'indent': 1})
    f_alert = wb.add_format({'bg_color': '#EA9999', 'font_color': '#980000', 'bold': True, 'border': 1, 'align': 'center'})
    f_header_count = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#B4A7D6', 'align': 'center'})
    f_header_hhee = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#E06666', 'align': 'center', 'font_color': 'white'})
    
    f_nac = wb.add_format({'bg_color': '#FCE4D6', 'border': 1, 'align': 'center'})
    f_int = wb.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
    f_sep_col = wb.add_format({'bg_color': '#E0E0E0', 'border': 1, 'align': 'center', 'font_size': 9})
    f_hhee_ok = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center', 'bold': True})
    f_hhee_bad = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'bold': True})
    
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        '3': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center', 'font_size': 8, 'text_wrap': True}),
        '4': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'bold': True, 'font_size': 8, 'text_wrap': True}),
        'C': wb.add_format({'bg_color': '#C6E0B4', 'border': 1, 'align': 'center'}),
        'HHEE': wb.add_format({'bg_color': '#7030A0', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'}),
        'Nac': f_nac,
        'Int': f_int
    }
    
    all_dates = sorted(df['Fecha'].unique())
    dates = [d for d in all_dates if start_d <= d.date() <= end_d]
    df_staff = df[df['Rol'] != 'HHEE'].drop_duplicates(subset=['Nombre', 'Fecha'])
    df_hhee = df[df['Rol'] == 'HHEE'].drop_duplicates(subset=['Nombre', 'Fecha'])
    
    df_staff_sorted = df_staff[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    df_hhee_sorted = df_hhee[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Nombre'])
    
    ws_data = wb.add_worksheet("Datos_Validacion")
    unique_roles = sorted([str(x) for x in df['Rol'].unique() if x != 'HHEE'])
    unique_names = sorted([str(x) for x in df['Nombre'].unique() if 'HHEE' not in str(x)])
    ws_data.write_column(1, 0, unique_roles)
    ws_data.write_column(1, 1, unique_names)
    ws_data.hide()
    
    name_range = f"Datos_Validacion!$B$2:$B${len(unique_names)+1}"
    role_range = f"Datos_Validacion!$A$2:$A${len(unique_roles)+1}"

    ws_teorico = wb.add_worksheet("Plan_Teorico")
    ws_teorico.write(0, 0, "ID") 
    
    teorico_row = 1
    for _, p in df_staff_sorted.iterrows():
        n = p['Nombre']
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            key = f"{n}_{d_iso}"
            ws_teorico.write(teorico_row, 0, key)
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            for h in range(24):
                try:
                    val = subset[subset['Hora'] == h]
                    if not val.empty: task = str(val.iloc[0]['Tarea'])
                    else: task = ""
                    ws_teorico.write_string(teorico_row, 1+h, task)
                except: ws_teorico.write_string(teorico_row, 1+h, "")
            teorico_row += 1
            
    for _, p in df_hhee_sorted.iterrows():
        n = p['Nombre']
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            key = f"{n}_{d_iso}"
            ws_teorico.write(teorico_row, 0, key)
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            for h in range(24):
                try:
                    val = subset[subset['Hora'] == h]
                    if not val.empty: task = str(val.iloc[0]['Tarea'])
                    else: task = ""
                    ws_teorico.write_string(teorico_row, 1+h, task)
                except: ws_teorico.write_string(teorico_row, 1+h, "")
            teorico_row += 1
    ws_teorico.hide()

    ws_bit = wb.add_worksheet("Bitacora_Incidencias")
    headers_bit = ["Tipo Colaborador", "Nombre Colaborador", "Fecha (YYYY-MM-DD)", "Tipo Incidencia", "Hora Inicio (0-23)", "Hora Fin (0-23)"]
    for i, h in enumerate(headers_bit): ws_bit.write(0, i, h, f_cabify)
    ws_bit.data_validation('A2:A1000', {'validate': 'list', 'source': role_range})
    ws_bit.data_validation('B2:B1000', {'validate': 'list', 'source': name_range})
    ws_bit.data_validation('D2:D1000', {'validate': 'list', 'source': ['Inasistencia', 'Atraso', 'Salida Anticipada']})
    ws_bit.write(0, 7, "GUÍA:", f_cabify)
    ws_bit.write(1, 7, "1. INASISTENCIA: Marque el tipo y NO se preocupe por las horas.")

    ws_real = wb.add_worksheet("Plan_Operativo")
    
    ws_real.write(9, 0, "Colaborador", f_cabify)
    ws_real.write(9, 1, "Rol", f_cabify)
    ws_real.freeze_panes(10, 2)
    
    col = 2
    d_map = {}
    
    ws_real.write(2, 0, "DOTACIÓN AGENTES", f_header_count)
    ws_real.write(3, 0, "DOTACIÓN COORDINADORES", f_header_count)
    ws_real.write(4, 0, "DOTACIÓN ANFITRIONES", f_header_count)
    
    ws_real.write(5, 0, "HHEE AGENTES", f_header_hhee)
    ws_real.write(6, 0, "HHEE COORDINACIÓN", f_header_hhee)
    ws_real.write(7, 0, "HHEE ANFITRIONES", f_header_hhee)
    ws_real.write(8, 0, "TOTAL HHEE", f_header_hhee)
    
    for d in dates:
        d_str = pd.to_datetime(d).strftime("%d-%b")
        d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
        ws_real.merge_range(0, col, 0, col+25, d_str, f_date)
        ws_real.write(9, col, "Turno", f_cabify)
        ws_real.write(9, col+1, "Lugar", f_cabify)
        d_map[d_iso] = col
        
        for h in range(24):
            ws_real.write(9, col+2+h, h, f_cabify)
            col_idx = col+2+h
            col_let = xlsxwriter.utility.xl_col_to_name(col_idx)
            
            f_ag = f'=COUNTIFS($B$11:$B$1000,"Agente",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"<>Libre",{col_let}11:{col_let}1000,"<>*Ausente*", {col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(2, col_idx, f_ag, f_header_count)
            f_co = f'=COUNTIFS($B$11:$B$1000,"Coordinador",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(3, col_idx, f_co, f_header_count)
            f_an = f'=COUNTIFS($B$11:$B$1000,"Anfitrion",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(4, col_idx, f_an, f_header_count)
            
            f_h_ag = f'=COUNTIFS($A$11:$A$1000,"HHEE Agente*",{col_let}11:{col_let}1000,"HHEE")'
            ws_real.write_formula(5, col_idx, f_h_ag, f_header_hhee)
            f_h_co = f'=COUNTIFS($A$11:$A$1000,"HHEE Coordinación",{col_let}11:{col_let}1000,"HHEE")'
            ws_real.write_formula(6, col_idx, f_h_co, f_header_hhee)
            f_h_an = f'=COUNTIFS($A$11:$A$1000,"HHEE Anfitriones*",{col_let}11:{col_let}1000,"HHEE")'
            ws_real.write_formula(7, col_idx, f_h_an, f_header_hhee)
            
            col_l = xlsxwriter.utility.xl_col_to_name(col_idx)
            f_total = f'=SUM({col_l}6:{col_l}8)'
            ws_real.write_formula(8, col_idx, f_total, f_header_hhee)

        col += 26
    
    hhee_range = f"C6:{xlsxwriter.utility.xl_col_to_name(col-1)}9"
    ws_real.conditional_format(hhee_range, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': f_hhee_ok})
    ws_real.conditional_format(hhee_range, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': f_hhee_bad})

    row = 10
    curr_group = ""
    
    for _, p in df_staff_sorted.iterrows():
        n = p['Nombre']
        r = p['Rol']
        grp = p['Sub_Group']
        
        grp_label = f"{r.upper()}"
        if r == "Agente": grp_label += f" - {grp}"
        
        if grp_label != curr_group:
            ws_real.merge_range(row, 0, row, col-1, grp_label, f_group)
            row += 1
            curr_group = grp_label
            
        ws_real.write(row, 0, n, f_base)
        ws_real.write(row, 1, r, f_base)
        
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            c_start = d_map[d_iso]
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            
            if subset.empty:
                ws_real.write(row, c_start, "-", f_base)
                ws_real.write(row, c_start+1, "", f_sep_col)
                for h in range(24): 
                    key = f"{n}_{d_iso}"
                    col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                    formula = f'T(INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0))) & ""'
                    ws_real.write_formula(row, c_start+2+h, formula, f_base)
            else:
                t_raw = subset.iloc[0]['Turno_Raw']
                if "libre" in str(t_raw).lower(): t_raw = ""
                try: lugar = subset.iloc[0]['Base_Diaria']
                except: lugar = "?"
                ws_real.write(row, c_start, str(t_raw), f_sep_col)
                fmt_lugar = f_sep_col
                if r == "Anfitrion":
                    if "Nac" in str(lugar): fmt_lugar = f_nac
                    elif "Int" in str(lugar): fmt_lugar = f_int
                ws_real.write(row, c_start+1, str(lugar), fmt_lugar)
            
                key = f"{n}_{d_iso}"
                for h in range(24):
                    col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                    formula = (
                        f'=IF(AND(COUNTIFS(Bitacora_Incidencias!$B:$B,"{n}",Bitacora_Incidencias!$C:$C,"{d_iso}",'
                        f'Bitacora_Incidencias!$D:$D,"Inasistencia")>0, INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0))<>""),"FALTA",'
                        f'IF(COUNTIFS(Bitacora_Incidencias!$B:$B,"{n}",Bitacora_Incidencias!$C:$C,"{d_iso}",'
                        f'Bitacora_Incidencias!$E:$E,"<={h}",Bitacora_Incidencias!$F:$F,">={h}")>0,"INCIDENCIA",'
                        f'INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0)) & ""))'
                    )
                    ws_real.write_formula(row, c_start+2+h, formula, f_base)
        row += 1
    
    ws_real.merge_range(row, 0, row, col-1, "REQUERIMIENTOS HHEE (AUTOMÁTICO)", f_group)
    row += 1
    
    for _, p in df_hhee_sorted.iterrows():
        n = p['Nombre']
        r = p['Rol']
        ws_real.write(row, 0, n, f_base)
        ws_real.write(row, 1, r, f_base)
        
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            c_start = d_map[d_iso]
            ws_real.write(row, c_start, "Demanda", f_sep_col)
            try: 
                lugar = df[(df['Nombre']==n)&(df['Fecha']==d)].iloc[0]['Base_Diaria']
            except: lugar = "?"
            ws_real.write(row, c_start+1, str(lugar), f_sep_col)
            
            key = f"{n}_{d_iso}"
            for h in range(24):
                col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                formula = f'INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0)) & ""'
                ws_real.write_formula(row, c_start+2+h, formula, f_base)
        row += 1

    end_col_let = xlsxwriter.utility.xl_col_to_name(col-1)
    data_range = f"C11:{end_col_let}{row}"
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"FALTA"', 'format': f_alert})
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"INCIDENCIA"', 'format': f_alert})
    
    st_map_cond = {
        '3:': st_map['3'], '4:': st_map['4'], 'Cubrir': st_map['4'], 
        'C': st_map['C'], '2': st_map['2'], 'HHEE': st_map['HHEE'],
        'Nac': st_map['Nac'], 'Int': st_map['Int']
    }
    for k, fmt in st_map_cond.items():
        criteria = 'begins with' if len(k) > 2 else 'equal to'
        val = k if len(k) > 2 else f'"{k}"'
        ws_real.conditional_format(data_range, {'type': 'text' if len(k)>2 else 'cell', 'criteria': criteria, 'value': val, 'format': fmt})

    ws_res = wb.add_worksheet("Resumen_Estadistico")
    f_bold = wb.add_format({'bold': True})
    df_res = df[df['Fecha'].apply(lambda x: start_d <= x.date() <= end_d)].copy()
    
    ws_res.write(0, 0, "TOTAL HORAS POR COUNTER (EJECUTIVOS)", f_bold)
    ag_work = df_res[(df_res['Rol'] == 'Agente') & (df_res['Tarea'].astype(str).str.contains(r'^(1|3|4|Cubrir|Apoyo)', regex=True))]
    if not ag_work.empty:
        ag_pivot = ag_work.groupby(['Nombre', 'Counter']).size().unstack(fill_value=0)
        r_idx = 2
        ws_res.write(r_idx, 0, "Nombre", f_bold)
        for i, c in enumerate(ag_pivot.columns): ws_res.write(r_idx, i+1, c, f_bold)
        r_idx += 1
        for name, row_data in ag_pivot.iterrows():
            ws_res.write(r_idx, 0, name)
            for i, val in enumerate(row_data): ws_res.write(r_idx, i+1, val)
            r_idx += 1
    else: r_idx = 5

    r_idx += 2
    ws_res.write(r_idx, 0, "ESTADÍSTICAS HHEE (HORAS TOTALES)", f_bold)
    r_idx += 2
    hhee_active = df_res[(df_res['Rol'] == 'HHEE') & (df_res['Tarea'] == 'HHEE')]
    
    ws_res.write(r_idx, 0, "Total HHEE Requeridas:", f_bold)
    ws_res.write(r_idx, 1, len(hhee_active))
    r_idx += 2
    
    ws_res.write(r_idx, 0, "Por Franja Horaria", f_bold)
    ws_res.write(r_idx, 1, "Cantidad", f_bold)
    by_hour = hhee_active['Hora'].value_counts().sort_index()
    for h, count in by_hour.items():
        r_idx += 1
        ws_res.write(r_idx, 0, f"{h}:00 - {h+1}:00")
        ws_res.write(r_idx, 1, count)
        
    wb.close()
    return out

st.sidebar.markdown("---")
if st.sidebar.button("🚀 Generar Planificación V53"):
    if not uploaded_sheets: st.error("Carga archivos.")
    elif not (start_d and end_d): st.error("Define fechas.")
    else:
        with st.spinner("Procesando Escenarios..."):
            dfs = []
            for role, (key) in [("Agente","exec"),("Coordinador","coord"),("Anfitrion","host"),("Supervisor","sup")]:
                if key in uploaded_sheets:
                    f, s = uploaded_sheets[key]
                    dfs.append(process_file_sheet(f, s, role, start_d, end_d))
            full = pd.concat(dfs)
            
            if full.empty: st.error("Sin datos.")
            else:
                final = logic_engine(full, agents_no_tica)
                st.success("¡Listo! Descarga la Suite Operativa V53.")
                st.download_button("📥 Descargar Suite (V53)", make_excel(final, start_d, end_d), f"Planificacion_Operativa.xlsx")
