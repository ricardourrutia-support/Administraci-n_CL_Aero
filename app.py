import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Gestor de Turnos: V41 (Excel Inteligente con Bitácora)")
st.markdown("""
**Sistema Operativo V41:**
* **Motor Lógico:** Basado en V37 (Conteo de horas reales y cobertura correcta).
* **Excel Dinámico:** * Hoja `Bitacora`: Para escribir incidencias manualmente.
    * Hoja `Operacion_Real`: Se actualiza sola mediante fórmulas de Excel al escribir en la bitácora.
    * **Contadores Vivos:** Los encabezados de dotación en el Excel cambian si marcas una falta.
""")

# --- INICIALIZACIÓN ---
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
        # Carga ampliada para turnos noche
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
                if pd.isna(shift_val): shift_val = "Libre"
                extracted_data.append({
                    'Nombre': clean_name, 'Rol': role, 'Fecha': date_obj, 'Turno_Raw': shift_val
                })
    except Exception as e: st.error(f"Error en {role}: {e}")
    return pd.DataFrame(extracted_data)

# --- UI LATERAL ---
st.sidebar.header("1. Periodo")
today = datetime.now()
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

# --- MOTOR LÓGICO V37 (Base Solida) ---
def logic_engine(df, no_tica_list):
    rows = []
    
    # 1. Clasificación
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

    # 2. Expandir
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
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': -1, 'Base_Diaria': '-'})
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
    
    # 3. Asignación Base (Memoria V37)
    main_counters_aire = ["T1 AIRE", "T2 AIRE"]
    main_counters_tierra = ["T1 TIERRA", "T2 TIERRA"]
    daily_assignments = {} 
    anf_base_assignments = {}
    last_ag_counter = {}
    last_anf_zone = {}
    sorted_dates = sorted(df_h['Fecha'].unique())
    
    for d in sorted_dates:
        # Agentes
        df_d_ag = df_h[(df_h['Fecha'] == d) & (df_h['Rol'] == 'Agente') & (df_h['Hora'] != -1)]
        active_ag = df_d_ag['Nombre'].unique()
        load_ag = {c: 0 for c in main_counters_aire + main_counters_tierra}
        continuers_ag = []; starters_ag = []
        for name in active_ag:
            works_midnight = not df_d_ag[(df_d_ag['Nombre'] == name) & (df_d_ag['Hora'] == 0)].empty
            if works_midnight and name in last_ag_counter: continuers_ag.append(name)
            else: starters_ag.append(name)
        for name in continuers_ag:
            prev = last_ag_counter[name]; daily_assignments[(name, d)] = prev; load_ag[prev] += 1
        starters_ag.sort()
        for name in starters_ag:
            has_tica = name not in no_tica_list
            chosen = sorted(main_counters_tierra if not has_tica else main_counters_aire + main_counters_tierra, key=lambda c: load_ag[c])[0]
            daily_assignments[(name, d)] = chosen; load_ag[chosen] += 1
        for name in active_ag: last_ag_counter[name] = daily_assignments[(name, d)]
            
        # Anfitriones
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

    # 4. Coordinadores OFF
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

    # HHEE Lists
    hhee_counters = []; hhee_coord = []; hhee_anf = []

    # 5. Proceso Hora
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        idx_ag = g[g['Rol']=='Agente'].index.tolist()
        idx_co = g[g['Rol']=='Coordinador'].index.tolist()
        idx_an = g[g['Rol']=='Anfitrion'].index.tolist()
        idx_su = g[g['Rol']=='Supervisor'].index.tolist()
        
        for idx in idx_ag:
            df_h.at[idx, 'Counter'] = daily_assignments.get((df_h.at[idx, 'Nombre'], d), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_an:
            df_h.at[idx, 'Counter'] = anf_base_assignments.get((df_h.at[idx, 'Nombre'], d), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_su:
            df_h.at[idx, 'Counter'] = "General"; df_h.at[idx, 'Tarea'] = '1'

        def apply_break(indices, start_range, slots):
            cands = [i for i in indices if start_range[0] <= df_h.at[i, 'Start_H'] <= start_range[1]]
            for i, idx in enumerate(cands):
                if h == slots[i % len(slots)]: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
        
        apply_break(idx_ag, (0, 11), [13, 14]); apply_break(idx_ag, (12, 23), [2, 3])
        apply_break(idx_an, (0, 11), [13, 14, 15]); apply_break(idx_an, (12, 23), [2, 3, 4])

        # Quiebres
        active_counts = {c: 0 for c in main_counters_aire + main_counters_tierra}
        donors = [] 
        for idx in idx_ag:
            if df_h.at[idx, 'Tarea'] == '1':
                c = df_h.at[idx, 'Counter']
                if c in active_counts: active_counts[c] += 1; donors.append(idx)
        
        empty = [c for c, count in active_counts.items() if count == 0]
        
        # Cobertura Agentes
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
                avail_c = [i for i in idx_co if df_h.at[i, 'Tarea'] != 'C']
                cand = None
                # Logic simplificada Coord
                if avail_c: cand = avail_c[0]
                if cand is not None: df_h.at[cand, 'Tarea'] = f"4: Cubrir {target_cnt}"; df_h.at[cand, 'Counter'] = target_cnt; covered = True
            
            if not covered and idx_an:
                avail_a = [i for i in idx_an if df_h.at[i, 'Tarea'] != 'C']
                if len(avail_a) > 2:
                    ix = avail_a[0]; df_h.at[ix, 'Tarea'] = f"4: Cubrir {target_cnt}"; df_h.at[ix, 'Counter'] = target_cnt; covered = True
            
            if not covered: hhee_counters.append({'Fecha': d, 'Hora': h, 'Counter': target_cnt})

        # Cobertura Anfitriones (Normalizada V34)
        anf_avail = [idx for idx in idx_an if df_h.at[idx, 'Tarea'] in ['?', '1']]
        zones_assigned = {'Nac': [], 'Int': []}
        for i, idx in enumerate(anf_avail):
            z = df_h.at[idx, 'Counter']
            z_key = 'Nac' if 'Nac' in z else 'Int'
            zones_assigned[z_key].append(idx)
            
        missing_nac = len(zones_assigned['Nac']) == 0
        missing_int = len(zones_assigned['Int']) == 0
        target_zone_name = "Nac" if missing_nac else ("Int" if missing_int else None)
        
        if target_zone_name:
            filled = False
            other = 'Int' if target_zone_name == 'Nac' else 'Nac'
            if len(zones_assigned[other]) > 1:
                cand = zones_assigned[other][0]
                df_h.at[cand, 'Tarea'] = f"3: Cubrir {target_zone_name}"; df_h.at[cand, 'Counter'] = f"Zona {target_zone_name}"; filled = True
            
            if not filled: # Coord
                avail_c = [i for i in idx_co if df_h.at[i, 'Tarea'] != 'C']
                if avail_c:
                    cand = avail_c[0]
                    df_h.at[cand, 'Tarea'] = f"Cubrir {target_zone_name}"; filled = True
            
            if not filled:
                for d_idx in donors:
                    orig = daily_assignments.get((df_h.at[d_idx, 'Nombre'], d))
                    if active_counts.get(orig, 0) > 1:
                        df_h.at[d_idx, 'Tarea'] = f"Cubrir {target_zone_name}"; active_counts[orig] -= 1; donors.remove(d_idx); filled = True; break
            
            if not filled: hhee_anf.append({'Fecha': d, 'Hora': h, 'Counter': f"Zona {target_zone_name}"})

        # Finales Coord/Sup
        coords_t1 = [i for i in idx_co if df_h.at[i, 'Tarea'] == '1']
        if not coords_t1: hhee_coord.append({'Fecha': d, 'Hora': h, 'Counter': 'Supervisión'})
        for idx in idx_co:
            if df_h.at[idx, 'Tarea'] == '1': df_h.at[idx, 'Counter'] = 'General'
        for idx in idx_su:
            df_h.at[idx, 'Counter'] = 'General'; df_h.at[idx, 'Tarea'] = '1'

    # Normalización Bases
    final_bases = {}
    grouped = df_h[df_h['Hora'] != -1].groupby(['Nombre', 'Fecha'])
    for (name, date), group in grouped:
        valid_counts = group[~group['Counter'].isin(['Casino', 'Oficina', 'General'])]['Counter'].value_counts()
        real_base = valid_counts.index[0] if not valid_counts.empty else group['Counter'].mode()[0]
        final_bases[(name, date)] = real_base
        for idx in group.index:
            task = str(df_h.at[idx, 'Tarea']); cnt = str(df_h.at[idx, 'Counter'])
            if "Cubrir" in task and (real_base in task or real_base in cnt): df_h.at[idx, 'Tarea'] = '1'

    for idx, row in df_h.iterrows():
        if row['Hora'] != -1: df_h.at[idx, 'Base_Diaria'] = final_bases.get((row['Nombre'], row['Fecha']), "-")

    # HHEE Rows
    unique_dates = sorted(df_h['Fecha'].unique())
    hhee_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_counters)
    for cnt in ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]:
        base_row = {'Nombre': f"HHEE {cnt}", 'Rol': 'HHEE', 'Sub_Group': 'Counter', 'Role_Rank': 900, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': cnt}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h, cnt) in hhee_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)
            
    # HHEE Coord/Anf
    hhee_c_set = set((x['Fecha'], x['Hora']) for x in hhee_coord)
    base_row = {'Nombre': "HHEE COORDINACIÓN", 'Rol': 'HHEE', 'Sub_Group': 'Supervisión', 'Role_Rank': 901, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': 'General'}
    for d in unique_dates:
        for h in range(24):
            task_val = "HHEE" if (d, h) in hhee_c_set else ""
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
        df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)
        
    hhee_a_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_anf)
    for z in ['Zona Int', 'Zona Nac']:
        base_row = {'Nombre': f"HHEE {z.upper()}", 'Rol': 'HHEE', 'Sub_Group': 'Losa', 'Role_Rank': 902, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': z}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h, z) in hhee_a_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)

    return df_h

# --- EXCEL GENERATOR CON FÓRMULAS V41 ---
def make_excel(df, start_d, end_d):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    
    # 1. PLAN BASE (Oculto/Referencia)
    ws_base = wb.add_worksheet("Plan_Base")
    # Estilos básicos
    f_date = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'align': 'center'})
    
    # Filtrar fechas
    all_dates = sorted(df['Fecha'].unique())
    dates = [d for d in all_dates if start_d <= d.date() <= end_d]
    
    # Escribir Base
    # Estructura: Nombre | Fecha | Hora 0..23
    # Vamos a aplanar para que sea fácil de referenciar: Una fila por persona-día?
    # No, mantengamos estructura sábana.
    
    # Encabezados
    col = 0
    ws_base.write(0, col, "ID_Unico") # Clave para buscar: Nombre_Fecha
    col += 1
    for h in range(24): ws_base.write(0, col+h, f"H{h}")
    
    # Datos Plan Base
    df_sorted = df[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    
    # Mapa para saber en qué fila queda cada persona en la hoja Operacion
    person_row_map = {} 
    current_row = 1
    
    for _, p in df_sorted.iterrows():
        n = p['Nombre']
        for d in dates:
            d_str = pd.to_datetime(d).strftime("%Y-%m-%d")
            key = f"{n}_{d_str}"
            
            # Datos de este dia
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            
            # Escribir ID
            ws_base.write(current_row, 0, key)
            
            # Escribir Tareas
            for h in range(24):
                try:
                    val = subset[subset['Hora'] == h]
                    if not val.empty:
                        task = str(val.iloc[0]['Tarea'])
                    else:
                        task = ""
                    ws_base.write(current_row, 1+h, task)
                except: pass
            current_row += 1
    
    # 2. BITÁCORA
    ws_bit = wb.add_worksheet("Bitacora")
    ws_bit.write(0, 0, "Nombre Colaborador")
    ws_bit.write(0, 1, "Fecha (AAAA-MM-DD)")
    ws_bit.write(0, 2, "Tipo Incidencia")
    ws_bit.write(0, 3, "Hora Inicio (0-23)")
    ws_bit.write(0, 4, "Hora Fin (0-23)")
    
    # Validacion de datos para Nombres
    names_list = df_sorted['Nombre'].tolist()
    # Escribir nombres en columna oculta para validación si son muchos, o directo si pocos.
    # Limitación xlsxwriter: validación lista directa max 255 chars. Usaremos un rango en hoja config si fuera pro.
    # Aquí simplificamos: el usuario debe escribir bien el nombre.
    
    # 3. OPERACIÓN REAL (La Sábana con Fórmulas)
    ws_op = wb.add_worksheet("Operacion_Real")
    
    # Formatos
    f_head = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#44546A', 'font_color': 'white', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'text_wrap': True})
    f_alert = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'}) # Rojo incidencias
    
    # Estructura Sábana Visual
    ws_op.write(6, 0, "Colaborador", f_head)
    ws_op.write(6, 1, "Rol", f_head)
    ws_op.freeze_panes(7, 2)
    
    col = 2
    d_map = {} # Map fecha -> columna inicio
    
    # Encabezados Superiores (Contadores con Fórmulas Excel)
    ws_op.write(2, 0, "DOTACIÓN REAL", f_head)
    
    for d in dates:
        d_str = pd.to_datetime(d).strftime("%d-%b")
        d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
        ws_op.merge_range(0, col, 0, col+25, d_str, f_head)
        ws_op.write(6, col, "Turno", f_head)
        ws_op.write(6, col+1, "Lugar", f_head)
        
        # Guardar columna de fecha para formulas
        d_map[d_iso] = col 
        
        for h in range(24): 
            ws_op.write(6, col+2+h, h, f_head)
            # Aquí podríamos poner formulas de conteo en filas 2,3,4 sumando la columna actual
            # =CONTAR.SI(RangoColumna, "1") + ...
            col_letter = xlsxwriter.utility.xl_col_to_name(col+2+h)
            formula_count = f'=COUNTIFS({col_letter}8:{col_letter}1000,"*1*") + COUNTIFS({col_letter}8:{col_letter}1000,"*3*")'
            ws_op.write_formula(2, col+2+h, formula_count, f_base)

        col += 26
        
    # Filas de Personas con Fórmulas Mágicas
    row = 7
    
    for _, p in df_sorted.iterrows():
        n = p['Nombre']
        r = p['Rol']
        
        ws_op.write(row, 0, n, f_base)
        ws_op.write(row, 1, r, f_base)
        
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            c_start = d_map[d_iso]
            
            # Obtener datos base para Turno y Lugar (Textos fijos)
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            if subset.empty:
                t_raw = "-"
                lugar = "Libre"
            else:
                t_raw = subset.iloc[0]['Turno_Raw']
                try: lugar = subset.iloc[0]['Base_Diaria']
                except: lugar = "?"
            
            ws_op.write(row, c_start, str(t_raw), f_base)
            ws_op.write(row, c_start+1, str(lugar), f_base)
            
            # Celdas Horarias (FORMULAS)
            key = f"{n}_{d_iso}"
            
            # Buscar fila en Plan_Base (ineficiente en python puro, pero simulamos)
            # En Excel real usariamos VLOOKUP(key, Plan_Base!A:Z, column, 0)
            # Como sabemos el orden (es el mismo que df_sorted), podemos calcular la fila si Plan_Base sigue el mismo orden.
            # Sí, en step 1 iteramos df_sorted igual.
            # Plan_Base row = (row - 7) * len(dates) + date_index + 1 ?? No, Plan_Base es flat.
            # Mejor usamos INDEX/MATCH.
            
            for h in range(24):
                # Formula Lógica:
                # 1. Buscar si hay incidencia en Bitacora para (Nombre, Fecha) y que la Hora esté en rango [Ini, Fin]
                #    =CONTAR.SI.CONJUNTO(Bitacora!A:A, Nombre, Bitacora!B:B, Fecha, Bitacora!D:D, "<="&h, Bitacora!E:E, ">="&h)
                # 2. Si > 0, escribir "FALTA" (o Tipo)
                # 3. Si 0, traer valor de Plan_Base usando INDICE/COINCIDIR
                
                # Excel Date Serial no es trivial in string formulas, pasamos fecha como texto o numero.
                # Simplificación: Asumimos fecha en texto en Bitacora columna B.
                
                # Referencia a celda actual en Plan_Base para el valor por defecto
                # Usaremos LOOKUP: INDICE(Plan_Base!ColH, COINCIDIR(Key, Plan_Base!A:A, 0))
                # Columna en Plan_Base: A=0, B=H0 (idx 1), C=H1 (idx 2)... -> H + 1
                col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                
                formula = (
                    f'=IF(COUNTIFS(Bitacora!$A:$A, "{n}", Bitacora!$B:$B, "{d_iso}", '
                    f'Bitacora!$D:$D, "<={h}", Bitacora!$E:$E, ">={h}")>0, "INCIDENCIA", '
                    f'INDEX(Plan_Base!{col_plan_letter}:{col_plan_letter}, MATCH("{key}", Plan_Base!$A:$A, 0)))'
                )
                
                # Escribir formula
                ws_op.write_formula(row, c_start+2+h, formula, f_base)
                
                # Conditional Formatting para "INCIDENCIA"
                ws_op.conditional_format(row, c_start+2+h, row, c_start+2+h, 
                                        {'type': 'cell', 'criteria': 'equal to', 'value': '"INCIDENCIA"', 'format': f_alert})

        row += 1

    wb.close()
    return out

if st.button("🚀 Generar Planificación V41"):
    if not uploaded_sheets: st.error("Carga archivos.")
    elif not (start_d and end_d): st.error("Define fechas.")
    else:
        with st.spinner("Procesando..."):
            dfs = []
            for role, (key) in [("Agente","exec"),("Coordinador","coord"),("Anfitrion","host"),("Supervisor","sup")]:
                if key in uploaded_sheets:
                    f, s = uploaded_sheets[key]
                    dfs.append(process_file_sheet(f, s, role, start_d, end_d))
            full = pd.concat(dfs)
            
            if full.empty: st.error("Sin datos.")
            else:
                final = logic_engine(full, agents_no_tica)
                st.success("¡Listo!")
                st.download_button("📥 Descargar Excel Inteligente (V41)", make_excel(final, start_d, end_d), f"Planificacion_V41.xlsx")
