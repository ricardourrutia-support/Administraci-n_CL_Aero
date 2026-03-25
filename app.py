import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto (V69)", layout="wide")
st.title("✈️ Gestor de Turnos: V69 (Blindaje Operativo)")
st.markdown("""
**Optimizaciones V69:**
1. **Regla Coordinador:** Nunca se moverán a cubrir si eso implica dejar la Tarea 1 (Supervisión) en 0.
2. **HHEE Anfitriones:** Garantía de mínimo 1 persona por zona; si no hay nadie en ninguna, detona HHEE.
3. **Supervisores Visibles:** Se agregó un contador total de horas para Supervisores y Coordinadores en el Resumen.
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

def get_prior_night_agents(df_raw, start_date):
    agents = set()
    prior_date = start_date - timedelta(days=1)
    for _, r in df_raw.iterrows():
        c_dt = r['Fecha'].date() if isinstance(r['Fecha'], datetime) else r['Fecha']
        if c_dt == prior_date:
            hours, start_h = parse_shift_time(r['Turno_Raw'])
            if hours and start_h >= 18:
                agents.add(r['Nombre'])
    return sorted(list(agents))

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

st.sidebar.markdown("---")
st.sidebar.header("4. Continuidad (Madrugada)")
init_counters = {}
if 'exec' in uploaded_sheets and start_d:
    f_exec, s_exec = uploaded_sheets['exec']
    try:
        df_temp = process_file_sheet(f_exec, s_exec, "Agente", start_d, end_d)
        midnight_agents = get_prior_night_agents(df_temp, start_d)
        if midnight_agents:
            st.sidebar.info(f"Agentes de la noche anterior ({ (start_d - timedelta(days=1)).strftime('%d/%m') }). Indique en qué counter amanecieron:")
            for ag in midnight_agents:
                init_counters[ag] = st.sidebar.selectbox(ag, ["T2 AIRE", "T2 TIERRA", "T1 AIRE", "T1 TIERRA"], key=f"init_{ag}")
    except: pass

# --- MOTOR LÓGICO ---
def logic_engine(df, no_tica_list, initial_counters):
    rows = []
    raw_shifts_map = {}
    for _, r in df.iterrows():
        raw_shifts_map[(r['Nombre'], r['Fecha'])] = r['Turno_Raw']

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
            rows.append({
                **r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '', 
                'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': -1, 
                'Base_Diaria': '', 'Shift_Date': r['Fecha'] 
            })
        else:
            shift_date = r['Fecha'] 
            for h in hours:
                current_date = r['Fecha']
                if start_h >= 18 and h < 12: current_date = current_date + timedelta(days=1)
                rows.append({
                    'Nombre': r['Nombre'], 'Rol': r['Rol'], 'Turno_Raw': r['Turno_Raw'],
                    'Fecha': current_date, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                    'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': start_h, 
                    'Base_Diaria': '?', 'Shift_Date': shift_date 
                })
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h, raw_shifts_map, {}

    main_counters_aire = ["T1 AIRE", "T2 AIRE"]
    main_counters_tierra = ["T1 TIERRA", "T2 TIERRA"]
    pref_order = ["T2 AIRE", "T2 TIERRA", "T1 AIRE", "T1 TIERRA"] 
    
    shift_assignments = {} 
    shift_status = {} 
    anf_shift_assignments = {}
    last_ag_counter = {}
    
    for nm, cnt in initial_counters.items():
        last_ag_counter[nm] = cnt
        
    sorted_shift_dates = sorted(df_h['Shift_Date'].unique())
    
    for s_d in sorted_shift_dates:
        try:
            s_d_date = pd.to_datetime(s_d).date()
            is_prior_day = (s_d_date == (start_d - timedelta(days=1)))
        except:
            is_prior_day = False
            
        # --- AGENTES ---
        df_sd_ag = df_h[(df_h['Shift_Date'] == s_d) & (df_h['Rol'] == 'Agente') & (df_h['Hora'] != -1)]
        if not df_sd_ag.empty:
            active_ag_names = df_sd_ag['Nombre'].unique()
            am_starters = []
            pm_starters = []
            for nm in active_ag_names:
                sh = df_sd_ag[df_sd_ag['Nombre'] == nm].iloc[0]['Start_H']
                if sh < 12: am_starters.append((nm, sh))
                else: pm_starters.append((nm, sh))
                
            load_ag = {c: 0 for c in main_counters_aire + main_counters_tierra}
            
            def assign_group_sd(starters_list, is_am):
                fixed_assigned = set()
                starters_list.sort(key=lambda x: x[1]) 
                
                for name, sh in starters_list:
                    has_tica = name not in no_tica_list
                    assigned_c = None
                    status = 2 
                    
                    if is_prior_day and not is_am and name in initial_counters:
                        assigned_c = initial_counters[name]
                        if len(fixed_assigned) < 4:
                            fixed_assigned.add(assigned_c)
                            status = 1
                    else:
                        if len(fixed_assigned) < 4:
                            for fc in ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]:
                                if fc not in fixed_assigned:
                                    if not has_tica and fc == "T1 AIRE": continue
                                    assigned_c = fc
                                    break
                        
                        if assigned_c:
                            fixed_assigned.add(assigned_c)
                            status = 1 
                        else:
                            status = 2 
                            valid_prefs = [c for c in pref_order if has_tica or c != "T1 AIRE"]
                            valid_prefs.sort(key=lambda c: (load_ag.get(c, 0), pref_order.index(c)))
                            assigned_c = valid_prefs[0]
                    
                    shift_assignments[(name, s_d)] = assigned_c
                    shift_status[(name, s_d)] = status
                    load_ag[assigned_c] = load_ag.get(assigned_c, 0) + 1
                    
            assign_group_sd(am_starters, True)
            assign_group_sd(pm_starters, False)

        # --- ANFITRIONES ---
        df_sd_anf = df_h[(df_h['Shift_Date'] == s_d) & (df_h['Rol'] == 'Anfitrion') & (df_h['Hora'] != -1)]
        if not df_sd_anf.empty:
            active_anf_names = df_sd_anf['Nombre'].unique()
            zone_load = {'Zona Int': 0, 'Zona Nac': 0}
            for name in sorted(active_anf_names):
                z = 'Zona Int' if zone_load['Zona Int'] <= zone_load['Zona Nac'] else 'Zona Nac'
                anf_shift_assignments[(name, s_d)] = z
                zone_load[z] += 1

    # --- COORDINADORES Y SUPERVISORES ---
    coords_active = df_h[(df_h['Rol'].isin(['Coordinador', 'Supervisor'])) & (df_h['Hora'] != -1)]
    for idx, row in coords_active.iterrows():
        if row['Rol'] == 'Supervisor':
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'
            continue
            
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

    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        idx_ag = g[g['Rol']=='Agente'].index.tolist()
        idx_co = g[g['Rol']=='Coordinador'].index.tolist()
        idx_an = g[g['Rol']=='Anfitrion'].index.tolist()
        idx_su = g[g['Rol']=='Supervisor'].index.tolist()
        
        for idx in idx_ag:
            df_h.at[idx, 'Counter'] = shift_assignments.get((df_h.at[idx, 'Nombre'], df_h.at[idx, 'Shift_Date']), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_an:
            df_h.at[idx, 'Counter'] = anf_shift_assignments.get((df_h.at[idx, 'Nombre'], df_h.at[idx, 'Shift_Date']), "General")
            df_h.at[idx, 'Tarea'] = '1'
        for idx in idx_co:
            if df_h.at[idx, 'Tarea'] not in ['2', 'C']:
                df_h.at[idx, 'Counter'] = "General"; df_h.at[idx, 'Tarea'] = '1'

        for idx in idx_ag:
            sh = df_h.at[idx, 'Start_H']
            break_h = -1
            if 0 <= sh <= 11:
                if sh <= 8: break_h = 12
                elif sh == 9: break_h = 13
                elif sh == 10: break_h = 14
                else: break_h = 15
            elif 18 <= sh <= 23:
                if sh <= 20: break_h = 2
                elif sh == 21: break_h = 3
                else: break_h = 4
            else: 
                break_h = (sh + 4) % 24
            if h == break_h:
                df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                
        def apply_break_anf(indices, start_range, slots):
            cands = [i for i in indices if start_range[0] <= df_h.at[i, 'Start_H'] <= start_range[1]]
            for i, idx in enumerate(cands):
                if h == slots[i % len(slots)]: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
        apply_break_anf(idx_an, (0, 11), [13, 14, 15]); apply_break_anf(idx_an, (12, 23), [2, 3, 4])

        # V69: NUEVA LÓGICA DE COORDINADOR (Protección Tarea 1)
        def find_coord_cover():
            c_t1 = [i for i in idx_co if str(df_h.at[i, 'Tarea']) == '1']
            c_t2 = [i for i in idx_co if str(df_h.at[i, 'Tarea']) == '2']
            
            # Solo usa un T2 si al menos hay alguien en T1 operando
            if len(c_t1) >= 1 and len(c_t2) > 0:
                return c_t2[0]
            # Solo usa un T1 si hay MÁS de 1 en T1 (para nunca dejar T1 en cero)
            if len(c_t1) > 1:
                return c_t1[0]
            return None

        needs_coverage = []
        for idx in idx_ag:
            nm = df_h.at[idx, 'Nombre']
            sd = df_h.at[idx, 'Shift_Date']
            status = shift_status.get((nm, sd), 2)
            if status == 1 and df_h.at[idx, 'Tarea'] == 'C':
                target_cnt = shift_assignments.get((nm, sd))
                needs_coverage.append(target_cnt)

        active_counts = {c: 0 for c in main_counters_aire + main_counters_tierra}
        for idx in idx_ag:
            if df_h.at[idx, 'Tarea'] == '1':
                active_counts[df_h.at[idx, 'Counter']] += 1
        
        for c, count in active_counts.items():
            if count == 0 and c not in needs_coverage:
                needs_coverage.append(c)
                
        available_flotantes = []
        for idx in idx_ag:
            nm = df_h.at[idx, 'Nombre']
            sd = df_h.at[idx, 'Shift_Date']
            status = shift_status.get((nm, sd), 2)
            if status == 2 and df_h.at[idx, 'Tarea'] == '1':
                available_flotantes.append(idx)
                
        uncovered = []
        for target_cnt in needs_coverage:
            covered = False
            for f_idx in available_flotantes:
                nm = df_h.at[f_idx, 'Nombre']
                if target_cnt == "T1 AIRE" and nm in no_tica_list: continue
                f_base = shift_assignments.get((nm, df_h.at[f_idx, 'Shift_Date']))
                if f_base == target_cnt:
                    df_h.at[f_idx, 'Tarea'] = f"3: Cubrir {target_cnt}"
                    df_h.at[f_idx, 'Counter'] = target_cnt
                    available_flotantes.remove(f_idx)
                    covered = True
                    break
            if not covered: uncovered.append(target_cnt)
            
        still_uncovered = []
        for target_cnt in uncovered:
            covered = False
            for f_idx in available_flotantes:
                nm = df_h.at[f_idx, 'Nombre']
                if target_cnt == "T1 AIRE" and nm in no_tica_list: continue
                df_h.at[f_idx, 'Tarea'] = f"3: Cubrir {target_cnt}"
                df_h.at[f_idx, 'Counter'] = target_cnt
                available_flotantes.remove(f_idx)
                covered = True
                break
            if not covered: still_uncovered.append(target_cnt)
                
        for target_cnt in still_uncovered:
            covered = False
            cand = find_coord_cover()
            if cand is not None:
                df_h.at[cand, 'Tarea'] = f"4: Cubrir {target_cnt}"
                df_h.at[cand, 'Counter'] = target_cnt
                covered = True
                
            if not covered:
                hhee_counters.append({'Fecha': d, 'Hora': h, 'Counter': target_cnt})

        # V69: CORRECCIÓN HHEE ANFITRIONES (Cuando ambas zonas caen a 0)
        active_nac = sum(1 for idx in idx_an if df_h.at[idx, 'Tarea'] == '1' and df_h.at[idx, 'Counter'] == 'Zona Nac')
        active_int = sum(1 for idx in idx_an if df_h.at[idx, 'Tarea'] == '1' and df_h.at[idx, 'Counter'] == 'Zona Int')
        
        target_zs = []
        if active_nac == 0: target_zs.append('Zona Nac')
        if active_int == 0: target_zs.append('Zona Int')
        
        for tz in target_zs:
            covered = False
            other_z = 'Zona Int' if tz == 'Zona Nac' else 'Zona Nac'
            active_other = sum(1 for idx in idx_an if df_h.at[idx, 'Tarea'] == '1' and df_h.at[idx, 'Counter'] == other_z)
            
            # Solo pide a un compañero si el otro lado tiene a más de 1
            if active_other > 1:
                for a_idx in idx_an:
                    if df_h.at[a_idx, 'Tarea'] == '1' and df_h.at[a_idx, 'Counter'] == other_z:
                        df_h.at[a_idx, 'Counter'] = tz 
                        covered = True
                        break
            
            if not covered:
                cand = find_coord_cover()
                if cand is not None:
                    df_h.at[cand, 'Tarea'] = f"Cubrir {tz}"
                    df_h.at[cand, 'Counter'] = tz
                    covered = True
            
            if not covered:
                hhee_anf.append({'Fecha': d, 'Hora': h, 'Counter': tz})

    for (name, sd), group in df_h[df_h['Rol'] == 'Anfitrion'].groupby(['Nombre', 'Shift_Date']):
        valid_locs = group[~group['Counter'].isin(['Casino'])]['Counter']
        if not valid_locs.empty:
            true_base = valid_locs.mode()[0] 
        else:
            true_base = 'Zona Int'
            
        anf_shift_assignments[(name, sd)] = true_base 
        
        for idx in group.index:
            loc = df_h.at[idx, 'Counter']
            if loc == 'Casino': df_h.at[idx, 'Tarea'] = 'C'
            elif loc == true_base: df_h.at[idx, 'Tarea'] = '1'
            else: df_h.at[idx, 'Tarea'] = f"3: Cubrir {loc}" 

    for idx, row in df_h.iterrows():
        if row['Hora'] != -1: 
            if row['Rol'] == 'Coordinador': df_h.at[idx, 'Base_Diaria'] = "General"
            elif row['Rol'] == 'Supervisor': df_h.at[idx, 'Base_Diaria'] = "General"
            elif row['Rol'] == 'Agente': df_h.at[idx, 'Base_Diaria'] = shift_assignments.get((row['Nombre'], row['Shift_Date']), "")
            elif row['Rol'] == 'Anfitrion': df_h.at[idx, 'Base_Diaria'] = anf_shift_assignments.get((row['Nombre'], row['Shift_Date']), "")

    if 'incidencias' in st.session_state and st.session_state.incidencias:
        df_h = apply_incidents(df_h, st.session_state.incidencias)

    unique_dates = sorted(df_h['Fecha'].unique())
    hhee_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_counters)
    
    for cnt in ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]:
        base_row = {'Nombre': f"HHEE Agente - {cnt}", 'Rol': 'HHEE', 'Sub_Group': 'Counter', 'Role_Rank': 900, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': cnt}
        for d in unique_dates:
            for h in range(24):
                task_val = "REQ" if (d, h, cnt) in hhee_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val, 'Shift_Date': d}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-', 'Shift_Date': d}])], ignore_index=True)
            
    hhee_c_set = set((x['Fecha'], x['Hora']) for x in hhee_coord)
    if hhee_c_set:
        base_row = {'Nombre': "HHEE Coordinación", 'Rol': 'HHEE', 'Sub_Group': 'Supervisión', 'Role_Rank': 901, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': 'General'}
        for d in unique_dates:
            for h in range(24):
                task_val = "REQ" if (d, h) in hhee_c_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val, 'Shift_Date': d}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-', 'Shift_Date': d}])], ignore_index=True)
        
    hhee_a_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_anf)
    for z in ['Zona Int', 'Zona Nac']:
        base_row = {'Nombre': f"HHEE Anfitriones - {z.upper()}", 'Rol': 'HHEE', 'Sub_Group': 'Losa', 'Role_Rank': 902, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Base_Diaria': z}
        for d in unique_dates:
            for h in range(24):
                task_val = "REQ" if (d, h, z) in hhee_a_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val, 'Shift_Date': d}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-', 'Shift_Date': d}])], ignore_index=True)

    eligible_hhee = {}
    supervisores = df_h[df_h['Rol'] == 'Supervisor']['Nombre'].unique().tolist()
    
    for d in unique_dates:
        for h in range(24):
            eligibles = list(supervisores)
            prev_h = 23 if h == 0 else h - 1
            prev_d = d - timedelta(days=1) if h == 0 else d
            
            working_prev = set(df_h[(df_h['Fecha'] == prev_d) & (df_h['Hora'] == prev_h) & (df_h['Rol'].isin(['Agente', 'Anfitrion', 'Coordinador'])) & (df_h['Tarea'] != 'Ausente')]['Nombre'])
            working_curr = set(df_h[(df_h['Fecha'] == d) & (df_h['Hora'] == h) & (df_h['Rol'].isin(['Agente', 'Anfitrion', 'Coordinador']))]['Nombre'])
            
            just_finished = working_prev - working_curr
            for nm in just_finished:
                if nm not in eligibles: eligibles.append(nm)
                
            eligible_hhee[(d, h)] = eligibles

    return df_h, raw_shifts_map, eligible_hhee

# --- EXCEL GENERATOR (V69) ---
def make_excel(df, raw_shifts_map, start_d, end_d, eligible_hhee):
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
    f_sep_col = wb.add_format({'bg_color': '#1C1C1C', 'border': 1, 'align': 'center'}) 
    
    f_hhee_req = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'bold': True})
    f_hhee_bad = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'bold': True})
    f_hhee_ok = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center', 'bold': True})
    
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        '3': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center', 'font_size': 8, 'text_wrap': True}),
        '4': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'bold': True, 'font_size': 8, 'text_wrap': True}),
        'C': wb.add_format({'bg_color': '#C6E0B4', 'border': 1, 'align': 'center'}),
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
    
    opciones_manuales = ['1', '2', 'C', '3: Cubrir T1 AIRE', '3: Cubrir T1 TIERRA', '3: Cubrir T2 AIRE', '3: Cubrir T2 TIERRA', 'Zona Int', 'Zona Nac', 'General']
    ws_data.write_column(1, 0, unique_roles)
    ws_data.write_column(1, 1, unique_names)
    ws_data.write_column(1, 2, opciones_manuales)
    ws_data.hide()
    
    name_range = f"Datos_Validacion!$B$2:$B${len(unique_names)+1}"
    role_range = f"Datos_Validacion!$A$2:$A${len(unique_roles)+1}"
    manual_range = f"Datos_Validacion!$C$2:$C${len(opciones_manuales)+1}"

    ws_hhee_val = wb.add_worksheet("HHEE_Val")
    ws_hhee_val.hide()
    
    val_map = {} 
    v_col = 0
    for d in dates:
        for h in range(24):
            ws_hhee_val.write(0, v_col, f"{d.strftime('%Y-%m-%d')}_{h}")
            eligibles = eligible_hhee.get((d, h), [])
            if not eligibles: eligibles = ["-"]
            ws_hhee_val.write_column(1, v_col, eligibles)
            val_map[(d, h)] = xlsxwriter.utility.xl_col_to_name(v_col)
            v_col += 1

    ws_teorico = wb.add_worksheet("Plan_Teorico")
    ws_shiftdate = wb.add_worksheet("Plan_ShiftDate") 
    ws_teorico.write(0, 0, "ID") 
    ws_shiftdate.write(0, 0, "ID")
    
    teorico_row = 1
    for _, p in pd.concat([df_staff_sorted, df_hhee_sorted]).iterrows():
        n = p['Nombre']
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            key = f"{n}_{d_iso}"
            ws_teorico.write(teorico_row, 0, key)
            ws_shiftdate.write(teorico_row, 0, key)
            
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            for h in range(24):
                try:
                    val = subset[subset['Hora'] == h]
                    if not val.empty:
                        task = str(val.iloc[0]['Tarea'])
                        s_date = pd.to_datetime(val.iloc[0]['Shift_Date']).strftime("%Y-%m-%d")
                    else:
                        task = ""; s_date = ""
                    ws_teorico.write_string(teorico_row, 1+h, task)
                    ws_shiftdate.write_string(teorico_row, 1+h, s_date)
                except:
                    ws_teorico.write_string(teorico_row, 1+h, "")
                    ws_shiftdate.write_string(teorico_row, 1+h, "")
            teorico_row += 1
    ws_teorico.hide()
    ws_shiftdate.hide()

    ws_bit = wb.add_worksheet("Bitacora_Incidencias")
    headers_bit = ["Tipo Colaborador", "Nombre Colaborador", "Fecha (YYYY-MM-DD)", "Tipo Incidencia", "Hora Inicio (0-23)", "Hora Fin (0-23)"]
    for i, h in enumerate(headers_bit): ws_bit.write(0, i, h, f_cabify)
    ws_bit.data_validation('A2:A1000', {'validate': 'list', 'source': role_range})
    ws_bit.data_validation('B2:B1000', {'validate': 'list', 'source': name_range})
    ws_bit.data_validation('D2:D1000', {'validate': 'list', 'source': ['Inasistencia', 'Atraso', 'Salida Anticipada']})
    ws_bit.write(0, 7, "GUÍA OPERATIVA V69:", f_cabify)
    ws_bit.write(1, 7, "INASISTENCIA: Marque el día de inicio del turno. Borrará automáticamente la madrugada siguiente.")

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
    ws_real.write(8, 0, "TOTAL HHEE REQ.", f_header_hhee)
    
    días_es = {0: 'Lun', 1: 'Mar', 2: 'Mié', 3: 'Jue', 4: 'Vie', 5: 'Sáb', 6: 'Dom'}
    
    for d in dates:
        d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
        d_str = f"{días_es[d.weekday()]} {pd.to_datetime(d).strftime('%d-%b')}"
        
        ws_real.set_column(col, col, 1.5)
        ws_real.write(9, col, "", f_sep_col)
        
        ws_real.merge_range(0, col+1, 0, col+26, d_str, f_date)
        ws_real.write(9, col+1, "Turno", f_cabify)
        ws_real.write(9, col+2, "Lugar", f_cabify)
        
        d_map[d_iso] = col + 1
        
        for h in range(24):
            ws_real.write(9, col+3+h, h, f_cabify)
            col_idx = col+3+h
            col_let = xlsxwriter.utility.xl_col_to_name(col_idx)
            
            f_ag = f'=COUNTIFS($B$11:$B$1000,"Agente",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"<>Libre",{col_let}11:{col_let}1000,"<>*Ausente*", {col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(2, col_idx, f_ag, f_header_count)
            f_co = f'=COUNTIFS($B$11:$B$1000,"Coordinador",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(3, col_idx, f_co, f_header_count)
            f_an = f'=COUNTIFS($B$11:$B$1000,"Anfitrion",{col_let}11:{col_let}1000,"<>FALTA",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(4, col_idx, f_an, f_header_count)
            
            f_h_ag = f'=COUNTIFS($A$11:$A$1000,"HHEE Agente*",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(5, col_idx, f_h_ag, f_header_hhee)
            f_h_co = f'=COUNTIFS($A$11:$A$1000,"HHEE Coordinación",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(6, col_idx, f_h_co, f_header_hhee)
            f_h_an = f'=COUNTIFS($A$11:$A$1000,"HHEE Anfitriones*",{col_let}11:{col_let}1000,"?*")'
            ws_real.write_formula(7, col_idx, f_h_an, f_header_hhee)
            
            f_total = f'=SUM({col_let}6:{col_let}8)'
            ws_real.write_formula(8, col_idx, f_total, f_header_hhee)

        col += 27
    
    hhee_tot_range = f"D9:{xlsxwriter.utility.xl_col_to_name(col-1)}9"
    ws_real.conditional_format(hhee_tot_range, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': f_hhee_ok})
    ws_real.conditional_format(hhee_tot_range, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': f_hhee_bad})

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
            
            ws_real.write(row, c_start-1, "", f_sep_col) 
            
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            
            if subset.empty:
                t_raw_original = raw_shifts_map.get((n, d), "")
                if "libre" in str(t_raw_original).lower(): t_raw_original = ""
                ws_real.write(row, c_start, str(t_raw_original), f_base)
                ws_real.write(row, c_start+1, "", f_base)
                
                for h in range(24): 
                    key = f"{n}_{d_iso}"
                    col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                    formula = f'T(INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0))) & ""'
                    ws_real.write_formula(row, c_start+2+h, formula, f_base)
            else:
                t_raw_original = raw_shifts_map.get((n, d), "")
                if "libre" in str(t_raw_original).lower(): t_raw_original = ""
                try: lugar = subset.iloc[0]['Base_Diaria']
                except: lugar = "?"
                ws_real.write(row, c_start, str(t_raw_original), f_base)
                
                fmt_lugar = f_base
                if r == "Anfitrion":
                    if "Nac" in str(lugar): fmt_lugar = f_nac
                    elif "Int" in str(lugar): fmt_lugar = f_int
                ws_real.write(row, c_start+1, str(lugar), fmt_lugar)
            
                key = f"{n}_{d_iso}"
                for h in range(24):
                    col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                    formula = (
                        f'=IF(INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0))&""="", "",'
                        f'IF(COUNTIFS(Bitacora_Incidencias!$B:$B,"{n}",Bitacora_Incidencias!$C:$C,INDEX(Plan_ShiftDate!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_ShiftDate!$A:$A,0)),Bitacora_Incidencias!$D:$D,"Inasistencia")>0,"FALTA",'
                        f'IF(COUNTIFS(Bitacora_Incidencias!$B:$B,"{n}",Bitacora_Incidencias!$C:$C,"{d_iso}",Bitacora_Incidencias!$E:$E,"<={h}",Bitacora_Incidencias!$F:$F,">={h}",Bitacora_Incidencias!$D:$D,"<>Inasistencia")>0,"INCIDENCIA",'
                        f'INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0)) & "")))'
                    )
                    ws_real.write_formula(row, c_start+2+h, formula, f_base)
                    ws_real.data_validation(row, c_start+2+h, row, c_start+2+h, {'validate': 'list', 'source': manual_range, 'show_error': False})
        row += 1
    
    hhee_start_row = row + 1
    ws_real.merge_range(row, 0, row, col-1, "REQUERIMIENTOS HHEE (SELECCIONE NOMBRE)", f_group)
    row += 1
    
    for _, p in df_hhee_sorted.iterrows():
        n = p['Nombre']
        r = p['Rol']
        ws_real.write(row, 0, n, f_base)
        ws_real.write(row, 1, r, f_base)
        
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            c_start = d_map[d_iso]
            ws_real.write(row, c_start-1, "", f_sep_col)
            ws_real.write(row, c_start, "Demanda", f_base)
            try: 
                lugar = df[(df['Nombre']==n)&(df['Fecha']==d)].iloc[0]['Base_Diaria']
            except: lugar = "?"
            ws_real.write(row, c_start+1, str(lugar), f_base)
            
            key = f"{n}_{d_iso}"
            for h in range(24):
                col_plan_letter = xlsxwriter.utility.xl_col_to_name(h + 1)
                formula = f'=IF(INDEX(Plan_Teorico!{col_plan_letter}:{col_plan_letter},MATCH("{key}",Plan_Teorico!$A:$A,0))="REQ", "REQ", "")'
                ws_real.write_formula(row, c_start+2+h, formula, f_base)
                
                val_col_let = val_map.get((d, h))
                if val_col_let:
                    val_source = f'=HHEE_Val!${val_col_let}$2:${val_col_let}$100'
                    ws_real.data_validation(row, c_start+2+h, row, c_start+2+h, {'validate': 'list', 'source': val_source, 'show_error': False})

        row += 1
    hhee_end_row = row

    end_col_let = xlsxwriter.utility.xl_col_to_name(col-1)
    data_range = f"D11:{end_col_let}{row}"
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"FALTA"', 'format': f_alert})
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"INCIDENCIA"', 'format': f_alert})
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"REQ"', 'format': f_hhee_req})
    
    st_map_cond = {
        '3:': st_map['3'], '4:': st_map['4'], 'Cubrir': st_map['4'], 
        'C': st_map['C'], '2': st_map['2'], 'Nac': st_map['Nac'], 'Int': st_map['Int']
    }
    for k, fmt in st_map_cond.items():
        criteria = 'begins with' if len(k) > 2 else 'equal to'
        val = k if len(k) > 2 else f'"{k}"'
        ws_real.conditional_format(data_range, {'type': 'text' if len(k)>2 else 'cell', 'criteria': criteria, 'value': val, 'format': fmt})

    hhee_box_range = f"D{hhee_start_row+1}:{end_col_let}{hhee_end_row}"
    ws_real.conditional_format(hhee_box_range, {'type': 'formula', 'criteria': f'=AND(D{hhee_start_row+1}<>"", D{hhee_start_row+1}<>"REQ")', 'format': f_hhee_ok})

    # ------------------------------------------------
    # HOJA 4: RESUMEN V69
    # ------------------------------------------------
    ws_res = wb.add_worksheet("Resumen_Estadistico")
    f_bold = wb.add_format({'bold': True})
    
    df_res = df[df['Fecha'].apply(lambda x: start_d <= x.date() <= end_d)].copy()
    
    # TABLA 1: HORAS DE EJECUTIVOS
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

    # TABLA 2: HORAS DE SUPERVISORES Y OTROS ROLES (V69 NUEVO)
    r_idx += 2
    ws_res.write(r_idx, 0, "HORAS TRABAJADAS (OTROS ROLES)", f_bold)
    r_idx += 1
    other_work = df_res[(df_res['Rol'].isin(['Anfitrion', 'Coordinador', 'Supervisor'])) & (df_res['Hora'] != -1) & (~df_res['Tarea'].astype(str).str.contains('Ausente|Falta', na=False, case=False))]
    if not other_work.empty:
        totals = other_work.groupby(['Rol', 'Nombre']).size().reset_index(name='Horas')
        ws_res.write(r_idx, 0, "Rol", f_bold)
        ws_res.write(r_idx, 1, "Nombre", f_bold)
        ws_res.write(r_idx, 2, "Horas Totales (Inc. Colación)", f_bold)
        r_idx += 1
        for _, row_data in totals.iterrows():
            ws_res.write(r_idx, 0, row_data['Rol'])
            ws_res.write(r_idx, 1, row_data['Nombre'])
            ws_res.write(r_idx, 2, row_data['Horas'])
            r_idx += 1

    # TABLA 3: ESTADÍSTICA DE HHEE
    r_idx += 2
    ws_res.write(r_idx, 0, "ESTADÍSTICAS HHEE (REQUERIDAS POR SISTEMA)", f_bold)
    r_idx += 2
    hhee_active = df_res[(df_res['Rol'] == 'HHEE') & (df_res['Tarea'] == 'REQ')]
    
    ws_res.write(r_idx, 0, "Total HHEE Requeridas (Teórico):", f_bold)
    ws_res.write(r_idx, 1, len(hhee_active))
    r_idx += 2
    
    ws_res.write(r_idx, 0, "HHEE ASIGNADAS MANUALMENTE", f_bold)
    ws_res.write(r_idx+1, 0, "Colaborador", f_bold)
    ws_res.write(r_idx+1, 1, "Horas Asignadas", f_bold)
    r_idx += 2
    
    for _, p in df_staff_sorted.iterrows():
        n = p['Nombre']
        ws_res.write(r_idx, 0, n)
        f_count = f'=COUNTIF(Plan_Operativo!$D${hhee_start_row+1}:${end_col_let}${hhee_end_row}, "{n}")'
        ws_res.write_formula(r_idx, 1, f_count)
        r_idx += 1
        
    wb.close()
    return out

st.sidebar.markdown("---")
if st.sidebar.button("🚀 Generar Planificación V69"):
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
            
            if full.empty: st.error("Sin datos válidos (revise el formato del Excel o las fechas).")
            else:
                final, raw_map, el_hhee = logic_engine(full, agents_no_tica, init_counters)
                st.success("¡Listo! Descarga la Suite Operativa V69.")
                st.download_button("📥 Descargar Suite (V69)", make_excel(final, raw_map, start_d, end_d, el_hhee), f"Planificacion_Operativa.xlsx")
