import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Gestor de Turnos: V15 (Anfitriones Estables + HHEE Totales)")
st.markdown("""
**Novedades V15:**
1. **Turnos Anfitriones:** Parser mejorado para asegurar las 11 horas continuas.
2. **HHEE Coordinación:** Alerta si no hay nadie en Tarea 1.
3. **HHEE Anfitriones:** Alerta si hay menos de 2 en Losa (Nac/Int).
4. **Protección:** Un Anfitrión no cubre a un Agente si eso rompe su propio mínimo.
""")

# --- PARSEO ROBUSTO ---
def parse_shift_time(shift_str):
    """
    Intenta extraer horas de inicio y fin de strings sucios.
    Soporta: '08:00 - 19:00', '8-19', '08:00 a 19:00', etc.
    """
    if pd.isna(shift_str): return [], None
    s = str(shift_str).lower().strip()
    
    # Palabras clave de inactividad
    if s == "" or any(x in s for x in ['libre', 'nan', 'l', 'x', 'vacaciones', 'licencia', 'falla', 'domingos libres', 'festivo', 'feriado']):
        return [], None
    
    # 1. Limpieza agresiva: dejar solo números y separadores clave
    # Reemplazar palabras comunes por espacios
    s_clean = s.replace("hrs", "").replace("horas", "").replace("de", "").replace("a", "-").replace(":", ".")
    
    # Buscar patrones de números
    # Buscamos floats o ints (ej: 8.30, 8, 20.00)
    numbers = re.findall(r"(\d+(?:\.\d+)?)", s_clean)
    
    start_h = -1
    end_h = -1
    
    try:
        if len(numbers) >= 2:
            # Tomamos el primero y el ultimo (o el segundo) asumiendo Inicio - Fin
            n1 = float(numbers[0])
            n2 = float(numbers[1])
            
            # Normalizar (ej: si es 8.30, lo tratamos como hora 8)
            start_h = int(n1)
            end_h = int(n2)
            
            # Validación básica de rango horario
            if not (0 <= start_h <= 24) or not (0 <= end_h <= 24):
                return [], None
                
            # Generar rango
            if start_h < end_h:
                hours = list(range(start_h, end_h))
            elif start_h > end_h: # Turno noche
                hours = list(range(start_h, 24)) + list(range(0, end_h))
            else:
                return [], None # Mismo inicio y fin
            
            return hours, start_h
    except:
        return [], None
        
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
            if pd.isna(name_val): continue
            s_name = str(name_val).strip()
            # Filtros básicos de filas basura
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
start_d, end_d = (date_range[0], date_range[1]) if len(date_range)==2 else (None, None)

st.sidebar.markdown("---")
st.sidebar.header("2. Archivos")
uploaded_sheets = {} 
files_info = [("Agente", "exec"), ("Coordinador", "coord"), ("Anfitrion", "host"), ("Supervisor", "sup")]

for label, key in files_info:
    f = st.sidebar.file_uploader(f"{label}", type=["xlsx"], key=key)
    if f and start_d:
        try:
            xl = pd.ExcelFile(f)
            m_names = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            curr = m_names[start_d.month]
            def_ix = next((i for i, s in enumerate(xl.sheet_names) if curr.lower() in s.lower()), 0)
            sel_sheet = st.sidebar.selectbox(f"Hoja ({label})", xl.sheet_names, index=def_ix, key=f"{key}_sheet")
            uploaded_sheets[key] = (f, sel_sheet)
        except: pass

st.sidebar.markdown("---")
st.sidebar.header("3. TICA")
agents_no_tica = []
if 'exec' in uploaded_sheets and start_d:
    f_exec, s_exec = uploaded_sheets['exec']
    try:
        df_temp = process_file_sheet(f_exec, s_exec, "Agente", start_d, end_d)
        if not df_temp.empty:
            unique_names = sorted(df_temp['Nombre'].unique().tolist())
            agents_no_tica = st.sidebar.multiselect("Agentes SIN TICA", unique_names)
    except: pass

# --- MOTOR LÓGICO V15 ---
def logic_engine(df, no_tica_list):
    rows = []
    
    # 1. CLASIFICACIÓN
    agent_class = {}
    df_agentes = df[df['Rol'] == 'Agente']
    for name, group in df_agentes.groupby('Nombre'):
        am = 0
        pm = 0
        for _, r in group.iterrows():
            _, start_h = parse_shift_time(r['Turno_Raw'])
            if start_h is not None:
                if start_h < 12: am += 1
                else: pm += 1
        agent_class[name] = "Nocturno" if pm > am else "Diurno"

    # 2. EXPANDIR
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
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 
                         'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': -1})
        else:
            for h in hours:
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                             'Role_Rank': role_rank, 'Sub_Group': sub_group, 'Start_H': start_h})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    # 3. COUNTERS DIARIOS AGENTES
    main_counters_aire = ["T1 AIRE", "T2 AIRE"]
    main_counters_tierra = ["T1 TIERRA", "T2 TIERRA"]
    daily_assignments = {} 
    
    for d in df_h['Fecha'].unique():
        active = df_h[(df_h['Fecha'] == d) & (df_h['Rol'] == 'Agente') & (df_h['Hora'] != -1)]['Nombre'].unique()
        load = {c: 0 for c in main_counters_aire + main_counters_tierra}
        for ag_name in active:
            has_tica = ag_name not in no_tica_list
            chosen = sorted(main_counters_tierra if not has_tica else main_counters_aire + main_counters_tierra, key=lambda c: load[c])[0]
            load[chosen] += 1
            daily_assignments[(ag_name, d)] = chosen

    # 4. PRE-PROCESO COORDINADORES (OFF)
    coords_active = df_h[(df_h['Rol'] == 'Coordinador') & (df_h['Hora'] != -1)]
    for idx, row in coords_active.iterrows():
        st_h = row['Start_H']
        h = row['Hora']
        nm = row['Nombre']
        is_odd = hash(nm) % 2 != 0
        
        if st_h == 10:
            if h == 10: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
            elif h in [14, 15]:
                if (h == 14 and is_odd) or (h == 15 and not is_odd):
                    df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else:
                    df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
        elif st_h == 5:
            if h in [11, 12, 13]:
                if h == 12: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'
        elif st_h == 21:
            if h in [5, 6, 7]:
                if h == 6: df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
                else: df_h.at[idx, 'Tarea'] = '2'; df_h.at[idx, 'Counter'] = 'Oficina'

    # LISTAS PARA RASTREAR HHEE
    hhee_counters = []
    hhee_coord = []
    hhee_anf = []

    # 5. PROCESO HORA A HORA
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        idx_ag = g[g['Rol']=='Agente'].index.tolist()
        idx_co = g[g['Rol']=='Coordinador'].index.tolist()
        idx_an = g[g['Rol']=='Anfitrion'].index.tolist()
        idx_su = g[g['Rol']=='Supervisor'].index.tolist()
        
        # A. Agentes (Base y Colación)
        for idx in idx_ag:
            base = daily_assignments.get((df_h.at[idx, 'Nombre'], d), "General")
            df_h.at[idx, 'Counter'] = base
            df_h.at[idx, 'Tarea'] = '1'

        def apply_break(indices, start_range, slots):
            cands = [i for i in indices if start_range[0] <= df_h.at[i, 'Start_H'] <= start_range[1]]
            for i, idx in enumerate(cands):
                if h == slots[i % len(slots)]:
                    df_h.at[idx, 'Tarea'] = 'C'; df_h.at[idx, 'Counter'] = 'Casino'
        
        apply_break(idx_ag, (0, 11), [13, 14]) 
        apply_break(idx_ag, (12, 23), [2, 3]) 
        apply_break(idx_an, (0, 11), [13, 14, 15])
        apply_break(idx_an, (12, 23), [2, 3])

        # B. Detectar Quiebres Agentes
        active_counts = {c: 0 for c in main_counters_aire + main_counters_tierra}
        donors = [] 
        for idx in idx_ag:
            if df_h.at[idx, 'Tarea'] == '1':
                c = df_h.at[idx, 'Counter']
                if c in active_counts:
                    active_counts[c] += 1
                    donors.append(idx)
        
        empty = [c for c, count in active_counts.items() if count == 0]
        
        # C. Cobertura (Jerarquía)
        for target_cnt in empty:
            covered = False
            
            # 1. Agente Flotante
            possible = []
            for d_idx in donors:
                orig = df_h.at[d_idx, 'Counter']
                if active_counts.get(orig, 0) > 1:
                    d_nm = df_h.at[d_idx, 'Nombre']
                    if "AIRE" in target_cnt and (d_nm in no_tica_list): continue
                    possible.append(d_idx)
            if possible:
                best = possible[0]
                df_h.at[best, 'Tarea'] = f"3: Cubrir {target_cnt}"
                df_h.at[best, 'Counter'] = target_cnt
                active_counts[daily_assignments[(df_h.at[best, 'Nombre'], d)]] -= 1
                active_counts[target_cnt] += 1
                donors.remove(best)
                covered = True
            
            # 2. Coordinador
            if not covered:
                # Solo si queda al menos 1 coord en Tarea 1
                avail_c = [i for i in idx_co if df_h.at[i, 'Tarea'] != 'C']
                
                # Contamos cuantos hay en T1 ahora mismo
                current_t1_count = len([i for i in avail_c if df_h.at[i, 'Tarea'] == '1'])
                
                candidate = None
                # Prioridad: Tomar de Tarea 2 para no afectar T1
                t2_candidates = [i for i in avail_c if df_h.at[i, 'Tarea'] == '2']
                t1_candidates = [i for i in avail_c if df_h.at[i, 'Tarea'] == '1']
                
                if t2_candidates:
                    candidate = t2_candidates[0]
                elif t1_candidates and current_t1_count > 1: # Solo si sobran en T1
                    candidate = t1_candidates[0]
                
                if candidate is not None:
                    df_h.at[candidate, 'Tarea'] = f"4: Cubrir {target_cnt}"
                    df_h.at[candidate, 'Counter'] = target_cnt
                    covered = True
            
            # 3. Anfitrión
            if not covered and idx_an:
                avail_a = [i for i in idx_an if df_h.at[i, 'Tarea'] != 'C']
                # REGLA ANFITRION: Solo cubre si hay al menos 2 activos (para que queden 2)
                # O sea, necesitamos al menos 3 disponibles para tomar 1.
                # Si la regla es "quedan 2", entonces len(avail_a) > 2.
                # Si la regla es "hay 2, tomo 1", entonces len(avail_a) >= 2.
                # El usuario dijo: "solo en caso que existan 2 anfitriones activos... uno pasaría".
                # Interpretación: Si hay 2, tomo 1 y queda 1.
                # Corrección: "Debe haber a lo menos un anfitrión siempre en Losa Nacional y uno en Int".
                # Eso significa que SIEMPRE deben quedar 2 (1 Nac + 1 Int).
                # Por lo tanto, necesito tener 3 para usar 1.
                
                if len(avail_a) > 2:
                    ix = avail_a[0]
                    df_h.at[ix, 'Tarea'] = f"4: Cubrir {target_cnt}"
                    df_h.at[ix, 'Counter'] = target_cnt
                    covered = True
            
            # 4. HHEE Counter
            if not covered:
                hhee_counters.append({'Fecha': d, 'Hora': h, 'Counter': target_cnt})

        # D. Chequeo Final de Reglas de Oro (Para HHEE Generales)
        
        # HHEE COORDINADORES: Al menos 1 en Tarea 1
        active_coords_t1 = [i for i in idx_co if df_h.at[i, 'Tarea'] == '1']
        if len(active_coords_t1) == 0:
            hhee_coord.append({'Fecha': d, 'Hora': h, 'Counter': 'Supervisión'})

        # HHEE ANFITRIONES: Al menos 2 en Losa (1 Nac + 1 Int)
        # Los que no están comiendo ni cubriendo (Tarea 1)
        active_anf_t1 = [i for i in idx_an if df_h.at[i, 'Tarea'] == '1']
        
        # Asignar lugares a los que quedan
        for i, idx in enumerate(active_anf_t1):
            df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'
            
        if len(active_anf_t1) < 2:
            hhee_anf.append({'Fecha': d, 'Hora': h, 'Counter': 'Losa Minima'})

        # E. Finalizar resto
        for idx in idx_co:
            if df_h.at[idx, 'Tarea'] == '1': df_h.at[idx, 'Counter'] = 'General'
        for idx in idx_su:
            df_h.at[idx, 'Counter'] = 'General'; df_h.at[idx, 'Tarea'] = '1'

    # --- INSERTAR FILAS HHEE ---
    unique_dates = df_h['Fecha'].unique()
    
    # 1. HHEE Counters
    hhee_set = set((x['Fecha'], x['Hora'], x['Counter']) for x in hhee_counters)
    for cnt in ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]:
        base_row = {'Nombre': f"HHEE {cnt}", 'Rol': 'HHEE', 'Sub_Group': 'Counter', 'Role_Rank': 900, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Counter': cnt}
        for d in unique_dates:
            for h in range(24):
                task_val = "HHEE" if (d, h, cnt) in hhee_set else ""
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)

    # 2. HHEE Coordinación
    hhee_c_set = set((x['Fecha'], x['Hora']) for x in hhee_coord)
    base_row = {'Nombre': "HHEE COORDINACIÓN", 'Rol': 'HHEE', 'Sub_Group': 'Supervisión', 'Role_Rank': 901, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Counter': 'General'}
    for d in unique_dates:
        for h in range(24):
            task_val = "HHEE" if (d, h) in hhee_c_set else ""
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
        df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)

    # 3. HHEE Anfitriones
    hhee_a_set = set((x['Fecha'], x['Hora']) for x in hhee_anf)
    base_row = {'Nombre': "HHEE ANFITRIONES", 'Rol': 'HHEE', 'Sub_Group': 'Losa', 'Role_Rank': 902, 'Turno_Raw': 'Demanda', 'Start_H': -1, 'Counter': 'Losa'}
    for d in unique_dates:
        for h in range(24):
            task_val = "HHEE" if (d, h) in hhee_a_set else ""
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': task_val}])], ignore_index=True)
        df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-'}])], ignore_index=True)

    return df_h

# --- EXCEL ---
def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana V15")
    
    f_head = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#44546A', 'font_color': 'white', 'align': 'center'})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'text_wrap': True})
    f_libre = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'bg_color': '#F2F2F2', 'font_color': '#808080'})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#BFBFBF', 'align': 'left', 'indent': 1})
    
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        '3': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center', 'font_size': 8, 'text_wrap': True}),
        '4': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'bold': True, 'font_size': 8, 'text_wrap': True}),
        'C': wb.add_format({'bg_color': '#C6E0B4', 'border': 1, 'align': 'center'}),
        'HHEE': wb.add_format({'bg_color': '#7030A0', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'}),
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
        for h in range(24): ws.write(1, col+2+h, h, f_head)
        d_map[d] = col
        col += 26
        
    df_sorted = df[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    
    row = 2
    curr_group = ""
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    
    for _, p in df_sorted.iterrows():
        n, r, grp = p['Nombre'], p['Rol'], p['Sub_Group']
        
        grp_label = f"{r.upper()}"
        if r == "Agente": grp_label += f" - {grp}"
        if r == "HHEE": grp_label = "REQUERIMIENTOS HHEE"
        
        if grp_label != curr_group:
            ws.merge_range(row, 0, row, col-1, grp_label, f_group)
            row += 1
            curr_group = grp_label
            
        ws.write(row, 0, n, f_base)
        ws.write(row, 1, r, f_base)
        
        for d in dates:
            if d not in d_map: continue
            c = d_map[d]
            
            subset = df[(df['Nombre']==n) & (df['Fecha']==d)]
            if subset.empty:
                 ws.write(row, c, "-", f_libre)
                 ws.write(row, c+1, "Libre", f_libre)
                 for h in range(24): ws.write(row, c+2+h, "", f_libre)
                 continue
            
            # Raw shift
            t_raw = subset.iloc[0]['Turno_Raw']
            # If HHEE row
            if r == "HHEE": t_raw = "Demanda"
            ws.write(row, c, str(t_raw), f_base)
            
            # Place
            active = subset[subset['Hora'] != -1]
            if active.empty and r != "HHEE":
                ws.write(row, c+1, "Libre", f_libre)
            else:
                try:
                    if r == "HHEE": ws.write(row, c+1, subset.iloc[0]['Counter'], f_base)
                    else: ws.write(row, c+1, active['Counter'].mode()[0], f_base)
                except: ws.write(row, c+1, "?", f_base)
            
            # Hours
            for h in range(24):
                try:
                    val = df_idx.loc[(n, d, h)]
                    if isinstance(val, pd.DataFrame): val = val.iloc[0]
                    task = str(val['Tarea'])
                    
                    style = f_base
                    if task.startswith('3'): style = st_map['3']
                    elif task.startswith('4'): style = st_map['4']
                    elif task == 'C': style = st_map['C']
                    elif task == '2': style = st_map['2']
                    elif task == 'HHEE': style = st_map['HHEE']
                    
                    ws.write(row, c+2+h, task, style)
                except: ws.write(row, c+2+h, "", f_libre)
        row += 1
    wb.close()
    return out

if st.button("🚀 Generar Planificación V15"):
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
                st.download_button("📥 Descargar Excel", make_excel(final), f"Planificacion_V15.xlsx")
