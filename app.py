import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto (V71 - Básico)", layout="wide")
st.title("✈️ Gestor de Turnos: V71 (Asignación Manual)")
st.markdown("""
**Versión Esencial:**
1. **Sábana Limpia:** El sistema procesa los turnos y coloca los bloques de horas (Tarea `1`).
2. **Colaciones Automáticas:** Asigna la tarea `C` basada en la hora de ingreso para ahorrar tiempo.
3. **Control del Supervisor:** La asignación de Counters y Zonas ("Por Asignar") se hace directamente en el Excel.
4. **Cero Automatización de Cobertura:** No se mueven agentes ni se generan HHEE automáticamente. Todo es control manual.
""")

# --- INICIALIZACIÓN ---
if 'incidencias' not in st.session_state:
    st.session_state.incidencias = []

today = datetime.now()
uploaded_sheets = {}
start_d = None
end_d = None

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

# --- MOTOR LÓGICO BÁSICO (V71) ---
def logic_engine_basic(df):
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

    # Expandir turnos y asociar fecha de origen
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
                    'Base_Diaria': 'Por Asignar', 'Shift_Date': shift_date 
                })
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h, raw_shifts_map

    # PROCESAR HORA A HORA (SOLO ASIGNACIÓN DE COLACIONES Y TAREA 1)
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        for idx in g.index:
            rol = df_h.at[idx, 'Rol']
            sh = df_h.at[idx, 'Start_H'] # Aquí definimos la variable "sh"
            
            # Asignar tarea 1 básica
            if rol in ['Coordinador', 'Supervisor']:
                df_h.at[idx, 'Base_Diaria'] = 'General'
                # Lógica fija supervisores/coords
                nm = df_h.at[idx, 'Nombre']
                is_odd = hash(nm) % 2 != 0
                # CORRECCIÓN V71: Usar "sh" en lugar de "st_h"
                if sh == 10:
                    if h == 10: df_h.at[idx, 'Tarea'] = '2'
                    elif h in [14, 15]:
                        if (h == 14 and is_odd) or (h == 15 and not is_odd): df_h.at[idx, 'Tarea'] = 'C'
                        else: df_h.at[idx, 'Tarea'] = '2'
                elif sh == 5:
                    if h in [11, 12, 13]:
                        if h == 12: df_h.at[idx, 'Tarea'] = 'C'
                        else: df_h.at[idx, 'Tarea'] = '2'
                elif sh == 21:
                    if h in [5, 6, 7]:
                        if h == 6: df_h.at[idx, 'Tarea'] = 'C'
                        else: df_h.at[idx, 'Tarea'] = '2'
                else:
                    df_h.at[idx, 'Tarea'] = '1'
            else:
                df_h.at[idx, 'Base_Diaria'] = 'Por Asignar'
                df_h.at[idx, 'Tarea'] = '1'

                # Colaciones (Cálculo matemático)
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
                    df_h.at[idx, 'Tarea'] = 'C'

    if 'incidencias' in st.session_state and st.session_state.incidencias:
        df_h = apply_incidents(df_h, st.session_state.incidencias)

    # CREAR FILAS GENÉRICAS PARA HHEE MANUALES (Al final del Excel)
    unique_dates = sorted(df_h['Fecha'].unique())
    hhee_labels = ["HHEE Agente 1", "HHEE Agente 2", "HHEE Agente 3", "HHEE Anfitrion 1", "HHEE Coord/Sup 1"]
    
    for lbl in hhee_labels:
        base_row = {'Nombre': lbl, 'Rol': 'HHEE', 'Sub_Group': 'Asignación Manual', 'Role_Rank': 900, 'Turno_Raw': 'Manual', 'Start_H': -1, 'Base_Diaria': 'Por Asignar'}
        for d in unique_dates:
            for h in range(24):
                df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': h, 'Tarea': '', 'Shift_Date': d}])], ignore_index=True)
            df_h = pd.concat([df_h, pd.DataFrame([{**base_row, 'Fecha': d, 'Hora': -1, 'Tarea': '-', 'Shift_Date': d}])], ignore_index=True)

    return df_h, raw_shifts_map

# --- EXCEL GENERATOR (V71) ---
def make_excel(df, raw_shifts_map, start_d, end_d):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    
    f_cabify = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#7145D6', 'font_color': 'white', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9, 'text_wrap': True})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#F3F3F3', 'align': 'center'})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#EFEFEF', 'align': 'left', 'indent': 1})
    f_alert = wb.add_format({'bg_color': '#EA9999', 'font_color': '#980000', 'bold': True, 'border': 1, 'align': 'center'})
    f_header_count = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#B4A7D6', 'align': 'center'})
    
    f_sep_col = wb.add_format({'bg_color': '#1C1C1C', 'border': 1, 'align': 'center'}) 
    
    # Colores base
    st_map = {
        '2': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}),
        'C': wb.add_format({'bg_color': '#C6E0B4', 'border': 1, 'align': 'center'}),
        'T1 AIRE': wb.add_format({'bg_color': '#BDD7EE', 'border': 1, 'align': 'center', 'font_size': 8}),
        'T1 TIERRA': wb.add_format({'bg_color': '#F8CBAD', 'border': 1, 'align': 'center', 'font_size': 8}),
        'T2 AIRE': wb.add_format({'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_size': 8}),
        'T2 TIERRA': wb.add_format({'bg_color': '#FCE4D6', 'border': 1, 'align': 'center', 'font_size': 8}),
    }
    
    all_dates = sorted(df['Fecha'].unique())
    dates = [d for d in all_dates if start_d <= d.date() <= end_d]
    df_staff = df[df['Rol'] != 'HHEE'].drop_duplicates(subset=['Nombre', 'Fecha'])
    df_hhee = df[df['Rol'] == 'HHEE'].drop_duplicates(subset=['Nombre', 'Fecha'])
    
    df_staff_sorted = df_staff[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    df_hhee_sorted = df_hhee[['Nombre', 'Rol', 'Sub_Group', 'Role_Rank']].drop_duplicates().sort_values(['Nombre'])
    
    # ------------------------------------------------
    # HOJAS OCULTAS DE VALIDACIÓN
    # ------------------------------------------------
    ws_data = wb.add_worksheet("Datos_Validacion")
    unique_roles = sorted([str(x) for x in df['Rol'].unique() if x != 'HHEE'])
    unique_names = sorted([str(x) for x in df['Nombre'].unique() if 'HHEE' not in str(x)])
    
    opciones_grilla = ['1', '2', 'C', 'T1 AIRE', 'T1 TIERRA', 'T2 AIRE', 'T2 TIERRA', 'Zona Int', 'Zona Nac', 'General']
    opciones_lugar = ['Por Asignar', 'T1 AIRE', 'T1 TIERRA', 'T2 AIRE', 'T2 TIERRA', 'Zona Int', 'Zona Nac', 'General']
    
    ws_data.write_column(1, 0, unique_roles)
    ws_data.write_column(1, 1, unique_names)
    ws_data.write_column(1, 2, opciones_grilla)
    ws_data.write_column(1, 3, opciones_lugar)
    ws_data.hide()
    
    name_range = f"Datos_Validacion!$B$2:$B${len(unique_names)+1}"
    role_range = f"Datos_Validacion!$A$2:$A${len(unique_roles)+1}"
    grilla_range = f"Datos_Validacion!$C$2:$C${len(opciones_grilla)+1}"
    lugar_range = f"Datos_Validacion!$D$2:$D${len(opciones_lugar)+1}"

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

    # ------------------------------------------------
    # HOJA 2: BITÁCORA
    # ------------------------------------------------
    ws_bit = wb.add_worksheet("Bitacora_Incidencias")
    headers_bit = ["Tipo Colaborador", "Nombre Colaborador", "Fecha (YYYY-MM-DD)", "Tipo Incidencia", "Hora Inicio (0-23)", "Hora Fin (0-23)"]
    for i, h in enumerate(headers_bit): ws_bit.write(0, i, h, f_cabify)
    ws_bit.data_validation('A2:A1000', {'validate': 'list', 'source': role_range})
    ws_bit.data_validation('B2:B1000', {'validate': 'list', 'source': name_range})
    ws_bit.data_validation('D2:D1000', {'validate': 'list', 'source': ['Inasistencia', 'Atraso', 'Salida Anticipada']})
    ws_bit.write(0, 7, "GUÍA OPERATIVA V71:", f_cabify)
    ws_bit.write(1, 7, "INASISTENCIA: Marque el día de inicio del turno. Borrará automáticamente la madrugada siguiente.")

    # ------------------------------------------------
    # HOJA 3: PLAN OPERATIVO
    # ------------------------------------------------
    ws_real = wb.add_worksheet("Plan_Operativo")
    ws_real.write(5, 0, "Colaborador", f_cabify)
    ws_real.write(5, 1, "Rol", f_cabify)
    ws_real.freeze_panes(6, 2)
    
    col = 2
    d_map = {}
    
    ws_real.write(2, 0, "DOTACIÓN ACTIVA", f_header_count)
    
    días_es = {0: 'Lun', 1: 'Mar', 2: 'Mié', 3: 'Jue', 4: 'Vie', 5: 'Sáb', 6: 'Dom'}
    
    for d in dates:
        d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
        d_str = f"{días_es[d.weekday()]} {pd.to_datetime(d).strftime('%d-%b')}"
        
        ws_real.set_column(col, col, 1.5)
        ws_real.write(5, col, "", f_sep_col)
        
        ws_real.merge_range(0, col+1, 0, col+26, d_str, f_date)
        ws_real.write(5, col+1, "Turno", f_cabify)
        ws_real.write(5, col+2, "Lugar (Supervisor)", f_cabify)
        
        d_map[d_iso] = col + 1
        
        for h in range(24):
            ws_real.write(5, col+3+h, h, f_cabify)
            col_idx = col+3+h
            col_let = xlsxwriter.utility.xl_col_to_name(col_idx)
            
            f_ag = f'=COUNTIFS($B$7:$B$1000,"<>HHEE",{col_let}7:{col_let}1000,"<>FALTA",{col_let}7:{col_let}1000,"<>Libre",{col_let}7:{col_let}1000,"<>*Ausente*", {col_let}7:{col_let}1000,"?*")'
            ws_real.write_formula(2, col_idx, f_ag, f_header_count)

        col += 27

    row = 6
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
                
                ws_real.write(row, c_start+1, str(lugar), f_base)
                ws_real.data_validation(row, c_start+1, row, c_start+1, {'validate': 'list', 'source': lugar_range, 'show_error': False})
            
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
                    ws_real.data_validation(row, c_start+2+h, row, c_start+2+h, {'validate': 'list', 'source': grilla_range, 'show_error': False})
        row += 1
    
    ws_real.merge_range(row, 0, row, col-1, "HHEE Y COBERTURAS MANUALES", f_group)
    row += 1
    
    for _, p in df_hhee_sorted.iterrows():
        n = p['Nombre']
        r = p['Rol']
        ws_real.data_validation(row, 0, row, 0, {'validate': 'list', 'source': name_range, 'show_error': False})
        ws_real.write(row, 0, n, f_base)
        ws_real.write(row, 1, r, f_base)
        
        for d in dates:
            d_iso = pd.to_datetime(d).strftime("%Y-%m-%d")
            c_start = d_map[d_iso]
            ws_real.write(row, c_start-1, "", f_sep_col)
            ws_real.write(row, c_start, "Manual", f_base)
            ws_real.write(row, c_start+1, "Por Asignar", f_base)
            ws_real.data_validation(row, c_start+1, row, c_start+1, {'validate': 'list', 'source': lugar_range, 'show_error': False})
            
            for h in range(24):
                ws_real.write(row, c_start+2+h, "", f_base)
                ws_real.data_validation(row, c_start+2+h, row, c_start+2+h, {'validate': 'list', 'source': grilla_range, 'show_error': False})

        row += 1

    end_col_let = xlsxwriter.utility.xl_col_to_name(col-1)
    data_range = f"D7:{end_col_let}{row}"
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"FALTA"', 'format': f_alert})
    ws_real.conditional_format(data_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"INCIDENCIA"', 'format': f_alert})
    
    for k, fmt in st_map.items():
        criteria = 'begins with' if len(k) > 2 else 'equal to'
        val = k if len(k) > 2 else f'"{k}"'
        ws_real.conditional_format(data_range, {'type': 'text' if len(k)>2 else 'cell', 'criteria': criteria, 'value': val, 'format': fmt})

    wb.close()
    return out

st.sidebar.markdown("---")
if st.sidebar.button("🚀 Generar Planificación V71 (Manual)"):
    if not uploaded_sheets: st.error("Carga archivos.")
    elif not (start_d and end_d): st.error("Define fechas.")
    else:
        with st.spinner("Construyendo grilla base..."):
            dfs = []
            for role, (key) in [("Agente","exec"),("Coordinador","coord"),("Anfitrion","host"),("Supervisor","sup")]:
                if key in uploaded_sheets:
                    f, s = uploaded_sheets[key]
                    dfs.append(process_file_sheet(f, s, role, start_d, end_d))
            full = pd.concat(dfs)
            
            if full.empty: st.error("Sin datos válidos.")
            else:
                final, raw_map = logic_engine_basic(full)
                st.success("¡Listo! Descarga la Suite Operativa V71.")
                st.download_button("📥 Descargar Suite (V71)", make_excel(final, raw_map, start_d, end_d), f"Planificacion_Operativa_Manual.xlsx")
