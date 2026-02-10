import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Generador de Sábana de Turnos (Reglas V2)")
st.markdown("Genera la planificación base aplicando las reglas de negocio estrictas (TICA, Horarios Coordinadores, Colaciones dinámicas).")

# --- PARSEO DE TURNOS ---
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
        
        # Generar rango de horas
        if start_h < end_h: hours = list(range(start_h, end_h))
        elif start_h > end_h: hours = list(range(start_h, 24)) + list(range(0, end_h))
        else: hours = [start_h]
        
        return hours, start_h
    except: return [], None

# --- LECTURA ROBUSTA ---
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
        
        # Buscar columna nombre
        name_col = df.columns[0]
        for col in df.columns:
            if "nombre" in str(col).lower() or "cargo" in str(col).lower() or "supervisor" in str(col).lower():
                name_col = col
                break
        
        # Mapear fechas
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

# --- UI ---
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

# --- MOTOR DE REGLAS V2 ---
def logic_engine(df):
    # 1. Expandir a Horas y calcular Hora Ingreso
    rows = []
    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        
        # Determinar si es Diurno/Nocturno basado en hora inicio
        # Diurno: Ingreso 05:00 - 14:00. Nocturno: 15:00 - 04:00
        sh_type = "Diurno"
        if start_h is not None and (start_h >= 20 or start_h < 5): 
            sh_type = "Nocturno"
        elif start_h == 5:
            sh_type = "Madrugada" # Especial Coordinadores

        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 'Tipo': sh_type, 'Start_H': -1})
        else:
            for h in hours:
                # Inicializar
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 'Tipo': sh_type, 'Start_H': start_h})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    # 2. Asignación Hora por Hora
    main_counters = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    # Iterar por bloques horarios
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        execs_idx = g[g['Rol']=='Ejecutivo'].index.tolist()
        coords_idx = g[g['Rol']=='Coordinador'].index.tolist()
        hosts_idx = g[g['Rol']=='Anfitriones'].index.tolist()
        sups_idx = g[g['Rol']=='Supervisor'].index.tolist()
        
        # --- REGLA 1: EJECUTIVOS Y COLACIÓN ---
        avail_execs = []
        for idx in execs_idx:
            start_h = df_h.at[idx, 'Start_H']
            is_col = False
            
            # REGLA COLACION 
            # Diurnos (ingreso 8-10): Colación 12-15
            if 8 <= start_h <= 10:
                if h in [13, 14]: # Rango heurístico dentro de 12-15
                    if hash(df_h.at[idx, 'Nombre']) % 2 == (h % 2): is_col = True
            
            # Nocturnos (ingreso 20-22): Colación 02-04
            elif (20 <= start_h <= 22):
                if h in [2, 3]: # Rango exacto 02:00-03:00 o 03:00-04:00
                    if hash(df_h.at[idx, 'Nombre']) % 2 == (h % 2): is_col = True

            if is_col:
                df_h.at[idx, 'Tarea'] = 'C'
                df_h.at[idx, 'Counter'] = 'Casino'
            else:
                avail_execs.append(idx)
        
        # --- REGLA 2: EJECUTIVOS A COUNTERS (SIN TICA vs CON TICA) ---
        # Nota: Como no tenemos dato TICA, asumimos todos tienen.
        # Si tuviéramos columna TICA, filtraríamos aquí.
        
        spare_execs = []
        
        # Llenar Counters (Min 1 por counter) 
        for i, idx in enumerate(avail_execs):
            if i < 4:
                cnt = main_counters[i]
                df_h.at[idx, 'Counter'] = cnt
                df_h.at[idx, 'Tarea'] = '1'
            else:
                spare_execs.append(idx)
        
        # --- REGLA 3: QUIEBRES Y TAREA 3 (FLOTANTE) ---
        counters_status = {c: 0 for c in main_counters}
        # Verificar cobertura real
        for idx in avail_execs:
            c_assigned = df_h.at[idx, 'Counter']
            if c_assigned in counters_status: counters_status[c_assigned] = 1
            
        for cnt_name, status in counters_status.items():
            if status == 0: # Quiebre
                covered = False
                
                # REGLA: Tarea 3 Flotante (Aire->Aire, Tierra->Tierra) 
                if spare_execs:
                    best_candidate = -1
                    # Buscamos si alguno "sobrante" sería mejor por afinidad (Lógica simplificada aquí)
                    best_candidate = spare_execs.pop(0) 
                    
                    df_h.at[best_candidate, 'Counter'] = cnt_name
                    df_h.at[best_candidate, 'Tarea'] = '3'
                    covered = True
                
                # REGLA: Tarea 4 (Coordinador) 
                if not covered and coords_idx:
                    # Solo si no está ocupado en Tarea 2 crítica
                    coord_candidato = coords_idx[0] 
                    # (En una versión más avanzada verificaríamos si el coord está en su hora de Tarea 2 obligatoria)
                    df_h.at[coord_candidato, 'Counter'] = cnt_name
                    df_h.at[coord_candidato, 'Tarea'] = '4'
                    coords_idx.pop(0) # Ya lo usamos
                    covered = True
                    
                # REGLA: Tarea 4 (Anfitrión) 
                if not covered and hosts_idx:
                    host_candidato = hosts_idx.pop(0)
                    df_h.at[host_candidato, 'Counter'] = cnt_name
                    df_h.at[host_candidato, 'Tarea'] = '4'
        
        # --- REGLA 4: EXCESOS (EJECUTIVOS) ---
        # Priorizar Aire 
        for i, idx in enumerate(spare_execs):
            df_h.at[idx, 'Tarea'] = '1'
            # Alternar entre T1 Aire y T2 Aire
            df_h.at[idx, 'Counter'] = "T1 AIRE" if i % 2 == 0 else "T2 AIRE"
            
        # --- REGLA 5: COORDINADORES (ADMINISTRACION DE TAREAS) ---
        for idx in coords_idx:
            start_h = df_h.at[idx, 'Start_H']
            task_assigned = '1'
            counter_assigned = 'Superv. Piso'
            
            # REGLA DE INGRESOS 
            
            # Ingreso 10:00 -> Tarea 2 (10-11) y (14-16)
            if start_h == 10:
                if h == 10: 
                    task_assigned = '2'; counter_assigned = 'Oficina'
                elif h in [14, 15]: # Bloque tarde
                    task_assigned = '2'; counter_assigned = 'Oficina'
                    
            # Ingreso 05:00 -> Colación 11-14 (aprox) + 2 bloques Tarea 2
            elif start_h == 5:
                if h in [12]: # Colacion medio dia
                    task_assigned = 'C'; counter_assigned = 'Casino'
                elif h in [6, 7]: # Tarea 2 temprano
                    task_assigned = '2'; counter_assigned = 'Oficina'
            
            # Ingreso 21:00 -> Colación + Tarea 2 en 05-08
            elif start_h >= 21:
                if h == 2: # Colacion noche
                     task_assigned = 'C'; counter_assigned = 'Casino'
                elif h in [5, 6]: # Tarea 2 madrugada
                     task_assigned = '2'; counter_assigned = 'Oficina'

            df_h.at[idx, 'Tarea'] = task_assigned
            df_h.at[idx, 'Counter'] = counter_assigned

        # --- REGLA 6: ANFITRIONES ---
        # 
        for i, idx in enumerate(hosts_idx):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'

        # --- REGLA 7: SUPERVISORES ---
        # Siempre Tarea 1 
        for idx in sups_idx:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'

    return df_h

# --- EXCEL (Mismo formato previo) ---
def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    
    ws = wb.add_worksheet("Sábana Turnos")
    fmt_h = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#404040', 'font_color': 'white', 'align': 'center'})
    fmt_d = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
    fmt_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    
    fmt_t2 = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'}) 
    fmt_t3 = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#DDEBF7'}) 
    fmt_t4 = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#F4CCCC', 'font_color': '#990000', 'bold': True})
    fmt_col = wb.add_format({'border': 1, 'align': 'center', 'bg_color': '#E2EFDA', 'font_color': '#38761D', 'bold': True})

    ws.write(1, 0, "Colaborador", fmt_h)
    ws.write(1, 1, "Rol", fmt_h)
    ws.freeze_panes(2, 2) 
    
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
        ws.write(row, 0, n, fmt_base)
        ws.write(row, 1, r, fmt_base)
        
        for d in dates:
            if d not in d_map: continue
            c = d_map[d]
            try:
                t_raw = df_base.loc[(n, d), 'Turno_Raw']
                ws.write(row, c, str(t_raw), fmt_base)
            except: ws.write(row, c, "-", fmt_base)
            
            for h in range(24):
                try:
                    val = df_idx.loc[(n, d, h)]
                    if isinstance(val, pd.DataFrame): val = val.iloc[0]
                    
                    task = val['Tarea']
                    cnt = val['Counter']
                    
                    if h == 12: ws.write(row, c+1, cnt if cnt!='?' else '-', fmt_base)
                    
                    style = fmt_base
                    if task == '2': style = fmt_t2
                    if task == '3': style = fmt_t3
                    if task == '4': style = fmt_t4
                    if task == 'C': style = fmt_col
                    
                    ws.write(row, c+2+h, task, style)
                except: pass
        row += 1
        
    ws2 = wb.add_worksheet("Bitácora Incidencias")
    headers = ["Fecha", "Colaborador", "Tipo Incidencia", "Hora Inicio", "Hora Fin", "Comentario Supervisor", "Aplica HHEE"]
    for i, h in enumerate(headers): ws2.write(0, i, h, fmt_h)
    ws2.set_column(0, 6, 20)
        
    wb.close()
    return out

# --- EJECUCIÓN ---
if st.button("🚀 Generar Sábana Maestra (V2)"):
    if not uploaded_sheets:
        st.error("Carga los archivos primero.")
    else:
        with st.spinner("Aplicando Reglas de Negocio..."):
            dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                dfs.append(process_file_sheet(uf, us, role, target_month, sel_year))
            full = pd.concat(dfs)
            if full.empty: st.error("No hay datos.")
            else:
                final = logic_engine(full)
                st.success(f"Planificación Completa. {len(final['Nombre'].unique())} colaboradores.")
                st.download_button("📥 Descargar Sábana", make_excel(final), f"Planificacion_ReglasV2_{sel_month_name}.xlsx")
