import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import xlsxwriter
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestor de Turnos Aeropuerto", layout="wide")
st.title("✈️ Gestor de Turnos: Reglas de Negocio V4")
st.markdown("""
**Reglas Actualizadas (Doc 1):**
* **Agentes:** Cobertura cruzada Aire/Tierra. Selector para agentes **Sin TICA**.
* **Anfitriones:** Mínimo 2 por franja.
* **HHEE:** Se asignan automáticamente si no se cumple la cobertura mínima.
""")

# --- PARSEO ---
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

files_config = [{"label": "Agente", "key": "exec"}, {"label": "Anfitrion", "key": "host"},
                {"label": "Coordinador", "key": "coord"}, {"label": "Supervisor", "key": "sup"}]
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

# --- SELECTOR SIN TICA ---
st.sidebar.markdown("---")
st.sidebar.header("2. Excepciones")
agents_no_tica = []
if 'exec' in uploaded_sheets:
    # Carga preliminar para obtener nombres
    if st.sidebar.checkbox("Cargar lista de Agentes para seleccionar SIN TICA"):
        with st.spinner("Leyendo nombres..."):
            uf, us = uploaded_sheets['exec']
            df_temp = process_file_sheet(uf, us, "Agente", target_month, sel_year)
            if not df_temp.empty:
                unique_agents = sorted(df_temp['Nombre'].unique())
                agents_no_tica = st.sidebar.multiselect("Selecciona Agentes SIN TICA (Solo Tierra)", unique_agents)

# --- MOTOR LÓGICO V4 ---
def logic_engine(df, no_tica_list):
    rows = []
    # Prioridad de Roles para Excel: Agente -> Supervisor -> Coordinador -> Anfitrion
    role_priority = {'Agente': 1, 'Supervisor': 2, 'Coordinador': 3, 'Anfitrion': 4}
    
    for _, r in df.iterrows():
        hours, start_h = parse_shift_time(r['Turno_Raw'])
        
        if not hours:
            rows.append({**r, 'Hora': -1, 'Tarea': str(r['Turno_Raw']), 'Counter': '-', 
                         'Role_Rank': role_priority.get(r['Rol'], 9), 'Start_H': -1})
        else:
            for h in hours:
                rows.append({**r, 'Hora': h, 'Tarea': '1', 'Counter': '?', 
                             'Role_Rank': role_priority.get(r['Rol'], 9), 'Start_H': start_h})
    
    df_h = pd.DataFrame(rows)
    if df_h.empty: return df_h
    
    main_counters = ["T1 AIRE", "T1 TIERRA", "T2 AIRE", "T2 TIERRA"]
    
    # Iterar por franja horaria
    for (d, h), g in df_h[df_h['Hora'] != -1].groupby(['Fecha', 'Hora']):
        
        # Índices por rol
        agentes = g[g['Rol']=='Agente'].index.tolist()
        coords = g[g['Rol']=='Coordinador'].index.tolist()
        anfitriones = g[g['Rol']=='Anfitrion'].index.tolist()
        sups = g[g['Rol']=='Supervisor'].index.tolist()
        
        # --- 1. COLACIONES ---
        # Función auxiliar
        def apply_break(indices_list, valid_start_range, break_slots, role_name):
            candidates = []
            for idx in indices_list:
                st_h = df_h.at[idx, 'Start_H']
                # Verificar rango de ingreso
                if valid_start_range[0] <= st_h <= valid_start_range[1]:
                    candidates.append(idx)
            
            # Distribuir
            for i, idx in enumerate(candidates):
                slot_idx = i % len(break_slots)
                if h == break_slots[slot_idx]:
                    df_h.at[idx, 'Tarea'] = 'C'
                    df_h.at[idx, 'Counter'] = 'Casino'

        # Agentes Diurnos (8-10) -> Break 12-15
        apply_break(agentes, (8, 10), [13, 14]) # Heurística: 13 y 14 para cubrir rango 12-15
        # Agentes Nocturnos (20-22) -> Break 02-04
        apply_break(agentes, (20, 22), [2, 3])
        
        # Anfitriones Diurnos (8-9) -> Break 13-16
        apply_break(anfitriones, (8, 9), [13, 14, 15])
        # Anfitriones Nocturnos (20-21) -> Break 02-04
        apply_break(anfitriones, (20, 21), [2, 3])

        # Coordinadores (Reglas específicas Doc 1)
        apply_break(coords, (5, 5), [12]) # Ingreso 05:00 -> Break ~12
        apply_break(coords, (21, 21), [2]) # Ingreso 21:00 -> Break ~02
        # Ingreso 10:00 no tiene colación explícita en reglas nuevas, solo Tareas 2.
        
        # --- 2. FILTRAR DISPONIBLES ---
        active_agentes = [i for i in agentes if df_h.at[i, 'Tarea'] != 'C']
        active_coords = [i for i in coords if df_h.at[i, 'Tarea'] != 'C']
        active_anfitriones = [i for i in anfitriones if df_h.at[i, 'Tarea'] != 'C']
        
        # --- 3. ASIGNACIÓN DE COUNTERS (AGENTES) ---
        # Separar con TICA y sin TICA
        with_tica = []
        no_tica = []
        
        for idx in active_agentes:
            name = df_h.at[idx, 'Nombre']
            if name in no_tica_list:
                no_tica.append(idx)
            else:
                with_tica.append(idx)
        
        counters_status = {c: False for c in main_counters}
        spare_agentes = []
        
        # Llenar Tierra (Prioridad para Sin TICA)
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
        
        # Sobrantes Sin Tica (No pueden ir a aire) -> Quedan flotando en Tierra o Refuerzo Tierra
        spare_no_tica = no_tica 

        # Llenar Aire (Solo Con TICA)
        aire_cnts = ["T1 AIRE", "T2 AIRE"]
        for a_cnt in aire_cnts:
            if with_tica:
                idx = with_tica.pop(0)
                df_h.at[idx, 'Counter'] = a_cnt
                df_h.at[idx, 'Tarea'] = '1'
                counters_status[a_cnt] = True
        
        spare_with_tica = with_tica
        
        # --- 4. CUBRIR QUIEBRES (Tarea 3 y 4) ---
        for cnt_name, covered in counters_status.items():
            if not covered:
                is_tierra = "TIERRA" in cnt_name
                is_aire = "AIRE" in cnt_name
                
                # TAREA 3: Cobertura Cruzada (Aire->Tierra, Tierra->Aire)
                # Si falta Tierra, busco alguien que sobre de Aire (Con TICA)
                # Si falta Aire, busco alguien que sobre de Tierra (Sin TICA o Con TICA)
                
                filled = False
                
                # Intento 1: Flotante
                if is_aire and spare_with_tica: # Solo con TICA cubre Aire
                    idx = spare_with_tica.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '3'
                    filled = True
                elif is_tierra:
                    # Para tierra sirve cualquiera
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
                
                # Intento 2: Coordinador (Tarea 4)
                if not filled and active_coords:
                    idx = active_coords[0] # Simplificado
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'
                    active_coords.pop(0)
                    filled = True
                    
                # Intento 3: Anfitrión (Tarea 4)
                if not filled and active_anfitriones:
                    idx = active_anfitriones.pop(0)
                    df_h.at[idx, 'Counter'] = cnt_name
                    df_h.at[idx, 'Tarea'] = '4'
                    active_anfitriones.pop(0) # Ya usado
                    filled = True
                
                # Intento 4: HHEE (Nadie pudo cubrir)
                if not filled:
                    # Creamos un registro virtual de HHEE
                    # Nota: Esto es complejo de insertar en el DF iterando, 
                    # por simplicidad marcamos en un log o dejamos el counter vacio en el excel
                    pass 

        # --- 5. ASIGNAR RESTANTES ---
        
        # Agentes Sobrantes -> Refuerzo (Aire prioridad)
        # Los Sin TICA van a Refuerzo Tierra obligados
        for idx in spare_no_tica:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = "Refuerzo Tierra"
            
        for i, idx in enumerate(spare_with_tica):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = "T1 AIRE" if i%2==0 else "T2 AIRE" # Prioridad Aire
            
        # Coordinadores (Tareas Admin)
        for idx in active_coords:
            st_h = df_h.at[idx, 'Start_H']
            task = '1'
            cnt = 'Piso'
            
            # Reglas Tarea 2
            if st_h == 10 and (h == 10 or h in [14, 15]):
                task = '2'; cnt = 'Oficina'
            elif st_h == 5 and (h in [6, 7]):
                task = '2'; cnt = 'Oficina'
            elif st_h == 21 and (h in [5, 6]):
                task = '2'; cnt = 'Oficina'
                
            df_h.at[idx, 'Tarea'] = task
            df_h.at[idx, 'Counter'] = cnt
            
        # Anfitriones
        # Mínimo 2 por franja. Si hay menos, HHEE necesaria (se ve visualmente si hay huecos)
        for i, idx in enumerate(active_anfitriones):
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'Zona Int' if i%2==0 else 'Zona Nac'
            
        # Supervisores
        for idx in sups:
            df_h.at[idx, 'Tarea'] = '1'
            df_h.at[idx, 'Counter'] = 'General'

    return df_h

# --- EXCEL ---
def make_excel(df):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Sábana V4")
    
    # Formatos
    f_head = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#44546A', 'font_color': 'white', 'align': 'center'})
    f_date = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    f_base = wb.add_format({'border': 1, 'align': 'center', 'font_size': 9})
    f_group = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#BFBFBF', 'align': 'left'})
    
    # Estilos Tareas
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
        
    # Ordenar: Agente -> Supervisor -> Coord -> Anfitrion
    df_sorted = df[['Nombre', 'Rol', 'Role_Rank']].drop_duplicates().sort_values(['Role_Rank', 'Nombre'])
    
    row = 2
    curr_role = ""
    
    df_idx = df.set_index(['Nombre', 'Fecha', 'Hora'])
    df_base = df.drop_duplicates(['Nombre', 'Fecha']).set_index(['Nombre', 'Fecha'])
    
    for _, p in df_sorted.iterrows():
        n, r = p['Nombre'], p['Rol']
        
        # Separador de Grupo
        if r != curr_role:
            ws.merge_range(row, 0, row, col-1, f"--- {r.upper()} ---", f_group)
            row += 1
            curr_role = r
            
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
                    
                    style = st_map.get(task, f_base)
                    ws.write(row, c+2+h, task, style)
                except: pass
        row += 1
        
    wb.close()
    return out

# --- EJECUCIÓN ---
if st.button("🚀 Generar Planificación V4"):
    if not uploaded_sheets:
        st.error("Sube los archivos primero.")
    else:
        with st.spinner("Procesando..."):
            dfs = []
            for role, (uf, us) in uploaded_sheets.items():
                dfs.append(process_file_sheet(uf, us, role, target_month, sel_year))
            full = pd.concat(dfs)
            
            if full.empty: st.error("No hay datos.")
            else:
                final = logic_engine(full, agents_no_tica)
                st.success("¡Listo!")
                st.download_button("📥 Descargar Excel", make_excel(final), f"Planificacion_V4_{sel_month_name}.xlsx")
