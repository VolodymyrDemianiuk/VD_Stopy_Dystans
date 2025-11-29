import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ==========================================
# 1. LOGIKA BIZNESOWA (SKRYPT VD - ZAADAPTOWANY)
# ==========================================

# --- STAŁE ---
A_FRONT_TO_034 = 72.5     
B_034_TO_057   = 66.0     
C_058_TO_062   = 11.7     
CROSS_LANES    = 5.4      
START_STOP     = 9.0      
WRONG_ENTRY_PENALTY = 100000.0 
ENTRY_SIDE_SOFT_PENALTY = 0.0
U_TURN_PENALTY = 200.0         
NEIGHBOR_MIDDLE_PENALTY = 1.0 
BRIDGE_PENALTY = 0.0
FRONT = "front"           
P034  = "034"             
P057  = "057"             
P058  = "058"             
TRYBY_EXT_ALL = {680, 690}
TRYB_670_EXT  = {1022,1023,1024,1025,1026}
START_STOP_MAP = {
    602:(826,825), 601:(826,825), 610:(826,825),
    722:(826,825), 622:(826,825), 620:(826,825),
    630:(859,860), 640:(859,860), 641:(859,860),
    650:(859,860), 655:(859,860), 660:(859,860),
    670:(1020,1019), 680:(1020,1019), 690:(1020,1019),
}
AXIS_602_LANE   = 800
SS_TO_AXIS_602  = 70.2     
LENGTH_602_LOOP = 146.0    
NEEDED = {
 "Numer misji":["Numer misji","numer misji","misja","nr misji","id misji"],
 "Tryb Pracy":["Tryb Pracy","tryb pracy","tryb"],
 "SKU":["SKU","sku"],
 "numer lini":["numer lini","numer linii","linia","line"],
 "Regal":["Regal","Regał","regał","alejka","regal"],
 "Kolumna":["Kolumna","kolumna","kol"],
 "Poziom":["Poziom","poziom","level"],
 "miejsce":["miejsce","slot","miejsce pobrania"],
}

# --- FUNKCJE POMOCNICZE ---
def tryb_ma_rozszerzenie(tryb:int, lane:int=None)->bool:
    if tryb in TRYBY_EXT_ALL: return True
    if tryb == 670 and lane in TRYB_670_EXT: return True
    return False

def pair_key(reg:int)->int:
    reg=int(reg); return reg if reg % 2 == 1 else reg - 1

def normalize_rack(reg_num: int) -> int:
    if reg_num >= 901: return reg_num - 20
    return reg_num

def pair_gaps(a:int,b:int)->int:
    norm_a = normalize_rack(int(a)); norm_b = normalize_rack(int(b))
    return abs(norm_a - norm_b) // 2

def _to_tryb_int(v):
    if v is None or (isinstance(v,float) and np.isnan(v)): return None
    m=re.search(r'(\d+)',str(v)); return int(m.group(1)) if m else None

def _try_int(x):
    if pd.isna(x): return None
    m=re.search(r'(-?\d+)',str(x)); return int(m.group(1)) if m else None

def _find_col(df,names):
    cols=list(df.columns); low=[str(c).strip().lower() for c in cols]
    for cand in names:
        c=cand.lower()
        for i,h in enumerate(low):
            if c==h or c in h or h in c: return cols[i]
    return None

def _auto_map(df):
    m={k:_find_col(df,v) for k,v in NEEDED.items()}
    miss=[k for k in ("Numer misji","Regal","Kolumna") if m.get(k) is None]
    if miss: raise ValueError(f"Brakuje kolumn: {miss}. Nagłówki: {list(df.columns)}")
    return m

def fmt_loc(reg,col,poziom,miejsce):
    k=str(int(col)).zfill(3); p="00"
    if isinstance(poziom,str) and poziom.strip():
        mm=re.search(r'(\d+)',poziom); p=mm.group(1).zfill(2) if mm else "00"
    elif pd.notna(poziom): p=str(int(poziom)).zfill(2)
    m="C"
    if isinstance(miejsce,str) and miejsce.strip(): m=miejsce.strip()
    elif pd.notna(miejsce): m=str(miejsce).strip()
    return f"{int(reg)}.{k}{p}{m}"

def _pos(anchor):
    if anchor==FRONT: return 0.0
    if anchor==P034:  return A_FRONT_TO_034
    if anchor in (P057,P058): return A_FRONT_TO_034 + B_034_TO_057
    return 0.0

def _lane_distance(entry_anchor, exit_anchor, needA, needB, needC, has_ext):
    req = 0.0
    if needA: req = max(req, A_FRONT_TO_034)
    if needB: req = max(req, A_FRONT_TO_034 + B_034_TO_057)
    if needC and has_ext: req = max(req, A_FRONT_TO_034 + B_034_TO_057 + C_058_TO_062)
    a, b = 0.0, req
    s = _pos(entry_anchor)
    t = _pos(exit_anchor)
    if req == 0.0: return abs(t - s)
    return round((b - a) + min(abs(s - a) + abs(t - b), abs(s - b) + abs(t - a)), 1)

# --- ALGORYTMY SYMULACJI (BEZ ZMIAN W LOGICE) ---
def simulate_blocks(grp: pd.DataFrame, tryb:int, start_pair:int):
    det=[]; total=0.0
    g = grp.copy()
    blocks=[]; curp=None; cols=[]
    for _,r in g.iterrows():
        p=pair_key(int(r["Regal"])); c=int(r["Kolumna"])
        if curp is None: curp=p
        if p!=curp:
            if cols: blocks.append((curp,cols)); cols=[]
            curp=p
        cols.append(c)
    if cols: blocks.append((curp,cols))
    if not blocks: return 0.0, start_pair, []

    first_pair=blocks[0][0]
    hop = pair_gaps(start_pair, first_pair) * CROSS_LANES
    if hop>0:
        det.append((f"Front: front ({start_pair}) → para ({first_pair})", round(hop,1)))
        total += hop

    info=[]
    for p,cols_list in blocks:
        cols_int=[int(c) for c in cols_list]
        needA = any(c<=33 for c in cols_int)
        needB = any(c>=35 for c in cols_int)
        needC = any(c>=59 for c in cols_int)
        has_ext = tryb_ma_rozszerzenie(tryb,p)
        first_col = cols_int[0] if cols_int else 0
        info.append({'p': p, 'nA': needA, 'nB': needB, 'nC': needC, 'ext': has_ext, 'first_col': first_col})

    n=len(info)
    INF=10**9
    dp=[[INF,INF,INF] for _ in range(n)]
    prev=[[None,None,None] for _ in range(n)] 

    def get_m(state, nC, ext):
        if state==0: return 0.0
        if state==2: return A_FRONT_TO_034
        if nC and ext: return A_FRONT_TO_034 + B_034_TO_057 + C_058_TO_062
        return A_FRONT_TO_034 + B_034_TO_057

    def calc_aisle_cost_dp(state_in, state_out, data):
        nA, nB, nC, ext = data['nA'], data['nB'], data['nC'], data['ext']
        cost = INF
        anchor_back = (P058 if (ext and nC) else P057)
        if state_in == 0 and state_out == 0:
             cost = _lane_distance(FRONT, FRONT, nA, nB, nC, ext)
             if nA: cost += U_TURN_PENALTY 
        elif (state_in == 0 and state_out == 1) or (state_in == 1 and state_out == 0):
             cost = _lane_distance(FRONT, anchor_back, nA, nB, nC, ext)
        elif state_in == 1 and state_out == 1:
             cost = _lane_distance(anchor_back, anchor_back, nA, nB, nC, ext)
             if nB: cost += U_TURN_PENALTY
        elif (state_in == 0 and state_out == 2) or (state_in == 2 and state_out == 0):
            if not (nB or nC): cost = A_FRONT_TO_034
        elif (state_in == 1 and state_out == 2) or (state_in == 2 and state_out == 1):
            if not nA: cost = B_034_TO_057 + (C_058_TO_062 if (ext and nC) else 0.0)
        elif state_in == 2 and state_out == 2:
            if not (nA or nB or nC): cost = 0.0
        return cost

    def calc_aisle_cost_real(state_in, state_out, data):
        cost = calc_aisle_cost_dp(state_in, state_out, data)
        if cost >= INF: return cost
        if state_in==0 and state_out==0 and data['nA']: cost -= U_TURN_PENALTY
        if state_in==1 and state_out==1 and data['nB']: cost -= U_TURN_PENALTY
        return cost

    curr = info[0]
    dp[0][0] = calc_aisle_cost_real(0, 0, curr)
    dp[0][1] = calc_aisle_cost_real(0, 1, curr)
    dp[0][2] = calc_aisle_cost_real(0, 2, curr)

    for i in range(n - 1):
        curr = info[i]
        next_blk = info[i+1]
        gaps = pair_gaps(curr['p'], next_blk['p'])
        preferred_entry = None
        if next_blk['first_col'] > 0:
            if next_blk['first_col'] <= 34: preferred_entry = 0 
            elif next_blk['first_col'] >= 35: preferred_entry = 1

        for s_out in range(3):
            if dp[i][s_out] >= INF: continue
            for s_in in range(3):
                is_A_only_curr = curr['nA'] and not curr['nB']
                is_B_only_curr = curr['nB'] and not curr['nA']
                is_A_only_next = next_blk['nA'] and not next_blk['nB']
                is_B_only_next = next_blk['nB'] and not next_blk['nA']
                force_traversal = False
                if gaps == 1:
                    if (is_A_only_curr and is_B_only_next) or (is_B_only_curr and is_A_only_next):
                        force_traversal = True
                entry_penalty = 0.0
                if preferred_entry is not None and s_in != preferred_entry:
                    if gaps >= 2 and ((preferred_entry == 0 and s_in == 1) or (preferred_entry == 1 and s_in == 0)):
                        entry_penalty += WRONG_ENTRY_PENALTY
                    else:
                        entry_penalty += ENTRY_SIDE_SOFT_PENALTY
                cross_cost = INF
                method = 0
                if s_out == s_in:
                    cross_cost = gaps * CROSS_LANES
                    if gaps == 1 and s_out == 2:
                        if force_traversal: cross_cost += 10000.0
                        else: cross_cost += NEIGHBOR_MIDDLE_PENALTY
                    method = 0
                if gaps >= 2 and s_out != s_in:
                    m_out = get_m(s_out, curr['nC'], curr['ext'])
                    m_in  = get_m(s_in,  next_blk['nC'], next_blk['ext'])
                    drive_cost = abs(m_out - m_in)
                    cross_cost = gaps * CROSS_LANES + drive_cost + BRIDGE_PENALTY
                    method = 1
                if cross_cost >= INF: continue
                total_transition = cross_cost + entry_penalty
                for s_target in range(3):
                    pick_cost = calc_aisle_cost_dp(s_in, s_target, next_blk)
                    total_new = dp[i][s_out] + total_transition + pick_cost
                    if total_new < dp[i+1][s_target]:
                        dp[i+1][s_target] = total_new
                        prev[i+1][s_target] = (s_out, method, s_in)

    best_end = 0
    min_val = dp[-1][0]
    for s in range(1, 3):
        if dp[-1][s] < min_val: min_val = dp[-1][s]; best_end = s
    path = []
    curr_s = best_end
    for i in range(n-1, 0, -1):
        p_s, meth, in_s = prev[i][curr_s]
        path.append((p_s, meth, in_s, curr_s))
        curr_s = p_s
    path.reverse()
    path.insert(0, (None, 0, 0, curr_s))
    
    accumulated_debug = total
    def nname(idx): return "A" if idx==0 else "034" if idx==2 else "B"
    def label_cross(idx):
        n = nname(idx)
        return "Front" if n=="A" else "057" if n=="B" else "034"

    for i in range(n):
        curr = info[i]
        _, meth, s_in, s_out = path[i]
        if i > 0:
            p_out = path[i-1][3]
            prev_blk = info[i-1]
            gaps = pair_gaps(prev_blk['p'], curr['p'])
            lbl_start = label_cross(p_out)
            lbl_end   = label_cross(s_in)
            if meth == 0: 
                c_dist = gaps * CROSS_LANES
                det.append((f"{lbl_start}: para {prev_blk['p']}→{curr['p']}", round(c_dist, 1)))
                accumulated_debug += c_dist
            else: 
                mid_pair = (prev_blk['p'] + curr['p']) // 2
                dist_leg1 = pair_gaps(prev_blk['p'], mid_pair) * CROSS_LANES
                det.append((f"{lbl_start}: para {prev_blk['p']}→{mid_pair}", round(dist_leg1, 1)))
                lvl_dist = abs(get_m(p_out, prev_blk['nC'], prev_blk['ext']) - get_m(s_in, curr['nC'], curr['ext']))
                lc_start = nname(p_out); lc_end = nname(s_in)
                det.append((f"Alejka ({mid_pair}-{mid_pair+1}): {lc_start}→{lc_end}", round(lvl_dist, 1)))
                dist_leg2 = pair_gaps(mid_pair, curr['p']) * CROSS_LANES
                det.append((f"{lbl_end}: para {mid_pair}→{curr['p']}", round(dist_leg2, 1)))
                accumulated_debug += gaps * CROSS_LANES + lvl_dist

        cost_pick = calc_aisle_cost_real(s_in, s_out, curr)
        parts=[]
        if curr['nA']: parts.append("A=72,5")
        if curr['nB']: parts.append("B=66,0")
        if curr['nC'] and curr['ext']: parts.append("C=11,7")
        sl = nname(s_in); el = nname(s_out)
        arrow = f"{sl}→{el}"
        if sl == el and cost_pick > 0.1: arrow = f"{sl}→...→{el}"
        elif cost_pick < 0.1 and not parts: arrow = f"{sl} (przejazd)"
        det.append((f"Alejka ({curr['p']}-{curr['p']+1}): {arrow} ({' + '.join(parts) if parts else 'przejazd'})", cost_pick))
        accumulated_debug += cost_pick

    last_pair = info[-1]['p']
    ended_state = path[-1][3]
    if ended_state != 0: 
        ss_pair_tuple = START_STOP_MAP.get(tryb)
        ss_p = pair_key(min(ss_pair_tuple)) if ss_pair_tuple else last_pair
        norm_last = normalize_rack(last_pair); norm_ss = normalize_rack(ss_p)
        if norm_ss < norm_last:
            if last_pair == 901: neighbor = 879
            else: neighbor = last_pair - 2
        else:
            if last_pair == 879: neighbor = 901
            else: neighbor = last_pair + 2
        det.append((f"{label_cross(ended_state)}: para {last_pair}→{neighbor}", 5.4))
        accumulated_debug += 5.4
        ext_n = tryb_ma_rozszerzenie(tryb, neighbor)
        d_drive = get_m(ended_state, info[-1]['nC'], info[-1]['ext']) 
        det.append((f"Alejka ({neighbor}-{neighbor+1}): {nname(ended_state)}→A (powrót)", round(d_drive, 1)))
        accumulated_debug += d_drive
        gaps_back = pair_gaps(neighbor, ss_p)
        if gaps_back > 0:
             det.append((f"Front: para {neighbor}→{ss_p}", round(gaps_back * 5.4, 1)))
             accumulated_debug += gaps_back * 5.4
        last_pair = ss_p

    return round(accumulated_debug, 1), last_pair, det

def simulate_602(grp_602: pd.DataFrame, last_pair:int, ss_pair:int):
    det=[]; total=0.0
    axis_pair = pair_key(AXIS_602_LANE) 
    if last_pair == ss_pair:
        det.append((f"Front: front ({ss_pair}) → oś 602", SS_TO_AXIS_602))
        total += SS_TO_AXIS_602
    else:
        hop = pair_gaps(last_pair, axis_pair) * CROSS_LANES
        if hop>0:
            det.append((f"Front: front ({last_pair}) → oś 602", round(hop,1)))
            total += hop
    det.append((f"602: ścieżka oś→…→oś (2×73,0)", LENGTH_602_LOOP))
    total += LENGTH_602_LOOP
    return round(total,1), axis_pair, det


# ==========================================
# 2. LOGIKA PRZETWARZANIA DANYCH (Wrapper)
# ==========================================

@st.cache_data
def process_data(uploaded_file):
    """
    Funkcja przyjmuje plik, wykonuje logikę ze skryptu i zwraca struktury danych
    gotowe do wyświetlenia w Streamlit i pobrania jako Excel.
    """
    try:
        raw = pd.read_excel(uploaded_file, header=0, dtype=object)
        mapping = _auto_map(raw)
    except:
        try:
            raw = pd.read_excel(uploaded_file, header=1, dtype=object)
            mapping = _auto_map(raw)
        except Exception as e:
            return None, None, str(e)

    df=raw.rename(columns={mapping[k]:k for k in mapping if mapping[k]})

    # Czyszczenie
    if "Numer misji" in df.columns: df["Numer misji"] = df["Numer misji"].astype(str).str.strip()
    if "Tryb Pracy" in df.columns: df["Tryb Pracy"] = df["Tryb Pracy"].astype(str).str.strip()
    if "numer lini" not in df.columns:
        df["_ord"]=df.groupby("Numer misji").cumcount()+1
        df["numer lini"]=df["_ord"]*10
        df.drop(columns=["_ord"], inplace=True)

    df["Regal"]=df["Regal"].apply(_try_int)
    df["Kolumna"]=df["Kolumna"].apply(_try_int)
    df=df.dropna(subset=["Regal","Kolumna"]).copy()
    df["Regal"]=df["Regal"].astype(int); df["Kolumna"]=df["Kolumna"].astype(int)

    # Sortowanie
    df["__sort_nl__"] = pd.to_numeric(df["numer lini"], errors="coerce").fillna(9999999)
    df["__orig_idx__"] = range(len(df))

    rows_details=[]; rows_summary=[]; rows_debug=[]
    
    # Grupowanie i obliczanie
    groups = list(df.groupby(["Numer misji","Tryb Pracy"], dropna=False))
    
    # Progress bar w Streamlit
    progress_bar = st.progress(0)
    total_groups = len(groups)
    
    for i, ((misja,tryb),grp_misja) in enumerate(groups):
        tryb_int=_to_tryb_int(tryb)
        g = grp_misja.sort_values(["__sort_nl__", "__orig_idx__"], kind="mergesort")
        
        # Podział na bloki 602 / aisle
        blocks=[]; cur_t=None; buf=[]
        def flush():
            nonlocal blocks,cur_t,buf
            if buf: blocks.append((cur_t, pd.DataFrame(buf).reset_index(drop=True)))
            buf=[]
        for _,r in g.iterrows():
            t="602" if int(r["Regal"]) in (600,601) else "aisle"
            if cur_t is None: cur_t=t
            if t!=cur_t: flush(); cur_t=t
            buf.append(r.to_dict())
        flush()

        det=[]; total=0.0
        ss_pair_tuple=START_STOP_MAP.get(tryb_int)
        ss_pair=pair_key(min(ss_pair_tuple)) if ss_pair_tuple else pair_key(AXIS_602_LANE)
        ss_txt=f"({ss_pair_tuple[0]}-{ss_pair_tuple[1]})" if ss_pair_tuple else ""

        total+=START_STOP; det.append((f"S/S wejście {ss_txt}", START_STOP))
        last_pair=ss_pair

        for typ,bdf in blocks:
            if typ=="602":
                d,last_pair,ddet=simulate_602(bdf, last_pair, ss_pair)
                total+=d; det.extend(ddet)
            else:
                d,last_pair,ddet=simulate_blocks(bdf, tryb_int, last_pair)
                total+=d; det.extend(ddet)
        
        if ss_pair_tuple:
            hop = pair_gaps(last_pair, ss_pair) * CROSS_LANES
            if hop > 0:
                det.append((f"Front: front ({last_pair}) → S/S", round(hop,1)))
                total += hop

        det.append((f"S/S wyjście {ss_txt}", START_STOP))
        total += START_STOP
        dist = round(total, 1)
        stops_count=len(g.drop_duplicates(["Regal","Kolumna"]))

        # Zapis wyników
        for _,r in g.iterrows():
            out=r.to_dict()
            out["Pełna lokalizacja"]=fmt_loc(int(r["Regal"]), int(r["Kolumna"]), r.get("Poziom"), r.get("miejsce"))
            out["Ilość stopów w misji"]=stops_count
            out["Dystans (m)"]=dist
            out["Dystans (km)"]=round(dist/1000.0,3)
            for tmpc in ["__sort_nl__", "__orig_idx__"]:
                if tmpc in out: del out[tmpc]
            rows_details.append(out)

        rows_summary.append({"Numer misji":misja,"Tryb Pracy":tryb,"Start/Stop":ss_txt.strip("()"),
                             "Ilość stopów":stops_count,"Dystans (m)":dist,"Dystans (km)":round(dist/1000.0,3)})

        k=1
        for opis,metry in det:
            rows_debug.append({"Numer misji":misja,"Tryb Pracy":tryb,"Krok":k,"Opis":opis,"Metry":round(metry, 1)})
            k+=1

        progress_bar.progress(min((i+1)/total_groups, 1.0))

    df_det=pd.DataFrame(rows_details)
    df_sum=pd.DataFrame(rows_summary)
    progress_bar.empty()
    
    return df_det, df_sum, None

def generate_excel_download(df_det, df_sum):
    output = io.BytesIO()
    wb = Workbook()
    
    ws1=wb.active; ws1.title="Szczegóły"
    for r in dataframe_to_rows(df_det, index=False, header=True): ws1.append(r)

    ws2=wb.create_sheet("Podsumowanie")
    for r in dataframe_to_rows(df_sum, index=False, header=True): ws2.append(r)
    
    wb.save(output)
    return output.getvalue()


# ==========================================
# 3. INTERFEJS UŻYTKOWNIKA (STREAMLIT + PLOTLY)
# ==========================================

st.set_page_config(page_title="VD Stopy Dystans", layout="wide", page_icon="📦")

# Stylizacja CSS (Nowoczesna)
st.markdown("""
<style>
    .main {background-color: #f8f9fa;}
    .block-container {padding-top: 2rem;}
    h1 {color: #2c3e50;}
    div.stMetric {background-color: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);}
</style>
""", unsafe_allow_html=True)

st.title("📦 VD: Analiza Stopy i Dystans")
st.markdown("**Automatyczna analiza ścieżek kompletacyjnych i wizualizacja tras**")

# --- SIDEBAR (Upload) ---
st.sidebar.header("📂 Dane wejściowe")
uploaded_file = st.sidebar.file_uploader("Wgraj plik Excel (analiza.xlsx)", type=["xlsx"])

if uploaded_file:
    df_det, df_sum, error = process_data(uploaded_file)
    
    if error:
        st.error(f"Błąd przetwarzania pliku: {error}")
    else:
        # --- DASHBOARD GŁÓWNY ---
        st.sidebar.success("Plik przetworzony pomyślnie!")
        
        # KPI Top Level
        total_dist_km = df_sum["Dystans (km)"].sum()
        total_missions = df_sum["Numer misji"].nunique()
        avg_dist = df_sum["Dystans (m)"].mean()
        
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Całkowity Dystans", f"{total_dist_km:.2f} km")
        k2.metric("Liczba Misji", f"{total_missions}")
        k3.metric("Śr. Dystans / Misję", f"{avg_dist:.1f} m")
        k4.metric("Liczba Stopów", f"{df_sum['Ilość stopów'].sum()}")

        # --- SEKCJA DIGITAL TWIN (MAPA) ---
        st.subheader("🗺️ Wizualizacja Trasy (VD Viewer)")
        
        # Wybór misji do podglądu
        selected_mission = st.selectbox("Wybierz misję do wizualizacji:", df_det["Numer misji"].unique())
        
        # Filtrowanie danych dla wybranej misji
        mission_data = df_det[df_det["Numer misji"] == selected_mission].sort_values("numer lini")
        
        # --- TWORZENIE MAPY W PLOTLY ---
        
        def get_coords(regal, kolumna):
            """ Zamienia Regal/Kolumna na X/Y do wykresu """
            pair = pair_key(regal)
            offset = 0.2 if regal % 2 == 0 else -0.2
            x = pair + offset
            y = kolumna 
            return x, y

        # Generowanie punktów dla wybranej misji
        mission_coords = []
        for _, row in mission_data.iterrows():
            x, y = get_coords(row["Regal"], row["Kolumna"])
            mission_coords.append({"x": x, "y": y, "label": row["Pełna lokalizacja"], "sku": row.get("SKU", "")})
        
        df_coords = pd.DataFrame(mission_coords)
        
        fig = go.Figure()

        # Tło (Szkielet magazynu)
        max_pair = df_det["Regal"].apply(pair_key).max()
        max_col = df_det["Kolumna"].max()
        
        bg_x = []
        bg_y = []
        active_pairs = sorted(df_det["Regal"].apply(pair_key).unique())
        for p in active_pairs:
            for c in range(0, int(max_col)+1, 5): 
                bg_x.append(p - 0.2); bg_y.append(c)
                bg_x.append(p + 0.2); bg_y.append(c)
                
        fig.add_trace(go.Scatter(
            x=bg_x, y=bg_y, mode='markers',
            marker=dict(color='#e0e0e0', size=4),
            name='Struktura Magazynu', hoverinfo='none'
        ))

        # Ścieżka (Linia łącząca punkty)
        if not df_coords.empty:
            fig.add_trace(go.Scatter(
                x=df_coords["x"], y=df_coords["y"],
                mode='lines+markers',
                line=dict(color='#0066cc', width=3, dash='dot'),
                marker=dict(size=10, color='#ff9900', symbol='square'),
                name='Trasa Zbiórki',
                text=df_coords["label"],
                hoverinfo='text+x+y'
            ))
            
            # Start (zmieniono symbol na triangle-right)
            fig.add_trace(go.Scatter(
                x=[df_coords.iloc[0]["x"]], y=[df_coords.iloc[0]["y"]],
                mode='markers', marker=dict(size=14, color='green', symbol='triangle-right'), name='Start'
            ))
            # Koniec (zmieniono symbol na square)
            fig.add_trace(go.Scatter(
                x=[df_coords.iloc[-1]["x"]], y=[df_coords.iloc[-1]["y"]],
                mode='markers', marker=dict(size=14, color='red', symbol='square'), name='Koniec'
            ))

        fig.update_layout(
            title=f"Wizualizacja Misji: {selected_mission}",
            xaxis_title="Numer Alei (Pary Regałów)",
            yaxis_title="Kolumna (Głębokość)",
            height=600,
            plot_bgcolor='white',
            showlegend=True,
            xaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
            yaxis=dict(showgrid=True, gridcolor='#f0f0f0')
        )
        
        st.plotly_chart(fig, use_container_width=True)

        # --- TABELE SZCZEGÓŁOWE ---
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("📋 Szczegóły Misji")
            st.dataframe(mission_data[["numer lini", "Pełna lokalizacja", "SKU", "Regal", "Kolumna"]], hide_index=True)
        
        with c2:
            st.subheader("📊 Statystyki Misji")
            mission_sum = df_sum[df_sum["Numer misji"] == selected_mission].iloc[0]
            st.write(f"**Tryb:** {mission_sum['Tryb Pracy']}")
            st.write(f"**Dystans:** {mission_sum['Dystans (m)']} m")
            st.write(f"**Liczba stopów:** {mission_sum['Ilość stopów']}")

        # --- DOWNLOAD ---
        st.markdown("---")
        st.header("📥 Pobierz Raport")
        excel_data = generate_excel_download(df_det, df_sum)
        st.download_button(
            label="Pobierz Raport Excel (VD Stopy Dystans)",
            data=excel_data,
            file_name="Raport_Stopy_Dystans.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    # Ekran startowy
    st.info("👈 Wgraj plik 'analiza.xlsx' w panelu bocznym.")
    st.markdown("""
    ### VD Stopy Dystans - Instrukcja
    1. Aplikacja wczytuje Twój plik z danymi.
    2. Przelicza trasy zgodnie z algorytmem (strefy A/B/C, 602, kary).
    3. Wizualizuje trasę na mapie 2D.
    4. Pozwala pobrać gotowy raport w formacie Excel.
    """)