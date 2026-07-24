import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import time
import io
import os
import sqlite3
from datetime import datetime, date

# 1. ตั้งค่าหน้าจอ
st.set_page_config(layout="wide", page_title="Xbar R Control Chart", initial_sidebar_state="collapsed")
st.markdown("""
    <style>
        .block-container { padding-top: 3.5rem; padding-bottom: 0rem; padding-left: 1.5rem; padding-right: 1.5rem; }
        div[data-testid="stMetricValue"] { font-size: 1.3rem; }
        div[data-testid="stMetricLabel"] { font-size: 0.85rem; }
        h4 { margin-top: 0px; }
        div[data-testid="column"] > div > div > div > div > div { padding-bottom: 0rem; }
    </style>
""", unsafe_allow_html=True)

# 🌟 หัวเว็บ
st.markdown("<h3 style='color: #0033A0; margin-bottom: 20px;'>📊 SPC - XBarR Chart (Network Multi-User System)</h3>", unsafe_allow_html=True)

# ==========================================
# 🌟 0. ค่าคงที่ Xbar-R Chart ตามขนาดตัวอย่าง (n = 2..10)
# ==========================================
MAX_N = 10  # จำนวนคอลัมน์ค่าวัดสูงสุดที่รองรับ (x1..x10)

XBAR_R_CONSTANTS = {
    2:  {"A2": 1.880, "d2": 1.128, "D3": 0.000, "D4": 3.267},
    3:  {"A2": 1.023, "d2": 1.693, "D3": 0.000, "D4": 2.574},
    4:  {"A2": 0.729, "d2": 2.059, "D3": 0.000, "D4": 2.282},
    5:  {"A2": 0.577, "d2": 2.326, "D3": 0.000, "D4": 2.114},
    6:  {"A2": 0.483, "d2": 2.534, "D3": 0.000, "D4": 2.004},
    7:  {"A2": 0.419, "d2": 2.704, "D3": 0.076, "D4": 1.924},
    8:  {"A2": 0.373, "d2": 2.847, "D3": 0.136, "D4": 1.864},
    9:  {"A2": 0.337, "d2": 2.970, "D3": 0.184, "D4": 1.816},
    10: {"A2": 0.308, "d2": 3.078, "D3": 0.223, "D4": 1.777},
}

def get_constants(n):
    """คืนค่าคงที่ A2, d2, D3, D4 ตามขนาดตัวอย่าง n (จำกัดช่วง 2-10, ถ้านอกช่วงใช้ n=3 แทน)"""
    n_clamped = int(n) if pd.notna(n) and int(n) in XBAR_R_CONSTANTS else 3
    c = XBAR_R_CONSTANTS[n_clamped]
    return c["A2"], c["d2"], c["D3"], c["D4"], n_clamped

# ==========================================
# 🌟 1. กำหนด Network Path และระบบ SQLite Database ส่วนกลาง
# ==========================================
NETWORK_PATH = r"Z:"  
ALT_PATH = r"\\pbpr0d\Ricoh Scan\S.Watcharaporn\00.Web App\spc_project"

IS_NETWORK_STORAGE = False

if os.path.exists(NETWORK_PATH):
    DB_FILE_PATH = os.path.join(NETWORK_PATH, "spc_master_network.db")
    excel_file = os.path.join(NETWORK_PATH, "master_data.xlsx")
    IS_NETWORK_STORAGE = True
elif os.path.exists(ALT_PATH):
    DB_FILE_PATH = os.path.join(ALT_PATH, "spc_master_network.db")
    excel_file = os.path.join(ALT_PATH, "master_data.xlsx")
    IS_NETWORK_STORAGE = True
else:
    DB_FILE_PATH = "spc_master_network.db"
    excel_file = "master_data.xlsx"
    IS_NETWORK_STORAGE = False

DB_BUSY_TIMEOUT_MS = 10000  # 10 วินาที
DB_MAX_RETRY = 3

def get_db_connection():
    conn = sqlite3.connect(DB_FILE_PATH, timeout=DB_BUSY_TIMEOUT_MS / 1000, check_same_thread=False)
    conn.execute(f"PRAGMA busy_timeout = {DB_BUSY_TIMEOUT_MS};")
    if IS_NETWORK_STORAGE:
        conn.execute("PRAGMA journal_mode = DELETE;")
    else:
        conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")

    x_cols_def = ", ".join([f"x{i} REAL" for i in range(1, MAX_N + 1)])
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS spc_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model TEXT,
            die_no TEXT,
            dimension TEXT,
            date TEXT,
            time TEXT,
            inspector TEXT,
            year TEXT,
            mc TEXT,
            n INTEGER,
            {x_cols_def},
            action TEXT
        )
    ''')

    existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(spc_records)").fetchall()}
    new_cols = {"year": "TEXT", "mc": "TEXT", "n": "INTEGER"}
    for i in range(1, MAX_N + 1):
        new_cols[f"x{i}"] = "REAL"
    for col, col_type in new_cols.items():
        if col not in existing_cols:
            conn.execute(f"ALTER TABLE spc_records ADD COLUMN {col} {col_type};")
    conn.commit()
    return conn

def _run_with_retry(fn):
    last_err = None
    for attempt in range(DB_MAX_RETRY):
        try:
            return fn()
        except sqlite3.OperationalError as e:
            last_err = e
            if "locked" in str(e).lower():
                time.sleep(0.5 * (attempt + 1))
                continue
            raise
    raise last_err

X_COLS = [f"x{i}" for i in range(1, MAX_N + 1)]
X_COLS_DISPLAY = [f"X{i}" for i in range(1, MAX_N + 1)]

def load_die_data_db(model_name, die_no):
    try:
        conn = get_db_connection()
        x_select = ", ".join([f"{c} as '{d}'" for c, d in zip(X_COLS, X_COLS_DISPLAY)])
        query = f"""
            SELECT id, model as 'Model', die_no as 'Die no.', dimension as 'Dimension',
                   date as 'Date', time as 'Time', inspector as 'Inspector',
                   year as 'Year', mc as 'M/C', n as 'n',
                   {x_select}, action as 'Action'
            FROM spc_records
            WHERE model = ? AND die_no = ?
        """
        df = pd.read_sql_query(query, conn, params=(model_name, die_no))
        conn.close()
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูลจาก Database: {e}")
        cols = ['id', 'Model', 'Die no.', 'Dimension', 'Date', 'Time', 'Inspector', 'Year', 'M/C', 'n'] + X_COLS_DISPLAY + ['Action']
        return pd.DataFrame(columns=cols)

def save_single_record_db(model_name, die_no, dim, date_str, time_str, inspector, year_str, mc_str, values, action=""):
    def _do():
        conn = get_db_connection()
        cursor = conn.cursor()
        padded_values = list(values) + [None] * (MAX_N - len(values))
        placeholders = ", ".join(["?"] * MAX_N)
        cursor.execute(f'''
            INSERT INTO spc_records (model, die_no, dimension, date, time, inspector, year, mc, n, {", ".join(X_COLS)}, action)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, {placeholders}, ?)
        ''', (model_name, die_no, dim, date_str, time_str, inspector, year_str, mc_str, len(values), *padded_values, action))
        conn.commit()
        conn.close()
        return True
    try:
        return _run_with_retry(_do)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล (ลองใหม่ {DB_MAX_RETRY} ครั้งแล้ว): {e}")
        return False

def update_records_db(df_edited):
    """🌟 ฟังก์ชันอัปเดตข้อมูลที่มีการแก้ไขใน st.data_editor กลับลง SQLite Database"""
    def _do():
        conn = get_db_connection()
        cursor = conn.cursor()
        
        for _, row in df_edited.iterrows():
            record_id = row['id']
            # แปลงค่า X1-X10
            x_vals = [row[col] if pd.notna(row[col]) else None for col in X_COLS_DISPLAY]
            
            sql = f'''
                UPDATE spc_records 
                SET date=?, time=?, inspector=?, year=?, mc=?, action=?,
                    x1=?, x2=?, x3=?, x4=?, x5=?, x6=?, x7=?, x8=?, x9=?, x10=?
                WHERE id=?
            '''
            params = (
                str(row['Date']), str(row['Time']), str(row['Inspector']), 
                str(row['Year']), str(row['M/C']), str(row['Action']),
                *x_vals, record_id
            )
            cursor.execute(sql, params)
            
        conn.commit()
        conn.close()
        return True
        
    try:
        return _run_with_retry(_do)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการบันทึกการแก้ไข: {e}")
        return False

def delete_records_by_ids(id_list):
    def _do():
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.executemany("DELETE FROM spc_records WHERE id = ?", [(idx,) for idx in id_list])
        conn.commit()
        conn.close()
        return True
    try:
        return _run_with_retry(_do)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการลบข้อมูล (ลองใหม่ {DB_MAX_RETRY} ครั้งแล้ว): {e}")
        return False

# ==========================================
# 🌟 2. ระบบดึง Master Data
# ==========================================
@st.cache_data
def load_master_data(file_path, modified_time):
    if os.path.exists(file_path):
        df_master = pd.read_excel(file_path)

        expected_cols = ['Model', 'Part Name', 'Die no.', 'Dimension', 'Unit', 'Gage name',
                          'Frequency', 'Process', 'LSL', 'USL', 'Sample size (n)']
        for col in expected_cols:
            if col not in df_master.columns:
                if col in ['LSL', 'USL']:
                    df_master[col] = 0.0
                elif col == 'Sample size (n)':
                    df_master[col] = 3
                else:
                    df_master[col] = "-"

        master_dict = {}
        for _, row in df_master.iterrows():
            model = str(row['Model']).strip()
            die_no = str(row['Die no.']).strip()
            dim = str(row['Dimension']).strip()

            if model not in master_dict: master_dict[model] = {}
            if die_no not in master_dict[model]: master_dict[model][die_no] = {}

            try:
                raw_n = int(row['Sample size (n)'])
                sample_n = raw_n if 2 <= raw_n <= MAX_N else 3
            except (ValueError, TypeError):
                sample_n = 3

            master_dict[model][die_no][dim] = {
                "Part Name": str(row['Part Name']),
                "Unit": str(row['Unit']),
                "Gage name": str(row['Gage name']),
                "Frequency": str(row['Frequency']),
                "Process": str(row['Process']),
                "USL": float(row['USL']),
                "LSL": float(row['LSL']),
                "Sample size (n)": sample_n
            }
        return master_dict
    else:
        st.warning(f"⚠️ ไม่พบไฟล์ {file_path} (กำลังใช้ข้อมูลจำลองแทน)")
        return {
            "MD-101": {
                "1": {
                    "●Length": {
                        "Part Name": "PART-A01", "Unit": "mm.", "Gage name": "Micrometer",
                        "Frequency": "4hr/1pcs", "Process": "Stamping",
                        "USL": 10.50, "LSL": 9.50, "Sample size (n)": 3
                    }
                }
            }
        }

file_mod_time = os.path.getmtime(excel_file) if os.path.exists(excel_file) else 0
MASTER_DATA = load_master_data(excel_file, file_mod_time)

if 'ooc_alert' not in st.session_state:
    st.session_state.ooc_alert = None
if 'last_inspector' not in st.session_state:
    st.session_state.last_inspector = ""

# ==========================================
# 🌟 3. ส่วนหัว: Dropdown 3 ระดับ
# ==========================================
col_sel1, col_sel2, col_sel3 = st.columns(3)

with col_sel1:
    selected_model = st.selectbox("🔍 1. เลือก Model", list(MASTER_DATA.keys()))

die_list = list(MASTER_DATA[selected_model].keys())
with col_sel2:
    selected_die = st.selectbox("🎲 2. เลือก Die no.", die_list)

dim_list = list(MASTER_DATA[selected_model][selected_die].keys())
with col_sel3:
    selected_dim = st.selectbox("📏 3. เลือก Dimension (Characteristic)", dim_list)

info = MASTER_DATA[selected_model][selected_die][selected_dim]
spec_max = info["USL"]
spec_min = info["LSL"]
sample_n = info["Sample size (n)"]
A2, d2, D3, D4, sample_n = get_constants(sample_n)

df_spc = load_die_data_db(selected_model, selected_die)

# ==========================================
# 🌟 4. แสดงตาราง Master Data
# ==========================================
with st.expander("📌 Master Data & Specification (อ้างอิงจาก Excel)", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.text_input("Part Name", info['Part Name'], disabled=True)
        st.text_input("Characteristic", selected_dim, disabled=True)
        st.text_input("Unit", info['Unit'], disabled=True)
    with c2:
        st.text_input("Model", selected_model, disabled=True)
        st.text_input("Die no.", selected_die, disabled=True)
        st.text_input("Gage name", info['Gage name'], disabled=True)
        st.text_input("Sample size (n)", str(sample_n), disabled=True)
    with c3:
        st.text_input("Frequency", info['Frequency'], disabled=True)
        f_year = st.text_input("Year", str(date.today().year), key="year_input")
        st.text_input("Process", info['Process'], disabled=True)
        f_mc = st.text_input("M/C", "", key="mc_input")
    with c4:
        st.markdown("**Specification**")
        st.text_input("Max. (USL)", f"{spec_max:.3f}", disabled=True)
        st.text_input("Min. (LSL)", f"{spec_min:.3f}", disabled=True)

st.markdown("---")

# ==========================================
# 🌟 5. ระบบคำนวณสถิติ
# ==========================================
df_raw = df_spc[df_spc['Dimension'] == selected_dim].copy()

x_db_bar = r_bar = sigma = ucl_x = lcl_x = ucl_r = lcl_r = cp = cpu = cpl = cpk = 0

if len(df_raw) > 0:
    df_raw[X_COLS_DISPLAY] = df_raw[X_COLS_DISPLAY].apply(pd.to_numeric, errors='coerce')
    df_raw['Xbar'] = df_raw[X_COLS_DISPLAY].mean(axis=1, skipna=True)
    df_raw['R'] = df_raw[X_COLS_DISPLAY].max(axis=1, skipna=True) - df_raw[X_COLS_DISPLAY].min(axis=1, skipna=True)

    df_50 = df_raw.tail(50)
    x_db_bar, r_bar = df_50['Xbar'].mean(), df_50['R'].mean()
    sigma = r_bar / d2
    ucl_x, lcl_x = x_db_bar + (A2 * r_bar), x_db_bar - (A2 * r_bar)
    ucl_r, lcl_r = D4 * r_bar, D3 * r_bar

    if sigma > 0:
        cp = (spec_max - spec_min) / (6 * sigma)
        cpu, cpl = (spec_max - x_db_bar) / (3 * sigma), (x_db_bar - spec_min) / (3 * sigma)
        cpk = min(cpu, cpl)

    conditions = [
        (df_raw['Xbar'] > ucl_x) | (df_raw['Xbar'] < lcl_x),
        (df_raw['R'] > ucl_r)
    ]
    df_raw['OCS'] = np.select(conditions, ['OOC (Xbar)', 'OOC (R)'], default='OK')

df_plot = df_raw.tail(50)

if st.session_state.ooc_alert:
    st.error(f"🚨 **แจ้งเตือน:** ข้อมูลลอตล่าสุดพบความผิดปกติประเภท **{st.session_state.ooc_alert}** กรุณาดำเนินการแก้ไข!")
    if st.button("รับทราบ และซ่อนการแจ้งเตือน"):
        st.session_state.ooc_alert = None
        st.rerun()

# ==========================================
# 🌟 6. ส่วนแท็บและ UI
# ==========================================
tab1, tab2, tab3 = st.tabs(["📈 Dashboard (บันทึกและดูกราฟ)", "🗂️ จัดการข้อมูล (แก้ไข/ลบ/ใส่ Action)", "📊 Summary Report (สรุปภาพรวม)"])

# ---------------------------------------------------------
# TAB 1: DASHBOARD
# ---------------------------------------------------------
with tab1:
    col_stats, col_form, col_charts = st.columns([1.2, 1.3, 2.5])

    with col_stats:
        st.markdown("**🧮 Statistics**")
        s1, s2 = st.columns(2)
        s1.metric("Total (n)", len(df_raw))
        s2.metric("s (σ)", f"{sigma:.3f}")
        s1.metric("X̿", f"{x_db_bar:.3f}")
        s2.metric("R-bar", f"{r_bar:.3f}")
        s1.metric("Cp", f"{cp:.2f}")
        if cpk >= 1.33: s2.success(f"Cpk: {cpk:.2f}")
        elif len(df_raw) > 0: s2.error(f"Cpk: {cpk:.2f}")
        else: s2.metric("Cpk", "0.00")

        st.markdown("**📏 Control Limits**")
        l1, l2 = st.columns(2)
        l1.metric("UCL (Xbar)", f"{ucl_x:.3f}")
        l2.metric("LCL (Xbar)", f"{lcl_x:.3f}")
        l1.metric("UCL (R)", f"{ucl_r:.3f}")
        l2.metric("LCL (R)", f"{lcl_r:.3f}")

        st.markdown("**📋 Data History**")
        if len(df_plot) > 0:
            st.dataframe(df_plot[['Time', 'Xbar', 'R', 'OCS']].iloc[::-1], height=150, use_container_width=True)
        else:
            st.caption("No data.")

    with col_form:
        st.markdown("**📝 Data Entry**")
        with st.form("input_form", clear_on_submit=True):
            c_dt1, c_dt2 = st.columns(2)
            f_date = c_dt1.date_input("Date")
            f_time = c_dt2.time_input("Time")

            f_insp = st.text_input("Inspector", value=st.session_state.last_inspector)

            st.write(f"Measurements (X1 - X{sample_n}) [{selected_model} (Die {selected_die}) : {selected_dim}]:")
            values = []
            per_row = 5
            for row_start in range(0, sample_n, per_row):
                row_ids = list(range(row_start, min(row_start + per_row, sample_n)))
                cols = st.columns(len(row_ids))
                for col, idx in zip(cols, row_ids):
                    v = col.number_input(f"{idx + 1}", value=0.00, step=0.01, key=f"meas_{idx}")
                    values.append(v)

            submitted = st.form_submit_button("💾 Save", use_container_width=True)

            if submitted:
                st.session_state.last_inspector = f_insp
                current_xbar = sum(values) / sample_n
                current_r = max(values) - min(values)

                ocs_status = "OK"
                if len(df_raw) > 0:
                    if current_xbar > ucl_x or current_xbar < lcl_x: ocs_status = "OOC (Xbar)"
                    elif current_r > ucl_r: ocs_status = "OOC (R)"

                success = save_single_record_db(
                    selected_model, selected_die, selected_dim,
                    f_date.strftime("%Y-%m-%d"), f_time.strftime("%H:%M"),
                    f_insp, f_year, f_mc, values
                )

                if success:
                    if ocs_status != "OK":
                        st.session_state.ooc_alert = ocs_status
                        st.toast(f"🚨 พบความผิดปกติ ({ocs_status})!", icon="🚨")
                        time.sleep(1.5)
                    else:
                        st.session_state.ooc_alert = None
                        st.toast("✅ บันทึกลง Network Database สำเร็จ", icon="✅")
                        time.sleep(1)
                st.rerun()

    with col_charts:
        st.markdown(f"**📈 Control Charts: {selected_model} (Die {selected_die}) - {selected_dim}**")
        if len(df_plot) > 0:
            def safe_format_date(row):
                d_str = str(row['Date']) if pd.notna(row['Date']) else ""
                t_str = str(row['Time']) if pd.notna(row['Time']) else ""
                short_date = d_str[5:] if len(d_str) >= 5 else d_str
                return f"{short_date}\n{t_str}"

            x_labels = df_plot.apply(safe_format_date, axis=1)
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 5.5))

            ax1.plot(x_labels, df_plot['Xbar'], marker='o', color='#0033A0')
            ax1.axhline(x_db_bar, color='blue', linestyle='-', label='CL')
            ax1.axhline(ucl_x, color='red', linestyle='--', label='UCL')
            ax1.axhline(lcl_x, color='red', linestyle='--', label='LCL')
            ax1.axhline(spec_max, color='purple', linestyle=':', label='USL')
            ax1.axhline(spec_min, color='purple', linestyle=':', label='LSL')
            ax1.set_title("Xbar Chart", fontsize=10)
            ax1.legend(loc="upper right", fontsize=7, ncol=5)
            ax1.grid(True, alpha=0.3)
            ax1.tick_params(axis='x', rotation=0, labelsize=8)

            ax2.plot(x_labels, df_plot['R'], marker='s', color='#D22630')
            ax2.axhline(r_bar, color='blue', linestyle='-', label='CL')
            ax2.axhline(ucl_r, color='red', linestyle='--', label='UCL')
            ax2.axhline(lcl_r, color='red', linestyle='--', label='LCL')
            ax2.set_title("R Chart", fontsize=10)
            ax2.legend(loc="upper right", fontsize=7, ncol=3)
            ax2.grid(True, alpha=0.3)
            ax2.tick_params(axis='x', rotation=0, labelsize=8)

            plt.tight_layout(pad=1.0)

            img_buf = io.BytesIO()
            fig.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
            img_buf.seek(0)

            st.download_button(
                label="📷 ดาวน์โหลดรูปกราฟ (PNG)",
                data=img_buf,
                file_name=f"Chart_{selected_model}_Die{selected_die}_{time.strftime('%Y%m%d_%H%M')}.png",
                mime="image/png",
                use_container_width=False
            )

            st.pyplot(fig)
        else:
            st.info("💡 ยังไม่มีข้อมูลสำหรับชิ้นงาน/Die/จุดวัดนี้ กรุณาบันทึกข้อมูล")

# ---------------------------------------------------------
# TAB 2: DATA MANAGEMENT
# ---------------------------------------------------------
with tab2:
    st.markdown(f"#### 🛠️ จัดการฐานข้อมูล (เฉพาะ Model: {selected_model} | Die no.: {selected_die})")

    edit_columns = ['Model', 'Die no.', 'Dimension', 'Date', 'Time', 'Inspector', 'Year', 'M/C', 'n'] + X_COLS_DISPLAY + ['Action']

    df_for_editor = df_spc.copy()
    df_for_editor.insert(0, '🗑️ เลือกเพื่อลบ', False)

    display_columns = ['🗑️ เลือกเพื่อลบ', 'id'] + edit_columns
    df_display = df_for_editor[display_columns].iloc[::-1].reset_index(drop=True)

    edited_df = st.data_editor(
        df_display,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=['id', 'Model', 'Die no.', 'n'],
        height=350
    )

    # 🌟 ปุ่มบันทึกการแก้ไข และ ปุ่มลบแถวที่เลือก
    col_save_btn, col_del_btn, _ = st.columns([1.5, 1.5, 3])
    
    with col_save_btn:
        if st.button("💾 บันทึกการแก้ไขข้อมูล", type="primary", use_container_width=True):
            if update_records_db(edited_df):
                st.success("บันทึกการแก้ไขลง Database สำเร็จ!")
                time.sleep(1)
                st.rerun()

    with col_del_btn:
        if st.button("🗑️ ยืนยันการลบแถวที่เลือก", use_container_width=True):
            to_delete = edited_df[edited_df['🗑️ เลือกเพื่อลบ'] == True]
            if len(to_delete) > 0:
                ids_to_del = to_delete['id'].tolist()
                if delete_records_by_ids(ids_to_del):
                    st.success("ลบข้อมูลออกจาก Database สำเร็จ!")
                    time.sleep(1)
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 📥 ดาวน์โหลดข้อมูล (Export Data)")

    if len(df_spc) > 0:
        df_spc_dates = pd.to_datetime(df_spc['Date'], errors='coerce')
        valid_dates = df_spc_dates.dropna()

        if len(valid_dates) > 0:
            min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
        else:
            min_date = max_date = date.today()

        st.markdown("**📅 เลือกช่วงวันที่สำหรับดาวน์โหลด**")
        dc1, dc2, dc3 = st.columns([1.5, 1.5, 2])
        with dc1:
            dl_start = st.date_input("จากวันที่", value=min_date, min_value=min_date, max_value=max_date, key="dl_start")
        with dc2:
            dl_end = st.date_input("ถึงวันที่", value=max_date, min_value=min_date, max_value=max_date, key="dl_end")
        with dc3:
            st.write("")
            st.write("")
            include_all_dims = st.checkbox("รวมทุก Dimension ของ Die นี้ (ไม่ใช่แค่จุดวัดที่เลือกอยู่)", value=True)

        if dl_start > dl_end:
            st.warning("⚠️ 'จากวันที่' ต้องไม่มากกว่า 'ถึงวันที่'")
            df_export = df_spc.iloc[0:0]
        else:
            date_mask = (df_spc_dates.dt.date >= dl_start) & (df_spc_dates.dt.date <= dl_end)
            df_export = df_spc[date_mask].copy()
            if not include_all_dims:
                df_export = df_export[df_export['Dimension'] == selected_dim]

        st.caption(f"พบข้อมูล {len(df_export)} แถว ในช่วงวันที่ {dl_start} ถึง {dl_end}")

        if len(df_export) > 0:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_export[edit_columns].to_excel(writer, index=False, sheet_name=f'SPC_{selected_model}_Die{selected_die}')

            col_dl, _ = st.columns([2, 3])
            with col_dl:
                st.download_button(
                    label=f"📊 คลิกเพื่อดาวน์โหลดไฟล์ Excel ({selected_model} - Die {selected_die})",
                    data=buffer.getvalue(),
                    file_name=f"SPC_Report_{selected_model}_Die{selected_die}_{dl_start}_to_{dl_end}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.caption("ไม่พบข้อมูลในช่วงวันที่ที่เลือก")
    else:
        st.caption("ยังไม่มีข้อมูลสำหรับดาวน์โหลด")


# ---------------------------------------------------------
# TAB 3: SUMMARY REPORT
# ---------------------------------------------------------
with tab3:
    st.markdown("#### 📊 Summary Report (สรุปภาพรวม Capability ทุก Model & Die no.)")

    summary_data = []
    try:
        conn = get_db_connection()
        all_db_df = pd.read_sql_query("SELECT * FROM spc_records", conn)
        conn.close()
    except Exception:
        all_db_df = pd.DataFrame()

    if len(all_db_df) > 0:
        grouped = all_db_df.groupby(['model', 'die_no', 'dimension'])
        for (m_key, d_key, g_dim), group in grouped:
            if m_key in MASTER_DATA and d_key in MASTER_DATA[m_key] and g_dim in MASTER_DATA[m_key][d_key]:
                g_info = MASTER_DATA[m_key][d_key][g_dim]
                g_spec_max = g_info["USL"]
                g_spec_min = g_info["LSL"]
                g_A2, g_d2, g_D3, g_D4, g_n = get_constants(g_info["Sample size (n)"])

                df_calc = group.tail(50).copy()
                df_calc[X_COLS] = df_calc[X_COLS].apply(pd.to_numeric, errors='coerce')
                df_calc['Xbar'] = df_calc[X_COLS].mean(axis=1, skipna=True)
                df_calc['R'] = df_calc[X_COLS].max(axis=1, skipna=True) - df_calc[X_COLS].min(axis=1, skipna=True)

                g_x_db_bar = df_calc['Xbar'].mean()
                g_r_bar = df_calc['R'].mean()
                g_sigma = g_r_bar / g_d2

                if g_sigma > 0:
                    g_cp = (g_spec_max - g_spec_min) / (6 * g_sigma)
                    g_cpu = (g_spec_max - g_x_db_bar) / (3 * g_sigma)
                    g_cpl = (g_x_db_bar - g_spec_min) / (3 * g_sigma)
                    g_cpk = min(g_cpu, g_cpl)
                else:
                    g_cp = g_cpk = 0

                summary_data.append({
                    "Model": m_key,
                    "Die no.": d_key,
                    "Dimension": g_dim,
                    "n (sample size)": g_n,
                    "Data Count (n)": len(group),
                    "Cp": round(g_cp, 2),
                    "Cpk": round(g_cpk, 2),
                    "Status": "✅ Pass" if g_cpk >= 1.33 else "❌ Action Required"
                })

    if summary_data:
        df_summary = pd.DataFrame(summary_data)
        def highlight_cpk(val):
            if isinstance(val, (int, float)):
                color = 'green' if val >= 1.33 else 'red'
                return f'color: {color}; font-weight: bold;'
            return ''

        styled_summary = df_summary.style.map(highlight_cpk, subset=['Cpk', 'Cp'])
        st.dataframe(styled_summary, use_container_width=True, hide_index=True)
    else:
        st.info("💡 ยังไม่มีข้อมูลสำหรับสรุปภาพรวม")
