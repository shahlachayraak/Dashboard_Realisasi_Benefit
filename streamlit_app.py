import datetime
import os
import re
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu

# 1. Setup Halaman Streamlit
st.set_page_config(
    page_title="Dashboard Realisasi Benefit Prohire & Rehire",
    page_icon="📊",
    layout="wide",
)

EXCEL_PATH = "dashboard/raw.xlsx"


# 2. Custom CSS Tampilan Modern & Ukuran Card Konsisten
st.markdown(
    """
    <style>
    .stApp {
        background-color: #94D2FF;
    }
    
    /* CSS Card Status (4 Kotak Bawah) */
    .custom-card {
        padding: 14px 16px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
        color: #1e293b;
        min-height: 70px;          /* DIUBAH: Gunakan min-height agar fleksibel */
        height: 95px !important;     /* DIUBAH: Tinggi menyesuaikan konten */
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-sizing: border-box;
    }
    .custom-card-title {
        font-size: 18px;             /* DIUBAH: Diperkecil sedikit agar pas di layar kecil */
        font-weight: 700;
        line-height: 1;
        height: 34px;                  /* KUNCI: Membuat area judul selalu 2 baris untuk semua kotak */
        display: flex;
        align-items: center;            /* Agar judul 1 baris tetap rapi di tengah */
        word-wrap: break-word;
    }
    .custom-card-value {
        font-size: 26px;
        font-weight: 700;
        margin: 0;
        line-height: 1;
    }
    .card-green { background-color: #DCFCE7; border: 1px solid #BBF7D0; }
    .card-amber { background-color: #FFA1A1; border: 1px solid #FF5757; }
    .card-emerald { background-color: #DCFCE7; border: 1px solid #A7F3D0; }
    .card-rose { background-color: #FFA1A1; border: 1px solid #FF5757; }

    /* CSS Tabel Compact Inside Cards */
    .card-table-container {
        margin-top: 10px;
        border-radius: 6px;
    }
    .card-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        background-color: #ffffff;
    }
    .card-table th {
        background-color: #002060;
        color: #ffffff;
        text-align: left;
        padding: 6px 8px;
        font-size: 12px;
    }
    .card-table td {
        padding: 6px 8px;
        border-bottom: 1px solid #e2e8f0;
        color: #111;
    }
    .card-table tr:hover {
        background-color: #f8fafc;
    }

    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# Helper Fungsi Tanggal & Angka
def parse_date_value(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "NaT", "-"]:
        return datetime.date.today()
    try:
        dt = pd.to_datetime(val)
        return dt.date()
    except Exception:
        return datetime.date.today()


def parse_number(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "-"]:
        return 0.0
    cleaned = re.sub(r"[^\d.,]", "", str(val))
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except Exception:
        return 0.0


# Fungsi Membaca Data Excel
def load_data():
    if os.path.exists(EXCEL_PATH):
        df = pd.read_excel(EXCEL_PATH, dtype=str)
        df.columns = df.columns.str.strip()
        df = df.astype(object)

        cols_to_drop = [
            c for c in df.columns if str(c).strip().upper() in ["NO", "NO."]
        ]
        if cols_to_drop:
            df = df.drop(columns=cols_to_drop)

        if not any("AGAMA" in str(c).upper() for c in df.columns):
            df["Agama"] = "-"

        if not any(
            "STATUS PEGAWAI" in str(c).upper() or "STATUS_PEGAWAI" in str(c).upper()
            for c in df.columns
        ):
            df["Status Pegawai"] = "Aktif"

        df = df.reset_index(drop=True)
        return df
    else:
        st.error(f"File {EXCEL_PATH} tidak ditemukan di folder 'dashboard/'.")
        return pd.DataFrame()


# Fungsi Menyimpan Data ke Excel
def save_data(df):
    cols_to_drop = [
        c
        for c in df.columns
        if str(c).strip().upper() in ["NO", "NO.", "LABEL_DROPDOWN", "LABEL_SIMPEL", "ROW_INDEX", "DROPDOWN_TRACK"]
    ]
    df_clean = df.drop(columns=cols_to_drop, errors="ignore")
    df_clean.to_excel(EXCEL_PATH, index=False)


# LOGIKA H-14 UPCT SEBELUM END DATE KONTRAK
def is_upct_due_soon(row, col_upct, kw_selesai):
    status_upct = str(row.get(col_upct, ""))
    if re.search(kw_selesai, status_upct, re.IGNORECASE):
        return False, None, None

    status_peg = str(row.get("Status Pegawai", "Aktif")).lower()
    if status_peg != "aktif":
        return False, None, None

    today = datetime.date.today()
    end_date_val = parse_date_value(row.get("end_date", ""))

    selisih_hari = (end_date_val - today).days
    if 0 <= selisih_hari <= 14:  # H-14 Sebelum End Date
        return True, selisih_hari, end_date_val

    return False, None, None


# LOGIKA H-21 (3 MINGGU) MENJELANG HARI RAYA KEAGAMAAN
def is_thr_due_soon(row, col_thr, col_agama, kw_selesai, hari_raya_dates):
    status_thr = str(row.get(col_thr, ""))
    if re.search(kw_selesai, status_thr, re.IGNORECASE):
        return False, None, None, None

    status_peg = str(row.get("Status Pegawai", "Aktif")).lower()
    if status_peg != "aktif":
        return False, None, None, None

    agama = str(row.get(col_agama, "")).strip()
    today = datetime.date.today()

    # Cari tanggal hari raya berdasarkan agama
    tgl_raya = None
    nama_raya = ""
    for ag, (nama_h, d_str) in hari_raya_dates.items():
        if ag.lower() in agama.lower():
            tgl_raya = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
            nama_raya = nama_h
            break

    if not tgl_raya:
        return False, None, None, None

    selisih_hari = (tgl_raya - today).days
    if 0 <= selisih_hari <= 21:  # H-21 (3 Minggu) Sebelum Hari Raya
        return True, selisih_hari, tgl_raya, nama_raya

    return False, None, None, None


# Pop-Up Dialog Reminder
@st.dialog("🔔 Reminder Pengajuan UPCT & THR Pegawai")
def show_reminder_dialog(df_upct_pending, list_thr_due, col_nama_param, col_nip_param):
    
    # 1. Daftar Pegawai UPCT
    list_upct_peg = []
    if df_upct_pending is not None and not df_upct_pending.empty:
        for _, row in df_upct_pending.iterrows():
            nama = str(row.get(col_nama_param, "-")).strip()
            nip = str(row.get(col_nip_param, "-")).strip()
            list_upct_peg.append(f"<li><b>{nama}</b> (NIP: <code>{nip}</code>)</li>")

    # 2. Daftar Pegawai THR
    list_thr_peg = []
    if len(list_thr_due) > 0:
        for item in list_thr_due:
            nama = str(item.get("Nama", "-")).strip()
            nip = str(item.get("NIP", "-")).strip()
            list_thr_peg.append(f"<li><b>{nama}</b> (NIP: <code>{nip}</code>)</li>")

    # TAMPILKAN POP-UP DENGAN KOTAK BERWARNA
    with st.container():
        # Kotak UPCT (Kuning)
        if list_upct_peg:
            items_upct_html = "".join(sorted(list_upct_peg))
            st.markdown(
                f"""
                <div style="background-color: #FEFCE8; border-left: 5px solid #EAB308; padding: 14px 16px; border-radius: 8px; margin-bottom: 16px;">
                    <div style="color: #854D0E; font-weight: 700; font-size: 15px; margin-bottom: 8px;">
                        ⚠️ UPCT pegawai berikut sudah harus diproses!!
                    </div>
                    <ul style="margin: 0; padding-left: 20px; color: #713F12; line-height: 1.6;">
                        {items_upct_html}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )

        # Kotak THR (Hijau)
        if list_thr_peg:
            items_thr_html = "".join(sorted(list_thr_peg))
            st.markdown(
                f"""
                <div style="background-color: #F0FDF4; border-left: 5px solid #22C55E; padding: 14px 16px; border-radius: 8px; margin-bottom: 16px;">
                    <div style="color: #166534; font-weight: 700; font-size: 15px; margin-bottom: 8px;">
                        🎁 THR pegawai berikut sudah harus diproses!!
                    </div>
                    <ul style="margin: 0; padding-left: 20px; color: #14532D; line-height: 1.6;">
                        {items_thr_html}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )

        if not list_upct_peg and not list_thr_peg:
            st.info("Tidak ada pengajuan UPCT/THR yang perlu diproses saat ini.")

    if st.button("Tutup & Lanjutkan Ke Dashboard", type="primary"):
        st.session_state["reminder_shown"] = True
        st.rerun()


# --- BACA DATA MASTER ---
df = load_data()

# List Pilihan
LIST_AGAMA = ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Khonghucu", "-"]
LIST_STATUS_PEGAWAI = ["Aktif", "Selesai Kontrak / Tidak Aktif"]

# MASTER TANGGAL HARI RAYA KETETAPAN
HARI_RAYA_DATES = {
    "Islam": ("Idul Fitri", "2026-03-20"),
    "Kristen": ("Natal", "2026-12-25"),
    "Katolik": ("Natal", "2026-12-25"),
    "Hindu": ("Nyepi", "2026-03-19"),
    "Buddha": ("Waisak", "2026-05-31"),
    "Khonghucu": ("Imlek", "2026-02-17"),
}

# MENU NAVIGASI HORIZONTAL
menu = option_menu(
    menu_title=None,
    options=[
        "Dashboard & Tracking",
        "Tambah Pegawai Baru",
        "Edit Data Pegawai",
        "Hapus Data Pegawai",
    ],
    icons=["speedometer2", "person-plus-fill", "pencil-square", "trash-fill"],
    menu_icon="cast",
    default_index=0,
    orientation="horizontal",
    styles={
        "container": {
            "padding": "8px 10px!important",
            "background-color": "#ffffff",
            "border-radius": "12px",
            "border": "1px solid #e5e7eb",
            "margin-bottom": "20px",
            "height": "auto!important",
        },
        "icon": {"color": "#1447E6", "font-size": "16px"},
        "nav-link": {
            "font-size": "13px",
            "text-align": "center",
            "margin": "0px 4px",
            "padding": "10px 12px",
            "border-radius": "8px",
            "white-space": "nowrap",
            "--hover-color": "#f1f5f9",
        },
        "nav-link-selected": {
            "background-color": "#FFF100",
            "color": "#1447E6",
            "font-weight": "600",
        },
    },
)

# Helper untuk mendapatkan daftar pegawai unik (1 nama = 1 entitas dropdown)
def get_unique_pegawai_options(data_frame, col_n, col_p, col_s_peg):
    unique_names = sorted(list(set(data_frame[col_n].dropna().astype(str).str.strip())))
    options = []
    map_name_to_records = {}

    for name in unique_names:
        sub_df = data_frame[data_frame[col_n].astype(str).str.strip() == name]
        
        active_sub = sub_df[sub_df[col_s_peg].astype(str).str.lower() == "aktif"]
        if not active_sub.empty:
            active_nip = str(active_sub.iloc[-1][col_p]).strip()
        else:
            active_nip = str(sub_df.iloc[-1][col_p]).strip()

        label = f"{name} - {active_nip}"
        options.append(label)
        map_name_to_records[label] = name

    return options, map_name_to_records


if not df.empty:
    col_nip = next((c for c in df.columns if "NIP" in str(c).upper()), df.columns[0])
    col_perner = next((c for c in df.columns if "PERNER" in str(c).upper()), col_nip)
    col_nama = next((c for c in df.columns if "NAMA" in str(c).upper()), df.columns[1])
    col_jenis = next((c for c in df.columns if "JENIS" in str(c).upper()), "Jenis")
    col_agama = next((c for c in df.columns if "AGAMA" in str(c).upper()), "Agama")
    col_status_peg = next((c for c in df.columns if "STATUS PEGAWAI" in str(c).upper()), "Status Pegawai")
    col_sk = next((c for c in df.columns if "SK" in str(c).upper()), "no_sk")

    col_upct = next((c for c in df.columns if "STATUS" in c.upper() and "UPCT" in c.upper()), "status_upct")
    col_thr = next((c for c in df.columns if "STATUS" in c.upper() and "THR" in c.upper()), "status_thr")
    col_cat_upct = next((c for c in df.columns if "UPCT" in c.upper() and any(kw in str(c).upper() for kw in ["KENDALA", "CATATAN", "KET"])), "ket_upct")
    col_cat_thr = next((c for c in df.columns if "THR" in c.upper() and any(kw in str(c).upper() for kw in ["KENDALA", "CATATAN", "KET"])), "ket_thr")

    col_tgl_upct = next((c for c in df.columns if any(k in str(c).upper() for k in ["TGL", "TANGGAL"]) and "UPCT" in str(c).upper()), "tgl_realisasi_upct")
    col_tgl_thr = next((c for c in df.columns if any(k in str(c).upper() for k in ["TGL", "TANGGAL"]) and "THR" in str(c).upper()), "tgl_realisasi_thr")

    df["label_simpel"] = df[col_nama].astype(str).str.strip()

    # MENU 1: DASHBOARD MONITORING & TRACKING INTERAKTIF
    if menu == "Dashboard & Tracking":
        kw_selesai = "Sudah|Terlaksana|Selesai|Cair|Done|Ya"
        st.title("📊 Dashboard Monitoring Benefit")

        df_active = df[df[col_status_peg].astype(str).str.lower() == "aktif"].copy()

        # REMINDER THR KEMBALI MENGGUNAKAN H-21 SESUAI AGAMA
        list_thr_due = []
        for idx_row, row in df_active.iterrows():
            is_due_t, s_hari, tgl_r, nama_r = is_thr_due_soon(row, col_thr, col_agama, kw_selesai, HARI_RAYA_DATES)
            if is_due_t:
                list_thr_due.append({
                    "NIP": row.get(col_nip, "-"),
                    "Nama": row.get(col_nama, "-"),
                    "No SK": row.get(col_sk, "-"),
                    "Agama": row.get(col_agama, "-"),
                    "Hari Raya": f"{nama_r} ({tgl_r})",
                    "Sisa Hari": f"H-{s_hari}",
                    "Status THR": row.get(col_thr, "Belum/Gagal"),
                })

        upct_due_indices = []
        for idx_row, row in df_active.iterrows():
            is_due_u, _, _ = is_upct_due_soon(row, col_upct, kw_selesai)
            if is_due_u:
                upct_due_indices.append(idx_row)

        df_upct_pending = df_active.loc[upct_due_indices].copy() if upct_due_indices else pd.DataFrame()

        if (not df_upct_pending.empty or len(list_thr_due) > 0) and "reminder_shown" not in st.session_state:
            show_reminder_dialog(df_upct_pending, list_thr_due, col_nama, col_nip)

        # METRIK & CARD RINGKASAN
        st.subheader("📈 Ringkasan Realisasi Benefit (Pegawai Aktif)")

        df_prohire_active = df_active[df_active[col_jenis].astype(str).str.contains("prohire", case=False, na=False)]
        df_rehire_active = df_active[df_active[col_jenis].astype(str).str.contains("rehire", case=False, na=False)]

        def build_employee_table_html(df_sub):
            if df_sub.empty:
                return "<p style='margin-top:10px; color:#666; font-style:italic;'>Tidak ada data pegawai aktif</p>"

            rows_list = []
            for _, row in df_sub.iterrows():
                nip = str(row.get(col_nip, "-"))
                nama = str(row.get(col_nama, "-"))
                agama = str(row.get(col_agama, "-"))
                status_u = str(row.get(col_upct, "-"))
                status_t = str(row.get(col_thr, "-"))

                icon_u = "✅" if re.search(kw_selesai, status_u, re.IGNORECASE) else "❌"
                icon_t = "✅" if re.search(kw_selesai, status_t, re.IGNORECASE) else "❌"

                rows_list.append(
                    f"<tr>"
                    f"<td style='color: #444;'>{nip}</td>"
                    f"<td style='font-weight: 600;'>{nama}</td>"
                    f"<td>{agama}</td>"
                    f"<td style='text-align: center; font-size: 14px;'>{icon_u}</td>"
                    f"<td style='text-align: center; font-size: 14px;'>{icon_t}</td>"
                    f"</tr>"
                )

            rows_html = "".join(rows_list)

            table_html = (
                f'<div class="card-table-container">'
                f'<table class="card-table">'
                f'<thead><tr><th>NIP</th><th>Nama</th><th>Agama</th><th style="text-align: center;">UPCT</th><th style="text-align: center;">THR</th></tr></thead>'
                f'<tbody>{rows_html}</tbody>'
                f'</table></div>'
            )
            return table_html

        prohire_count = len(df_prohire_active)
        rehire_count = len(df_rehire_active)

        html_prohire = build_employee_table_html(df_prohire_active)
        html_rehire = build_employee_table_html(df_rehire_active)

        st.markdown(
            f"""<div style="display: flex; gap: 16px; width: 100%; align-items: stretch;">
<div style="flex: 1; background-color: #FDFF9C; padding: 16px; border-radius: 12px; border: 1px solid #BAE6FD; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08); display: flex; flex-direction: column;">
<div style="font-size: 18px; font-weight: 600; color: #1e293b; opacity: 0.85;">👤 Pegawai ProHire</div>
<div style="font-size: 26px; font-weight: 700; color: #1e293b; margin-bottom: 2px;">{prohire_count} Orang</div>
<div style="flex: 1;">
{html_prohire}
</div>
</div>
<div style="flex: 1; background-color: #FDFF9C; padding: 16px; border-radius: 12px; border: 1px solid #C7D2FE; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08); display: flex; flex-direction: column;">
<div style="font-size: 18px; font-weight: 600; color: #1e293b; opacity: 0.85;">🔄 Pegawai ReHire</div>
<div style="font-size: 26px; font-weight: 700; color: #1e293b; margin-bottom: 2px;">{rehire_count} Orang</div>
<div style="flex: 1;">
{html_rehire}
</div>
</div>
</div>""",
            unsafe_allow_html=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)

        total_pegawai = len(df_active)
        upct_selesai = df_active[col_upct].astype(str).str.contains(kw_selesai, case=False, na=False).sum()
        upct_belum = total_pegawai - upct_selesai

        thr_selesai = df_active[col_thr].astype(str).str.contains(kw_selesai, case=False, na=False).sum()
        thr_belum = total_pegawai - thr_selesai

        m3, m4, m5, m6 = st.columns(4)
        with m3:
            st.markdown(f"""<div class="custom-card card-green"><div class="custom-card-title">UPCT Terlaksana</div><div class="custom-card-value">{upct_selesai} Orang</div></div>""", unsafe_allow_html=True)
        with m4:
            st.markdown(f"""<div class="custom-card card-amber"><div class="custom-card-title">UPCT Belum Selesai</div><div class="custom-card-value">{upct_belum} Orang</div></div>""", unsafe_allow_html=True)
        with m5:
            st.markdown(f"""<div class="custom-card card-emerald"><div class="custom-card-title">THR Terlaksana</div><div class="custom-card-value">{thr_selesai} Orang</div></div>""", unsafe_allow_html=True)
        with m6:
            st.markdown(f"""<div class="custom-card card-rose"><div class="custom-card-title">THR Belum Selesai</div><div class="custom-card-value">{thr_belum} Orang</div></div>""", unsafe_allow_html=True)

        st.divider()

        # TRACKING (Bagian yang Diperbaiki)
        st.subheader("⚡ Update Status Tracking & Realisasi Pegawai (Aktif)")
        df_active["dropdown_track"] = df_active.apply(lambda r: f"{r.get(col_nama, '-')} | NIP: {r.get(col_nip, '-')}", axis=1)
        dropdown_options = df_active["dropdown_track"].tolist()

        if dropdown_options:
            selected_label = st.selectbox("Pilih Pegawai Aktif:", dropdown_options, key="tracking_select")
            idx = df_active[df_active["dropdown_track"] == selected_label].index[0]
            options_status = ["Terlaksana", "Belum Terlaksana"]

            tab_upct, tab_thr = st.tabs(["🌴 Tracking Benefit UPCT", "🎁 Tracking Benefit THR"])

            with tab_upct:
                with st.container(border=True):
                    if "flash_msg_upct" in st.session_state:
                        st.success(st.session_state["flash_msg_upct"])
                        del st.session_state["flash_msg_upct"]

                    st.markdown(f"**Update Status UPCT untuk:** `{df.loc[idx, col_nama]}` | NIP: `{df.loc[idx, col_nip]}` | SK: `{df.loc[idx, col_sk]}`")
                    
                    val_upct = df.loc[idx, col_upct]
                    idx_upct = options_status.index(val_upct) if val_upct in options_status else 1
                    
                    val_cat_upct = df.loc[idx, col_cat_upct]
                    curr_cat_upct = str(val_cat_upct) if pd.notna(val_cat_upct) and str(val_cat_upct).lower() != "nan" else ""
                    curr_tgl_upct = parse_date_value(df.loc[idx, col_tgl_upct])

                    new_upct = st.radio("Pilih Status UPCT:", options_status, index=idx_upct, horizontal=True, key=f"radio_upct_{idx}")
                    new_tgl_upct = st.date_input("Tanggal Realisasi UPCT:", value=curr_tgl_upct, format="YYYY-MM-DD", key=f"tgl_upct_{idx}")
                    new_cat_upct = st.text_input("Catatan UPCT:", value=curr_cat_upct, key=f"text_upct_{idx}")

                    if st.button("💾 Simpan Status UPCT", type="primary"):
                        df.loc[idx, col_upct] = new_upct
                        df.loc[idx, col_cat_upct] = new_cat_upct
                        df.loc[idx, col_tgl_upct] = new_tgl_upct.strftime("%Y-%m-%d") if new_tgl_upct else ""
                        save_data(df)
                        st.session_state["flash_msg_upct"] = f"✅ Status UPCT {df.loc[idx, col_nama]} ({df.loc[idx, col_sk]}) berhasil diperbarui!"
                        st.rerun()

            with tab_thr:
                with st.container(border=True):
                    if "flash_msg_thr" in st.session_state:
                        st.success(st.session_state["flash_msg_thr"])
                        del st.session_state["flash_msg_thr"]

                    st.markdown(f"**Update Status THR untuk:** `{df.loc[idx, col_nama]}` | NIP: `{df.loc[idx, col_nip]}` | SK: `{df.loc[idx, col_sk]}`")
                    
                    val_thr = df.loc[idx, col_thr]
                    idx_thr = options_status.index(val_thr) if val_thr in options_status else 1
                    
                    val_cat_thr = df.loc[idx, col_cat_thr]
                    curr_cat_thr = str(val_cat_thr) if pd.notna(val_cat_thr) and str(val_cat_thr).lower() != "nan" else ""
                    curr_tgl_thr = parse_date_value(df.loc[idx, col_tgl_thr])

                    new_thr = st.radio("Pilih Status THR:", options_status, index=idx_thr, horizontal=True, key=f"radio_thr_{idx}")
                    new_tgl_thr = st.date_input("Tanggal Realisasi THR:", value=curr_tgl_thr, format="YYYY-MM-DD", key=f"tgl_thr_{idx}")
                    new_cat_thr = st.text_input("Catatan THR:", value=curr_cat_thr, key=f"text_thr_{idx}")

                    if st.button("💾 Simpan Status THR", type="primary"):
                        df.loc[idx, col_thr] = new_thr
                        df.loc[idx, col_cat_thr] = new_cat_thr
                        df.loc[idx, col_tgl_thr] = new_tgl_thr.strftime("%Y-%m-%d") if new_tgl_thr else ""
                        save_data(df)
                        st.session_state["flash_msg_thr"] = f"✅ Status THR {df.loc[idx, col_nama]} ({df.loc[idx, col_sk]}) berhasil diperbarui!"
                        st.rerun()

        st.divider()

        # TABEL MASTER DATA
        st.subheader("📋 Master Data Pegawai (Termasuk Riwayat SK Lama)")
        df_display = df.copy()

        for c in df_display.columns:
            if any(kw in str(c).upper() for kw in ["DATE", "TGL", "TANGGAL", "START", "END"]):
                df_display[c] = df_display[c].apply(
                    lambda val: parse_date_value(val).strftime("%Y-%m-%d")
                    if pd.notna(val) and str(val).strip() not in ["", "nan", "NaT", "-"]
                    else "-"
                )

        col_sort_start = next((c for c in df_display.columns if "START" in str(c).upper()), None)
        if col_nama in df_display.columns:
            sort_cols = [col_nama]
            sort_asc = [True]
            if col_sort_start:
                sort_cols.append(col_sort_start)
                sort_asc.append(False)
            df_display = df_display.sort_values(by=sort_cols, ascending=sort_asc).reset_index(drop=True)

        df_display[col_nama] = df_display[col_nama].mask(df_display[col_nama].duplicated(), "")

        priority_front = [col_nama, col_nip, col_perner, col_jenis, col_agama, "Status Pegawai", "gaji_pokok"]
        priority_back = [col_sk, "start_date", "end_date", "masa_kontrak"]

        all_cols = list(df_display.columns)
        front_cols = [c for c in priority_front if c in all_cols]
        back_cols = [c for c in priority_back if c in all_cols]

        middle_cols = [
            c for c in all_cols
            if c not in front_cols
            and c not in back_cols
            and c not in ["label_dropdown", "label_simpel", "dropdown_track"]
        ]

        ordered_cols = front_cols + middle_cols + back_cols
        df_display = df_display[ordered_cols]

        df_display.insert(0, "No", range(1, len(df_display) + 1))
        st.dataframe(df_display, use_container_width=True, hide_index=True)

    # MENU 2: TAMBAH DATA BARU
    elif menu == "Tambah Pegawai Baru":
        st.title("➕ Tambah Record Pegawai Baru")

        if "flash_msg" in st.session_state:
            st.success(st.session_state["flash_msg"])
            del st.session_state["flash_msg"]

        col1, col2 = st.columns(2)
        with col1:
            nip = st.text_input("NIP Baru", key="add_nip")
            perner = st.text_input("PERNER", key="add_perner")
            nama = st.text_input("Nama Lengkap Pegawai", key="add_nama")
            jenis = st.selectbox("Jenis Pegawai", ["ProHire", "ReHire"], key="add_jenis")
            agama = st.selectbox("Agama", LIST_AGAMA, key="add_agama")
            gaji_pokok = st.number_input("Gaji Pokok (Rp)", min_value=0.0, step=100000.0, format="%.0f", key="add_gaji")

        with col2:
            bobot_upct = st.number_input("Bobot UPCT", min_value=0.0, max_value=10.0, value=0.5, step=0.5, key="add_b_upct")
            bobot_thr = st.number_input("Bobot THR", min_value=0.0, max_value=10.0, value=1.0, step=0.5, key="add_b_thr")
            no_sk = st.text_input("No. Dokumen SK", key="add_sk")
            start_date = st.date_input("Start Date", value=datetime.date.today(), key="add_start")
            end_date = st.date_input("End Date", value=datetime.date.today(), key="add_end")

            selisih_hari = (end_date - start_date).days
            masa_kontrak_auto = f"{round(selisih_hari / 30)} Bulan ({selisih_hari} Hari)" if selisih_hari > 0 else "0 Hari"
            st.text_input("Masa Kontrak (Otomatis Terhitung):", value=masa_kontrak_auto, disabled=True)

        if st.button("💾 Simpan Pegawai Baru", type="primary"):
            if str(nip).strip() != "" or str(nama).strip() != "":
                new_data = {
                    col_nip: str(nip).strip(),
                    col_perner: str(perner).strip(),
                    col_nama: str(nama).strip(),
                    "Jenis": str(jenis),
                    "Agama": str(agama),
                    "Status Pegawai": "Aktif",
                    "gaji_pokok": gaji_pokok,
                    "bobot_upct": bobot_upct,
                    "nilai_upct": gaji_pokok * bobot_upct,
                    "bobot_thr": bobot_thr,
                    "nilai_thr": gaji_pokok * bobot_thr,
                    col_sk: str(no_sk).strip(),
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "masa_kontrak": masa_kontrak_auto,
                }
                for col in df.columns:
                    if col not in new_data and col not in ["label_dropdown", "label_simpel"]:
                        new_data[col] = "Belum/Gagal" if "STATUS" in str(col).upper() else "-"

                new_row = pd.DataFrame([new_data])
                df_updated = pd.concat([df, new_row], ignore_index=True)
                save_data(df_updated)

                st.session_state["flash_msg"] = f"🎉 Data Pegawai Baru ({nama}) BERHASIL DISIMPAN!"
                st.rerun()
            else:
                st.error("⚠️ Kolom NIP dan Nama Wajib Diisi!")

    # MENU 3: EDIT DATA PEGAWAI
    elif menu == "Edit Data Pegawai":
        st.title("✏️ Edit Data & Perpanjang Kontrak Pegawai")

        if "flash_msg" in st.session_state:
            st.success(st.session_state["flash_msg"])
            del st.session_state["flash_msg"]

        pegawai_options, map_name = get_unique_pegawai_options(df, col_nama, col_nip, col_status_peg)

        if pegawai_options:
            selected_pegawai_label = st.selectbox(
                "Pilih Pegawai (Nama / NIP):",
                options=pegawai_options,
                key="edit_pegawai_select"
            )

            selected_nama_extracted = map_name[selected_pegawai_label]
            df_pegawai = df[df[col_nama].astype(str).str.strip() == selected_nama_extracted].copy()

            sk_options = df_pegawai.apply(
                lambda r: f"SK: {r.get(col_sk, '-')} | NIP: {r.get(col_nip, '-')} ({r.get('Status Pegawai', 'Aktif')})",
                axis=1,
            ).tolist()

            selected_sk_label = st.selectbox(
                "Pilih Riwayat SK:",
                options=sk_options,
                key="edit_sk_select"
            )

            selected_idx = df_pegawai.index[sk_options.index(selected_sk_label)]
            row_data = df.iloc[selected_idx]

            st.info(
                f"✍️ Menampilkan Data Pegawai: **{row_data.get(col_nama, '-')}** "
                f"| NIP SK Ini: `{row_data.get(col_nip, '-')}` | SK: `{row_data.get(col_sk, '-')}` | Status: **{row_data.get('Status Pegawai', 'Aktif')}**"
            )

            col_e1, col_e2 = st.columns(2)

            with col_e1:
                val_nip = st.text_input("NIP", value=str(row_data.get(col_nip, "")), key=f"e_nip_{selected_idx}")
                val_perner = st.text_input("PERNER", value=str(row_data.get(col_perner, "")), key=f"e_perner_{selected_idx}")
                val_nama = st.text_input("Nama Lengkap", value=str(row_data.get(col_nama, "")), key=f"e_nama_{selected_idx}")

                curr_j = str(row_data.get("Jenis", "ProHire"))
                val_jenis = st.selectbox("Jenis Pegawai", ["ProHire", "ReHire"], index=0 if "pro" in curr_j.lower() else 1, key=f"e_jenis_{selected_idx}")

                curr_a = str(row_data.get("Agama", "-"))
                val_agama = st.selectbox("Agama", LIST_AGAMA, index=LIST_AGAMA.index(curr_a) if curr_a in LIST_AGAMA else len(LIST_AGAMA) - 1, key=f"e_agama_{selected_idx}")

                curr_sp = str(row_data.get("Status Pegawai", "Aktif"))
                val_status_peg = st.selectbox("Status Record SK Ini", LIST_STATUS_PEGAWAI, index=0 if "aktif" in curr_sp.lower() and "tidak" not in curr_sp.lower() and "selesai" not in curr_sp.lower() else 1, key=f"e_sp_{selected_idx}")

            with col_e2:
                val_gaji = st.number_input("Gaji Pokok (Rp)", value=parse_number(row_data.get("gaji_pokok", 0)), step=100000.0, format="%.0f", key=f"e_gaji_{selected_idx}")
                val_b_upct = st.number_input("Bobot UPCT", value=parse_number(row_data.get("bobot_upct", 0.5)), step=0.5, key=f"e_b_upct_{selected_idx}")
                val_b_thr = st.number_input("Bobot THR", value=parse_number(row_data.get("bobot_thr", 1.0)), step=0.5, key=f"e_b_thr_{selected_idx}")
                val_sk = st.text_input("No. Dokumen SK", value=str(row_data.get(col_sk, "")), key=f"e_sk_{selected_idx}")
                
                start_val = parse_date_value(row_data.get("start_date", ""))
                end_val = parse_date_value(row_data.get("end_date", ""))
                
                val_start = st.date_input("Start Date", value=start_val, key=f"e_start_{selected_idx}")
                val_end = st.date_input("End Date", value=end_val, key=f"e_end_{selected_idx}")

                selisih_hari = (val_end - val_start).days
                val_masa_kontrak = f"{round(selisih_hari / 30)} Bulan ({selisih_hari} Hari)" if selisih_hari > 0 else "0 Hari"
                st.text_input("Masa Kontrak (Otomatis):", value=val_masa_kontrak, disabled=True, key=f"e_masa_{selected_idx}")

            if st.button("💾 Simpan Perubahan Data", type="primary"):
                df.loc[selected_idx, col_nip] = str(val_nip).strip()
                df.loc[selected_idx, col_perner] = str(val_perner).strip()
                df.loc[selected_idx, col_nama] = str(val_nama).strip()
                df.loc[selected_idx, "Jenis"] = str(val_jenis)
                df.loc[selected_idx, "Agama"] = str(val_agama)
                df.loc[selected_idx, "Status Pegawai"] = str(val_status_peg)
                df.loc[selected_idx, "gaji_pokok"] = val_gaji
                df.loc[selected_idx, "bobot_upct"] = val_b_upct
                df.loc[selected_idx, "nilai_upct"] = val_gaji * val_b_upct
                df.loc[selected_idx, "bobot_thr"] = val_b_thr
                df.loc[selected_idx, "nilai_thr"] = val_gaji * val_b_thr
                df.loc[selected_idx, col_sk] = str(val_sk).strip()
                df.loc[selected_idx, "start_date"] = val_start.strftime("%Y-%m-%d")
                df.loc[selected_idx, "end_date"] = val_end.strftime("%Y-%m-%d")
                df.loc[selected_idx, "masa_kontrak"] = val_masa_kontrak

                save_data(df)
                st.session_state["flash_msg"] = f"✅ Perubahan Data SK {val_sk} BERHASIL DISIMPAN!"
                st.rerun()

    # MENU 4: HAPUS DATA PEGAWAI
    elif menu == "Hapus Data Pegawai":
        st.title("🗑️ Hapus Data Pegawai")

        if "flash_msg" in st.session_state:
            st.success(st.session_state["flash_msg"])
            del st.session_state["flash_msg"]

        pegawai_options, map_name = get_unique_pegawai_options(df, col_nama, col_nip, col_status_peg)

        if pegawai_options:
            selected_pegawai_label = st.selectbox(
                "Pilih Pegawai yang Akan Dihapus:",
                options=pegawai_options,
                key="del_pegawai_select"
            )

            selected_nama_extracted = map_name[selected_pegawai_label]
            df_pegawai = df[df[col_nama].astype(str).str.strip() == selected_nama_extracted].copy()

            st.warning(f"⚠️ Menemukan {len(df_pegawai)} record riwayat SK untuk pegawai ini.")
            
            hapus_mode = st.radio(
                "Pilihan Penghapusan:",
                ["Hapus SK Spesifik", "Hapus Seluruh Data Pegawai Ini"],
                horizontal=True
            )

            if hapus_mode == "Hapus SK Spesifik":
                sk_options = df_pegawai.apply(
                    lambda r: f"SK: {r.get(col_sk, '-')} | NIP: {r.get(col_nip, '-')} ({r.get('Status Pegawai', 'Aktif')})",
                    axis=1,
                ).tolist()

                selected_sk_label = st.selectbox(
                    "Pilih Dokumen SK yang Akan Dihapus:",
                    options=sk_options,
                    key="del_sk_select"
                )

                selected_idx = df_pegawai.index[sk_options.index(selected_sk_label)]

                if st.button("❌ Hapus Record SK Ini", type="primary"):
                    df_updated = df.drop(index=selected_idx).reset_index(drop=True)
                    save_data(df_updated)
                    st.session_state["flash_msg"] = f"✅ Record SK berhasil dihapus!"
                    st.rerun()

            else:
                if st.button("Hapus SELURUH Data Pegawai Ini", type="primary"):
                    df_updated = df[df[col_nama].astype(str).str.strip() != selected_nama_extracted].reset_index(drop=True)
                    save_data(df_updated)
                    st.session_state["flash_msg"] = f"✅ Seluruh record untuk {selected_nama_extracted} berhasil dihapus!"
                    st.rerun()