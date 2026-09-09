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


# 2. Custom CSS Tampilan Modern
st.markdown(
    """
    <style>
    .stApp {
        background-color: #94D2FF;
    }
    .custom-card {
        padding: 16px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
        margin-bottom: 10px;
        color: #1e293b;
    }
    .custom-card-title {
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 6px;
        opacity: 0.85;
    }
    .custom-card-value {
        font-size: 28px;
        font-weight: 700;
        margin: 0;
    }
    .card-blue { background-color: #FDFF9C; border: 1px solid #BAE6FD; }
    .card-indigo { background-color: #FDFF9C; border: 1px solid #C7D2FE; }
    .card-green { background-color: #DCFCE7; border: 1px solid #BBF7D0; }
    .card-amber { background-color: #FFA1A1; border: 1px solid #FF5757; }
    .card-emerald { background-color: #DCFCE7; border: 1px solid #A7F3D0; }
    .card-rose { background-color: #FFA1A1; border: 1px solid #FF5757; }

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


# -------------------------------------------------------------
# LOGIKA TANGGAL HARI RAYA & NOTIFIKASI H-21
# -------------------------------------------------------------
def get_hari_raya_date(agama, current_year):
    agama_clean = str(agama).strip().lower()

    if "kristen" in agama_clean or "katolik" in agama_clean:
        return datetime.date(current_year, 12, 25)
    elif "islam" in agama_clean:
        idul_fitri_dates = {
            2025: datetime.date(2025, 3, 31),
            2026: datetime.date(2026, 3, 20),
            2027: datetime.date(2027, 3, 10),
        }
        return idul_fitri_dates.get(
            current_year, datetime.date(current_year, 3, 20)
        )

    return None


def is_thr_due_soon(row, col_thr, kw_selesai):
    status_thr = str(row.get(col_thr, ""))
    if re.search(kw_selesai, status_thr, re.IGNORECASE):
        return False, None, None

    agama = row.get("Agama", "-")
    status_peg = str(row.get("Status Pegawai", "Aktif")).lower()
    if status_peg != "aktif":
        return False, None, None

    today = datetime.date.today()
    target_date = get_hari_raya_date(agama, today.year)

    if target_date:
        if today > target_date:
            target_date = get_hari_raya_date(agama, today.year + 1)

        if target_date:
            selisih_hari = (target_date - today).days
            if 0 <= selisih_hari <= 21:
                return True, selisih_hari, target_date

    return False, None, None


# Pop-Up Dialog Reminder
@st.dialog("🔔 Reminder Pengajuan UPCT & THR Pegawai")
def show_reminder_dialog(df_upct_pending, list_thr_due):
    st.warning("⚠️ **Terdapat perhatian khusus untuk realisasi benefit:**")

    if not df_upct_pending.empty:
        st.markdown(
            f"• **UPCT Pending**: Ada **{len(df_upct_pending)}** pegawai yang UPCT-nya masih *Belum Selesai*."
        )

    if len(list_thr_due) > 0:
        st.markdown(
            f"• **THR Mendekati Hari Raya (H-21)**: Ada **{len(list_thr_due)}** pegawai yang akan merayakan Hari Raya dalam ≤ 3 minggu ke depan dan THR-nya belum diproses!"
        )

    st.write("---")
    st.markdown("#### 📋 Daftar Pegawai Perlu Tindak Lanjut:")

    if len(list_thr_due) > 0:
        st.markdown("**1. Pengajuan THR (Mendekati H-21 Hari Raya):**")
        df_thr_show = pd.DataFrame(list_thr_due)
        st.dataframe(df_thr_show, use_container_width=True, hide_index=True)

    if not df_upct_pending.empty:
        st.markdown("**2. Pengajuan UPCT Belum Selesai:**")
        cols_show = [
            c
            for c in df_upct_pending.columns
            if any(
                kw in str(c).upper()
                for kw in ["NIP", "PERNER", "NAMA", "JENIS", "AGAMA", "STATUS"]
            )
        ]
        st.dataframe(
            df_upct_pending[list(dict.fromkeys(cols_show))],
            use_container_width=True,
            hide_index=True,
        )

    if st.button("Tutup & Lanjutkan Ke Dashboard", type="primary"):
        st.session_state["reminder_shown"] = True
        st.rerun()


# --- BACA DATA MASTER ---
df = load_data()

# List Pilihan
LIST_AGAMA = ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Khonghucu", "-"]
LIST_STATUS_PEGAWAI = ["Aktif", "Selesai Kontrak / Tidak Aktif"]

# -------------------------------------------------------------
# MENU NAVIGASI HORIZONTAL
# -------------------------------------------------------------
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

    # Grouping kunci identitas berdasarkan NAMA agar satu nama selalu menyatu
    df["label_simpel"] = df[col_nama].astype(str).str.strip()

    # -------------------------------------------------------------
    # MENU 1: DASHBOARD MONITORING & TRACKING INTERAKTIF
    # -------------------------------------------------------------
    if menu == "Dashboard & Tracking":
        kw_selesai = "Sudah|Terlaksana|Selesai|Cair|Done|Ya"
        st.title("📊 Dashboard Monitoring Benefit")

        df_active = df[df[col_status_peg].astype(str).str.lower() == "aktif"].copy()

        # REMINDER
        list_thr_due = []
        for idx_row, row in df_active.iterrows():
            is_due, days_left, tgl_hr = is_thr_due_soon(row, col_thr, kw_selesai)
            if is_due:
                list_thr_due.append({
                    "NIP": row.get(col_nip, "-"),
                    "Nama": row.get(col_nama, "-"),
                    "No SK": row.get(col_sk, "-"),
                    "Agama": row.get(col_agama, "-"),
                    "Status THR": row.get(col_thr, "-"),
                    "Estimasi Hari Raya": tgl_hr.strftime("%d-%m-%Y"),
                    "Sisa Waktu": f"H-{days_left} Hari Lagi",
                })

        df_upct_pending = df_active[~df_active[col_upct].astype(str).str.contains(kw_selesai, case=False, na=False)]

        if (not df_upct_pending.empty or len(list_thr_due) > 0) and "reminder_shown" not in st.session_state:
            show_reminder_dialog(df_upct_pending, list_thr_due)

        if len(list_thr_due) > 0:
            st.error(f"🚨 **REMINDER PENGAJUAN THR (H-21):** Terdapat **{len(list_thr_due)}** pegawai yang akan merayakan Hari Raya (≤ 3 Minggu) dan THR-nya **Belum Dicairkan/Diajukan**!")

        # METRIK
        st.subheader("📈 Ringkasan Realisasi Benefit (Pegawai Aktif)")
        total_pegawai = len(df_active)
        total_prohire = df_active[col_jenis].astype(str).str.contains("prohire", case=False, na=False).sum()
        total_rehire = df_active[col_jenis].astype(str).str.contains("rehire", case=False, na=False).sum()

        upct_selesai = df_active[col_upct].astype(str).str.contains(kw_selesai, case=False, na=False).sum()
        upct_belum = total_pegawai - upct_selesai

        thr_selesai = df_active[col_thr].astype(str).str.contains(kw_selesai, case=False, na=False).sum()
        thr_belum = total_pegawai - thr_selesai

        m1, m2 = st.columns(2)
        with m1:
            st.markdown(f"""<div class="custom-card card-blue"><div class="custom-card-title">👤 Pegawai ProHire</div><div class="custom-card-value">{total_prohire} Orang</div></div>""", unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class="custom-card card-indigo"><div class="custom-card-title">🔄 Pegawai ReHire</div><div class="custom-card-value">{total_rehire} Orang</div></div>""", unsafe_allow_html=True)

        m3, m4, m5, m6 = st.columns(4)
        with m3:
            st.markdown(f"""<div class="custom-card card-green"><div class="custom-card-title">🌴 UPCT Terlaksana</div><div class="custom-card-value">{upct_selesai} Orang</div></div>""", unsafe_allow_html=True)
        with m4:
            st.markdown(f"""<div class="custom-card card-amber"><div class="custom-card-title">⏳ UPCT Belum Selesai</div><div class="custom-card-value">{upct_belum} Orang</div></div>""", unsafe_allow_html=True)
        with m5:
            st.markdown(f"""<div class="custom-card card-emerald"><div class="custom-card-title">🎁 THR Terlaksana</div><div class="custom-card-value">{thr_selesai} Orang</div></div>""", unsafe_allow_html=True)
        with m6:
            st.markdown(f"""<div class="custom-card card-rose"><div class="custom-card-title">⏳ THR Belum Selesai</div><div class="custom-card-value">{thr_belum} Orang</div></div>""", unsafe_allow_html=True)

        st.divider()

        # TRACKING
        st.subheader("⚡ Update Status Tracking & Realisasi Pegawai (Aktif)")
        df_active["dropdown_track"] = df_active.apply(lambda r: f"{r.get(col_nama, '-')} | NIP: {r.get(col_nip, '-')} | SK: {r.get(col_sk, '-')}", axis=1)
        dropdown_options = df_active["dropdown_track"].tolist()
        
        if dropdown_options:
            selected_label = st.selectbox("Pilih Pegawai Aktif:", dropdown_options, key="tracking_select")
            idx = df_active[df_active["dropdown_track"] == selected_label].index[0]
            options_status = ["Terlaksana", "On Progress", "Belum/Gagal"]

            tab_upct, tab_thr = st.tabs(["🌴 Tracking Benefit UPCT", "🎁 Tracking Benefit THR"])

            with tab_upct:
                with st.container(border=True):
                    if "flash_msg_upct" in st.session_state:
                        st.success(st.session_state["flash_msg_upct"])
                        del st.session_state["flash_msg_upct"]

                    st.markdown(f"**Update Status UPCT untuk:** `{df.loc[idx, col_nama]}` | NIP: `{df.loc[idx, col_nip]}` | SK: `{df.loc[idx, col_sk]}`")
                    curr_upct = df.loc[idx, col_upct] if df.loc[idx, col_upct] in options_status else "Belum/Gagal"
                    curr_cat_upct = str(df.loc[idx, col_cat_upct])
                    curr_tgl_upct = parse_date_value(df.loc[idx, col_tgl_upct])

                    new_upct = st.radio("Pilih Status UPCT:", options_status, index=(options_status.index(curr_upct) if curr_upct in options_status else 2), horizontal=True, key=f"radio_upct_{idx}")
                    new_tgl_upct = st.date_input("Tanggal Realisasi UPCT:", value=curr_tgl_upct, format="YYYY-MM-DD", key=f"tgl_upct_{idx}")
                    new_cat_upct = st.text_input("Catatan UPCT:", value=curr_cat_upct, key=f"text_upct_{idx}")

                    if st.button("💾 Simpan Status UPCT", type="primary"):
                        df.loc[idx, col_upct] = new_upct
                        df.loc[idx, col_cat_upct] = new_cat_upct
                        df.loc[idx, col_tgl_upct] = new_tgl_upct.strftime("%Y-%m-%d")
                        save_data(df)
                        st.session_state["flash_msg_upct"] = f"✅ Status UPCT {df.loc[idx, col_nama]} ({df.loc[idx, col_sk]}) berhasil diperbarui!"
                        st.rerun()

            with tab_thr:
                with st.container(border=True):
                    if "flash_msg_thr" in st.session_state:
                        st.success(st.session_state["flash_msg_thr"])
                        del st.session_state["flash_msg_thr"]

                    st.markdown(f"**Update Status THR untuk:** `{df.loc[idx, col_nama]}` | NIP: `{df.loc[idx, col_nip]}` | SK: `{df.loc[idx, col_sk]}`")
                    curr_thr = df.loc[idx, col_thr] if df.loc[idx, col_thr] in options_status else "Belum/Gagal"
                    curr_cat_thr = str(df.loc[idx, col_cat_thr])
                    curr_tgl_thr = parse_date_value(df.loc[idx, col_tgl_thr])

                    new_thr = st.radio("Pilih Status THR:", options_status, index=(options_status.index(curr_thr) if curr_thr in options_status else 2), horizontal=True, key=f"radio_thr_{idx}")
                    new_tgl_thr = st.date_input("Tanggal Realisasi THR:", value=curr_tgl_thr, format="YYYY-MM-DD", key=f"tgl_thr_{idx}")
                    new_cat_thr = st.text_input("Catatan THR:", value=curr_cat_thr, key=f"text_thr_{idx}")

                    if st.button("💾 Simpan Status THR", type="primary"):
                        df.loc[idx, col_thr] = new_thr
                        df.loc[idx, col_cat_thr] = new_cat_thr
                        df.loc[idx, col_tgl_thr] = new_tgl_thr.strftime("%Y-%m-%d")
                        save_data(df)
                        st.session_state["flash_msg_thr"] = f"✅ Status THR {df.loc[idx, col_nama]} ({df.loc[idx, col_sk]}) berhasil diperbarui!"
                        st.rerun()

        st.divider()

        # TABEL MASTER DATA (NAMA DISAMARKAN PADA BARIS DUPLIKAT)
        st.subheader("📋 Master Data Pegawai (Termasuk Riwayat SK Lama)")
        df_display = df.copy()

        # 1. Format Tanggal
        for c in df_display.columns:
            if any(kw in str(c).upper() for kw in ["DATE", "TGL", "TANGGAL", "START", "END"]):
                df_display[c] = df_display[c].apply(
                    lambda val: parse_date_value(val).strftime("%Y-%m-%d")
                    if pd.notna(val) and str(val).strip() not in ["", "nan", "NaT", "-"]
                    else "-"
                )

        # 2. Urutkan berdasarkan Nama Pegawai dan Tanggal SK Terbaru
        col_sort_start = next((c for c in df_display.columns if "START" in str(c).upper()), None)
        if col_nama in df_display.columns:
            sort_cols = [col_nama]
            sort_asc = [True]
            if col_sort_start:
                sort_cols.append(col_sort_start)
                sort_asc.append(False)
            df_display = df_display.sort_values(by=sort_cols, ascending=sort_asc).reset_index(drop=True)

        # 3. Kosongkan teks nama jika nama di baris tersebut sama dengan baris sebelumnya
        df_display[col_nama] = df_display[col_nama].mask(df_display[col_nama].duplicated(), "")

        # 4. Susun Ulang Kolom
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

    # -------------------------------------------------------------
    # MENU 2: TAMBAH DATA BARU
    # -------------------------------------------------------------
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

    # -------------------------------------------------------------
    # MENU 3: EDIT DATA & PERPANJANG KONTRAK
    # -------------------------------------------------------------
    elif menu == "Edit Data Pegawai":
        st.title("✏️ Edit Data & Perpanjang Kontrak Pegawai")

        if "flash_msg" in st.session_state:
            st.success(st.session_state["flash_msg"])
            del st.session_state["flash_msg"]

        # 1. Pilih Pegawai Unique (Berdasarkan NAMA)
        unique_pegawai = df["label_simpel"].drop_duplicates().tolist()

        selected_pegawai_label = st.selectbox(
            "Pilih Nama Pegawai:",
            options=unique_pegawai,
            key="edit_pegawai_select"
        )

        df_pegawai = df[df["label_simpel"] == selected_pegawai_label].copy()

        # 2. Dropdown Sekunder Pilih SK/NIP
        sk_options = df_pegawai.apply(
            lambda r: f"SK: {r.get(col_sk, '-')} | NIP: {r.get(col_nip, '-')} [{r.get('Status Pegawai', 'Aktif')}] ({parse_date_value(r.get('start_date', '')).strftime('%Y-%m-%d')} s.d {parse_date_value(r.get('end_date', '')).strftime('%Y-%m-%d')})",
            axis=1,
        ).tolist()

        selected_sk_label = st.selectbox(
            "Pilih Riwayat SK / NIP yang Ingin Di-edit / Diperpanjang:",
            options=sk_options,
            key="edit_sk_select"
        )

        selected_idx = df_pegawai.index[sk_options.index(selected_sk_label)]
        row_data = df.iloc[selected_idx]

        st.info(
            f"✍️ Menampilkan Data: **{row_data.get(col_nama, '-')}** "
            f"| SK Selected: `{row_data.get(col_sk, '-')}` | Status SK: **{row_data.get('Status Pegawai', 'Aktif')}**"
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
            val_no_sk = st.text_input("No. SK / Addendum Baru/Lama", value=str(row_data.get(col_sk, "")), key=f"e_sk_{selected_idx}")

            start_d_val = parse_date_value(row_data.get("start_date", ""))
            end_d_val = parse_date_value(row_data.get("end_date", ""))

            val_start = st.date_input("Start Date", value=start_d_val, key=f"e_start_{selected_idx}")
            val_end = st.date_input("End Date", value=end_d_val, key=f"e_end_{selected_idx}")

            selisih = (val_end - val_start).days
            m_kontrak_auto = f"{round(selisih / 30)} Bulan ({selisih} Hari)" if selisih > 0 else "0 Hari"
            st.text_input("Masa Kontrak (Terhitung Otomatis):", value=m_kontrak_auto, disabled=True, key=f"e_mk_{selected_idx}")

        st.markdown("<br>", unsafe_allow_html=True)

        btn_col1, btn_col2 = st.columns(2)

        # TOMBOL 1: PERBARUI RECORD DIPILIH
        with btn_col1:
            if st.button("🔄 Perbarui Record SK Ini Saja", use_container_width=True):
                old_nama_val = str(row_data.get(col_nama, "")).strip()
                indices_to_sync = df[df[col_nama].astype(str).str.strip() == old_nama_val].index

                # Sinkronkan Nama untuk semua record pegawai tersebut
                for idx_sync in indices_to_sync:
                    df.at[idx_sync, col_nama] = str(val_nama).strip()

                df.at[selected_idx, col_nip] = str(val_nip).strip()
                df.at[selected_idx, col_perner] = str(val_perner).strip()
                df.at[selected_idx, "Jenis"] = str(val_jenis)
                df.at[selected_idx, "Agama"] = str(val_agama)
                df.at[selected_idx, "Status Pegawai"] = "Aktif" if "aktif" in val_status_peg.lower() and "tidak" not in val_status_peg.lower() and "selesai" not in val_status_peg.lower() else "Selesai Kontrak / Tidak Aktif"
                df.at[selected_idx, "gaji_pokok"] = val_gaji
                df.at[selected_idx, "bobot_upct"] = val_b_upct
                df.at[selected_idx, "nilai_upct"] = val_gaji * val_b_upct
                df.at[selected_idx, "bobot_thr"] = val_b_thr
                df.at[selected_idx, "nilai_thr"] = val_gaji * val_b_thr
                df.at[selected_idx, col_sk] = str(val_no_sk).strip()
                df.at[selected_idx, "start_date"] = val_start.strftime("%Y-%m-%d")
                df.at[selected_idx, "end_date"] = val_end.strftime("%Y-%m-%d")
                df.at[selected_idx, "masa_kontrak"] = m_kontrak_auto

                save_data(df)
                st.session_state["flash_msg"] = f"✅ Data SK {val_no_sk} ({val_nama}) berhasil diperbarui!"
                st.rerun()

        # TOMBOL 2: PERPANJANG KONTRAK BARU
        with btn_col2:
            if st.button("➕ Simpan Sebagai Perpanjangan SK Baru", type="primary", use_container_width=True):
                df.at[selected_idx, "Status Pegawai"] = "Selesai Kontrak / Tidak Aktif"

                new_row_data = df.iloc[selected_idx].to_dict()

                new_row_data[col_nip] = str(val_nip).strip()
                new_row_data[col_perner] = str(val_perner).strip()
                new_row_data[col_nama] = str(val_nama).strip()
                new_row_data["Jenis"] = str(val_jenis)
                new_row_data["Agama"] = str(val_agama)

                new_row_data["Status Pegawai"] = "Aktif"
                new_row_data["gaji_pokok"] = val_gaji
                new_row_data["bobot_upct"] = val_b_upct
                new_row_data["nilai_upct"] = val_gaji * val_b_upct
                new_row_data["bobot_thr"] = val_b_thr
                new_row_data["nilai_thr"] = val_gaji * val_b_thr
                new_row_data[col_sk] = str(val_no_sk).strip()
                new_row_data["start_date"] = val_start.strftime("%Y-%m-%d")
                new_row_data["end_date"] = val_end.strftime("%Y-%m-%d")
                new_row_data["masa_kontrak"] = m_kontrak_auto

                new_row_data[col_upct] = "Belum/Gagal"
                new_row_data[col_thr] = "Belum/Gagal"
                new_row_data[col_cat_upct] = "-"
                new_row_data[col_cat_thr] = "-"
                new_row_data[col_tgl_upct] = "-"
                new_row_data[col_tgl_thr] = "-"

                new_row_data.pop("label_dropdown", None)
                new_row_data.pop("label_simpel", None)

                new_df_row = pd.DataFrame([new_row_data])
                df_updated = pd.concat([df, new_df_row], ignore_index=True)
                save_data(df_updated)

                st.session_state["flash_msg"] = f"🎉 SK Lama ({row_data.get(col_sk, '-')}) diset SELESAI KONTRAK, dan SK Baru ({val_no_sk}) atas nama {val_nama} BERHASIL DITAMBAHKAN!"
                st.rerun()

    # -------------------------------------------------------------
    # MENU 4: HAPUS DATA PEGAWAI
    # -------------------------------------------------------------
    elif menu == "Hapus Data Pegawai":
        st.title("🗑️ Hapus Data Pegawai")

        if "flash_msg" in st.session_state:
            st.success(st.session_state["flash_msg"])
            del st.session_state["flash_msg"]

        unique_pegawai = df["label_simpel"].drop_duplicates().tolist()

        selected_pegawai_label = st.selectbox(
            "Pilih Nama Pegawai:",
            options=unique_pegawai,
            key="delete_pegawai_select"
        )

        df_pegawai = df[df["label_simpel"] == selected_pegawai_label].copy()

        sk_options = df_pegawai.apply(
            lambda r: f"SK: {r.get(col_sk, '-')} | NIP: {r.get(col_nip, '-')} [{r.get('Status Pegawai', 'Aktif')}] ({parse_date_value(r.get('start_date', '')).strftime('%Y-%m-%d')} s.d {parse_date_value(r.get('end_date', '')).strftime('%Y-%m-%d')})",
            axis=1,
        ).tolist()

        selected_sk_label = st.selectbox(
            "Pilih Record SK / Addendum yang Ingin Dihapus:",
            options=sk_options,
            key="delete_sk_select"
        )

        selected_delete_idx = df_pegawai.index[sk_options.index(selected_sk_label)]
        target_row = df.iloc[selected_delete_idx]

        nip_terhapus = target_row.get(col_nip, "-")
        nama_terhapus = target_row.get(col_nama, "-")
        sk_terhapus = target_row.get(col_sk, "-")

        st.warning(
            f"⚠️ Anda akan menghapus record SK: **{sk_terhapus}** atas nama **{nama_terhapus}** (NIP: `{nip_terhapus}`). "
            f"Tindakan ini tidak dapat dibatalkan!"
        )

        if st.button("🚨 Hapus Permanen Record SK Ini", type="primary"):
            df_updated = df.drop(index=selected_delete_idx)
            save_data(df_updated)

            st.session_state["flash_msg"] = f"🗑️ Data Record SK {sk_terhapus} atas nama {nama_terhapus} (NIP: {nip_terhapus}) BERHASIL DIHAPUS PERMANEN!"
            st.rerun()