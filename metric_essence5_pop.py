import streamlit as st
import requests
import pandas as pd
from datetime import datetime, time, timedelta
import time as tm
import threading
from datetime import date
import re
import math
import uuid
def run() : 
    
    # page = st.sidebar.button("🔙 Kembali ke Halaman Utama")
    # if page == "🔙 Kembali ke Halaman Utama":
    #     import  metric_essence5
    #     metric_essence5.rerun()

    # URL dari Google Apps Script Web App
    APPS_SCRIPT_URL_ES = "https://script.google.com/macros/s/AKfycbwr-tJCErX2dl8Rj4fGjQDq1jNZlkjjLse6IboY0n5upMa6KfK7HYhNjBgXRkE-4WHpoA/exec"
    APPS_SCRIPT_URL_PR = "https://script.google.com/macros/s/AKfycbz5y58ApIInu1mL03bYcNS7jhBwHLhVCXnw8cPnBqDs-OzRn4BDB0axVk5BMobIog-YoQ/exec"
    APPS_SCRIPT_URL_BM = "https://script.google.com/macros/s/AKfycbyqYaOEQXUr2lM9kE4Pn34Uc-UnKEHFfCfmZs89lq4nfq_t3YFgsofod7JdWT26bNwH/exec"
    ############################################ FUNCTION APPS SCRIPT  #########################################################
    # Fungsi untuk mendapatkan semua data dari Google Sheets
    def get_all_data(url):
        try:
            response = requests.get(url, params={"action": "get_data"}, timeout=10)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            st.error(f"Terjadi kesalahan saat mengambil data: {e}")
            return []

    # Fungsi untuk mendapatkan opsi dari Google Sheets
    def get_options(url):
        try:
            response = requests.get(url, params={"action": "get_options"}, timeout=10)
            response.raise_for_status()
            options = response.json()
            return options
        except requests.exceptions.RequestException as e:
            st.error(f"Terjadi kesalahan saat mengambil data: {e}")
            return {}

    # Fungsi untuk mengirim data ke Google Sheets
    def add_data(form_data):
        try:
            response = requests.post(APPS_SCRIPT_URL_ES, json=form_data, timeout=10)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {"status": "error", "error": str(e)}

    # Fungsi untuk mendapatkan nomor SPK otomatis
    def generate_spk_number(selected_date):
        bulan_romawi = {
            1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI",
            7: "VII", 8: "VIII", 9: "IX", 10: "X", 11: "XI", 12: "XII"
        }
        all_data = get_all_data(APPS_SCRIPT_URL_ES)
        selected_month = selected_date.month
        selected_year = selected_date.year
        selected_month_romawi = bulan_romawi[selected_month]  # Konversi bulan ke Romawi

        # Ambil semua nomor SPK untuk bulan dan tahun ini
        spk_numbers = [
            row[0] for row in all_data
            if len(row) > 0 and f"/{selected_month_romawi}/{selected_year}" in row[0]
        ]

        if spk_numbers:
            # Ambil nomor terakhir dan tambahkan 1
            last_spk = max(spk_numbers)  # Ambil nomor terbesar
            last_number = int(last_spk.split("/")[0])  # Ambil angka sebelum "/PR/"
            new_number = last_number + 1
        else:
    # Jika belum ada SPK bulan ini, mulai dari 1
            new_number = 1

        # Format nomor SPK baru
        return f"{str(new_number).zfill(2)}/ESS/{selected_month_romawi}/{selected_year}"
    ################################################ FUNCTION GET DATA###################################
    # Ambil data dari Google Sheets
    st.session_state.setdefault("form_tanggal_pop", date.today())
    options_pr = get_options(APPS_SCRIPT_URL_PR)
    options_es = get_options(APPS_SCRIPT_URL_ES)

    defaults = {
        "form_nospk_pr_pop": "",
        "form_nospk_es_pop": generate_spk_number(st.session_state["form_tanggal_pop"]), 
        "form_tanggal_pop": st.session_state["form_tanggal_pop"], 
        "form_jenisproduk_pop": "", 
        "form_item_pop" : "",
        "form_komponen_pop" : "",
        "form_rencana_total_pemakaian_pop": "",
        "form_flv_pop_pop" : "",
        "form_rencana_flv_pop" : "",
        "form_realisasi_total_pemakaian_pop": "",
        "form_realisasi_flv_pop": "" ,
        "form_keterangan_pop" : "",
    }
    # Pastikan semua nilai default ada di session state tanpa overwrite jika sudah ada
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

    # Pastikan form_add_reset_pop ada di session_state
    st.session_state.setdefault("form_add_reset_pop", False)

    # Reset nilai form jika form_add_reset_pop bernilai True
    if st.session_state.form_add_reset_pop:
        st.session_state.update(defaults)
        st.session_state.form_add_reset_pop = False  # Kembalikan ke False setelah reset

    ################################################## OVERVIEW DATA PRODUKSI ###########################################################################################
    def filter_dataframe_inpop(df):
        """
        Adds a UI on top of a dataframe to let viewers filter columns.

        Args:
            df (pd.DataFrame): Original dataframe

        Returns:
            pd.DataFrame: Filtered dataframe
        """
        modify = st.checkbox("Tambah Filter",key='form_modify_pop')

        if not modify:
            return df

        df_filtered = df.copy()  # Salin dataframe agar tidak mengubah aslinya

        # Konversi tipe data untuk filter

        # Filter UI
        with st.container():
            to_filter_columns = st.multiselect("Pilih kolom untuk filter", df_filtered.columns)
            for column in to_filter_columns:
                left, right = st.columns((1, 20))
                left.write("↳")

                if column == "Tanggal":  # Slider untuk tanggal
                    # Ubah format tanggalnya dulu agar bisa di sort
                    bulan_indo_to_eng = {
                        "Januari": "January", "Februari": "February", "Maret": "March", "April": "April",
                        "Mei": "May", "Juni": "June", "Juli": "July", "Agustus": "August",
                        "September": "September", "Oktober": "October", "November": "November", "Desember": "December"
                    }
                    # Hilangkan nama hari dan ubah nama bulan ke bahasa Inggris
                    df_filtered[column] = df_filtered[column].apply(lambda x: re.sub(r"^\w+, ", "", x))  # Hapus nama hari
                    df_filtered[column] = df_filtered[column].replace(bulan_indo_to_eng, regex=True)  # Ubah nama bulan

                    # Parsing ke datetime
                    df_filtered[column] = pd.to_datetime(df_filtered[column], format="%d %B %Y")

                    min_date, max_date = df_filtered[column].min().date(), df_filtered[column].max().date()

                    # Ambil input tanggal dari pengguna
                    date_range = right.date_input(
                        f"Filter {column}",
                        min_value=min_date,
                        max_value=max_date,
                        value=(min_date, max_date),  # Default ke rentang min-max
                    )

                    # Pastikan date_range selalu dalam bentuk yang benar
                    if isinstance(date_range, tuple) and len(date_range) == 2:
                        start_date, end_date = date_range
                    else:
                        start_date = end_date = date_range[0]  # Ambil elemen pertama kalau masih tuple

                    # Filter data dengan hanya membandingkan tanggal (tanpa waktu)
                    df_filtered = df_filtered[
                        (df_filtered[column].dt.date >= start_date) & 
                        (df_filtered[column].dt.date <= end_date)
                    ]

                    # Ubah format ke "Hari, Tanggal Bulan Tahun"
                    # Kamus Nama Hari Inggris → Indonesia
                    hari_eng_to_indo = {
                        "Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Kamis",
                        "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"
                    }
                        # *Tambahkan Nama Hari dan Format ke "Hari, Tanggal Bulan Tahun"*
                    df_filtered["Hari"] = df_filtered[column].dt.strftime('%A').map(hari_eng_to_indo)  # Konversi nama hari
                    df_filtered[column] = df_filtered[column].dt.strftime('%d %B %Y')  # Format tanggal biasa
                    df_filtered[column] = df_filtered["Hari"] + ", " + df_filtered[column]  # Gabungkan Nama Hari
                    df_filtered.drop(columns=["Hari"], inplace=True)  # Hapus kolom tambahan

                elif column in ["Jam Start", "Jam Stop",'Total Hour']:

                    df_filtered[column] = pd.to_datetime(df_filtered[column], errors='coerce').dt.time
                    # Pastikan kolom tidak kosong
                    if df_filtered[column].dropna().empty:
                        st.warning(f"Tidak ada data untuk {column}.")
                        continue  # Lewati filter ini jika tidak ada data

                    min_time, max_time = df_filtered[column].dropna().min(), df_filtered[column].dropna().max()
                    # Tambahkan validasi jika min_time == max_time
                    if min_time == max_time:
                        min_time = (datetime.combine(datetime.today(), min_time) - timedelta(minutes=30)).time()
                        max_time = (datetime.combine(datetime.today(), max_time) + timedelta(minutes=30)).time()
                    start_time, end_time = right.slider(
                        f"Filter {column}",
                        min_value=min_time,
                        max_value=max_time,
                        value=(min_time, max_time),
                        format="HH:mm"
                    )

                    df_filtered = df_filtered[
                        (df_filtered[column] >= start_time) &
                        (df_filtered[column] <= end_time)
                    ]
                    df_filtered[column] = df_filtered[column].apply(lambda x: x.strftime("%H:%M") if pd.notnull(x) else "")

                elif column in ["Nomor SPK", "BU", "Jenis Produk", "Line", "Speed(kg/jam)",
                                "Rencana Total Output (kg)", "Rencana Total Output (Batch)", 
                                "Inner (roll)", "SM", "Alasan"]:  # Filter kategori
                    unique_values = df_filtered[column].unique()
                    selected_values = right.multiselect(
                        f"Filter {column}",
                        options=unique_values,
                        default=[],
                    )
                    if selected_values:
                        df_filtered = df_filtered[df_filtered[column].isin(selected_values)]
        return df_filtered
    def overview(key_prefix="default"):
        st.title("Data Overview")
        data = get_all_data(APPS_SCRIPT_URL_PR)
        columns = [
            "Nomor SPK", "Tanggal", "BU", "Jenis Produk", "Line",
            "Jam Start", "Jam Stop", "Total Hour", "Speed(kg/jam)", 
            "Rencana Total Output (kg)", "Rencana Total Output (Batch)", 
            "Inner (roll)", "SM", "Alasan"
        ]
        
        df = pd.DataFrame(data[1:], columns=columns) if data else pd.DataFrame(columns=columns)

        if data:
            # Balik urutan baris → data terbaru di atas
            df = df.iloc[::-1].reset_index(drop=True)

        key_unique = str(uuid.uuid4())
        if st.button("Muat Ulang Data", key=f"{key_prefix}_{key_unique}_btn_muat_ulang_data"):
            st.cache_data.clear()
            st.rerun()

        st.dataframe(filter_dataframe_inpop(df))

    overview(key_prefix="overview_loly")
    ########################### FUNCTION DEPENDENT DROPDOWN DARI SPK PRODUKSI ##############################
    # data dari options appscripct / spreadsheet 
    data_pr = [row for row in options_pr.get("Data Table", []) if isinstance(row, list) and len(row) > 0] 
    data_es = [row for row in options_es.get("HC", []) if isinstance(row, list) and len(row) > 2]  
    #################### Fungsi untuk mendapatkan daftar unik No. SPK Produksi ########################
    def extract_unique_produk_es(data): # fungsi untuk ambil item unik di spreadsheet HC Essence 
        try:
            return sorted(set(row[0] for row in data if row[0]))  
        except Exception as e:
            st.error(f"Error saat mengekstrak Item: {e}")
            return []

    def filter_spk_by_produk(data_produksi, item_es_options):
        try:
            seen = set()
            result = []
            item_es_lower = [item.lower() for item in item_es_options if isinstance(item, str)]

            for row in reversed(data_produksi):
                spk = row[0]
                produk = row[3]

                if not isinstance(produk, str):
                    continue

                produk_lower = produk.lower()

                # Cek apakah produk match sebagian dari salah satu item ES
                if spk and any(produk_lower in item for item in item_es_lower) and spk not in seen:
                    seen.add(spk)
                    result.append(spk)

            return result
        except Exception as e:
            st.error(f"Error saat memfilter SPK berdasarkan produk ESSENCE: {e}")
            return []
    options_bm = get_options(APPS_SCRIPT_URL_BM)
    list_dropdown_bm = [row for row in options_bm.get("Dropdown List", []) if isinstance(row, list) and len(row) > 0]
    #################### FUNCTION DEPENDENT DROPDOWN ITEM, CASING #############
    # Fungsi untuk memfilter Item berdasarkan item yang dipilih
    def filter_by_item(data, selected_item, column_index):
        try:
            selected_item = selected_item.lower()
            return sorted(set(
                row[column_index]
                for row in data
                if isinstance(row[0], str)
                and selected_item in row[0].lower()
                and row[column_index]
            ))
        except Exception as e:
            st.error(f"Error saat memfilter berdasarkan item: {e}")
            return []
        
    def filter_by_item_casing(data, produk, casing, col_index):
        produk_lower = produk.lower()
        return sorted(set(
            row[col_index]
            for row in data
            if produk_lower in row[0].lower() and row[2] == casing
        ))
    # Untuk mengambil kolom FLV hingga ke kanan
    def filter_flavour(data, item, komponen, flavour_start_index=7):
        try:
            item = item.lower()
            komponen = komponen.lower()
            result = []

            for row in data:
                if (
                    isinstance(row[0], str)
                    and isinstance(row[2], str)
                    and item in row[0].lower()
                    and komponen in row[3].lower()
                ):
                    for i, val in enumerate(row[flavour_start_index:], start=flavour_start_index):
                        if val not in (None, "", 0, "0"):
                            flavour_name = row[i] if i < len(row) else f"Kolom {i}"
                            result.append({
                                "index": i,
                                "nama": flavour_name,
                                "persen": val
                            })
                    break  # stop di baris pertama yang cocok
            return result
        except Exception as e:
            st.error(f"❌ Error saat memfilter Flavour: {e}")
            return []

    # # Ambil daftar unik dari dataset
    item_es_options = extract_unique_produk_es(data_es)
    spk_pr_options = filter_spk_by_produk(data_pr, item_es_options)

    # Dropdown untuk No.SPK Produksi
    st.markdown("### 🗂 Pilih No. SPK Produksi & Isi Tanggal Input ")
    st.markdown("---")

    # Inisialisasi session_state untuk tanggal dan SPK jika belum ada
    if "form_tanggal_pop" not in st.session_state:
        st.session_state["form_tanggal_pop"] = datetime.date.today()

    if "form_nospk_pr_pop" not in st.session_state:
        st.session_state["form_nospk_pr_pop"] = ""
    col1, col2 , col3 = st.columns(3)  
    with col1:
        tanggal = st.date_input("Tanggal Input", value=st.session_state.get("form_tanggal_pop"), key="form_tanggal_pop")
            # Simpan tanggal baru jika berbeda
        if tanggal != st.session_state["form_tanggal_pop"]:
            st.session_state["form_tanggal_pop"] = tanggal
    with col2:
        spk_pr = st.selectbox("Pilih No. SPK Produksi", [""] + spk_pr_options, key="form_nospk_pr_pop")
        # Reset SPK jika nilainya tidak valid
    with col3:
        spk_es = st.text_input("No. SPK Essence", value=generate_spk_number(st.session_state["form_tanggal_pop"]), key="form_nospk_es_pop", disabled=True)
    ############################## SETELAH MEMILIH SPK PRODUKSI #########################################################
    if spk_pr:
        data_rencana =[]
        data_realisasi=[]
        output_permen_kg = filler_kg = filler_batch = 0
        try:
            selected_row = next(
                row for row in data_pr
                if row[0] == spk_pr
            )
            ######################### SPK BALLMILL X SPK PRODUKSI ##############################
            # Simpan ke session_state agar bisa dipakai di form
            jenis_produk_pr  = st.session_state["form_produk_pop"] = selected_row[3]
            # Ganti 'Blaster' di awal jadi 'Filler'
            if jenis_produk_pr .lower().startswith("blaster"):
                nama_filler = jenis_produk_pr .replace("Blaster", "Filler", 1)
            else:
                nama_filler = jenis_produk_pr 
            st.session_state["form_produk_pop"] = jenis_produk_pr 

            outputKg_pr = st.session_state["form_output(kg)_pop"] = selected_row[9]
            outputBatch_pr = st.session_state["form_output(batch)_pop"] = selected_row[10]

            df_bm = pd.DataFrame(list_dropdown_bm, columns=["Item", "Siklus (kg)", "Filler", "Bt","-","Operator","--","No. Mesin"])
            # Bersihkan dan ubah format data
            def clean_number_column(column):
                return (
                    column.astype(str)
                    .str.replace(",", "", regex=False)  # hapus koma (pem. ribuan)
                    .astype(float)                      # titik tetap sebagai desimal)
                )
            # Bersihkan angka pada kolom numerik
            for col in ["Siklus (kg)", "Filler", "Bt"]:
                df_bm[col] = clean_number_column(df_bm[col])

            # ✅ Baru filter berdasarkan jenis produk
            df_bm_filtered = df_bm[df_bm["Item"] == jenis_produk_pr]

            # Cari baris pertama dengan Siklus (kg) yang >= output SPK
            row_match = df_bm_filtered[df_bm_filtered["Siklus (kg)"] >= float(outputKg_pr)].sort_values("Siklus (kg)").head(1)
            
            if not row_match.empty:
                outputKg_bm = row_match["Siklus (kg)"].values[0]
                filler_bm = row_match["Filler"].values[0]
                batch_bm = row_match["Bt"].values[0]

                output_permen_kg = float(outputKg_pr)
                filler_kg = round((float(output_permen_kg) / outputKg_bm) * filler_bm, 2)
                
                def force_ceil_to_half(x):
                    return math.ceil(x * 2 + 1e-9) / 2

                filler_batch_raw = (float(output_permen_kg) / outputKg_bm) * batch_bm
                filler_batch = force_ceil_to_half(filler_batch_raw)
                
            #####################################################################################
            # Simpan ke session_state agar bisa dipakai di form
            outputKg_pr = st.session_state["form_output(kg)_pop"] = selected_row[9]
            outputBatch_pr = st.session_state["form_output(batch)_pop"] = selected_row[10]
            
            # MENAMPILKAN INFORMASI SPK PRODUKSI TERPILIH ##################################
            st.markdown("""
            <h3 style='text-align: center;'>💡 Informasi SPK Produksi Terpilih</h2>
            """, unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)

            # Buat 3 kolom sejajar
            col1, col2, col3 = st.columns(3)

            with col1:
                st.markdown("*Jenis Produk*")
                st.markdown(f"<div style='font-size:20px; font-weight:bold'>{jenis_produk_pr}</div>", unsafe_allow_html=True)

            with col2:
                st.metric(label="Rencana Total Output (Kg)", value=round(float(outputKg_pr), 2))

            with col3:
                st.metric(label="Rencana Total Output (Batch)", value=round(float(outputBatch_pr), 2))

            card_style = """
                background-color: #e6f2ff;
                padding: 20px;
                border-radius: 16px;
                text-align: center;
                box-shadow: 2px 2px 6px rgba(0,0,0,0.1);
                height: 160px;
                width: 100%;
                display: flex;
                flex-direction: column;
                justify-content: center;
            """
            label_style = "font-size: 18px; font-weight: 600; margin-bottom: 8px;"
            value_style = "font-size: 28px; font-weight: bold; color: #222;"

            ########### MULAI MEMILIH CASING ################################

            casing_options = filter_by_item(data_es, jenis_produk_pr, 2)
            
            if not casing_options:
                st.warning("❌ Casing tidak ditemukan untuk produk ini.")
                st.stop()
            if "form_casing_pop" not in st.session_state or st.session_state["form_casing_pop"] not in casing_options:
                st.session_state["form_casing_pop"] = casing_options[0]

            casing_terpilih = st.selectbox("Pilih Casing", casing_options, key="form_casing_pop")

            ################################## RENCANA ################################################
            st.markdown('---')
            st.markdown("""
            <h3 style='text-align: center;'>📄 RENCANA METRIC ESSENCE</h2>
            """, unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            ########### PERHITUNGAN KEBUTUHAN FLAVOUR #############
            selected_casing = next(row for row in data_es
                    if row[2] == casing_terpilih)
            ### IF FILLER ####################
            if casing_terpilih.strip().upper() in ['FILLER', 'FILLER FOR 03', 'FILLER FOR 07']:
                st.markdown(f"""
                <div style="
                    background-color: #f9fafb;
                    border: 1px solid #d1d5db;
                    border-radius: 8px;
                    padding: 10px 16px;
                    margin: 16px auto;
                    width: 60%;
                    text-align: center;
                    font-family: 'Segoe UI', sans-serif;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                ">
                    <div style="font-size: 13px; font-weight: 500; color: #4b5563; margin-bottom: 6px;">
                        🧮 Hasil Perhitungan Filler (Batch)
                    </div>
                    <div style="font-size: 20px; font-weight: 600; color: #111827;">
                        {filler_batch:,.2f}
                    </div>
                </div>
            """, unsafe_allow_html=True)

                # Ambil semua baris komponen yang cocok dengan produk
                komponen_list = filter_by_item_casing(data_es, jenis_produk_pr, casing_terpilih, 3)
    
                if not komponen_list:
                    st.warning("❌ Tidak ditemukan komponen untuk FILLER.")
                    st.stop()

                # List kosong untuk menyimpan semua data yang akan dikirim
                for komponen in komponen_list:
                    st.markdown(f"<h5 style='text-align:left; margin-top: 12px;'>🔹Komponen: <strong>{komponen}</strong></h5>", unsafe_allow_html=True)

                    selected_komponen = next(
                        row for row in data_es
                        if isinstance(row[0], str)
                        and jenis_produk_pr.lower() in row[0].lower()
                        and row[2] == casing_terpilih
                        and row[3] == komponen
                    )

                    quantity =st.session_state["form_quantity_pop"] = selected_komponen[4] 
                    rasio = st.session_state["form_rasio_pop"] = selected_komponen[5]
                    total_based_flv = st.session_state["form_total_based_flv_pop"] = selected_komponen[6]


                    # Ambil flavour list dari fungsi baru
                    flavour_list = filter_flavour(data_es, jenis_produk_pr, komponen)
                    # Inisialisasi total
                    total_persen_flavour_filler = 0
                    total_flavour_kg = 0

                    # Ambil baris header dari data_es
                    header_row = data_es[0] if data_es else []

                    # Layout 4 kolom per baris
                    # Tampilkan header
                    for i in range(0, len(flavour_list), 4):
                        cols = st.columns(4)
                        current_row = flavour_list[i:i+4]
                        for j, flavour in enumerate(current_row):
                            try:
                                index = flavour["index"]
                                nama_kolom = header_row[index] if index < len(header_row) else f"Kolom {index}"
                                nama = flavour["nama"]
                                persen = float(str(flavour["persen"]).replace(",", "."))
                                total_rencana= round(outputKg_pr * rasio, 4)
                                hasil_kg_rencana = round((persen*total_rencana)/total_based_flv, 4)

                            

                                ### SIMPAN DATA UNTUK DIKIRIM ###
                                # Simpan satu baris data untuk setiap FLV
                                data_rencana.append({
                                    "NomorSPK": spk_es,
                                    "Tanggal": tanggal.strftime("%Y-%m-%d") ,
                                    "Item": jenis_produk_pr,
                                    "Casing": casing_terpilih,
                                    "Komponen": komponen,
                                    "FLV": nama_kolom,
                                    "RencanaTotal": total_rencana,
                                    "RencanaFLV": hasil_kg_rencana,
                                    "RealisasiTotal": "",  # Kosong karena belum diisi
                                    "RealisasiFLV": "",    # Kosong karena belum diisi
                                })

                                with cols[j]:
                                    st.markdown(f"""
                                        <div style="
                                            background-color: #eef6ff;
                                            border-radius: 8px;
                                            padding: 8px;
                                            margin-bottom: 16px;
                                            text-align: center;
                                            box-shadow: 1px 1px 3px rgba(0,0,0,0.08);
                                            font-family: 'Segoe UI', sans-serif;
                                        ">
                                            <div style="font-size: 12px; font-weight: 600; color: #555;">{nama_kolom}</div>
                                            <div style="font-size: 14x; font-weight: bold; color: #222;">{hasil_kg_rencana:.4f} Kg</div>
                                        </div>
                                    """, unsafe_allow_html=True)

                                

                            except Exception as e:
                                st.warning(f"Gagal menghitung flavour {flavour}: {e}")

                        card_style = """
                            background-color: #f8fdf9;
                            padding: 12px 16px;
                            border-left: 5px solid #34a853;
                            border-radius: 8px;
                            text-align: center;
                            margin: auto;
                            width: 85%;
                            margin-bottom: 12px;
                            """

                    label_style = "font-size: 14px; font-weight: 500; color: #555; margin-bottom: 4px;"
                    value_style = "font-size: 18px; font-weight: 700; color: #1e4620;"

                    st.markdown(f"""
                    <div style="
                        background-color: #fef6c9  ;
                        padding: 6px 12px;
                        border-radius: 6px;
                        text-align: center;
                        margin: auto;
                        width: 50%;
                        margin-top: 6px;
                        margin-bottom: 10px;
                    ">
                        <div style="font-size: 16px; font-weight: 500; color: #334155;">
                            Total Kebutuhan Flavour (Kg)
                        </div>
                        <div style="font-size: 20px; font-weight: 600; color: #1e293b;">
                            {total_rencana:,.4f}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
    ################## MULAI DARI SINI ELIF ############
            elif casing_terpilih.strip().upper() in['CORE','STRIPE']: 
                # Ambil semua baris komponen yang cocok dengan produk
                komponen_list = filter_by_item_casing(data_es, jenis_produk_pr, casing_terpilih, 3)
    
                if not komponen_list:
                    st.warning("❌ Tidak ditemukan komponen untuk FILLER.")
                    st.stop()

                for komponen in komponen_list:
                    # st.markdown(f"<h5 style='text-align: center;'>Komponen: {komponen}</h4>", unsafe_allow_html=True)
                    st.markdown(f"<h5 style='text-align:center; margin-top: 12px;'>🔹 Komponen: <strong>{komponen}</strong></h5>", unsafe_allow_html=True)

                    selected_komponen = next(
                        row for row in data_es
                        if isinstance(row[0], str)
                        and jenis_produk_pr.lower() in row[0].lower()
                        and row[2] == casing_terpilih
                        and row[3] == komponen
                    )

                    quantity =st.session_state["form_quantity_pop"] = selected_komponen[4] 
                    rasio = st.session_state["form_rasio_pop"] = selected_komponen[5]
                    total_based_flv = st.session_state["form_total_based_flv_pop"] = selected_komponen[6]
                    outputKg = outputKg_pr
                    
                    # Ambil flavour list dari fungsi baru
                    flavour_list = filter_flavour(data_es, jenis_produk_pr, komponen)
                    # Inisialisasi total
                    total_persen_flavour_filler = 0
                    total_flavour_kg = 0

                    # Ambil baris header dari data_es
                    
                    header_row = data_es[0] if data_es else []

                    # Layout 4 kolom per baris
                    # Tampilkan header
                    for i in range(0, len(flavour_list), 4):
                        cols = st.columns(4)
                        current_row = flavour_list[i:i+4]
                        for j, flavour in enumerate(current_row):
                            try:
                                index = flavour["index"]
                                nama_kolom = header_row[index] if index < len(header_row) else f"Kolom {index}"
                                nama = flavour["nama"]
                                persen = float(str(flavour["persen"]).replace(",", "."))
                                total_rencana= round(outputKg * rasio, 4)
                                hasil_kg_rencana = round((total_rencana*persen)/total_based_flv, 4)

                                data_rencana.append({
                                    "NomorSPK": spk_es,
                                    "Tanggal": tanggal.strftime("%Y-%m-%d") ,
                                    "Item": jenis_produk_pr,
                                    "Casing": casing_terpilih,
                                    "Komponen": komponen,
                                    "FLV": nama_kolom,
                                    "RencanaTotal": total_rencana,
                                    "RencanaFLV": hasil_kg_rencana,
                                    "RealisasiTotal": "",  # Kosong karena belum diisi
                                    "RealisasiFLV": "",    # Kosong karena belum diisi
                                })

                                with cols[j]:
                                    st.markdown(f"""
                                        <div style="
                                            background-color: #eef6ff;
                                            border-radius: 8px;
                                            padding: 8px;
                                            margin-bottom: 16px;
                                            text-align: center;
                                            box-shadow: 1px 1px 3px rgba(0,0,0,0.08);
                                            font-family: 'Segoe UI', sans-serif;
                                        ">
                                            <div style="font-size: 12px; font-weight: 600; color: #555;">{nama_kolom}</div>
                                            <div style="font-size: 14x; font-weight: bold; color: #222;">{hasil_kg_rencana:.4f} Kg</div>
                                        </div>
                                    """, unsafe_allow_html=True)
                            except Exception as e:
                                st.warning(f"Gagal menghitung flavour {flavour}: {e}")

                        card_style = """
                            background-color: #f8fdf9;
                            padding: 12px 16px;
                            border-left: 5px solid #34a853;
                            border-radius: 8px;
                            text-align: center;
                            margin: auto;
                            width: 85%;
                            margin-bottom: 12px;
                            """

                    label_style = "font-size: 14px; font-weight: 500; color: #555; margin-bottom: 4px;"
                    value_style = "font-size: 18px; font-weight: 700; color: #1e4620;"

                    st.markdown(f"""
                    <div style="
                        background-color: #fef6c9  ;
                        padding: 6px 12px;
                        border-radius: 6px;
                        text-align: center;
                        margin: auto;
                        width: 50%;
                        margin-top: 6px;
                        margin-bottom: 10px;
                    ">
                        <div style="font-size: 16px; font-weight: 500; color: #334155;">
                            Total Kebutuhan Flavour (Kg)
                        </div>
                        <div style="font-size: 20px; font-weight: 600; color: #1e293b;">
                            {total_rencana:,.4f}
                        </div>
                    </div>
                """, unsafe_allow_html=True)


    ################## MULAI DARI SINI ELSE ############   
            else:
                komponen_options = filter_by_item_casing(data_es,jenis_produk_pr, casing_terpilih, 3)
                if "form_komponen_pop_rencana" not in st.session_state or st.session_state["form_komponen_pop_rencana"] not in komponen_options:
                    st.session_state["form_komponen_pop_rencana"] = komponen_options[0]

                komponen = st.selectbox("Pilih Komponen", komponen_options, key="form_komponen_pop_rencana")

                selected_komponen = next(
                    row for row in data_es
                    if isinstance(row[0], str)
                    and jenis_produk_pr.lower() in row[0].lower()
                    and row[2] == casing_terpilih
                    and row[3] == komponen
                )


                quantity =st.session_state["form_quantity_pop"] = selected_komponen[4] 
                rasio = st.session_state["form_rasio_pop"] = selected_komponen[5]
                total_based_flv = st.session_state["form_total_based_flv_pop"] = selected_komponen[6]

                outputKg = outputKg_pr

                # Ambil flavour list dari fungsi baru
                flavour_list = filter_flavour(data_es, jenis_produk_pr, komponen)

                # Tampilkan header
                st.markdown(f"<h3 style='text-align: center;'>🔸 Kebutuhan Flavour untuk Komponen: {komponen}</h3>", unsafe_allow_html=True)

                # Inisialisasi total
                total_persen_flavour = 0
                total_flavour_kg = 0

                # Ambil baris header dari data_es
                header_row = data_es[0] if data_es else []

                # Layout 4 kolom per baris
                for i in range(0, len(flavour_list), 4):
                    cols = st.columns(4)
                    current_row = flavour_list[i:i+4]
                    for j, flavour in enumerate(current_row):
                        try:
                            index = flavour["index"]
                            nama_kolom = header_row[index] if index < len(header_row) else f"Kolom {index}"
                            nama = flavour["nama"]
                            persen = float(str(flavour["persen"]).replace(",", "."))
                            total_rencana= round(outputKg * rasio, 4)
                            hasil_kg_rencana = round((total_rencana*persen)/total_based_flv, 4)

                            data_rencana.append({
                                    "NomorSPK": spk_es,
                                    "Tanggal": tanggal.strftime("%Y-%m-%d") ,
                                    "Item": jenis_produk_pr,
                                    "Casing": casing_terpilih,
                                    "Komponen": komponen,
                                    "FLV": nama_kolom,
                                    "RencanaTotal": total_rencana,
                                    "RencanaFLV": hasil_kg_rencana,
                                    "RealisasiTotal": "",  # Kosong karena belum diisi
                                    "RealisasiFLV": "",    # Kosong karena belum diisi
                                })
                            
                            with cols[j]:
                                st.markdown(f"""
                                    <div style="
                                        background-color: #eef6ff;
                                        border-radius: 8px;
                                        padding: 8px;
                                        margin-bottom: 16px;
                                        text-align: center;
                                        box-shadow: 1px 1px 3px rgba(0,0,0,0.08);
                                        font-family: 'Segoe UI', sans-serif;
                                    ">
                                        <div style="font-size: 12px; font-weight: 600; color: #555;">{nama_kolom}</div>
                                        <div style="font-size: 14x; font-weight: bold; color: #222;">{hasil_kg_rencana:.4f} Kg</div>
                                    </div>
                                """, unsafe_allow_html=True)
                        except Exception as e:
                            st.warning(f"Gagal menghitung flavour {flavour}: {e}")
                card_style = """
                    background-color: #f8fdf9;
                    padding: 12px 16px;
                    border-left: 5px solid #34a853;
                    border-radius: 8px;
                    text-align: center;
                    margin: auto;
                    width: 85%;
                    margin-bottom: 12px;
                """

                label_style = "font-size: 14px; font-weight: 500; color: #555; margin-bottom: 4px;"
                value_style = "font-size: 18px; font-weight: 700; color: #1e4620;"

                st.markdown(f"""
                <div style="
                    background-color: #fef6c9  ;
                    padding: 6px 12px;
                    border-radius: 6px;
                    text-align: center;
                    margin: auto;
                    width: 50%;
                    margin-top: 6px;
                    margin-bottom: 10px;
                ">
                    <div style="font-size: 16px; font-weight: 500; color: #334155;">
                        Total Kebutuhan Flavour (Kg)
                    </div>
                    <div style="font-size: 20px; font-weight: 600; color: #1e293b;">
                        {total_rencana:,.4f}
                    </div>
                </div>
            """, unsafe_allow_html=True)
    ######################################################## METRIC REALISASI #####################################################

            st.markdown('---')
            st.markdown("""
            <h3 style='text-align: center;'>📊 REALISASI METRIC ESSENCE</h3>
            """, unsafe_allow_html=True)
            if casing_terpilih.strip().upper() in ['FILLER', 'FILLER FOR 03', 'FILLER FOR 07', 'CORE','STRIPE']:
                # Ambil semua baris komponen yang cocok dengan produk
                komponen_list = filter_by_item_casing(data_es, jenis_produk_pr, casing_terpilih, 3)
    
                if not komponen_list:
                    st.warning("❌ Tidak ditemukan komponen untuk FILLER.")
                    st.stop()

                for komponen in komponen_list:
                    st.markdown(f"<h5 style='text-align:left; margin-top: 12px;'>🔹 Komponen: <strong>{komponen}</strong></h5>", unsafe_allow_html=True)
                    selected_komponen = next(
                        row for row in data_es
                        if isinstance(row[0], str)
                        and jenis_produk_pr.lower() in row[0].lower()
                        and row[2] == casing_terpilih
                        and row[3] == komponen
                    )

                    quantity = selected_komponen[4]
                    rasio = selected_komponen[5]
                    total_based_flv = selected_komponen[6]

                    # Ambil flavour list
                    flavour_list = filter_flavour(data_es, jenis_produk_pr, komponen)
                    total_flavour_kg = 0
                    header_row = data_es[0] if data_es else []
                    
                    total_realisasi = st.number_input(
                        f"Total Realisasi untuk {komponen}",
                        key=f"form_total_realisasi_pop_{komponen}"
                    )
                    if total_realisasi > 0:
                        # Tampilkan dulu flavour-nya
                        for i in range(0, len(flavour_list), 4):
                            cols = st.columns(4)
                            for j, flavour in enumerate(flavour_list[i:i+4]):
                                try:
                                    index = flavour["index"]
                                    nama_kolom = header_row[index] if index < len(header_row) else f"Kolom {index}"
                                    persen = float(str(flavour["persen"]).replace(",", "."))
                                    hasil_kg_realisasi = round((persen*total_realisasi)/total_based_flv, 4)

                                    # ✅ Tambahkan data ke data_to_send
                                    data_realisasi.append({
                                        "NomorSPK": spk_es,
                                        "Tanggal": tanggal.strftime("%Y-%m-%d") ,
                                        "Item": jenis_produk_pr,
                                        "Casing": casing_terpilih,
                                        "Komponen": komponen,
                                        "FLV": nama_kolom,
                                        "RencanaTotal": "",  # Kosong karena tidak dihitung di sini
                                        "RencanaFLV": "",    # Kosong karena tidak dihitung di sini
                                        "RealisasiTotal": total_realisasi,
                                        "RealisasiFLV": hasil_kg_realisasi,
                                    })

                                    with cols[j]:
                                        st.markdown(f"""
                                            <div style="
                                                background-color: #eef6ff;
                                                border-radius: 8px;
                                                padding: 8px;
                                                margin-bottom: 16px;
                                                text-align: center;
                                                box-shadow: 1px 1px 3px rgba(0,0,0,0.08);
                                                font-family: 'Segoe UI', sans-serif;
                                            ">
                                                <div style="font-size: 12px; font-weight: 600; color: #555;">{nama_kolom}</div>
                                                <div style="font-size: 14x; font-weight: bold; color: #222;">{hasil_kg_realisasi:.4f} Kg</div>
                                            </div>
                                        """, unsafe_allow_html=True)
                                except Exception as e:
                                    st.warning(f"Gagal menampilkan flavour {flavour}: {e}")

                    # Input total realisasi ditempatkan setelah flavour tampil
                        # Total realisasi flavour ditampilkan setelah input
                        st.markdown(f"""
                            <div style="
                                background-color: #fef6c9;
                                padding: 6px 12px;
                                border-radius: 6px;
                                text-align: center;
                                margin: auto;
                                width: 50%;
                                margin-top: 6px;
                                margin-bottom: 10px;
                            ">
                                <div style="font-size: 16px; font-weight: 500; color: #334155;">
                                    Total Kebutuhan Flavour (Kg)
                                </div>
                                <div style="font-size: 20px; font-weight: 600; color: #1e293b;">
                                    {total_realisasi:,.4f}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                            

    ################## MULAI DARI SINI ELSE REALISASI ############   
            else:
                selected_komponen = next(
                        row for row in data_es
                        if isinstance(row[0], str)
                        and jenis_produk_pr.lower() in row[0].lower()
                        and row[2] == casing_terpilih
                        and row[3] == st.session_state["form_komponen_pop_rencana"]
                    )

                quantity =st.session_state["form_quantity_pop"] = selected_komponen[4] 
                rasio = st.session_state["form_rasio_pop"] = selected_komponen[5]
                total_based_flv = st.session_state["form_total_based_flv_pop"] = selected_komponen[6]

                outputKg = outputKg_pr

                # Ambil flavour list dari fungsi baru
                flavour_list = filter_flavour(data_es, jenis_produk_pr, st.session_state["form_komponen_pop_rencana"])

                # Tampilkan header
                st.markdown(f"<h3 style='text-align: center;'>🔸 Kebutuhan Flavour untuk Komponen: {st.session_state["form_komponen_pop_rencana"]}</h3>", unsafe_allow_html=True)

                # Inisialisasi total
                total_persen_flavour = 0
                total_flavour_kg = 0

                # Ambil baris header dari data_es
                header_row = data_es[0] if data_es else []

                total_realisasi = st.number_input(
                        f"Total Realisasi untuk {st.session_state["form_komponen_pop_rencana"]}",
                        key=f"form_total_realisasi_pop_{st.session_state["form_komponen_pop_rencana"]}"
                    )
                if total_realisasi > 0:
                    # Tampilkan dulu flavour-nya
                    for i in range(0, len(flavour_list), 4):
                        cols = st.columns(4)
                        for j, flavour in enumerate(flavour_list[i:i+4]):
                            try:
                                index = flavour["index"]
                                nama_kolom = header_row[index] if index < len(header_row) else f"Kolom {index}"
                                persen = float(str(flavour["persen"]).replace(",", "."))
                                hasil_kg_realisasi = round((persen*total_realisasi)/total_based_flv, 4)

                                # ✅ Tambahkan data ke data_to_send
                                data_realisasi.append({
                                        "NomorSPK": spk_es,
                                        "Tanggal": tanggal.strftime("%Y-%m-%d") ,
                                        "Item": jenis_produk_pr,
                                        "Casing": casing_terpilih,
                                        "Komponen": komponen,
                                        "FLV": nama_kolom,
                                        "RencanaTotal": "",  # Kosong karena tidak dihitung di sini
                                        "RencanaFLV": "",    # Kosong karena tidak dihitung di sini
                                        "RealisasiTotal": total_realisasi,
                                        "RealisasiFLV": hasil_kg_realisasi,
                                    })
                                
                                with cols[j]:
                                    st.markdown(f"""
                                        <div style="
                                            background-color: #eef6ff;
                                            border-radius: 8px;
                                            padding: 8px;
                                            margin-bottom: 16px;
                                            text-align: center;
                                            box-shadow: 1px 1px 3px rgba(0,0,0,0.08);
                                            font-family: 'Segoe UI', sans-serif;
                                        ">
                                            <div style="font-size: 12px; font-weight: 600; color: #555;">{nama_kolom}</div>
                                            <div style="font-size: 14x; font-weight: bold; color: #222;">{hasil_kg_realisasi:.4f} Kg</div>
                                        </div>
                                    """, unsafe_allow_html=True)
                            except Exception as e:
                                st.warning(f"Gagal menampilkan flavour {flavour}: {e}")

                # Input total realisasi ditempatkan setelah flavour tampil
                    # Total realisasi flavour ditampilkan setelah input
                    st.markdown(f"""
                        <div style="
                            background-color: #fef6c9;
                            padding: 6px 12px;
                            border-radius: 6px;
                            text-align: center;
                            margin: auto;
                            width: 50%;
                            margin-top: 6px;
                            margin-bottom: 10px;
                        ">
                            <div style="font-size: 16px; font-weight: 500; color: #334155;">
                                Total Kebutuhan Flavour (Kg)
                            </div>
                            <div style="font-size: 20px; font-weight: 600; color: #1e293b;">
                                {total_realisasi:,.4f}
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                        
            keterangan = st.text_area('Keterangan', key='form_keterangan_pop')

        except StopIteration:
            st.warning("Data tidak ditemukan untuk No. SPK tersebut..")
        except Exception as e:
            st.error(f"Terjadi error saat mengambil data: {e}")

    ############################################# SUBMIT DATA ###############################################################################

        # form_completed = all(st.session_state.get(key) for key in [
        #                     "form_tanggal_pop"])
        # ,disabled=not form_completed
        st.write("")
        submit_button = st.button("💾 Simpan Data", use_container_width=True)

        # Gabungkan rencana dan realisasi berdasarkan Komponen & FLV
        merged_data = {}

        # Masukkan data rencana ke dalam dict
        for row in data_rencana:
            key = f"{row['Komponen']}|{row['FLV']}"
            merged_data[key] = {
                "NomorSPK": row["NomorSPK"],
                "Tanggal": row["Tanggal"],
                "Item": row["Item"],
                "Casing": row["Casing"],
                "Komponen": row["Komponen"],
                "FLV": row["FLV"],
                "RencanaTotal": row["RencanaTotal"],
                "RencanaFLV": row["RencanaFLV"],
                "RealisasiTotal": "",
                "RealisasiFLV": "",
            }

        # Update jika FLV & Komponen sama, tapi dari realisasi
        for row in data_realisasi:
            key = f"{row['Komponen']}|{row['FLV']}"
            if key in merged_data:
                merged_data[key]["RealisasiTotal"] = row["RealisasiTotal"]
                merged_data[key]["RealisasiFLV"] = row["RealisasiFLV"]
            else:
                merged_data[key] = {
                    "NomorSPK": row["NomorSPK"],
                    "Tanggal": row["Tanggal"],
                    "Item": row["Item"],
                    "Casing": row["Casing"],
                    "Komponen": row["Komponen"],
                    "FLV": row["FLV"],
                    "RencanaTotal": "",
                    "RencanaFLV": "",
                    "RealisasiTotal": row["RealisasiTotal"],
                    "RealisasiFLV": row["RealisasiFLV"],
                }

        # Ubah ke list
        data_to_send = list(merged_data.values())

        # Jika tombol "Simpan Data" ditekan
        if submit_button:
            try:
                formatted_tanggal = tanggal.strftime("%Y-%m-%d") 
                
                # Data yang akan dikirim ke Apps Script 
                for row in data_to_send:
                    data = {
                        "action": "add_data",
                        "NomorSPK": row["NomorSPK"],
                        "Tanggal": row["Tanggal"],
                        "Item": row["Item"],
                        "Casing": row["Casing"],
                        "Komponen": row["Komponen"],
                        "FLV": row["FLV"],
                        "RencanaTotal": row.get("RencanaTotal", ""),
                        "RencanaFLV": row.get("RencanaFLV", ""),
                        "RealisasiTotal": row.get("RealisasiTotal", ""),
                        "RealisasiFLV": row.get("RealisasiFLV", ""),
                        "Keterangan": keterangan
                    }

                    response = requests.post(APPS_SCRIPT_URL_ES, json=data)
                    result = response.json()

                if result.get("status") == "success":
                    st.success("Data berhasil ditambahkan!")
                    st.session_state.form_add_reset_pop = True
                    st.rerun() 

                else:
                    st.error(f"Terjadi kesalahan: {result.get('error')}")
            except Exception as e:
                st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    run()