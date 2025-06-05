import pandas as pd
import streamlit as st
from fpdf import FPDF
import requests

def run():
    APPS_SCRIPT_URL_ES = "https://script.google.com/macros/s/AKfycbxQVBTiXWkWjxaj1J4MY3FCzPIZbohYOuN7FungAmQPFLtIaaIJJbR5b-LvXPAKyC7fTw/exec"
    
    def get_all_data(url):
        try:
            response = requests.get(url, params={"action": "get_data"}, timeout=15)
            response.raise_for_status()
            json_data = response.json()
            return json_data if isinstance(json_data, list) else []
        except requests.exceptions.RequestException as e:
            st.error(f"Terjadi kesalahan saat mengambil data: {e}")
            return []
    
    # Fungsi untuk mendapatkan opsi dari Google Sheets
    def get_options():
        try:
            response = requests.get(APPS_SCRIPT_URL_ES, params={"action": "get_options"}, timeout=10)
            response.raise_for_status()
            options = response.json()
            return options
        except requests.exceptions.RequestException as e:
            st.error(f"Terjadi kesalahan saat mengambil data: {e}")
            return {}
            
    data = get_all_data(APPS_SCRIPT_URL_ES)
    options_data = get_options()
    data_operator = [row for row in options_data.get("OPERATOR", []) if isinstance(row, list) and row]
    data_ka_sie = [row for row in options_data.get("KA SIE", []) if isinstance(row, list) and row]
    ################################################### OVERVIEW #############################################################################
    
    header = [
        "NOMOR SPK", "TANGGAL", "ITEM", "CASING", "KOMPONEN", "FLV",
        "RENCANA TOTAL PEMAKAIAN", "RENCANA FLV", "REALISASI TOTAL PEMAKAIAN",
        "REALISASI FLV", "KETERANGAN"
    ]
    
    df_all = pd.DataFrame(data, columns=header)
    
    st.title("📊 Overview Data SPK Essence")
    
    # Mapping bulan Indonesia ke angka
    bulan_indo = {
        "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
        "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
        "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
    }
    
    # Fungsi parsing tanggal dari format "Senin, 26 Mei 2025"
    def parse_tanggal_indo(tanggal_str):
        try:
            if pd.isna(tanggal_str):
                return pd.NaT
            if ',' in tanggal_str:
                tanggal_str = tanggal_str.split(',')[1].strip()  # ambil setelah koma
            parts = tanggal_str.split(' ')
            if len(parts) != 3:
                return pd.NaT
            hari, bulan, tahun = parts
            bulan_num = bulan_indo.get(bulan)
            if not bulan_num:
                return pd.NaT
            tanggal_baru = f"{hari}-{bulan_num}-{tahun}"
            return pd.to_datetime(tanggal_baru, format="%d-%m-%Y")
        except:
            return pd.NaT
    
    # Parsing tanggal
    df_all["TANGGAL_PARSED"] = df_all["TANGGAL"].apply(parse_tanggal_indo)
    # Filter berdasarkan tanggal parsed
    if df_all["TANGGAL_PARSED"].notna().any():
        min_date = df_all["TANGGAL_PARSED"].min().date()
        max_date = df_all["TANGGAL_PARSED"].max().date()
        tanggal_range = st.date_input("Pilih rentang tanggal", value=(min_date, max_date))
    
        if isinstance(tanggal_range, tuple) and len(tanggal_range) == 2:
            start_date, end_date = tanggal_range
            df_filtered_overview = df_all[
                (df_all["TANGGAL_PARSED"] >= pd.to_datetime(start_date)) &
                (df_all["TANGGAL_PARSED"] <= pd.to_datetime(end_date))
            ]
            
            # Tampilkan semua kolom hasil filter
            st.subheader("📄 Tabel Lengkap Data SPK Essence (Hasil Filter):")
            st.dataframe(df_filtered_overview.drop(columns=["TANGGAL_PARSED"]))
        else:
            st.warning("Silakan pilih rentang tanggal yang valid.")
    else:
        st.warning("Tidak ada data tanggal yang valid untuk ditampilkan.")
    
    ############################################# CETAK PDF ###################################################################################
    st.markdown('----')
    st.title("🧾 Cetak PDF SPK Essence")
    
    if data and isinstance(data[0], list):
        header = [
            "NOMOR SPK", "TANGGAL", "ITEM", "CASING", "KOMPONEN", "FLV",
            "RENCANA TOTAL PEMAKAIAN", "RENCANA FLV", "REALISASI TOTAL PEMAKAIAN",
            "REALISASI FLV", "KETERANGAN"
        ]
        df = pd.DataFrame(data, columns=header)
    
        spk_options = df['NOMOR SPK'].dropna().unique()
        selected_spk = st.selectbox("📌 Pilih Nomor SPK", sorted(spk_options))
    
        if selected_spk:
            df_filtered = df[df["NOMOR SPK"] == selected_spk].copy()
            tanggal_spk = df_filtered["TANGGAL"].iloc[0] if not df_filtered.empty else "-"
            df_filtered["NAMA PAKET"] = df_filtered["KOMPONEN"].astype(str).str.strip().str.upper() + " " + df_filtered["ITEM"].astype(str).str.strip().str.upper()
    
            grouped = df_filtered.groupby("NAMA PAKET").agg({
                "FLV": lambda x: "\n".join(x.dropna()),
                "REALISASI FLV": lambda x: "\n".join(x.dropna().astype(str)),
                "REALISASI TOTAL PEMAKAIAN": "first"
            }).reset_index()
    
            grouped.insert(0, "NO", range(1, len(grouped) + 1))
            grouped["NO. LOT BAHAN"] = ""
            grouped["NOMER PAKET"] = ""
            grouped = grouped[["NO", "NAMA PAKET", "FLV", "NO. LOT BAHAN", "REALISASI FLV", "REALISASI TOTAL PEMAKAIAN", "NOMER PAKET"]]
            grouped.columns = ["NO", "NAMA PAKET", "BAHAN", "NO. LOT BAHAN", "PEMAKAIAN", "TOTAL PAKET", "NOMER PAKET"]
    
            st.markdown("### ✍ Tanda Tangan")
            nama_operator = st.selectbox("👨‍🔧 Pilih Operator Essence", [row[0] for row in data_operator])
            nama_ka_sie = st.selectbox("👩‍💼 Pilih Ka. Sie Essence", [row[0] for row in data_ka_sie])
    
            grouped['OPERATOR'] = nama_operator
            grouped['KA SIE'] = nama_ka_sie
    
            st.subheader("📋 Data yang akan dicetak:")
            st.dataframe(grouped)
    
            def buat_pdf_essence(df, nomor_spk, tanggal_spk, nama_operator, nama_ka_sie):
                pdf = FPDF("L", "mm", "A4")
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
    
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 4, "PT ULTRA PRIMA ABADI", ln=True)
                pdf.cell(0, 4, "DIVISI HARD / SOFT CANDY", ln=True)
                pdf.cell(0, 6, "SURABAYA", ln=True)
                pdf.set_font("Arial", "B", 14)
                pdf.cell(0, 10, "LAPORAN PEMBUATAN ESSENCE DAN PEWARNA", ln=True, align="C")
                pdf.set_font("Arial", "", 11)
                pdf.cell(0, 8, f"NOMOR SPK: {nomor_spk}", ln=True)
                pdf.cell(0, 8, f"TANGGAL: {tanggal_spk}", ln=True)
                pdf.ln(5)
    
                # Perhitungan lebar kolom (dinamis dan tetap)
                pdf.set_font("Arial", "", 10)
                page_width = 297 - 20  # A4 landscape minus margin
                fixed_cols = {"NO. LOT BAHAN": 40, "NOMER PAKET": 40}
                headers = ["NO", "NAMA PAKET", "BAHAN", "NO. LOT BAHAN", "PEMAKAIAN", "TOTAL PAKET", "NOMER PAKET"]
                header_labels = {
                    "TOTAL PAKET": "TOTAL\nPAKET",  # Hanya label visual
                }
                sample_lengths = {
                    "NO": 10,
                    "NAMA PAKET": min(max(pdf.get_string_width(str(val)) for val in df["NAMA PAKET"]) + 5, 65),
                    "BAHAN": max(pdf.get_string_width(str(val).split("\n")[0]) for val in df["BAHAN"]) + 5,
                    "PEMAKAIAN": max(pdf.get_string_width(str(val).split("\n")[0]) for val in df["PEMAKAIAN"]) + 5,
                    "TOTAL PAKET": max(pdf.get_string_width(str(val)) for val in df["TOTAL PAKET"]) + 5
                }
    
                total_flexible = sum(sample_lengths.values())
                available_width = page_width - sum(fixed_cols.values())
                scale = available_width / total_flexible
    
                col_widths = {
                    "NO": sample_lengths["NO"] * scale,
                    "NAMA PAKET": sample_lengths["NAMA PAKET"] * scale,
                    "BAHAN": sample_lengths["BAHAN"] * scale,
                    "NO. LOT BAHAN": fixed_cols["NO. LOT BAHAN"],
                    "PEMAKAIAN": sample_lengths["PEMAKAIAN"] * scale,
                    "TOTAL PAKET": sample_lengths["TOTAL PAKET"] * scale,
                    "NOMER PAKET": fixed_cols["NOMER PAKET"]
                }
    
                # Header tabel
                pdf.set_font("Arial", "B", 10)
                for col in headers:
                    label = header_labels.get(col, col)
                    if "\n" in label:
                        x_before = pdf.get_x()
                        y_before = pdf.get_y()
                        pdf.multi_cell(col_widths[col], 4, label, border=1, align="C")
                        pdf.set_xy(x_before + col_widths[col], y_before)
                    else:
                        pdf.cell(col_widths[col], 8, label, border=1, align="C")
                pdf.ln()
    
                # Data
                pdf.set_font("Arial", "", 10)
                for _, row in df.iterrows():
                    bahan_lines = str(row["BAHAN"]).split("\n")
                    pakai_lines = str(row["PEMAKAIAN"]).split("\n")
                    max_lines = max(len(bahan_lines), len(pakai_lines))
                    row_height = max_lines * 6
    
                    x = pdf.get_x()
                    y = pdf.get_y()
    
                    pdf.multi_cell(col_widths["NO"], row_height, str(row["NO"]), border=1, align="C")
                    pdf.set_xy(x + col_widths["NO"], y)
                    pdf.multi_cell(col_widths["NAMA PAKET"], row_height, str(row["NAMA PAKET"]), border=1)
                    pdf.set_xy(x + sum([col_widths[h] for h in headers[:2]]), y)
                    pdf.multi_cell(col_widths["BAHAN"], 6, str(row["BAHAN"]), border=1)
                    pdf.set_xy(x + sum([col_widths[h] for h in headers[:3]]), y)
                    pdf.multi_cell(col_widths["NO. LOT BAHAN"], row_height, "", border=1)
                    pdf.set_xy(x + sum([col_widths[h] for h in headers[:4]]), y)
                    pdf.multi_cell(col_widths["PEMAKAIAN"], 6, str(row["PEMAKAIAN"]), border=1, align="R")
                    pdf.set_xy(x + sum([col_widths[h] for h in headers[:5]]), y)
                    pdf.multi_cell(col_widths["TOTAL PAKET"], row_height, str(row["TOTAL PAKET"]), border=1, align="R")
                    pdf.set_xy(x + sum([col_widths[h] for h in headers[:6]]), y)
                    pdf.multi_cell(col_widths["NOMER PAKET"], row_height, "", border=1)
    
                    pdf.set_y(y + row_height)
    
                # Spasi ekstra sebelum tanda tangan
                pdf.ln(25)  
    
                pdf.set_font("Arial", "", 11)
                width_col = 80
                page_width = 297 - 20  # Total lebar kertas A4 - margin
    
                pdf.set_font("Arial", "", 11)
                pdf.cell(60, 8, "Operator Essence", ln=0)
                pdf.cell(80, 8, f": {nama_operator}", ln=1)
    
                pdf.cell(60, 8, "Ka. Sie Essence", ln=0)
                pdf.cell(80, 8, f": {nama_ka_sie}", ln=1)
    
                return pdf.output(dest="S").encode("latin-1")
    
            pdf_bytes = buat_pdf_essence(grouped, selected_spk, tanggal_spk, nama_operator, nama_ka_sie)
            st.download_button(
                label="📥 Download PDF",
                data=pdf_bytes,
                file_name=f"SPK_Essence_{selected_spk.replace('/', '_')}.pdf",
                mime="application/pdf"
            )
    else:
        st.error("Data dari Google Sheets kosong atau formatnya tidak sesuai.")
if __name__ == "__main__":
    run()
