import re
import os
import shutil
from pathlib import Path

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ============================================================
# CONFIGURASI AWAL (Update: Sekarang ada 2 Template)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent

# Mendefinisikan 2 jalur template
TEMPLATE_LOLOS = BASE_DIR / "template" / "template_lolos.docx"
TEMPLATE_TIDAK_LOLOS = BASE_DIR / "template" / "template_tidak_lolos.docx"

OUTPUT_DIR = BASE_DIR / "output"
ZIP_BASE_PATH = BASE_DIR / "Hasil_Surat_Automasi"
ZIP_FILE_PATH = BASE_DIR / "Hasil_Surat_Automasi.zip"

# ============================================================
# FUNGSI HELPER
# ============================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    if hasattr(value, "is_integer") and value.is_integer():
        return str(int(value))
    return str(value).strip()

def normalize(value):
    value = clean_text(value).upper()
    value = re.sub(r"\s+", " ", value)
    return value

def format_date(value):
    if pd.isna(value):
        return ""
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return clean_text(value)
    months = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{dt.day} {months[dt.month - 1]} {dt.year}"

def format_rupiah(value):
    if pd.isna(value) or clean_text(value) == "":
        return ""
    try:
        number = float(value)
        if number.is_integer():
            number = int(number)
            return f"Rp{number:,}".replace(",", ".")
    except (ValueError, TypeError):
        pass
    text = clean_text(value)
    digits = re.sub(r"[^\d]", "", text)
    if digits:
        return f"Rp{int(digits):,}".replace(",", ".")
    return text

def get_angka_murni(value):
    """Fungsi baru untuk mengambil angka gajinya saja untuk dihitung"""
    text = clean_text(value)
    digits = re.sub(r"[^\d]", "", text)
    if digits:
        return int(digits)
    return 0

def get_area_from_region(row):
    region = normalize(row.get("REGION", ""))
    target_column = f"AREA ({region})"
    
    if target_column in row.index:
        return clean_text(row[target_column])

    target_norm = normalize(target_column)
    for column in row.index:
        if normalize(column) == target_norm:
            return clean_text(row[column])
    return ""

def build_address_lookup(master_df):
    lookup = {}
    for _, item in master_df.iterrows():
        area = normalize(item["AREA"])
        branch = normalize(item["NAMA CABANG"])
        address = clean_text(item["NAMA JALAN"])
        key = (area, branch)
        if key not in lookup:
            lookup[key] = address
    return lookup

# ============================================================
# ANTARMUKA WEB
# ============================================================
def main():
    st.set_page_config(page_title="Automasi Surat Divisi", page_icon="📄", layout="centered")
    
    st.title("📄 Pembuat Dokumen Pengajuan (Otomatis Seleksi)")
    st.markdown("Sistem akan menyeleksi otomatis apakah nasabah menggunakan **Template Lolos** atau **Template Tidak Lolos** berdasarkan status payroll dan akseptasi pendapatan.")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        file_pengajuan = st.file_uploader("1. Upload Data Pengajuan", type=["xlsx"])
        
        # --- FITUR BARU: PILIH SHEET EXCEL ---
        if file_pengajuan:
            excel_file = pd.ExcelFile(file_pengajuan)
            pilihan_sheet = st.selectbox("Pilih Halaman (Sheet) yang mau diproses:", excel_file.sheet_names)
            
    with col2:
        file_master = st.file_uploader("2. Upload Master Alamat", type=["xlsx"])

    if st.button("🚀 Proses Dokumen", use_container_width=True, type="primary"):
        if not file_pengajuan or not file_master:
            st.error("Silakan unggah kedua file Excel terlebih dahulu!")
            return
        
        # Cek apakah kedua template sudah disiapkan
        if not TEMPLATE_LOLOS.exists() or not TEMPLATE_TIDAK_LOLOS.exists():
            st.error(f"Template tidak lengkap! Pastikan ada file 'template_lolos.docx' dan 'template_tidak_lolos.docx' di folder template.")
            return

        with st.spinner("Membaca dan menyiapkan data..."):
            try:
                # --- UPDATE: Baca Excel berdasarkan Sheet yang dipilih user ---
                pengajuan = pd.read_excel(file_pengajuan, sheet_name=pilihan_sheet)
                
                # ============================================================
                # FITUR BARU: SATPAM PENGECEK KOLOM EXCEL
                # ============================================================
                kolom_wajib = [
                    "STATUS PAYROLL NASABAH", 
                    "NAMA NASABAH",
                    "REGION",
                    "AREA",
                    "AKSEPTASI PENDAPATAN",
                    "NAMA CABANG",
                    "TARGET MARKET",
                    "INSTANSI NASABAH",
                    "PLAFOND PENGAJUAN NASABAH"
                ]
                
                # Cek adakah kolom wajib yang hilang atau beda ketikan
                kolom_hilang = [kol for kol in kolom_wajib if kol not in pengajuan.columns]
                
                if kolom_hilang:
                    st.warning(f"⚠️ Waduh! Ada nama kolom yang beda atau hilang di Excel Pengajuan:\n\n**{', '.join(kolom_hilang)}**")
                    st.info("💡 Pastikan nama kolom di atas sama persis dengan format yang ditentukan (jangan ada spasi tambahan atau salah ketik). Silakan perbaiki Excel-nya dan proses ulang ya!")
                    return # Hentikan proses jika kolom tidak sesuai agar tidak error merah
                # ============================================================

                master = pd.read_excel(file_master, sheet_name="Nama Jalan")
                address_lookup = build_address_lookup(master)

                if OUTPUT_DIR.exists():
                    shutil.rmtree(OUTPUT_DIR)
                OUTPUT_DIR.mkdir(exist_ok=True)
                
                if ZIP_FILE_PATH.exists():
                    os.remove(ZIP_FILE_PATH)

                data_testing = pengajuan 
                total_data = len(data_testing)

                berhasil = 0
                gagal = 0
                daftar_data_gagal = []
                dokumen_berhasil = [] 

            except Exception as e:
                st.error(f"Gagal membaca file Excel: {e}")
                return

        progress_bar = st.progress(0)
        status_text = st.empty()

        for index, row in data_testing.iterrows():
            urutan = index + 1
            region = ""
            area = ""
            branch = ""
            nama_nasabah = f"TanpaNama_{urutan}"

            try:
                region = clean_text(row.get("REGION", ""))
                area = clean_text(row.get("AREA", ""))
                branch = clean_text(row.get("NAMA CABANG", ""))
                nama_nasabah = clean_text(row.get("NAMA NASABAH", nama_nasabah))

                address_key = (normalize(area), normalize(branch))
                address = address_lookup.get(address_key)
                if address is None:
                    address = ""

                # ============================================================
                # LOGIKA "OTAK" PENENTUAN TEMPLATE (REVISI FINAL)
                # ============================================================
                # 1. Ambil teks status payroll dan jadikan huruf kecil semua
                status_payroll = clean_text(row.get("STATUS PAYROLL NASABAH")).lower()
                
                # 2. Ambil angka gajinya saja
                kolom_gaji = "AKSEPTASI PENDAPATAN"
                nominal_gaji = get_angka_murni(row.get(kolom_gaji))

                # KONDISI 1: Jika Committed -> FIX TIDAK LOLOS (Gaji diabaikan)
                if "ctp" in status_payroll or "commit" in status_payroll:
                    template_terpilih = TEMPLATE_TIDAK_LOLOS
                    label_status = "TIDAK_LOLOS"
                    
                # KONDISI 2 & 3: Jika Efektif -> Cek Nominal Gaji
                elif "payroll" in status_payroll:
                    if nominal_gaji > 5000000:
                        template_terpilih = TEMPLATE_LOLOS
                        label_status = "LOLOS"
                    else:
                        template_terpilih = TEMPLATE_TIDAK_LOLOS
                        label_status = "TIDAK_LOLOS"
                        
                # Default jika kolom kosong atau statusnya di luar 2 pilihan tersebut
                else:
                    template_terpilih = TEMPLATE_TIDAK_LOLOS
                    label_status = "TIDAK_LOLOS"
                # ============================================================

                mapping = {
                    "TANGGAL": format_date(row.get("Start Time")),
                    "NO_SURAT_CF2": clean_text(row.get("No Surat CF2")),
                    
                    # === INI VERSI TEKS BIASA (UNTUK DI LUAR BLOK / PARAGRAF) ===
                    # Pakai .title() agar huruf depannya saja yang besar
                    "REG": region.title(),          
                    "CAB": branch.title(),          
                    "AREA": area.title(),           
                    "NAMA": nama_nasabah.title(),   
                    "KETERANGAN_REKOMENDASI": clean_text(row.get("KETERANGAN REKOMENDASI")).title(),
                    # ============================================================
                    
                    # === INI VERSI BLOK JUDUL (CAPSLOCK SEMUA) ===
                    "REG_KAPITAL": region.upper(),
                    "AREA_KAPITAL": area.upper(),
                    "KETERANGAN_REKOMENDASI_KAPITAL": clean_text(row.get("KETERANGAN REKOMENDASI")).upper(),
                    # ============================================================
                    
                    "WISE": clean_text(row.get("NO APLIKASI WISE")),
                    "SEG": clean_text(row.get("TARGET MARKET")).title(), # Boleh dipakaikan .title() juga
                    "INST": clean_text(row.get("INSTANSI NASABAH")).upper(), # Instansi biasanya bagus huruf besar semua
                    "PLAF": format_rupiah(row.get("PLAFOND PENGAJUAN NASABAH")),
                    "AKS": format_rupiah(row.get(kolom_gaji)),
                    "KET": clean_text(row.get("STATUS PAYROLL NASABAH")).title(),
                    "ALAMAT_AREA": address,
                }
                # Proses menggunakan template yang sudah diseleksi oleh IF-ELSE di atas
                doc = DocxTemplate(template_terpilih)
                doc.render(mapping)

                # Nama file akan diberi embel-embel LOLOS / TIDAK_LOLOS agar mudah dibedakan saat di-download
                safe_nama = nama_nasabah.replace("/", "_").replace("\\", "_")
                nama_file_output = f"{urutan}_{label_status}_{safe_nama}.docx"
                output_file = OUTPUT_DIR / nama_file_output
                
                doc.save(output_file)
                
                berhasil += 1
                dokumen_berhasil.append(output_file)

            except Exception as e:
                gagal += 1
                daftar_data_gagal.append({
                    "Baris di Excel": urutan + 1,
                    "Nama Nasabah": nama_nasabah,
                    "Alasan Gagal": str(e)
                })

            percent_complete = int((urutan / total_data) * 100)
            progress_bar.progress(percent_complete)
            status_text.text(f"Memproses {urutan}/{total_data} dokumen...")

        if daftar_data_gagal:
            df_error = pd.DataFrame(daftar_data_gagal)
            file_error_path = OUTPUT_DIR / "Laporan_Data_Gagal.xlsx"
            df_error.to_excel(file_error_path, index=False)

        shutil.make_archive(str(ZIP_BASE_PATH), 'zip', str(OUTPUT_DIR))

        st.success(f"Proses Selesai! Berhasil: {berhasil} Dokumen | Gagal: {gagal} Dokumen.")

        st.markdown("---")
        st.markdown("### 📦 Download Seluruh Hasil")
        with open(ZIP_FILE_PATH, "rb") as fp:
            st.download_button(
                label="📥 Download Hasil (ZIP)",
                data=fp,
                file_name="Hasil_Surat_Automasi.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
            
        if dokumen_berhasil:
            st.markdown("### 📄 Download Dokumen Satuan")
            with st.expander("Klik di sini untuk melihat dan download surat satu per satu"):
                for doc_path in dokumen_berhasil:
                    with open(doc_path, "rb") as f_doc:
                        st.download_button(
                            label=f"⬇️ {doc_path.name}",
                            data=f_doc,
                            file_name=doc_path.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=doc_path.name
                        )

if __name__ == "__main__":
    main()