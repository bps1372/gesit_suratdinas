import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
import io
import datetime
import requests

# Konfigurasi Halaman
st.set_page_config(page_title="Generator Laporan Perjalanan Dinas", layout="wide")

st.title("Generator Laporan Perjalanan Dinas")

# URL TEMPLAT GITHUB (Ganti dengan URL Raw Anda)
GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/username/nama-repo/main/templat.docx"

def load_template_from_github(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except Exception as e:
        st.error(f"Gagal mengunduh templat: {e}")
        return None

# --- FUNGSI FORMAT TANGGAL INDONESIA ---
def format_tanggal_indo(date_obj, include_hari=False):
    bulan_indo = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    hari_indo = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    
    tanggal = date_obj.day
    bulan = bulan_indo[date_obj.month]
    tahun = date_obj.year
    
    if include_hari:
        nama_hari = hari_indo[date_obj.weekday()]
        return f"{nama_hari}, {tanggal} {bulan} {tahun}"
    return f"{tanggal} {bulan} {tahun}"

# --- Data Referensi ---
list_nama = ["siA", "siB", "siC", "Lainnya"]
list_jabatan = ["Kepala BPS", "Statistisi Ahli Madya", "Statistisi Ahli Muda", "Statistisi Ahli Pertama", "Statistisi Mahir", "Statistisi Terampil", "Pranata Komputer Ahli Pertama", "Pranata Komputer Ahli Muda", "Pranata Komputer Ahli Madya", "Staf BPS", "Staf Subbagian Umum", "Kepala Subbagian Umum", "APK APBN Ahli Pertama", "APK APBN Muda", "APK APBN Madya", "Lainnya"]
list_golongan = ["IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", "II/c", "IX", "VII", "V", "Lainnya"]

# --- Form Input ---
st.subheader("1. Informasi Perjalanan")
kegiatan = st.text_input("Kegiatan")
tujuan = st.text_input("Tujuan Perjalanan")

col1, col2, col3 = st.columns(3)
with col1:
    pilihan_nama = st.selectbox("Nama", list_nama)
    nama = st.text_input("Ketik Nama Anda") if pilihan_nama == "Lainnya" else pilihan_nama
with col2:
    pilihan_jabatan = st.selectbox("Jabatan", list_jabatan)
    jabatan = st.text_input("Ketik Jabatan Anda") if pilihan_jabatan == "Lainnya" else pilihan_jabatan
with col3:
    pilihan_golongan = st.selectbox("Pangkat/Golongan", list_golongan)
    golongan = st.text_input("Ketik Pangkat/Golongan Anda") if pilihan_golongan == "Lainnya" else pilihan_golongan

waktu = st.date_input("Waktu Perjalanan", [])

dates = []
if len(waktu) == 2:
    dates = [waktu[0] + datetime.timedelta(days=i) for i in range((waktu[1] - waktu[0]).days + 1)]
elif len(waktu) == 1:
    dates = [waktu[0]]

data_harian = []
if dates:
    st.subheader(f"2. Detail Kegiatan Harian ({len(dates)} Hari)")
    for i, dt in enumerate(dates):
        tanggal_str = format_tanggal_indo(dt, include_hari=True)
        with st.expander(f"Detail - {tanggal_str}", expanded=(i==0)):
            c1, c2 = st.columns(2)
            with c1: jam_mulai = st.time_input("Jam Mulai", key=f"jm_{i}")
            with c2: jam_akhir = st.time_input("Jam Akhir", key=f"ja_{i}")
            uraian = st.text_area("Uraian", key=f"ur_{i}")
            foto = st.file_uploader("Upload Dokumentasi", type=['png', 'jpg', 'jpeg'], key=f"ft_{i}")
            
            data_harian.append({
                "tanggal": tanggal_str,
                "jam_mulai": jam_mulai.strftime('%H:%M'),
                "jam_akhir": jam_akhir.strftime('%H:%M'),
                "uraian": uraian,
                "foto": foto
            })

st.markdown("---")
if st.button("Generate Laporan", type="primary"):
    if not dates:
        st.error("Pilih waktu perjalanan!")
    else:
        with st.spinner("Memproses dokumen & menyamakan font templat..."):
            template_bytes = load_template_from_github(GITHUB_TEMPLATE_URL)
            if template_bytes:
                try:
                    doc = Document(template_bytes)
                    
                    if len(dates) > 1:
                        waktu_str = f"{format_tanggal_indo(dates[0])} - {format_tanggal_indo(dates[-1])}"
                    else:
                        waktu_str = format_tanggal_indo(dates[0])
                    
                    replacements = {
                        "<kegiatan>": kegiatan.upper() if kegiatan else "", 
                        "<nama>": nama, 
                        "<jabatan>": jabatan,
                        "<golongan>": golongan, 
                        "<tujuan>": tujuan, 
                        "<waktu>": waktu_str
                    }
                    
                    # 1. REPLACE TEKS STATIS (Otomatis mempertahankan font & bold asli dari run di Word)
                    for p in doc.paragraphs:
                        for key, val in replacements.items():
                            for run in p.runs:
                                if key in run.text: 
                                    run.text = run.text.replace(key, str(val))
                    
                    # 2. PROSES TABEL DINAMIS DENGAN PENYAMAAN FONT
                    if len(doc.tables) > 0:
                        tabel_kegiatan = doc.tables[0]
                        
                        # Loop cari baris contoh (<haritanggal>)
                        row_to_delete = None
                        for r in tabel_kegiatan.rows:
                            if "<haritanggal>" in r.cells[0].text:
                                row_to_delete = r
                                break
                        
                        # Inisialisasi variabel penyimpan gaya font templat
                        saved_font_name = "Arial"  # Standar fallback jika di Word berupa 'None'
                        saved_font_size = Pt(11)   # Standar fallback
                        
                        if row_to_delete:
                            # Ambil sampel font dari text run pertama di sel pertama baris templat
                            if row_to_delete.cells[0].paragraphs[0].runs:
                                run_sampel = row_to_delete.cells[0].paragraphs[0].runs[0]
                                if run_sampel.font.name:
                                    saved_font_name = run_sampel.font.name
                                if run_sampel.font.size:
                                    saved_font_size = run_sampel.font.size
                            
                            # Hapus baris penanda templat asli
                            tabel_kegiatan._tbl.remove(row_to_delete._tr)
                        
                        # Tambahkan baris data baru dan paksa gunakan font hasil capture
                        for hari in data_harian:
                            row_cells = tabel_kegiatan.add_row().cells
                            
                            # Kolom 1: Hari/Tanggal
                            p0 = row_cells[0].paragraphs[0]
                            run0 = p0.add_run(hari['tanggal'])
                            run0.font.name = saved_font_name
                            run0.font.size = saved_font_size
                            
                            # Kolom 2: Jam
                            p1 = row_cells[1].paragraphs[0]
                            run1 = p1.add_run(f"{hari['jam_mulai']} - {hari['jam_akhir']}")
                            run1.font.name = saved_font_name
                            run1.font.size = saved_font_size
                            
                            # Kolom 3: Uraian Kegiatan
                            p2 = row_cells[2].paragraphs[0]
                            run2 = p2.add_run(hari['uraian'])
                            run2.font.name = saved_font_name
                            run2.font.size = saved_font_size
                            
                            # Kolom 4: Dokumentasi Foto
                            if hari['foto']:
                                p3 = row_cells[3].paragraphs[0]
                                run3 = p3.add_run()
                                run3.add_picture(hari['foto'], width=Inches(1.5))
                                
                        # Proses juga text replacement untuk tag statis di luar tabel / dalam komponen sel lain
                        for row in tabel_kegiatan.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    for key, val in replacements.items():
                                        for run in p.runs:
                                            if key in run.text:
                                                run.text = run.text.replace(key, str(val))
                    
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)
                    st.success("Laporan Berhasil Dibuat!")
                    st.download_button("📥 Download Laporan", bio, f"Laporan_Perdin_{nama.replace(' ','_')}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Error: {e}")
