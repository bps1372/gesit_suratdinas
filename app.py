import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import datetime

# Konfigurasi Halaman
st.set_page_config(page_title="Generator Laporan Perjalanan Dinas", layout="wide")

st.title("Generator Laporan Perjalanan Dinas")
st.markdown("Otomatisasi pembuatan surat laporan perjalanan dinas. Silakan isi form di bawah ini.")

# --- 1. Upload Template ---
st.subheader("1. Unggah Templat")
template_file = st.file_uploader("Upload file templat (.docx)", type=['docx'])

# --- 2. Data Referensi ---
list_nama = ["siA", "siB", "siC", "Lainnya"]

list_jabatan = [
    "Kepala BPS", "Statistisi Ahli Madya", "Statistisi Ahli Muda", 
    "Statistisi Ahli Pertama", "Statistisi Mahir", "Statistisi Terampil", 
    "Pranata Komputer Ahli Pertama", "Pranata Komputer Ahli Muda", 
    "Pranata Komputer Ahli Madya", "Staf BPS", "Staf Subbagian Umum", 
    "Kepala Subbagian Umum", "APK APBN Ahli Pertama", "APK APBN Muda", 
    "APK APBN Madya", "Lainnya"
]

list_golongan = [
    "IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", 
    "II/c", "IX", "VII", "V", "Lainnya"
]

# --- 3. Form Input Utama ---
st.subheader("2. Informasi Perjalanan")
kegiatan = st.text_input("Kegiatan")
tujuan = st.text_input("Tujuan Perjalanan")

col1, col2, col3 = st.columns(3)

with col1:
    pilihan_nama = st.selectbox("Nama", list_nama)
    if pilihan_nama == "Lainnya":
        nama = st.text_input("Ketik Nama Anda")
    else:
        nama = pilihan_nama

with col2:
    pilihan_jabatan = st.selectbox("Jabatan", list_jabatan)
    if pilihan_jabatan == "Lainnya":
        jabatan = st.text_input("Ketik Jabatan Anda")
    else:
        jabatan = pilihan_jabatan

with col3:
    pilihan_golongan = st.selectbox("Pangkat/Golongan", list_golongan)
    if pilihan_golongan == "Lainnya":
        golongan = st.text_input("Ketik Pangkat/Golongan Anda")
    else:
        golongan = pilihan_golongan

# --- 4. Input Kalender (Waktu Perjalanan) ---
waktu = st.date_input("Waktu Perjalanan (Pilih rentang tanggal)", [])

dates = []
if len(waktu) == 2:
    start_date, end_date = waktu
    delta = end_date - start_date
    num_days = delta.days + 1
    dates = [start_date + datetime.timedelta(days=i) for i in range(num_days)]
elif len(waktu) == 1:
    dates = [waktu[0]]

# --- 5. Dynamic Input Berdasarkan Hari ---
data_harian = []
if dates:
    st.subheader(f"3. Detail Kegiatan Harian ({len(dates)} Hari)")
    
    hari_indo = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    
    for i, dt in enumerate(dates):
        nama_hari = hari_indo[dt.weekday()]
        tanggal_str = f"{nama_hari}, {dt.strftime('%d-%m-%Y')}"
        
        with st.expander(f"Detail Hari {i+1} - {tanggal_str}", expanded=(i==0)):
            c1, c2 = st.columns(2)
            with c1:
                jam_mulai = st.time_input(f"Jam Mulai", key=f"jm_{i}")
            with c2:
                jam_akhir = st.time_input(f"Jam Akhir", key=f"ja_{i}")
            
            uraian = st.text_area(f"Uraian Kegiatan", key=f"ur_{i}")
            foto = st.file_uploader(f"Upload Dokumentasi (Gambar)", type=['png', 'jpg', 'jpeg'], key=f"ft_{i}")
            
            data_harian.append({
                "tanggal": tanggal_str,
                "jam_mulai": jam_mulai.strftime('%H:%M'),
                "jam_akhir": jam_akhir.strftime('%H:%M'),
                "uraian": uraian,
                "foto": foto
            })

# --- 6. Proses Pembuatan Dokumen ---
st.markdown("---")
if st.button("Generate Laporan", type="primary"):
    if not template_file:
        st.error("Mohon unggah file templat terlebih dahulu!")
    elif not dates:
        st.error("Mohon pilih waktu perjalanan!")
    else:
        try:
            # Buka dokumen dari file upload
            doc = Document(template_file)
            
            # Buat format string untuk <waktu>
            if len(dates) > 1:
                waktu_str = f"{dates[0].strftime('%d-%m-%Y')} s.d. {dates[-1].strftime('%d-%m-%Y')}"
            else:
                waktu_str = dates[0].strftime('%d-%m-%Y')
            
            # Dictionary penggantian tag
            replacements = {
                "<kegiatan>": kegiatan,
                "<nama>": nama,
                "<jabatan>": jabatan,
                "<golongan>": golongan,
                "<tujuan>": tujuan,
                "<waktu>": waktu_str
            }
            
            # Fungsi replace untuk Paragraf
            for p in doc.paragraphs:
                for key, val in replacements.items():
                    if key in p.text:
                        # Metode paling aman agar format teks utuh (bold/italic tidak hilang)
                        for run in p.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, str(val))
                        # Fallback jika tag terpotong antar 'run'
                        if key in p.text:
                            p.text = p.text.replace(key, str(val))
            
            # Fungsi replace untuk Tabel (bagian Non-Dinamis)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, val in replacements.items():
                                if key in p.text:
                                    for run in p.runs:
                                        if key in run.text:
                                            run.text = run.text.replace(key, str(val))
                                    if key in p.text:
                                        p.text = p.text.replace(key, str(val))
            
            # --- Proses Tabel Dinamis (Untuk Detail Harian) ---
            # Asumsi: Tabel detail kegiatan adalah tabel pertama (indeks 0) di docx
            if len(doc.tables) > 0:
                tabel_kegiatan = doc.tables[0]
                
                # Cari baris yang mengandung tag <haritanggal>
                row_to_delete = None
                for row in tabel_kegiatan.rows:
                    if "<haritanggal>" in row.cells[0].text:
                        row_to_delete = row
                        break
                
                # Hapus baris templat tersebut dari dalam XML docx
                if row_to_delete:
                    tabel_kegiatan._tbl.remove(row_to_delete._tr)
                
                # Tambahkan baris baru secara berulang berdasarkan data per hari
                for hari in data_harian:
                    row_cells = tabel_kegiatan.add_row().cells
                    row_cells[0].text = hari['tanggal']
                    row_cells[1].text = f"{hari['jam_mulai']} - {hari['jam_akhir']}"
                    row_cells[2].text = hari['uraian']
                    
                    # Sisipkan gambar ke sel ke-4 jika pengguna mengunggahnya
                    if hari['foto'] is not None:
                        paragraph = row_cells[3].paragraphs[0]
                        run = paragraph.add_run()
                        # Diatur lebar 1.5 inci agar rapi masuk dalam kolom tabel
                        run.add_picture(hari['foto'], width=Inches(1.5))
            
            # Simpan hasil modifikasi ke Byte stream agar bisa didownload
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.success("Berhasil membuat laporan perjalanan dinas!")
            
            st.download_button(
                label="📥 Download Hasil Laporan",
                data=bio,
                file_name=f"Laporan_Perdin_{nama.replace(' ','_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses dokumen: {e}")
