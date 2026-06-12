import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import datetime
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Generator Laporan Perjalanan", layout="centered")

# ==========================================
# DICTIONARY UNTUK FORMAT TANGGAL INDONESIA
# ==========================================
hari_id = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
bulan_id = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}

def format_tanggal_indonesia(date_obj):
    hari = hari_id[date_obj.weekday()]
    bulan = bulan_id[date_obj.month]
    return f"{hari}, {date_obj.day} {bulan} {date_obj.year}"

# ==========================================
# LIST DATA DROPDOWN
# ==========================================
list_nama = ["siA", "siB", "siC", "Lainnya"]
list_jabatan = [
    "Kepala BPS", "Statistisi Ahli Madya", "Statistisi Ahli Muda", "Statistisi Ahli Pertama",
    "Statistisi Mahir", "Statistisi Terampil", "Pranata Komputer Ahli Pertama", 
    "Pranata Komputer Ahli Muda", "Pranata Komputer Ahli Madya", "Staf BPS", 
    "Staf Subbagian Umum", "Kepala Subbagian Umum", "APK APBN Ahli Pertama", 
    "APK APBN Muda", "APK APBN Madya", "Lainnya"
]
list_golongan = ["IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", "II/c", "IX", "VII", "V", "Lainnya"]

# ==========================================
# ANTARMUKA STREAMLIT
# ==========================================
st.title("📄 Generator Laporan Perjalanan")
st.markdown("Isi form di bawah ini untuk menghasilkan dokumen laporan perjalanan secara otomatis.")

# 1. Input Teks Biasa
kegiatan = st.text_input("Kegiatan")
tujuan = st.text_input("Tujuan Perjalanan")

# 2. Dropdown dengan Opsi "Lainnya"
col1, col2, col3 = st.columns(3)

with col1:
    pilihan_nama = st.selectbox("Nama", list_nama)
    nama = st.text_input("Ketik Nama Baru:") if pilihan_nama == "Lainnya" else pilihan_nama

with col2:
    pilihan_jabatan = st.selectbox("Jabatan", list_jabatan)
    jabatan = st.text_input("Ketik Jabatan Baru:") if pilihan_jabatan == "Lainnya" else pilihan_jabatan

with col3:
    pilihan_golongan = st.selectbox("Pangkat/Golongan", list_golongan)
    golongan = st.text_input("Ketik Golongan Baru:") if pilihan_golongan == "Lainnya" else pilihan_golongan

# 3. Input Rentang Waktu (Tanggal)
st.subheader("📅 Waktu Perjalanan")
rentang_tanggal = st.date_input("Pilih Tanggal Mulai hingga Tanggal Selesai", [])

start_date, end_date = None, None
if len(rentang_tanggal) == 2:
    start_date, end_date = rentang_tanggal
elif len(rentang_tanggal) == 1:
    start_date = end_date = rentang_tanggal[0]

# 4. Form Dinamis Per Hari
items_data = []

if start_date and end_date:
    delta = end_date - start_date
    jumlah_hari = delta.days + 1
    
    st.write(f"**Total Perjalanan: {jumlah_hari} Hari**")
    
    # Looping untuk membuat input di setiap harinya
    for i in range(jumlah_hari):
        current_date = start_date + datetime.timedelta(days=i)
        tgl_format = format_tanggal_indonesia(current_date)
        
        with st.expander(f"Hari ke-{i+1}: {tgl_format}", expanded=True):
            jam = st.text_input(f"Jam", key=f"jam_{i}")
            uraian = st.text_area(f"Uraian Kegiatan", key=f"uraian_{i}")
            foto = st.file_uploader(f"Upload Dokumentasi", type=['png', 'jpg', 'jpeg'], key=f"foto_{i}")
            
            # Simpan data sementara dalam list dict
            items_data.append({
                "haritanggaltahun": tgl_format,
                "jam": jam,
                "uraian": uraian,
                "foto_file": foto # Disimpan sementara sebagai objek file stream
            })

# ==========================================
# PROSES GENERATE WORD
# ==========================================
st.divider()
if st.button("Generate Laporan Word", type="primary"):
    if not (kegiatan and tujuan and start_date):
        st.warning("Mohon lengkapi Kegiatan, Tujuan, dan Rentang Waktu terlebih dahulu!")
    else:
        try:
            # Load Template
            doc = DocxTemplate("templat.docx")
            
            # Format teks waktu perjalanan di kop surat
            waktu_teks = format_tanggal_indonesia(start_date)
            if start_date != end_date:
                waktu_teks += f" s.d. {format_tanggal_indonesia(end_date)}"
                
            # Proses Items List (termasuk gambar)
            items_render = []
            for item in items_data:
                # Logika agar foto bisa masuk ke Word menggunakan InlineImage
                if item["foto_file"] is not None:
                    gambar_doc = InlineImage(doc, item["foto_file"], width=Mm(45)) # Lebar foto 4.5 cm agar rapi di tabel
                else:
                    gambar_doc = "" # Kosong jika tidak ada foto
                    
                items_render.append({
                    "haritanggaltahun": item["haritanggaltahun"],
                    "jam": item["jam"],
                    "uraian": item["uraian"],
                    "foto": gambar_doc
                })

            # Siapkan Dictionary Konteks
            context = {
                "kegiatan": kegiatan,
                "nama": nama,
                "jabatan": jabatan,
                "golongan": golongan,
                "tujuan": tujuan,
                "waktu": waktu_teks,
                "items": items_render
            }
            
            # Render dokumen
            doc.render(context)
            
            # Simpan ke Buffer (Bisa langsung didownload tanpa save ke disk)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("🎉 Dokumen Laporan Perjalanan berhasil dibuat!")
            st.download_button(
                label="⬇️ Download Laporan (.docx)",
                data=buffer,
                file_name=f"Laporan_Perjalanan_{nama.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses template: {e}")
            st.info("Pastikan file 'templat.docx' berada di folder yang sama dengan aplikasi.")
