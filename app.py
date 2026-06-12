import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import timedelta, time
import io

st.title("Generator Laporan Perjalanan Dinas")

# --- DATA LIST ---
list_nama = ["siA", "siB", "siC", "Lainnya"]
list_jabatan = [
    "Kepala BPS", "Statistisi Ahli Madya", "Statistisi Ahli Muda", 
    "Statistisi Ahli Pertama", "Statistisi Mahir", "Statistisi Terampil", 
    "Pranata Komputer Ahli Pertama", "Pranata Komputer Ahli Muda", 
    "Pranata Komputer Ahli Madya", "Staf BPS", "Staf Subbagian Umum", 
    "Kepala Subbagian Umum", "APK APBN Ahli Pertama", "APK APBN Muda", 
    "APK APBN Madya", "Lainnya"
]
list_golongan = ["IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", "II/c", "IX", "VII", "V", "Lainnya"]

# --- INPUT UMUM ---
st.header("Informasi Umum")
kegiatan = st.text_input("Kegiatan")
tujuan = st.text_input("Tujuan Perjalanan")

# Logika Dropdown + Input Manual (Lainnya)
col1, col2, col3 = st.columns(3)

with col1:
    nama_sel = st.selectbox("Nama", list_nama)
    nama = st.text_input("Ketik Nama Manual") if nama_sel == "Lainnya" else nama_sel

with col2:
    jab_sel = st.selectbox("Jabatan", list_jabatan)
    jabatan = st.text_input("Ketik Jabatan Manual") if jab_sel == "Lainnya" else jab_sel

with col3:
    gol_sel = st.selectbox("Pangkat/Golongan", list_golongan)
    golongan = st.text_input("Ketik Golongan Manual") if gol_sel == "Lainnya" else gol_sel

# --- INPUT TANGGAL & TABEL DINAMIS ---
st.header("Waktu & Rincian Perjalanan")
st.info("Pilih rentang tanggal (klik tanggal mulai, lalu klik tanggal selesai).")
waktu_perjalanan = st.date_input("Waktu Perjalanan", [])

items_data = []

# Kamus Hari Bahasa Indonesia
hari_indo = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}

if len(waktu_perjalanan) == 2:
    start_date, end_date = waktu_perjalanan
    jumlah_hari = (end_date - start_date).days + 1
    
    waktu_str = f"{start_date.strftime('%d-%m-%Y')} s.d {end_date.strftime('%d-%m-%Y')}"
    
    st.write("---")
    st.write(f"**Total Perjalanan: {jumlah_hari} Hari**")
    
    # Looping input berdasarkan jumlah hari
    for i in range(jumlah_hari):
        current_date = start_date + timedelta(days=i)
        nama_hari = hari_indo[current_date.weekday()]
        tgl_format = f"{nama_hari}, {current_date.strftime('%d-%m-%Y')}"
        
        st.subheader(f"Hari {i+1}: {tgl_format}")
        
        # Penyesuaian lebar kolom agar input jam memiliki ruang yang cukup
        col_jam, col_uraian, col_foto = st.columns([1.5, 2, 1])
        
        with col_jam:
            st.write("Jam") # Label untuk bagian waktu
            col_mulai, col_selesai = st.columns(2)
            
            with col_mulai:
                jam_mulai = st.time_input("Mulai", value=time(7, 30), key=f"mulai_{i}", label_visibility="collapsed")
            with col_selesai:
                jam_selesai = st.time_input("Selesai", value=time(16, 0), key=f"selesai_{i}", label_visibility="collapsed")
                
            # Menggabungkan format jam dan mengubah (:) menjadi (.)
            jam_gabungan = f"{jam_mulai.strftime('%H.%M')} - {jam_selesai.strftime('%H.%M')}"

        with col_uraian:
            uraian = st.text_area("Uraian Kegiatan", key=f"uraian_{i}", label_visibility="collapsed", placeholder="Ketik uraian di sini...")
        
        with col_foto:
            foto = st.file_uploader("Upload Dokumentasi", type=['jpg', 'jpeg', 'png'], key=f"foto_{i}", label_visibility="collapsed")
            
        items_data.append({
            "haritanggaltahun": tgl_format,
            "jam": jam_gabungan,
            "uraian": uraian,
            "foto_file": foto
        })
        st.write("---")

# --- PROSES GENERATE WORD ---
if st.button("Generate Laporan", type="primary"):
    if len(waktu_perjalanan) != 2:
        st.warning("Mohon pilih rentang tanggal mulai dan selesai terlebih dahulu pada bagian Waktu Perjalanan.")
    else:
        try:
            # Load Template
            doc = DocxTemplate("Laporan Perjadin s.docx")
            
            # Siapkan context dictionary
            context = {
                "kegiatan": kegiatan,
                "nama": nama,
                "jabatan": jabatan,
                "golongan": golongan,
                "tujuan": tujuan,
                "waktu": waktu_str,
                "items": []
            }
            
            # Proses item tabel dan foto
            for item in items_data:
                row_dict = {
                    "haritanggaltahun": item["haritanggaltahun"],
                    "jam": item["jam"],
                    "uraian": item["uraian"],
                }
                
                # Proses gambar menjadi InlineImage jika ada yang diupload
                if item["foto_file"] is not None:
                    # Menyesuaikan ukuran lebar foto ke 40mm (4cm) di dokumen Word
                    img = InlineImage(doc, item["foto_file"], width=Mm(40)) 
                    row_dict["foto"] = img
                else:
                    row_dict["foto"] = "" # Kosongkan jika tidak ada foto
                    
                context["items"].append(row_dict)
            
            # Render dokumen
            doc.render(context)
            
            # Simpan ke dalam memory buffer agar bisa di-download langsung via browser
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("Dokumen berhasil di-generate!")
            st.download_button(
                label="Unduh Dokumen Laporan",
                data=buffer,
                file_name=f"Laporan_Perjadin_{nama}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
            st.info("Pastikan file 'Laporan Perjadin s.docx' sudah sesuai tag formatnya ( {{ variabel }} ) dan berada di folder yang sama dengan app.py.")
