# ================================================
# Program: Analisis Nilai Mahasiswa dari Excel
# Library: pandas dan openpyxl
# ================================================

import pandas as pd

# 1️⃣ Import file Excel
# menggunakan file data_mahasiswa.xlsx yang diberikan.
file_path = r"C:\Users\ASUS\OneDrive\Documents\Folder Baru\data_mahasiswa.xlsx"
data = pd.read_excel(file_path)

# 2️⃣ modifikasi data mahasiswa
# untuk memenuhi syarat tugas, kita akan mengubah nama kolom yang ada:
#'NIM Mahasiswa' -> 'NIM'
#'Nama Mahasiswa' -> 'Nama'
#'Nilai 1' -> 'Nilai Tugas'
#'Nilai 2' -> 'Nilai UTS'
#'Nilai 3' -> 'Nilai UAS'
data.rename(columns={
    'NIM Mahasiswa':'NIM',
    'Nama Mahasiswa':'Nama',
    'Nilai 1':'Nilai Tugas',
    'Nilai 2':'Nilai UTS',
    'Nilai 3':'Nilai UAS'
},inplace=True)

#Hitung nilai akhir dengan bobot 
#Tugas: 30%, UTS: 30%, UAS: 40% 
data['Nilai Akhir']= (data['Nilai Tugas']*0.3) + (data['Nilai UTS']*0.3) + (data['Nilai UAS']*0.4)
#tambahkan kolom status
data['Status'] = data['Nilai Akhir'].apply(lambda nilai: "Lulus" if nilai >= 75 else "Tidak Lulus")

# Urutkan data berdasarkan nilai akhir tertinggi
data_terurut = data.sort_values(by='Nilai Akhir', ascending=False)

# Tampilkan 5 mahasiswa dengan nilai tertinggi
print("=== 5 Mahasiswa dengan Nilai Tertinggi ===")
top_5_mahasiswa = data_terurut.head(5)

# Memformat tampilan nilai agar tidak terlalu banyak angka di belakang koma
top_5_mahasiswa_tampil = top_5_mahasiswa.copy()
top_5_mahasiswa_tampil['Nilai Akhir'] = top_5_mahasiswa_tampil['Nilai Akhir'].round(2)
print(top_5_mahasiswa_tampil[['NIM', 'Nama', 'Nilai Akhir', 'Status']].to_string(index=False))

# Simpan hasil yang sudah diurutkan ke file Excel baru
output_file = "rekap_nilai.xlsx"
data_terurut.to_excel(output_file, index=False)

print(f"\ File rekap telah berhasil disimpan ke: {output_file}")