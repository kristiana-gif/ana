
# ================================================
# Program: Analisis Nilai Mahasiswa dari Excel
# Library: pandas dan openpyxl
# ================================================

import pandas as pd

# 1️⃣ Import file Excel
file_path = r"C:\Users\ASUS\OneDrive\Documents\Folder Baru\data_mahasiswa.xlsx"
data = pd.read_excel(file_path)
# modifikasi data mahasiswa 
#'NIM Mahasiswa' -> 'NIM'
#'Nama Mahasiswa' -> 'Nama'
#'Nilai 1' -> 'Nilai Tugas'
#'Nilai 2'-> 'Nilai UTS'
#'Nilai 3' -> 'Nilai UAS'
# data.rename(columns={
#     'NIM Mahasiswa':'NIM',
#     'Nama Mahasiswa':'Nama',
#     'Nilai 1': 'Nilai Tugas',
#     'Nilai 2':'Nilai UTS',
#     'Nilai 3':'Nilai UAS',

# },inplace=True)

# 2️⃣ Hitung rata-rata untuk setiap mahasiswa
data['Rata-rata'] = data[['Nilai 1', 'Nilai 2', 'Nilai 3']].mean(axis=1)

# 3️⃣ Cari mahasiswa dengan nilai rata-rata tertinggi
nilai_tertinggi = data['Rata-rata'].max()
terbaik = data[data['Rata-rata'] == nilai_tertinggi]

# 4️⃣ Tampilkan hasil di console
print("=== Data Mahasiswa dengan Nilai Rata-rata ===")
print(data)
print("\nMahasiswa dengan Nilai Tertinggi:")
print(terbaik[['NIM', 'Nama Mahasiswa', 'Rata-rata']])

# 5️⃣ Simpan ke file Excel baru
output_file = "hasil_mahasiswa.xlsx"
data.to_excel(output_file, index=False)

print(f"\nFile hasil telah disimpan ke: {output_file}")

#pesan gaguna