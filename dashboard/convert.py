import pandas as pd

# Ganti 'data_benefit.xlsx' sesuai nama file Excel kamu
df = pd.read_excel('dashboard/raw.xlsx')

# Simpan ke format JSON
df.to_json('data_pegawai.json', orient='records', indent=4)
print("Berhasil Convert ke data_pegawai.json!")