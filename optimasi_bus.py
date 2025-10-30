import pandas as pd
import pulp

# === 1. Baca data Excel tanpa header otomatis ===
armada = pd.read_excel("data_bus.xlsx", sheet_name=0, header=None)
rute = pd.read_excel("data_bus.xlsx", sheet_name=1, header=None)

# === 2. Hapus baris kosong (jika ada) ===
armada = armada.dropna()
rute = rute.dropna()

# === 3. Potong data agar tidak ikut header ganda ===
start_armada = armada[armada.iloc[:, 0].astype(str).str.contains("Armada", case=False)].index
start_rute = rute[rute.iloc[:, 0].astype(str).str.contains("Rute", case=False)].index

if len(start_armada) > 0:
    armada = armada.loc[start_armada[0] + 1:]
if len(start_rute) > 0:
    rute = rute.loc[start_rute[0] + 1:]

# === 4. Tambahkan nama kolom manual ===
armada.columns = ["Armada", "Kapasitas", "BiayaPerKm"]
rute.columns = ["Rute", "Permintaan", "Jarak"]

# === 5. Reset index dan ubah ke tipe numerik ===
armada = armada.reset_index(drop=True)
rute = rute.reset_index(drop=True)
armada["Kapasitas"] = armada["Kapasitas"].astype(float)
armada["BiayaPerKm"] = armada["BiayaPerKm"].astype(float)
rute["Permintaan"] = rute["Permintaan"].astype(float)
rute["Jarak"] = rute["Jarak"].astype(float)

print("=== Data Armada ===")
print(armada)
print("\n=== Data Rute ===")
print(rute)

# === 6. Buat kombinasi (Armada, Rute) ===
kombinasi = [(a, r) for a in armada["Armada"] for r in rute["Rute"]]

# === 7. Buat model Linear Programming ===
model = pulp.LpProblem("Optimasi_Armada_Bus", pulp.LpMinimize)

# === 8. Variabel keputusan (berapa kali armada a melayani rute r) ===
x = pulp.LpVariable.dicts("x", kombinasi, lowBound=0, cat="Continuous")

# === 9. Fungsi tujuan: minimisasi total biaya perjalanan ===
model += pulp.lpSum(
    x[(a, r)] *
    float(armada.loc[armada["Armada"] == a, "BiayaPerKm"].iloc[0]) *
    float(rute.loc[rute["Rute"] == r, "Jarak"].iloc[0])
    for (a, r) in kombinasi
)

# === 10. Kendala kapasitas armada ===
for a in armada["Armada"]:
    model += pulp.lpSum(x[(a, r)] for r in rute["Rute"]) <= float(
        armada.loc[armada["Armada"] == a, "Kapasitas"].iloc[0]
    ), f"Kapasitas_{a}"

# === 11. Kendala permintaan rute ===
for r in rute["Rute"]:
    model += pulp.lpSum(x[(a, r)] for a in armada["Armada"]) == float(
        rute.loc[rute["Rute"] == r, "Permintaan"].iloc[0]
    ), f"Permintaan_{r}"

# === 12. Jalankan solver ===
model.solve(pulp.PULP_CBC_CMD(msg=False))

# === 13. Ambil hasil ===
hasil = []
for (a, r) in kombinasi:
    nilai = pulp.value(x[(a, r)])
    if nilai > 0:
        hasil.append([a, r, nilai])

df_hasil = pd.DataFrame(hasil, columns=["Armada", "Rute", "Jumlah_Perjalanan"])

# === 14. Hitung biaya total per kombinasi ===
def hitung_biaya(row):
    biaya_km = float(armada.loc[armada["Armada"] == row["Armada"], "BiayaPerKm"].iloc[0])
    jarak = float(rute.loc[rute["Rute"] == row["Rute"], "Jarak"].iloc[0])
    return row["Jumlah_Perjalanan"] * biaya_km * jarak

df_hasil["Total_Biaya"] = df_hasil.apply(hitung_biaya, axis=1)

# === 15. Total biaya keseluruhan ===
total_biaya = df_hasil["Total_Biaya"].sum()

# === 16. Tampilkan hasil akhir ===
print("\n=== HASIL OPTIMASI ===")
print(df_hasil.to_string(index=False))
print(f"\nTotal Biaya Minimum: Rp {total_biaya:,.0f}")

# === 17. Simpan hasil ke Excel ===
df_hasil.to_excel("hasil_optimasi_bus.xlsx", index=False)
print("\n✅ Hasil tersimpan di 'hasil_optimasi_bus.xlsx'")
