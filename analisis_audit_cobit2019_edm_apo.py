"""
ANALISIS AUDIT IT GOVERNANCE COBIT 2019 - STUDI KASUS SIMULASI
Objek: PT Nusantara Digital Logistik (fiktif)
Domain: EDM dan APO
Input: Dataset_Audit_COBIT2019_EDM_APO.xlsx, sheet Dataset_Audit

CATATAN METODOLOGI:
Aturan penentuan level pada skrip ini disederhanakan untuk demonstrasi akademik
sampai level 3. Penilaian formal harus mengikuti COBIT 2019 Performance
Management Model, termasuk praktik, aktivitas, atribut, dan bukti yang relevan.
"""

import pandas as pd
import matplotlib.pyplot as plt

FILE_EXCEL = "Dataset_Audit_COBIT2019_EDM_APO.xlsx"
SHEET = "Dataset_Audit"

# 1. Membaca dataset dari Excel
df = pd.read_excel(FILE_EXCEL, sheet_name=SHEET)

# 2. Memastikan kolom yang dibutuhkan tersedia
required = [
    "Kode Proses", "Nama Proses", "Domain",
    "PA1.1", "PA2.1", "PA2.2", "PA3.1", "PA3.2", "Target"
]
missing = [col for col in required if col not in df.columns]
if missing:
    raise ValueError(f"Kolom wajib tidak ditemukan: {missing}")

# 3. Validasi rating atribut
rating_rank = {"N": 0, "P": 1, "L": 2, "F": 3}
rating_cols = ["PA1.1", "PA2.1", "PA2.2", "PA3.1", "PA3.2"]

for col in rating_cols:
    df[col] = df[col].astype(str).str.strip().str.upper()
    invalid = sorted(set(df[col]) - set(rating_rank))
    if invalid:
        raise ValueError(f"Rating tidak valid pada kolom {col}: {invalid}")

# 4. Menghitung capability level dengan aturan demonstrasi
def hitung_capability_level(row):
    level = 0

    # Level 1: PA1.1 minimal L
    if rating_rank[row["PA1.1"]] >= rating_rank["L"]:
        level = 1

    # Level 2: PA1.1 = F; PA2.1 dan PA2.2 minimal L
    if (
        row["PA1.1"] == "F"
        and rating_rank[row["PA2.1"]] >= rating_rank["L"]
        and rating_rank[row["PA2.2"]] >= rating_rank["L"]
    ):
        level = 2

    # Level 3: atribut level sebelumnya F; PA3.1 dan PA3.2 minimal L
    if (
        row["PA1.1"] == "F"
        and row["PA2.1"] == "F"
        and row["PA2.2"] == "F"
        and rating_rank[row["PA3.1"]] >= rating_rank["L"]
        and rating_rank[row["PA3.2"]] >= rating_rank["L"]
    ):
        level = 3

    return level

df["As-Is"] = df.apply(hitung_capability_level, axis=1)
df["Gap"] = df["Target"] - df["As-Is"]

if (df["Gap"] < 0).any():
    raise ValueError("Target lebih rendah daripada As-Is. Periksa dataset.")

df["Prioritas"] = df["Gap"].apply(
    lambda gap: "Tinggi" if gap >= 2
    else ("Sedang" if gap == 1 else "Pemeliharaan")
)

# 5. Menampilkan hasil analisis
print("\\n=== HASIL ANALISIS AUDIT COBIT 2019 ===")
print(df[["Kode Proses", "Domain", "As-Is", "Target", "Gap", "Prioritas"]]
      .to_string(index=False))

print("\\n=== RINGKASAN ===")
print("Jumlah proses:", len(df))
print("Rata-rata As-Is:", round(df["As-Is"].mean(), 2))
print("Rata-rata Gap:", round(df["Gap"].mean(), 2))
print("Gap tertinggi:", int(df["Gap"].max()))

print("\\n=== PRIORITAS PERBAIKAN ===")
prioritas = df.sort_values(
    by=["Gap", "Kode Proses"], ascending=[False, True]
)
print(prioritas[["Kode Proses", "Gap", "Prioritas"]].to_string(index=False))

# 6. Ringkasan per domain EDM dan APO
ringkasan_domain = df.groupby("Domain").agg(
    Jumlah_Proses=("Kode Proses", "count"),
    Rata_As_Is=("As-Is", "mean"),
    Rata_Gap=("Gap", "mean")
).round(2)

print("\\n=== RINGKASAN PER DOMAIN ===")
print(ringkasan_domain)

# 7. Menyimpan tabel hasil ke Excel baru
with pd.ExcelWriter("Hasil_Analisis_Audit_COBIT2019.xlsx",
                    engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Hasil_Per_Proses", index=False)
    ringkasan_domain.to_excel(writer, sheet_name="Ringkasan_Domain")

# 8. Grafik perbandingan As-Is dan Target
ax = df.plot(
    x="Kode Proses",
    y=["As-Is", "Target"],
    kind="bar",
    figsize=(10, 5)
)
ax.set_title("Analisis Gap Capability Level COBIT 2019")
ax.set_xlabel("Proses pada Domain EDM dan APO")
ax.set_ylabel("Capability Level")
ax.set_ylim(0, 5)
ax.legend(["As-Is", "Target / To-Be"])
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("Grafik_Analisis_Gap_COBIT2019.png", dpi=300)
plt.show()

# 9. Grafik gap setiap proses
ax = df.plot(
    x="Kode Proses",
    y="Gap",
    kind="bar",
    legend=False,
    figsize=(10, 4)
)
ax.set_title("Besarnya Gap per Proses COBIT 2019")
ax.set_xlabel("Proses")
ax.set_ylabel("Gap (tingkat)")
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("Grafik_Gap_Per_Proses.png", dpi=300)
plt.show()

print("\\nAnalisis selesai. Periksa file Excel hasil dan grafik pada folder kerja.")
