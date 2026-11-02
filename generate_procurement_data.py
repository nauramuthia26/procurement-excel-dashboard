"""
Generator data dummy pengadaan (Procurement Spend Analysis).
Semua data SINTETIS, bukan data perusahaan asli.

Output: procurement_raw.xlsx dengan 3 sheet:
  - Raw_Data     : transaksi PO (sengaja kotor)
  - Vendor_Ref   : tabel referensi vendor
  - Category_Ref : tabel referensi kategori + SLA hari kerja

Jalankan:
    pip install pandas openpyxl
    python generate_procurement_data.py
"""
import random
from datetime import date, timedelta

import pandas as pd
from openpyxl.utils import get_column_letter

SEED = 42
N_ROWS = 1500
START = date(2025, 1, 1)
END = date(2025, 12, 31)
OUTPUT = "procurement_raw.xlsx"

random.seed(SEED)

# ---------------------------------------------------------------- referensi
# kategori: (SLA hari, harga min, harga max, qty min, qty max, daftar item)
CATEGORIES = {
    "IT Hardware":     (14, 3_000_000, 25_000_000, 1, 10,
                        ["Laptop", "Monitor", "Keyboard", "Router", "SSD 1TB"]),
    "Software License": (7, 1_500_000, 40_000_000, 1, 20,
                        ["Lisensi Office", "Antivirus", "Cloud Subscription", "BI Tool Seat"]),
    "Office Supplies": (5, 5_000, 250_000, 10, 200,
                        ["Kertas A4", "Pulpen", "Map Folder", "Tinta Printer", "Stapler"]),
    "Furniture":       (21, 500_000, 8_000_000, 1, 15,
                        ["Kursi Kerja", "Meja Kerja", "Lemari Arsip", "Partisi"]),
    "Maintenance":     (10, 300_000, 15_000_000, 1, 5,
                        ["Servis AC", "Servis Genset", "Cleaning Service", "Servis Lift"]),
    "Marketing":       (12, 500_000, 20_000_000, 1, 50,
                        ["Cetak Brosur", "Banner", "Merchandise", "Booth Event"]),
}

# vendor: (nama, tier, kota, rata-rata keterlambatan hari)
# keterlambatan tinggi = vendor kurang andal -> bikin insight di dashboard
VENDORS = [
    ("PT Maju Jaya Abadi", "Gold", "Jakarta", 0),
    ("PT Sinar Teknologi Nusantara", "Gold", "Jakarta", 1),
    ("CV Berkah Sentosa", "Silver", "Bandung", 4),
    ("PT Cakra Digital Indonesia", "Gold", "Tangerang", 0),
    ("PT Mitra Perkasa Utama", "Silver", "Surabaya", 3),
    ("CV Anugerah Prima", "Bronze", "Bekasi", 7),
    ("PT Global Solusi Mandiri", "Gold", "Jakarta", 1),
    ("PT Bumi Raya Furnitur", "Silver", "Semarang", 5),
    ("CV Karya Mandiri", "Bronze", "Depok", 8),
    ("PT Prima Office Supplies", "Silver", "Jakarta", 2),
    ("PT Nusa Data Persada", "Gold", "Tangerang", 1),
    ("CV Sumber Rejeki", "Bronze", "Bogor", 9),
    ("PT Harmoni Service Prima", "Silver", "Jakarta", 3),
    ("PT Kreasi Media Utama", "Silver", "Bandung", 2),
    ("PT Lintas Komputindo", "Gold", "Jakarta", 0),
    ("CV Tiga Saudara", "Bronze", "Bekasi", 6),
    ("PT Cahaya Abadi Sejahtera", "Silver", "Surabaya", 4),
    ("PT Dwi Karya Teknik", "Silver", "Tangerang", 3),
    ("PT Era Baru Printing", "Bronze", "Depok", 5),
    ("PT Fortuna Jaya Makmur", "Gold", "Jakarta", 1),
]

DEPARTMENTS = ["IT", "Finance", "HR", "Operations", "Marketing", "Legal", "Procurement"]
STATUSES = ["Delivered"] * 80 + ["In Transit"] * 8 + ["Pending"] * 7 + ["Cancelled"] * 5


def random_date(start: date, end: date) -> date:
    return start + timedelta(days=random.randint(0, (end - start).days))


def add_business_days(d: date, n: int) -> date:
    """Tambah n hari kerja (Senin-Jumat), tanpa libur nasional."""
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


# ---------------------------------------------------------------- data bersih
vendor_names = [v[0] for v in VENDORS]
vendor_delay = {v[0]: v[3] for v in VENDORS}

rows = []
for i in range(1, N_ROWS + 1):
    category = random.choice(list(CATEGORIES))
    sla, pmin, pmax, qmin, qmax, items = CATEGORIES[category]
    vendor = random.choice(vendor_names)
    po_date = random_date(START, END)
    status = random.choice(STATUSES)

    delivery = None
    if status == "Delivered":
        lead = max(1, sla + random.randint(-3, 2) + vendor_delay[vendor]
                   + random.choice([0, 0, 0, 1, 2, 4]))
        delivery = add_business_days(po_date, lead)

    rows.append({
        "PO_Number": f"PO-2025-{i:05d}",
        "PO_Date": po_date,
        "Vendor": vendor,
        "Category": category,
        "Item": random.choice(items),
        "Department": random.choice(DEPARTMENTS),
        "Qty": random.randint(qmin, qmax),
        "Unit_Price": round(random.randint(pmin, pmax), -3),
        "Status": status,
        "Delivery_Date": delivery,
    })

df = pd.DataFrame(rows)
df["PO_Date"] = df["PO_Date"].astype(object)
df["Delivery_Date"] = df["Delivery_Date"].astype(object)

# ---------------------------------------------------------------- data kotor
issues = {}

# 1. Nama vendor tidak konsisten (spasi, huruf besar/kecil, "PT." vs "PT")
idx = df.sample(frac=0.10, random_state=1).index
def mess_vendor(name: str) -> str:
    return random.choice([
        name.upper(), name.lower(), name + " ", " " + name,
        name.replace("PT ", "PT. ").replace("CV ", "CV. "),
    ])
df.loc[idx, "Vendor"] = df.loc[idx, "Vendor"].apply(mess_vendor)
issues["Vendor nama tidak konsisten"] = len(idx)

# 2. Departemen kosong
idx = df.sample(n=15, random_state=2).index
df.loc[idx, "Department"] = None
issues["Department kosong"] = len(idx)

# 3. Qty nol atau negatif
idx = df.sample(n=5, random_state=3).index
df.loc[idx, "Qty"] = [random.choice([0, -1, -5]) for _ in idx]
issues["Qty <= 0"] = len(idx)

# 4. Harga kosong
idx = df.sample(n=6, random_state=4).index
df.loc[idx, "Unit_Price"] = None
issues["Unit_Price kosong"] = len(idx)

# 5. PO_Date tersimpan sebagai teks dd/mm/yyyy (bukan tanggal sungguhan)
idx = df.sample(n=40, random_state=5).index
df.loc[idx, "PO_Date"] = df.loc[idx, "PO_Date"].apply(lambda d: d.strftime("%d/%m/%Y"))
issues["PO_Date format teks"] = len(idx)

# 6. Tanggal terima lebih awal dari tanggal PO (tidak logis)
delivered = df[df["Status"] == "Delivered"]
idx = delivered[delivered["PO_Date"].apply(lambda d: isinstance(d, date))].sample(
    n=8, random_state=6).index
df.loc[idx, "Delivery_Date"] = df.loc[idx, "PO_Date"].apply(
    lambda d: d - timedelta(days=random.randint(1, 3)))
issues["Delivery_Date sebelum PO_Date"] = len(idx)

# 7. Baris duplikat persis
dups = df.sample(n=25, random_state=7)
df = pd.concat([df, dups], ignore_index=True)
issues["Baris duplikat"] = len(dups)

# acak urutan supaya duplikat tidak berdekatan
df = df.sample(frac=1, random_state=8).reset_index(drop=True)

# ---------------------------------------------------------------- tabel referensi
vendor_ref = pd.DataFrame(
    [(f"V{i:03d}", v[0], v[1], v[2]) for i, v in enumerate(VENDORS, 1)],
    columns=["Vendor_ID", "Vendor_Name", "Vendor_Tier", "City"],
)
category_ref = pd.DataFrame(
    [(c, CATEGORIES[c][0]) for c in CATEGORIES],
    columns=["Category", "SLA_Days"],
)

# ---------------------------------------------------------------- tulis Excel
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Raw_Data", index=False)
    vendor_ref.to_excel(writer, sheet_name="Vendor_Ref", index=False)
    category_ref.to_excel(writer, sheet_name="Category_Ref", index=False)

    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[letter].width = min(width + 3, 40)

    # format tanggal hanya untuk sel yang benar-benar bertipe tanggal
    ws = writer.book["Raw_Data"]
    for col_letter in ("B", "J"):
        for cell in ws[col_letter][1:]:
            if isinstance(cell.value, (date,)) or hasattr(cell.value, "year"):
                cell.number_format = "yyyy-mm-dd"
    for cell in ws["H"][1:]:
        cell.number_format = "#,##0"

# ---------------------------------------------------------------- ringkasan
print(f"Selesai: {OUTPUT} ({len(df)} baris di Raw_Data)")
print("Masalah data yang sengaja disisipkan:")
for k, v in issues.items():
    print(f"  - {k}: {v}")
