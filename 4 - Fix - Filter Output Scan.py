import os
import re
import pandas as pd
from pathlib import Path

# ==================================================
# 1️⃣ KONFIGURASI
# ==================================================
FOLDER_SCAN = "3_scan_output"
EXCEL_INPUT = "DLPD_ACMT_32AMS_202601.xlsx"
FOLDER_FOTO = "2_images"  # Folder foto asli

# ==================================================
# 2️⃣ FUNGSI CARI FOTO
# ==================================================
def cari_foto(idpel, folder_foto):
    """Cari file foto berdasarkan IDPEL"""
    folder = Path(folder_foto)
    idpel = str(idpel).strip()
    
    # Cari dengan berbagai pola
    patterns = [
        f"{idpel}_*.jpg",
        f"{idpel}_*.jpeg",
        f"{idpel}_*.png",
        f"{idpel}.jpg",
        f"{idpel}.jpeg",
        f"{idpel}.png",
    ]
    
    for pattern in patterns:
        files = list(folder.glob(pattern))
        if files:
            return str(files[0])  # Return path file pertama yang ditemukan
    
    return ""

# ==================================================
# 3️⃣ AMBIL IDPEL DARI NAMA FILE SCAN
# ==================================================
def ambil_idpel_dari_filename(folder):
    idpel_set = set()

    for file in os.listdir(folder):
        nama_file, _ = os.path.splitext(file)
        match = re.search(r"\d+", nama_file)  # ambil angka IDPEL

        if match:
            idpel_set.add(match.group())

    return idpel_set

# ==================================================
# 4️⃣ DETEKSI KOLOM IDPEL DI EXCEL (CASE-INSENSITIVE)
# ==================================================
def cari_kolom_idpel(df):
    for col in df.columns:
        if "idpel" in col.lower():
            return col
    raise Exception("❌ Kolom IDPEL tidak ditemukan di file Excel!")

# ==================================================
# 5️⃣ MAIN PROCESS
# ==================================================
def main():
    print("🔍 Mengambil IDPEL dari folder scan...")
    idpel_folder = ambil_idpel_dari_filename(FOLDER_SCAN)

    if not idpel_folder:
        raise Exception("❌ IDPEL dari folder scan tidak ditemukan!")

    print(f"✅ Total IDPEL ditemukan: {len(idpel_folder)}")

    print("📄 Membaca file Excel input...")
    df = pd.read_excel(EXCEL_INPUT, dtype=str)

    kolom_idpel = cari_kolom_idpel(df)
    print(f"✅ Kolom IDPEL terdeteksi: {kolom_idpel}")

    print("🧹 Melakukan filter data Excel...")
    df_filtered = df[df[kolom_idpel].isin(idpel_folder)]
    print(f"✅ Total data setelah filter: {len(df_filtered)} baris")

    # ==================================================
    # 6️⃣ TAMBAHKAN LINK FOTO
    # ==================================================
    print("📸 Menambahkan link foto ke data...")
    
    # Tambahkan kolom untuk foto
    df_filtered["FOTO_SCAN"] = ""  # Foto dari folder scan
    df_filtered["FOTO_ASLI"] = ""  # Foto asli dari folder foto
    
    # Isi kolom foto scan
    for idx, row in df_filtered.iterrows():
        idpel = str(row[kolom_idpel]).strip()
        
        # Cari file di folder scan
        scan_file = Path(FOLDER_SCAN) / f"{idpel}.jpg"
        if scan_file.exists():
            df_filtered.at[idx, "FOTO_SCAN"] = str(scan_file)
        
        # Cari file asli di folder foto
        if os.path.exists(FOLDER_FOTO):
            foto_asli = cari_foto(idpel, FOLDER_FOTO)
            if foto_asli:
                df_filtered.at[idx, "FOTO_ASLI"] = foto_asli
    
    # ==================================================
    # 7️⃣ NAMA FILE OUTPUT OTOMATIS
    # ==================================================
    nama_awal = os.path.splitext(os.path.basename(EXCEL_INPUT))[0]
    excel_output = f"Output_Scan_{nama_awal}_WITH_FOTO.xlsx"

    print("💾 Menyimpan hasil ke Excel...")
    df_filtered.to_excel(excel_output, index=False)

    print(f"🎉 SELESAI → {excel_output}")
    
    # Buat summary
    total_foto_scan = df_filtered["FOTO_SCAN"].apply(lambda x: 1 if x else 0).sum()
    total_foto_asli = df_filtered["FOTO_ASLI"].apply(lambda x: 1 if x else 0).sum()
    
    print(f"\n📊 SUMMARY FOTO:")
    print(f"   Total data: {len(df_filtered)} baris")
    print(f"   Foto scan ditemukan: {total_foto_scan}")
    print(f"   Foto asli ditemukan: {total_foto_asli}")

# ==================================================
# 8️⃣ EXECUTE
# ==================================================
if __name__ == "__main__":
    main()