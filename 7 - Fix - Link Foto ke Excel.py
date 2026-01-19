#!/usr/bin/env python3
# ==========================================================
# LINK FOTO KE FILE EXCEL YANG SUDAH ADA
# ==========================================================

import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import threading

def cari_foto(idpel, blth, foto_folder):
    """Cari file foto berdasarkan pola"""
    foto_folder = Path(foto_folder)
    idpel = str(idpel).strip()
    
    # Pola pencarian berdasarkan BLTH
    patterns = [
        f"{idpel}_{blth}_1.*",           # IDPEL_202601_1.jpg
        f"{idpel}_{blth}_2.*",           # IDPEL_202601_2.jpg
        f"{idpel}_{blth}_photoke-1.*",   # IDPEL_202601_photoke-1.jpg
        f"{idpel}_{blth}_photoke-2.*",   # IDPEL_202601_photoke-2.jpg
        f"{idpel}_{blth}.*",             # IDPEL_202601.jpg
        f"{idpel}.*",                    # IDPEL.jpg
        f"*{idpel}*_1.*",                # *_IDPEL*_1.jpg
        f"*{idpel}*_2.*",                # *_IDPEL*_2.jpg
    ]
    
    # Cek semua pola
    for pattern in patterns:
        for ext in ['.jpg', '.jpeg', '.png', '.bmp', '.JPG', '.JPEG', '.PNG']:
            files = list(foto_folder.glob(f"*{idpel}*{ext}"))
            if files:
                # Prioritaskan yang mengandung BLTH
                for file in files:
                    if blth in str(file):
                        return str(file)
                # Jika tidak ada yang mengandung BLTH, ambil yang pertama
                return str(files[0])
    
    return ""

def link_foto_ke_excel():
    excel_file = excel_var.get()
    foto_folder = foto_var.get()
    blth = blth_var.get()
    
    if not excel_file or not foto_folder or not blth:
        messagebox.showerror("Error", "Harap isi semua field!")
        return
    
    if not os.path.exists(excel_file):
        messagebox.showerror("Error", f"File Excel tidak ditemukan:\n{excel_file}")
        return
    
    if not os.path.exists(foto_folder):
        messagebox.showerror("Error", f"Folder foto tidak ditemukan:\n{foto_folder}")
        return
    
    try:
        log_text.delete("1.0", tk.END)
        log("📁 Membaca file Excel...")
        
        # Baca Excel
        df = pd.read_excel(excel_file, dtype=str)
        
        # Cari kolom IDPEL
        idpel_col = None
        for col in df.columns:
            if "idpel" in str(col).lower():
                idpel_col = col
                break
        
        if not idpel_col:
            messagebox.showerror("Error", "Kolom IDPEL tidak ditemukan di file Excel!")
            return
        
        log(f"✅ Kolom IDPEL ditemukan: {idpel_col}")
        log(f"🔍 Mencari foto dengan BLTH: {blth}...")
        
        # Tambahkan kolom foto
        df["PATH_FOTO_1"] = ""
        df["PATH_FOTO_2"] = ""
        df["FILE_FOTO_1"] = ""
        df["FILE_FOTO_2"] = ""
        
        total_rows = len(df)
        ditemukan = 0
        
        for idx, row in df.iterrows():
            idpel = str(row[idpel_col]).strip()
            
            if idpel and idpel.lower() != "nan":
                # Cari foto 1
                foto1 = cari_foto(idpel, blth, foto_folder)
                if foto1:
                    df.at[idx, "PATH_FOTO_1"] = foto1
                    df.at[idx, "FILE_FOTO_1"] = os.path.basename(foto1)
                
                # Cari foto 2
                foto2 = cari_foto(f"{idpel}_{blth}_2", blth, foto_folder)
                if not foto2:
                    # Coba cari dengan pola lain
                    foto2 = cari_foto(idpel, blth, foto_folder)
                    # Jika foto2 sama dengan foto1, set kosong
                    if foto2 == foto1:
                        foto2 = ""
                
                if foto2:
                    df.at[idx, "PATH_FOTO_2"] = foto2
                    df.at[idx, "FILE_FOTO_2"] = os.path.basename(foto2)
                
                if foto1 or foto2:
                    ditemukan += 1
            
            # Update progress setiap 100 baris
            if idx % 100 == 0:
                log(f"  Diproses: {idx+1}/{total_rows} baris")
                progress_var.set(int((idx+1) / total_rows * 100))
        
        # Simpan file baru
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        output_file = f"{base_name}_WITH_FOTO.xlsx"
        
        log("💾 Menyimpan file Excel dengan link foto...")
        df.to_excel(output_file, index=False)
        
        progress_var.set(100)
        
        log("\n" + "="*60)
        log(f"✅ SELESAI!")
        log(f"File disimpan sebagai: {output_file}")
        log(f"Total baris: {total_rows:,}")
        log(f"IDPEL dengan foto ditemukan: {ditemukan:,}")
        
        # Buat summary
        summary_file = f"Summary_Foto_{blth}.txt"
        with open(summary_file, "w", encoding="utf-8") as f:
            f.write(f"SUMMARY LINK FOTO - {blth}\n")
            f.write("="*50 + "\n")
            f.write(f"File Excel input: {excel_file}\n")
            f.write(f"Folder foto: {foto_folder}\n")
            f.write(f"BLTH: {blth}\n")
            f.write(f"Total baris data: {total_rows}\n")
            f.write(f"IDPEL dengan foto ditemukan: {ditemukan}\n")
            f.write(f"File output: {output_file}\n")
            f.write("\n" + "="*50 + "\n")
            f.write("Pola nama file yang dicari:\n")
            f.write(f"1. {idpel}_{blth}_1.jpg\n")
            f.write(f"2. {idpel}_{blth}_2.jpg\n")
            f.write(f"3. {idpel}_{blth}_photoke-1.jpg\n")
            f.write(f"4. {idpel}_{blth}_photoke-2.jpg\n")
            f.write(f"5. {idpel}_*.jpg (pola lainnya)\n")
        
        log(f"📝 Summary disimpan: {summary_file}")
        
        messagebox.showinfo("Selesai", 
            f"File Excel dengan link foto berhasil dibuat!\n\n"
            f"Output: {output_file}\n"
            f"Foto ditemukan: {ditemukan:,} dari {total_rows:,} baris")
    
    except Exception as e:
        log(f"❌ ERROR: {str(e)}")
        messagebox.showerror("Error", f"Terjadi kesalahan:\n{str(e)}")

def browse_excel():
    file = filedialog.askopenfilename(
        title="Pilih File Excel",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if file:
        excel_var.set(file)

def browse_folder():
    folder = filedialog.askdirectory(title="Pilih Folder Foto")
    if folder:
        foto_var.set(folder)

def log(msg):
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)
    root.update_idletasks()

def test_cari_foto():
    """Test pencarian foto dengan sample"""
    foto_folder = foto_var.get()
    blth = blth_var.get()
    
    if not foto_folder or not blth:
        messagebox.showerror("Error", "Isi folder foto dan BLTH terlebih dahulu!")
        return
    
    if not os.path.exists(foto_folder):
        messagebox.showerror("Error", f"Folder foto tidak ditemukan!")
        return
    
    log_text.delete("1.0", tk.END)
    log("🔍 TEST PENCARIAN FOTO")
    log("="*60)
    
    # Ambil beberapa file dari folder untuk contoh
    files = list(Path(foto_folder).glob("*.jpg"))[:5]
    
    if not files:
        log("Tidak ada file .jpg di folder tersebut")
        return
    
    log(f"Sample file di folder:\n")
    for file in files:
        log(f"  - {file.name}")
    
    log("\n" + "="*60)
    log("Mencoba ekstrak IDPEL dari nama file...")
    
    for file in files:
        filename = file.stem  # tanpa ekstensi
        # Coba ekstrak IDPEL (12-14 digit di awal)
        import re
        match = re.search(r'(\d{12,14})', filename)
        if match:
            idpel = match.group(1)
            log(f"\nFile: {file.name}")
            log(f"  IDPEL terdeteksi: {idpel}")
            
            # Cari pasangan foto
            foto1 = cari_foto(f"{idpel}_{blth}_1", blth, foto_folder)
            foto2 = cari_foto(f"{idpel}_{blth}_2", blth, foto_folder)
            
            if foto1:
                log(f"  Foto 1: {os.path.basename(foto1)}")
            else:
                log(f"  Foto 1: Tidak ditemukan")
            
            if foto2:
                log(f"  Foto 2: {os.path.basename(foto2)}")
            else:
                log(f"  Foto 2: Tidak ditemukan")
    
    # Test dengan sample IDPEL
    log("\n" + "="*60)
    log("Test dengan sample IDPEL:")
    sample_idpels = ["321500871587", "321500707583", "123456789012"]
    
    for idpel in sample_idpels:
        log(f"\nIDPEL: {idpel}")
        foto1 = cari_foto(f"{idpel}_{blth}_1", blth, foto_folder)
        foto2 = cari_foto(f"{idpel}_{blth}_2", blth, foto_folder)
        
        if foto1:
            log(f"  Foto 1 ditemukan: {os.path.basename(foto1)}")
        else:
            log(f"  Foto 1: TIDAK DITEMUKAN")
        
        if foto2:
            log(f"  Foto 2 ditemukan: {os.path.basename(foto2)}")
        else:
            log(f"  Foto 2: TIDAK DITEMUKAN")
    
    log("\nTest selesai!")

# ===================== GUI =====================
root = tk.Tk()
root.title("Link Foto ke Excel")
root.geometry("800x600")

# Frame input
input_frame = tk.LabelFrame(root, text="Pengaturan", padx=10, pady=10)
input_frame.pack(fill="x", padx=10, pady=10)

# File Excel
tk.Label(input_frame, text="File Excel:").grid(row=0, column=0, sticky="w", pady=5)
excel_var = tk.StringVar()
excel_entry = ttk.Entry(input_frame, textvariable=excel_var, width=60)
excel_entry.grid(row=0, column=1, padx=5, pady=5)
ttk.Button(input_frame, text="Browse", command=browse_excel).grid(row=0, column=2, padx=5)

# Folder Foto
tk.Label(input_frame, text="Folder Foto:").grid(row=1, column=0, sticky="w", pady=5)
foto_var = tk.StringVar()
foto_entry = ttk.Entry(input_frame, textvariable=foto_var, width=60)
foto_entry.grid(row=1, column=1, padx=5, pady=5)
ttk.Button(input_frame, text="Browse", command=browse_folder).grid(row=1, column=2, padx=5)

# BLTH
tk.Label(input_frame, text="BLTH (YYYYMM):").grid(row=2, column=0, sticky="w", pady=5)
blth_var = tk.StringVar(value="202601")
blth_entry = ttk.Entry(input_frame, textvariable=blth_var, width=15)
blth_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)

# Progress bar
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(fill="x", padx=10, pady=5)

# Tombol Proses
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

ttk.Button(
    button_frame,
    text="▶ LINK FOTO KE EXCEL",
    command=lambda: threading.Thread(target=link_foto_ke_excel, daemon=True).start()
).pack(side=tk.LEFT, padx=5)

ttk.Button(
    button_frame,
    text="🔍 TEST PENCARIAN FOTO",
    command=lambda: threading.Thread(target=test_cari_foto, daemon=True).start()
).pack(side=tk.LEFT, padx=5)

# Log area
log_label = tk.Label(root, text="Log Proses:")
log_label.pack(anchor="w", padx=10, pady=(10, 0))

log_text = tk.Text(root, height=15)
log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

root.mainloop()