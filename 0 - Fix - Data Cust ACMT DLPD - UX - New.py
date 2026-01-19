#!/usr/bin/env python3
# ==========================================================
# DOWNLOAD PLN → XLSX → GABUNG SEMUA SHEET → OUTPUT 1 XLSX
# DENGAN PILIHAN SERVER
# ==========================================================

import os
import time
import tempfile
import threading
import urllib3
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from urllib3.exceptions import InsecureRequestWarning
from pathlib import Path

# ===================== KONFIG =====================
MAX_RETRY = 3
RETRY_DELAY = 5

UNITAP_DICT = {
    "32AMU": ["32010", "32020", "32030", "32040"],
    "32AMS": ["32111", "32121", "32131", "32141", "32151", "32161"],
    "32CWP": ["32210", "32240", "32250", "32260", "32270", "32280"],
    "32CKD": ["32320", "32330", "32340", "32350", "32360", "32370", "32380"],
    "32CPG": ["32410", "32420", "32430", "32440", "32450"],
    "32CPR": ["32510", "32520", "32530", "32540", "32550", "32560", "32570"],
    "32CPL": ["32610", "32620", "32630", "32640", "32650", "32660", "32680"],
    "32CBK": ["32710", "32720", "32730", "32740", "32750", "32760", "32770"],
    "32CBB": ["32810", "32820", "32830", "32840", "32850"],
    "32CMJ": ["32910", "32920", "32930", "32940", "32950", "32960"],
}

urllib3.disable_warnings(InsecureRequestWarning)

URL_TEMPLATE = (
    "{base}/birt-acmt/run?"
    "__report=rpt_icmo_DataDetail.rptdesign"
    "&up={up}&blth={blth}"
    "&rbm=TOTAL&jns=FG_DLPD_JAMNYALA"
    "&tglbaca=&jnf=0&jnt=99999999999"
    "&kdbaca=&kdklpk=&dlpd=4&ptgs="
    "&__format=xlsx"
)

# ===================== GUI =====================
root = tk.Tk()
root.title("Downloader PLN - XLSX Merge Sheets")
root.geometry("760x680")  # Diperbesar untuk fitur tambahan

# -------- SERVER --------
tk.Label(root, text="Server", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
server_var = tk.StringVar(value="INTRANET")

frame_server = tk.Frame(root)
frame_server.pack(anchor="w", padx=20)

tk.Radiobutton(
    frame_server,
    text="INTRANET (ap2t.pln.co.id)",
    variable=server_var,
    value="INTRANET"
).pack(anchor="w")

tk.Radiobutton(
    frame_server,
    text="INTERNET (portalapp.iconpln.co.id)",
    variable=server_var,
    value="INTERNET"
).pack(anchor="w")

# -------- INPUT --------
frame_input = tk.Frame(root)
frame_input.pack(fill="x", padx=10, pady=10)

tk.Label(frame_input, text="BLTH (YYYYMM)").grid(row=0, column=0, sticky="w")
blth_entry = ttk.Entry(frame_input, width=12)
blth_entry.grid(row=0, column=1, padx=5)
blth_entry.insert(0, "202512")

tk.Label(frame_input, text="UNITAP").grid(row=1, column=0, sticky="w")
unitap_var = tk.StringVar()
unitap_combo = ttk.Combobox(
    frame_input,
    textvariable=unitap_var,
    values=list(UNITAP_DICT.keys()),
    state="readonly",
    width=15
)
unitap_combo.grid(row=1, column=1, padx=5)
unitap_combo.current(0)

# -------- FOTO LINK SETTINGS --------
frame_foto = tk.LabelFrame(root, text="Pengaturan Link Foto", padx=10, pady=5)
frame_foto.pack(fill="x", padx=10, pady=10)

# Folder Foto
tk.Label(frame_foto, text="Folder Foto:").grid(row=0, column=0, sticky="w")
foto_folder_var = tk.StringVar()
foto_entry = ttk.Entry(frame_foto, textvariable=foto_folder_var, width=50)
foto_entry.grid(row=0, column=1, padx=5, pady=2)
foto_entry.insert(0, "2_images")

def browse_foto_folder():
    folder = filedialog.askdirectory(title="Pilih Folder Foto")
    if folder:
        foto_folder_var.set(folder)

tk.Button(frame_foto, text="Browse", command=browse_foto_folder).grid(row=0, column=2, padx=5)

# Pola Nama File
tk.Label(frame_foto, text="Pola Nama File:").grid(row=1, column=0, sticky="w")
pola_var = tk.StringVar(value="IDPEL_tahunbulan_photoke-")
pola_combo = ttk.Combobox(
    frame_foto,
    textvariable=pola_var,
    values=["IDPEL_tahunbulan_photoke-", "IDPEL_tahunbulan_", "IDPEL_"],
    state="readonly",
    width=30
)
pola_combo.grid(row=1, column=1, padx=5, pady=2)

# -------- LOG --------
tk.Label(root, text="Log Proses", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10)
log_box = tk.Text(root, height=20)
log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

def log(msg):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)
    root.update_idletasks()

# ===================== CORE =====================
def cari_foto(idpel, blth, foto_folder):
    """Cari file foto berdasarkan pola"""
    foto_folder = Path(foto_folder)
    
    # Pola 1: IDPEL_tahunbulan_photoke-
    pola1 = f"{idpel}_{blth}_photoke-"
    
    # Pola 2: IDPEL_tahunbulan_
    pola2 = f"{idpel}_{blth}_"
    
    # Pola 3: IDPEL_
    pola3 = f"{idpel}_"
    
    # Cari file dengan berbagai ekstensi
    exts = ['.jpg', '.jpeg', '.png', '.bmp']
    
    for ext in exts:
        # Coba pola 1 (bisa ada angka di akhir: 1, 2, dll)
        for i in range(1, 10):  # cari hingga _9
            for file in foto_folder.glob(f"{pola1}{i}{ext}"):
                return str(file)
            for file in foto_folder.glob(f"{pola1}{i}.*"):
                return str(file)
        
        # Coba pola 2
        for file in foto_folder.glob(f"{pola2}*{ext}"):
            return str(file)
        
        # Coba pola 3
        for file in foto_folder.glob(f"{pola3}*{ext}"):
            return str(file)
        
        # Coba hanya IDPEL.ext
        file_path = foto_folder / f"{idpel}{ext}"
        if file_path.exists():
            return str(file_path)
    
    return ""

def proses_download():
    log_box.delete("1.0", tk.END)

    blth = blth_entry.get().strip()
    unitap = unitap_var.get()
    ups = UNITAP_DICT[unitap]
    
    foto_folder = foto_folder_var.get().strip()
    pola_nama = pola_var.get()

    if not os.path.exists(foto_folder):
        log(f"⚠ Folder foto '{foto_folder}' tidak ditemukan!")
        log("Foto tidak akan dilink. Lanjut proses download...")

    base_url = (
        "https://ap2t.pln.co.id"
        if server_var.get() == "INTRANET"
        else "https://portalapp.iconpln.co.id"
    )

    output_file = f"DLPD_ACMT_{unitap}_{blth}_WITH_FOTO.xlsx"
    log(f"SERVER : {server_var.get()}")
    log(f"UNITAP : {unitap}")
    log(f"FOLDER FOTO: {foto_folder}")
    log("=" * 60)

    http = urllib3.PoolManager(cert_reqs="CERT_NONE")
    dfs_final = []

    for up in ups:
        log(f"\nUP {up}")
        for attempt in range(1, MAX_RETRY + 1):
            log(f"  Attempt {attempt}")
            tmp = None
            try:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                tmp.close()

                url = URL_TEMPLATE.format(base=base_url, up=up, blth=blth)
                resp = http.request("GET", url, preload_content=False)

                with open(tmp.name, "wb") as f:
                    for c in resp.stream(1024 * 64):
                        if not c:
                            break
                        f.write(c)

                resp.release_conn()

                # 🔥 BACA SEMUA SHEET
                sheets = pd.read_excel(tmp.name, sheet_name=None)

                rows = 0
                for df in sheets.values():
                    df["UNITAP"] = unitap
                    df["UP"] = up
                    
                    # Tambahkan kolom link foto jika ada kolom IDPEL
                    idpel_col = None
                    for col in df.columns:
                        if "idpel" in str(col).lower():
                            idpel_col = col
                            break
                    
                    if idpel_col and os.path.exists(foto_folder):
                        log(f"  Menambahkan link foto untuk kolom {idpel_col}...")
                        # Tambahkan kolom untuk link foto
                        df["LINK_FOTO_1"] = ""
                        df["LINK_FOTO_2"] = ""
                        
                        for idx, row in df.iterrows():
                            idpel = str(row[idpel_col]).strip()
                            if idpel and idpel != "nan":
                                # Cari foto 1
                                foto1 = cari_foto(f"{idpel}_{blth}_1", blth, foto_folder)
                                if not foto1:
                                    foto1 = cari_foto(f"{idpel}_{blth[:4]}{blth[4:]}_1", blth, foto_folder)
                                if foto1:
                                    df.at[idx, "LINK_FOTO_1"] = foto1
                                
                                # Cari foto 2
                                foto2 = cari_foto(f"{idpel}_{blth}_2", blth, foto_folder)
                                if not foto2:
                                    foto2 = cari_foto(f"{idpel}_{blth[:4]}{blth[4:]}_2", blth, foto_folder)
                                if foto2:
                                    df.at[idx, "LINK_FOTO_2"] = foto2
                    
                    dfs_final.append(df)
                    rows += len(df)

                log(f"  ✔ {len(sheets)} sheet | {rows:,} baris")
                break

            except Exception as e:
                log(f"  ✖ {e}")
                time.sleep(RETRY_DELAY)

            finally:
                if tmp and os.path.exists(tmp.name):
                    os.remove(tmp.name)

    if dfs_final:
        final_df = pd.concat(dfs_final, ignore_index=True)
        
        # Simpan summary foto
        foto_ditemukan = final_df["LINK_FOTO_1"].apply(lambda x: 1 if x else 0).sum()
        total_rows = len(final_df)
        
        final_df.to_excel(output_file, index=False)
        
        log("\n" + "=" * 60)
        log(f"SELESAI ✅ TOTAL BARIS: {len(final_df):,}")
        log(f"FOTO DITEMUKAN: {foto_ditemukan:,} dari {total_rows:,} baris")
        
        # Buat file HTML untuk preview
        buat_html_preview(final_df, output_file, unitap, blth)
        
        messagebox.showinfo("Selesai", f"{output_file}\n\nFoto ditemukan: {foto_ditemukan:,} dari {total_rows:,} baris")

def buat_html_preview(df, excel_file, unitap, blth):
    """Buat file HTML untuk preview dengan foto"""
    html_file = f"Preview_Foto_{unitap}_{blth}.html"
    
    # Ambil kolom IDPEL
    idpel_col = None
    for col in df.columns:
        if "idpel" in str(col).lower():
            idpel_col = col
            break
    
    if not idpel_col:
        return
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Preview Foto - {unitap} - {blth}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            img {{ max-width: 200px; max-height: 150px; }}
            .foto-container {{ display: flex; gap: 10px; }}
            .foto-item {{ text-align: center; }}
        </style>
    </head>
    <body>
        <h1>Preview Foto - {unitap} - {blth}</h1>
        <p>Total Data: {len(df):,} baris</p>
        <p>File Excel: {excel_file}</p>
        <table>
            <tr>
                <th>No</th>
                <th>{idpel_col}</th>
                <th>Foto 1</th>
                <th>Foto 2</th>
                <th>Status</th>
            </tr>
    """
    
    for idx, row in df.head(100).iterrows():  # Batasi 100 baris untuk preview
        idpel = str(row[idpel_col])
        foto1 = row.get("LINK_FOTO_1", "")
        foto2 = row.get("LINK_FOTO_2", "")
        
        status = "❌"
        if foto1:
            status = "✅" if foto2 else "⚠️ (1 foto)"
        
        html_content += f"""
            <tr>
                <td>{idx + 1}</td>
                <td>{idpel}</td>
                <td>
        """
        
        if foto1 and os.path.exists(foto1):
            html_content += f'<div class="foto-item"><img src="{foto1}" alt="Foto 1"><br>Foto 1</div>'
        else:
            html_content += "Tidak ditemukan"
        
        html_content += "</td><td>"
        
        if foto2 and os.path.exists(foto2):
            html_content += f'<div class="foto-item"><img src="{foto2}" alt="Foto 2"><br>Foto 2</div>'
        else:
            html_content += "Tidak ditemukan"
        
        html_content += f"""
                </td>
                <td>{status}</td>
            </tr>
        """
    
    html_content += """
        </table>
        <p><em>Catatan: Hanya menampilkan 100 baris pertama untuk preview</em></p>
    </body>
    </html>
    """
    
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    log(f"File preview HTML dibuat: {html_file}")

# -------- BUTTON --------
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

ttk.Button(
    button_frame,
    text="▶ DOWNLOAD & LINK FOTO",
    command=lambda: threading.Thread(target=proses_download, daemon=True).start()
).pack(side=tk.LEFT, padx=5)

ttk.Button(
    button_frame,
    text="🔧 TEST LINK FOTO",
    command=lambda: threading.Thread(target=test_link_foto, daemon=True).start()
).pack(side=tk.LEFT, padx=5)

def test_link_foto():
    """Test fungsi link foto dengan beberapa sample"""
    foto_folder = foto_folder_var.get().strip()
    blth = blth_entry.get().strip()
    
    if not os.path.exists(foto_folder):
        messagebox.showerror("Error", f"Folder foto tidak ditemukan:\n{foto_folder}")
        return
    
    # Sample IDPEL untuk testing
    sample_idpels = [
        "321500871587",  # Contoh dengan foto
        "321500707583",  # Contoh lain
        "321500829434",  # Contoh lain
        "123456789012",  # Contoh tanpa foto
    ]
    
    log("\n" + "="*60)
    log("🔧 TEST LINK FOTO")
    log("="*60)
    
    for idpel in sample_idpels:
        log(f"\nMencari foto untuk IDPEL: {idpel}")
        
        # Foto 1
        foto1 = cari_foto(f"{idpel}_{blth}_1", blth, foto_folder)
        if foto1:
            log(f"  Foto 1 ditemukan: {os.path.basename(foto1)}")
        else:
            log(f"  Foto 1: TIDAK DITEMUKAN")
        
        # Foto 2
        foto2 = cari_foto(f"{idpel}_{blth}_2", blth, foto_folder)
        if foto2:
            log(f"  Foto 2 ditemukan: {os.path.basename(foto2)}")
        else:
            log(f"  Foto 2: TIDAK DITEMUKAN")
    
    log("\nTest selesai!")

root.mainloop()