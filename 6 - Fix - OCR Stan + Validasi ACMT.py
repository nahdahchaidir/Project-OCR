# pip install pytesseract pillow opencv-python pandas numpy rapidfuzz

import cv2
import pytesseract
import pandas as pd
import numpy as np
from pathlib import Path
from rapidfuzz import fuzz

# ===================== KONFIG =====================
FOLDER_IMAGE = "3_scan_output"      # hasil scan NEG
EXCEL_ACMT = "Output_Scan_DLPD.xlsx"
OUTPUT_EXCEL = "Hasil_OCR_Validasi.xlsx"

OCR_LANG = "eng"
MAX_DIFF = 50        # toleransi selisih angka meter
FUZZY_MIN = 80       # toleransi kemiripan digit

# ===================== PREPROCESS =====================
def preprocess(img_path):
    img = cv2.imread(str(img_path))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # enhance contrast
    gray = cv2.equalizeHist(gray)

    # threshold
    _, th = cv2.threshold(gray, 0, 255,
                           cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return th

# ===================== OCR =====================
def ocr_meter(img):
    config = r"--psm 7 -c tessedit_char_whitelist=0123456789"
    text = pytesseract.image_to_string(img, config=config, lang=OCR_LANG)
    digits = "".join(filter(str.isdigit, text))
    return digits

# ===================== VALIDASI =====================
def validate(ocr_val, db_val):
    if not ocr_val or not db_val:
        return "INVALID"

    try:
        o = int(ocr_val)
        d = int(db_val)

        if abs(o - d) <= MAX_DIFF:
            return "VALID_SELISIH"

        sim = fuzz.ratio(ocr_val, db_val)
        if sim >= FUZZY_MIN:
            return "VALID_POLA"

        return "TIDAK_VALID"

    except:
        return "ERROR"

# ===================== MAIN =====================
def main():
    df = pd.read_excel(EXCEL_ACMT, dtype=str)

    # cari kolom IDPEL & STAN
    idpel_col = next(c for c in df.columns if "idpel" in c.lower())
    stan_col = next(c for c in df.columns if "stan" in c.lower())

    results = []

    for img_path in Path(FOLDER_IMAGE).glob("*.jpg"):
        idpel = img_path.stem
        row = df[df[idpel_col] == idpel]

        if row.empty:
            continue

        stan_db = row.iloc[0][stan_col]

        img = preprocess(img_path)
        ocr_val = ocr_meter(img)
        status = validate(ocr_val, stan_db)

        results.append({
            "IDPEL": idpel,
            "STAN_DB": stan_db,
            "STAN_OCR": ocr_val,
            "STATUS": status,
            "FILE": img_path.name
        })

        print(f"{idpel} | DB:{stan_db} OCR:{ocr_val} => {status}")

    out_df = pd.DataFrame(results)
    out_df.to_excel(OUTPUT_EXCEL, index=False)

    print("\n✅ SELESAI")
    print(f"Hasil: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
