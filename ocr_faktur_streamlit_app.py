
import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
from PIL import Image

st.set_page_config(page_title="OCR Faktur Pajak", layout="centered")

st.title("📄 OCR Faktur Pajak dari PDF Scan")
st.write("Upload PDF scan faktur pajak, lalu hasilnya akan ditampilkan dan bisa diunduh sebagai Excel.")

uploaded_files = st.file_uploader("Upload satu atau beberapa file PDF", type=["pdf"], accept_multiple_files=True)

def extract_text_from_pdf(pdf_bytes):
    images = convert_from_bytes(pdf_bytes, dpi=300)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang="ind") + "\n"
    return text

def parse_faktur_text(text, filename):
    rows = []
    tanggal = re.search(r'(\d{1,2} \w+ 20\d{2})', text)
    seri = re.search(r'Faktur Pajak.*?(\d{10,})', text)
    penjual = re.search(r'Nama.*?:\s*(.*?)\n.*?NPWP.*?:\s*(\d+)', text, re.DOTALL)
    pembeli = re.search(r'Pembeli.*?Nama.*?:\s*(.*?)\n.*?NPWP.*?:\s*(\d+)', text, re.DOTALL)
    dpp = re.search(r'Dasar Pengenaan Pajak\s*([\d.,]+)', text)
    ppn = re.search(r'Jumlah PPN.*?([\d.,]+)', text)

    items = re.findall(r'(\d+)\s+(.*?)\s+Rp\s*([\d.,]+)', text)

    for item in items:
        row = {
            "File": filename,
            "Tanggal Faktur": tanggal.group(1) if tanggal else '',
            "Nomor Seri": seri.group(1) if seri else '',
            "Nama Penjual": penjual.group(1).strip() if penjual else '',
            "NPWP Penjual": penjual.group(2).strip() if penjual else '',
            "Nama Pembeli": pembeli.group(1).strip() if pembeli else '',
            "NPWP Pembeli": pembeli.group(2).strip() if pembeli else '',
            "Nama Barang/Jasa": item[1].strip(),
            "Harga Jual": item[2].replace('.', '').replace(',', '.') if item[2] else '',
            "DPP": dpp.group(1).replace('.', '').replace(',', '.') if dpp else '',
            "PPN": ppn.group(1).replace('.', '').replace(',', '.') if ppn else ''
        }
        rows.append(row)
    return rows

all_data = []

if uploaded_files:
    with st.spinner("🔍 Memproses dokumen..."):
        for uploaded_file in uploaded_files:
            pdf_bytes = uploaded_file.read()
            text = extract_text_from_pdf(pdf_bytes)
            result = parse_faktur_text(text, uploaded_file.name)
            all_data.extend(result)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"✅ Berhasil membaca {len(uploaded_files)} file.")
        st.dataframe(df)

        # Unduh sebagai Excel
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Faktur Pajak")
        st.download_button("⬇ Download Excel", output.getvalue(), file_name="Hasil_OCR_Faktur.xlsx")
    else:
        st.warning("Tidak ditemukan data yang dapat diproses.")
