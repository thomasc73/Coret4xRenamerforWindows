import os
import re
import pdfplumber
from tkinter import Tk, filedialog, messagebox

def extract_fields(pdf_path, file_type):
    # Buka dan ekstrak teks dari PDF
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Gagal membuka {pdf_path}: {e}")
        return None

    # Ekstrak tanggal dengan mencari pola tanggal berformat "DD NamaBulan YYYY"
    date_match = re.search(r"(\d{1,2}\s+(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4})", text)
    date_str = date_match.group(1) if date_match else "TanggalUnknown"

    # Ekstrak Kode dan Nomor Seri Faktur Pajak
    kode_match = re.search(r"Kode dan Nomor Seri Faktur Pajak:\s*([\d]+)", text)
    kode_str = kode_match.group(1) if kode_match else "KodeUnknown"

    if file_type == "Output":
        # Untuk OutputTaxInvoice, ambil Nama Pembeli dari bagian "Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:"
        buyer_section = re.split(r"Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:", text, flags=re.IGNORECASE)
        if len(buyer_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", buyer_section[1])
            buyer = name_match.group(1).strip() if name_match else "PembeliUnknown"
        else:
            buyer = "PembeliUnknown"
        # Format nama baru: FPK-NAMA PEMBELI-TANGGAL TANDA TANGAN-Kode dan Nomor Seri Faktur Pajak
        new_name = f"FPK-{buyer}-{date_str}-{kode_str}.pdf"
        return new_name

    elif file_type == "Input":
        # Untuk InputTaxInvoice, ambil Nama Pengusaha Kena Pajak dari bagian "Pengusaha Kena Pajak:"
        seller_section = re.split(r"Pengusaha Kena Pajak:", text, flags=re.IGNORECASE)
        if len(seller_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", seller_section[1])
            seller = name_match.group(1).strip() if name_match else "PKPUnknown"
        else:
            seller = "PKPUnknown"
        # Format nama baru: FPM-Pengusaha Kena Pajak-TANGGAL TANDA TANGAN-Kode dan Nomor Seri Faktur Pajak
        new_name = f"FPM-{seller}-{date_str}-{kode_str}.pdf"
        return new_name

    else:
        return None

def rename_files_in_folder(folder_path):
    renamed = 0
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            if filename.startswith("OutputTaxInvoice-"):
                new_filename = extract_fields(file_path, "Output")
            elif filename.startswith("InputTaxInvoice-"):
                new_filename = extract_fields(file_path, "Input")
            else:
                continue

            if new_filename:
                new_file_path = os.path.join(folder_path, new_filename)
                try:
                    os.rename(file_path, new_file_path)
                    print(f"Renamed: {filename} --> {new_filename}")
                    renamed += 1
                except Exception as e:
                    print(f"Error saat merename {filename}: {e}")
            else:
                print(f"Gagal ekstrak data dari {filename}")
    return renamed

def main():
    root = Tk()
    root.withdraw()  # Sembunyikan jendela utama Tkinter
    folder_selected = filedialog.askdirectory(title="Pilih Folder yang berisi file PDF")
    if not folder_selected:
        messagebox.showinfo("Info", "Folder tidak dipilih. Aplikasi akan keluar.")
        return

    total_renamed = rename_files_in_folder(folder_selected)
    messagebox.showinfo("Selesai", f"Proses rename selesai.\nTotal file yang direname: {total_renamed}")

if __name__ == "__main__":
    main()
