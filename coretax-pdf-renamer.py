import os
import re
import pdfplumber
from tkinter import Tk, filedialog, messagebox, Label, Button

def extract_fields(pdf_path, file_type):
    # Open and extract text from PDF
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Gagal membuka folder {pdf_path}: {e}")
        return None

    # Dictionary to map Indonesian month names to numbers
    month_map = {
        'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
        'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
        'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
    }

    # Extract date by searching for pattern "DD MonthName YYYY"
    date_str = "TanggalUnknown"  # Default value
    date_pattern = r"(\d{1,2})\s+(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+(\d{4})"
    date_match = re.search(date_pattern, text)
    
    if date_match:
        day = date_match.group(1).zfill(2)  # Pad single digit days with leading zero
        month = month_map[date_match.group(2)]
        year = date_match.group(3)
        date_str = f"{year}-{month}-{day}"

    # Extract Tax Invoice Code and Serial Number
    kode_match = re.search(r"Kode dan Nomor Seri Faktur Pajak:\s*([\d]+)", text)
    kode_str = kode_match.group(1) if kode_match else "KodeUnknown"

    # Extract reference number
    ref_match = re.search(r"\(Referensi:\s*([^\)]+)\)", text)
    ref_str = ref_match.group(1).strip() if ref_match else "RefUnknown"

    if file_type == "Output":
        # For OutputTaxInvoice, get Buyer Name from "Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:" section
        buyer_section = re.split(r"Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:", text, flags=re.IGNORECASE)
        if len(buyer_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", buyer_section[1])
            buyer = name_match.group(1).strip() if name_match else "PembeliUnknown"
        else:
            buyer = "PembeliUnknown"
        
        # New format: FPK-DATE-KODE-BUYER-REF
        new_name = f"FPK-{date_str}-{kode_str}-{buyer}-{ref_str}.pdf"
        return new_name
    
    elif file_type == "Input":
        # For InputTaxInvoice, get Taxable Entrepreneur Name from "Pengusaha Kena Pajak:" section
        seller_section = re.split(r"Pengusaha Kena Pajak:", text, flags=re.IGNORECASE)
        if len(seller_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", seller_section[1])
            seller = name_match.group(1).strip() if name_match else "PKPUnknown"
        else:
            seller = "PKPUnknown"
        
        # New format: FPM-DATE-KODE-SELLER-REF
        new_name = f"FPM-{date_str}-{kode_str}-{seller}-{ref_str}.pdf"
        return new_name
    
    else:
        return None

def sanitize_filename(filename):
    # Replace invalid Windows filename characters with underscore
    invalid_chars = r'[<>:"/\\|?*]'
    return re.sub(invalid_chars, '_', filename)

def rename_files_in_folder(folder_path):
    renamed = 0
    failed = 0
    
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
                # Sanitize filename to remove invalid Windows characters
                new_filename = sanitize_filename(new_filename)
                new_file_path = os.path.join(folder_path, new_filename)
                
                try:
                    # Check if destination file already exists
                    if os.path.exists(new_file_path):
                        base, ext = os.path.splitext(new_filename)
                        i = 1
                        while os.path.exists(os.path.join(folder_path, f"{base}_{i}{ext}")):
                            i += 1
                        new_file_path = os.path.join(folder_path, f"{base}_{i}{ext}")
                        new_filename = f"{base}_{i}{ext}"
                    
                    os.rename(file_path, new_file_path)
                    print(f"Renamed: {filename} --> {new_filename}")
                    renamed += 1
                except Exception as e:
                    print(f"Error renaming {filename}: {e}")
                    failed += 1
            else:
                print(f"Gagal mengambil data dari  {filename}")
                failed += 1
    
    return renamed, failed

def main():
    root = Tk()
    root.title("CORETAX PDF RENAMER")
    
    # Set window size and position it in the center of the screen
    window_width = 400
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width/2)
    center_y = int(screen_height/2 - window_height/2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    # Add description label
    description = """Aplikasi ini akan membantu anda me-rename file PDF Faktur Pajak Masukan dan Keluaran
    secara otomatis dengan format FPK/FPM-tanggal-nomor FP-Nama WP-referensi """
    label = Label(root, text=description, wraplength=350, justify="center", pady=20)
    label.pack()
    
    # Create and pack the button
    def select_folder():
        folder_selected = filedialog.askdirectory(title="PILIH FOLDER YANG AKAN DIPROSES")
        if not folder_selected:
            messagebox.showinfo("Info", "TIDAK ADA FOLDER YANG DIPILIH.")
            return
            
        # Show progress message
        status_label.config(text="Sedang memproses... Harap menunggu...")
        root.update()
        
        total_renamed, total_failed = rename_files_in_folder(folder_selected)
        
        status_label.config(text=f"Complete! Files renamed: {total_renamed}, Failed: {total_failed}")
        
        # Ask if user wants to open the folder
        if total_renamed > 0:
            if messagebox.askyesno("Open Folder", "Apakah anda mau membuka folder dengan file yang telah diproses?"):
                os.startfile(folder_selected)  # Windows-specific command to open folder
    
    button = Button(root, text="Select Folder", command=select_folder, padx=20, pady=10)
    button.pack()
    
    # Add status label
    status_label = Label(root, text="", wraplength=350, justify="center", pady=20)
    status_label.pack()
    
    root.mainloop()

if __name__ == "__main__":
    main()
