import os
import re
import pdfplumber
from tkinter import Tk, filedialog, messagebox

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
        print(f"Failed to open {pdf_path}: {e}")
        return None

    # Extract date by searching for pattern "DD MonthName YYYY"
    date_match = re.search(r"(\d{1,2}\s+(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4})", text)
    date_str = date_match.group(1) if date_match else "TanggalUnknown"

    # Extract Tax Invoice Code and Serial Number
    kode_match = re.search(r"Kode dan Nomor Seri Faktur Pajak:\s*([\d]+)", text)
    kode_str = kode_match.group(1) if kode_match else "KodeUnknown"

    if file_type == "Output":
        # For OutputTaxInvoice, get Buyer Name from "Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:" section
        buyer_section = re.split(r"Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:", text, flags=re.IGNORECASE)
        if len(buyer_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", buyer_section[1])
            buyer = name_match.group(1).strip() if name_match else "PembeliUnknown"
        else:
            buyer = "PembeliUnknown"
        
        # Format new name: FPK-BUYER NAME-SIGNATURE DATE-Tax Invoice Code and Serial Number
        new_name = f"FPK-{buyer}-{date_str}-{kode_str}.pdf"
        return new_name
    
    elif file_type == "Input":
        # For InputTaxInvoice, get Taxable Entrepreneur Name from "Pengusaha Kena Pajak:" section
        seller_section = re.split(r"Pengusaha Kena Pajak:", text, flags=re.IGNORECASE)
        if len(seller_section) > 1:
            name_match = re.search(r"Nama\s*:\s*(.+)", seller_section[1])
            seller = name_match.group(1).strip() if name_match else "PKPUnknown"
        else:
            seller = "PKPUnknown"
        
        # Format new name: FPM-Taxable Entrepreneur-SIGNATURE DATE-Tax Invoice Code and Serial Number
        new_name = f"FPM-{seller}-{date_str}-{kode_str}.pdf"
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
                print(f"Failed to extract data from {filename}")
                failed += 1
    
    return renamed, failed

def main():
    root = Tk()
    root.withdraw()  # Hide main Tkinter window
    root.title("PDF Invoice Renamer")
    
    # Set icon for dialog boxes
    root.iconbitmap(default=None)  # You can set a custom icon here if you have one
    
    folder_selected = filedialog.askdirectory(title="Select folder containing PDF files")
    if not folder_selected:
        messagebox.showinfo("Info", "No folder selected. Application will exit.")
        return
    
    # Show progress message
    print(f"Processing files in: {folder_selected}")
    print("This may take a few moments, please wait...")
    
    total_renamed, total_failed = rename_files_in_folder(folder_selected)
    
    messagebox.showinfo("Complete", f"Renaming process complete.\n\nFiles renamed: {total_renamed}\nFiles failed: {total_failed}")
    
    # Ask if user wants to open the folder
    if total_renamed > 0:
        if messagebox.askyesno("Open Folder", "Do you want to open the folder with renamed files?"):
            os.startfile(folder_selected)  # Windows-specific command to open folder

if __name__ == "__main__":
    main()
