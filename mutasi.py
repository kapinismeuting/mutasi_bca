#!/usr/bin/env python3
"""
Extract BCA bank statement transactions to Excel.
No parsing of Keterangan or Mutasi - keep raw values.
"""

import pdfplumber
import re
import os
from openpyxl import Workbook

# ===== CONFIGURATION =====
PDF_FOLDER = os.path.expanduser("~/dev/appdev/mutasi_bca/Mutasi/2026")
OUTPUT_FOLDER = os.path.expanduser("~/dev/appdev/mutasi_bca/Mutasi/2026")
# =========================

SUMMARY_PATTERN = re.compile(
    r'^(SALDO AWAL|MUTASI CR|MUTASI DB|SALDO AKHIR)\s*:'
)

def parse_bca_transactions(pdf_path: str):
    transactions = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split('\n')
            i = 0
            
            while i < len(lines):
                line = lines[i].strip()
                if not line:
                    i += 1
                    continue
                
                # ===== TRANSACTION ROW =====
                date_match = re.match(r'^(\d{2})/(\d{2})', line)
                if date_match:
                    day = int(date_match.group(1))
                    first_line = line
                    
                    desc_lines = [line[5:].strip()]
                    j = i + 1
                    
                    while j < len(lines):
                        next_line = lines[j].strip()
                        
                        if not next_line:
                            j += 1
                            continue
                        
                        # STOP at next transaction OR summary
                        if re.match(r'^\d{2}/\d{2}', next_line) or SUMMARY_PATTERN.match(next_line):
                            break
                        
                        desc_lines.append(next_line)
                        j += 1
                    
                    full_desc = "\n".join(desc_lines).strip()
                    
                    # ===== Extract DB / CR =====
                    db = ""
                    cr = ""
                    
                    db_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*DB', full_desc)
                    if db_match:
                        db = db_match.group(1)
                    else:
                        # safer CR: only take FIRST number from FIRST LINE
                        first_line_numbers = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', first_line)
                        if first_line_numbers:
                            cr = first_line_numbers[0]
                    
                    # ===== Extract Saldo =====
                    saldo = ""
                    saldo_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*$', first_line)
                    if saldo_match:
                        saldo = saldo_match.group(1)
                    
                    transactions.append({
                        'tanggal': day,
                        'keterangan': full_desc,
                        'db': db,
                        'cr': cr,
                        'saldo': saldo
                    })
                    
                    i = j
                    continue
                
                # ===== SUMMARY ROW =====
                if SUMMARY_PATTERN.match(line):
                    amount_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})', line)
                    amount = amount_match.group(1) if amount_match else ""
                    
                    transactions.append({
                        'tanggal': "",
                        'keterangan': line,
                        'db': "",
                        'cr': "",
                        'saldo': amount
                    })
                
                i += 1
    
    return transactions

def process_single_pdf(pdf_path: str, output_folder: str) -> None:
    os.makedirs(output_folder, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = os.path.join(output_folder, f"{base_name}.xlsx")
    
    print(f"Processing: {os.path.basename(pdf_path)}")
    try:
        transactions = parse_bca_transactions(pdf_path)
        wb = Workbook()
        ws = wb.active
        if ws is None:
            raise RuntimeError("Failed to create worksheet")
        
        ws.title = "Mutasi Rekening"
        headers = ['Tanggal', 'Keterangan', 'DB','CR', 'Saldo']
        ws.append(headers)
        
        for tx in transactions:
            ws.append([
                tx['tanggal'],
                tx['keterangan'],
                tx['db'],
                tx['cr'],
                tx['saldo']
            ])
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 80)
        
        wb.save(output_path)
        print(f"  ✅ Saved {len(transactions)} rows to {output_path}")
    except Exception as e:
        print(f"  ❌ Error: {e}")

def process_all_pdfs(pdf_folder: str, output_folder: str) -> None:
    if not os.path.isdir(pdf_folder):
        print(f"❌ Folder not found: {pdf_folder}")
        return
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"⚠️ No PDFs found in {pdf_folder}")
        return
    for filename in pdf_files:
        pdf_path = os.path.join(pdf_folder, filename)
        process_single_pdf(pdf_path, output_folder)
    print(f"\n✅ All done. Output folder: {output_folder}")

if __name__ == "__main__":
    process_all_pdfs(PDF_FOLDER, OUTPUT_FOLDER)