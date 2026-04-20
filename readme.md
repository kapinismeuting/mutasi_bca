## CONFIG: 

PDF_FOLDER = os.path.expanduser("home/user/your/path/to/Mutasi/2016")
OUTPUT_FOLDER = os.path.expanduser("home/user/your/path/to/Mutasi_Excel")

PDF_Folder = path ke folder file-file mutasi
OUTPUT_FOLDER = path hasil extract data ke xlsx

## INSTALLATION
1. Install required dependencies:
   ```bash
   pip install pdfplumber openpyxl
   ```

## USAGE
### Run mutasi.py (per-file processing)
```bash
python mutasi.py
```
- Processes each PDF in PDF_FOLDER into individual Excel files in OUTPUT_FOLDER

### Run mutasi_by_year.py (yearly consolidation)
```bash
python mutasi_by_year.py
```
- Combines all PDFs into a single Excel file with separate sheets for each year

## CONFIGURATION NOTES
1. Set PDF_FOLDER to your actual PDF directory
2. Set OUTPUT_FOLDER to your desired output location
3. Ensure PDFs are named consistently (e.g., "Mutasi_202301.pdf")

## NOTES
- No parsing of Keterangan or Mutasi fields - keeps raw values
- Requires PDFs to have consistent transaction formatting
- Output files will be created automatically in OUTPUT_FOLDER