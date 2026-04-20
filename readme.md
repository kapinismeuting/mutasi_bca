--
Config: 

PDF_FOLDER = os.path.expanduser("~/dev/appdev/Mutasi/2016")
OUTPUT_FOLDER = os.path.expanduser("~/dev/appdev/Mutasi_Excel")
---

PDF_Folder = path ke folder file-file mutasi
OUTPUT_FOLDER = path hasil extract data ke xlsx

mutasi_by_year: many file pdf mutasi = 1 file xlsx with sheets for each file
mutasi: 1 file pdf mutasi = 1 file xlsx