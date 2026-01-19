# Usage

* Export the XLSX from Google Drive to e.g. the working directory here
* Run `uv run make_documents.py -f html Karaoke.xlsx` to generate a HTML segment
* Run `uv run make_documents.py -f typst Karaoke.xlsx` to generate a Typst document
  * Run `typst compile karaoke.typ` to generate a PDF from that Typst document
