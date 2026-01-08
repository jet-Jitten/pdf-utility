from pypdf import PdfWriter, PdfReader
from pathlib import Path

def merge_pdfs(pdf_files, output_file):
    writer = PdfWriter()

    for pdf in pdf_files:
        reader = PdfReader(pdf)
        for page in reader.pages:
            writer.add_page(page)

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "wb") as f:
        writer.write(f)
