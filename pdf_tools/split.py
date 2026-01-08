from pypdf import PdfReader, PdfWriter
from pathlib import Path

def split_pdf(input_pdf, start_page, end_page, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    total_pages = len(reader.pages)

    if start_page < 1 or end_page > total_pages or start_page > end_page:
        raise ValueError(
            f"Invalid page range. PDF has {total_pages} pages."
        )

    # Convert to 0-based index
    for page_num in range(start_page - 1, end_page):
        writer.add_page(reader.pages[page_num])

    output_path = Path(output_pdf)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "wb") as f:
        writer.write(f)
