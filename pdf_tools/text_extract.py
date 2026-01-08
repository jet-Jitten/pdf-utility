from pypdf import PdfReader
from pathlib import Path
from pdf_tools.extract import parse_ranges

def extract_text(input_pdf, range_str, output_txt):
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)

    pages = parse_ranges(range_str, total_pages)

    text_content = []

    for page_num in pages:
        page = reader.pages[page_num]
        text = page.extract_text()
        if text:
            text_content.append(text)

    output_path = Path(output_txt)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(text_content))
