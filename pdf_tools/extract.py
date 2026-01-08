from pypdf import PdfReader, PdfWriter
from pathlib import Path

def parse_ranges(range_str, total_pages):
    pages = []

    ranges = range_str.split(",")

    for r in ranges:
        r = r.strip()
        if "-" in r:
            start, end = r.split("-")
            start = int(start)
            end = int(end)

            if start < 1 or end > total_pages or start > end:
                raise ValueError(f"Invalid range: {r}")

            pages.extend(range(start - 1, end))
        else:
            page = int(r)
            if page < 1 or page > total_pages:
                raise ValueError(f"Invalid page: {page}")
            pages.append(page - 1)

    return pages


def extract_pages(input_pdf, range_str, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    total_pages = len(reader.pages)
    pages = parse_ranges(range_str, total_pages)

    for page_num in pages:
        writer.add_page(reader.pages[page_num])

    output_path = Path(output_pdf)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "wb") as f:
        writer.write(f)
