import sys
from pdf_tools.merge import merge_pdfs
from pdf_tools.split import split_pdf
from pdf_tools.extract import extract_pages
from pdf_tools.text_extract import extract_text
from pdf_tools.img_to_pdf import images_to_pdf

def main():
    if len(sys.argv) < 2:
        print("Commands:")
        print("  merge output.pdf input1.pdf input2.pdf ...")
        print("  split input.pdf start_page end_page output.pdf")
        print("  extract input.pdf ranges output.pdf")
        print("  text input.pdf ranges output.txt")
        print("  img2pdf output.pdf image1.jpg image2.png ...")
        print("    example ranges: 1-3,5,9-10")
        return

    command = sys.argv[1].lower()

    try:
        if command == "merge":
            output = sys.argv[2]
            inputs = sys.argv[3:]
            merge_pdfs(inputs, output)
            print("✅ PDFs merged successfully")

        elif command == "split":
            input_pdf = sys.argv[2]
            start = int(sys.argv[3])
            end = int(sys.argv[4])
            output = sys.argv[5]
            split_pdf(input_pdf, start, end, output)
            print("✅ PDF split successfully")

        elif command == "extract":
            input_pdf = sys.argv[2]
            ranges = sys.argv[3]
            output = sys.argv[4]
            extract_pages(input_pdf, ranges, output)
            print("✅ Pages extracted successfully")

        elif command == "text":
            input_pdf = sys.argv[2]
            ranges = sys.argv[3]
            output = sys.argv[4]
            extract_text(input_pdf, ranges, output)
            print("✅ Text extracted successfully")

        elif command == "img2pdf":
            output = sys.argv[2]
            images = sys.argv[3:]
            images_to_pdf(output, images)
            print("✅ Images converted to PDF")

        else:
            print(f"Unknown command: {command}")

    except Exception as e:
        print("❌ Error:", e)

if __name__ == "__main__":
    main()
