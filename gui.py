import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

from pypdf import PdfReader, PdfWriter
from pdf_tools.extract import parse_ranges, extract_pages
from pdf_tools.text_extract import extract_text
from pdf_tools.img_to_pdf import images_to_pdf


class PDFApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Utility")
        self.geometry("800x500")
        self.resizable(False, False)

        self._setup_style()
        self._create_tabs()

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

    def _create_tabs(self):
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        notebook.add(self._merge_tab(notebook), text="Merge PDFs")
        notebook.add(self._extract_tab(notebook), text="Extract Pages")
        notebook.add(self._text_tab(notebook), text="Extract Text")
        notebook.add(self._img_tab(notebook), text="Images → PDF")

    # ---------------- MERGE TAB ----------------
    def _merge_tab(self, parent):
        frame = ttk.Frame(parent)

        files = []
        page_ranges = []

        tree = ttk.Treeview(
            frame,
            columns=("pages", "Selected Pages"),
            show="tree headings",
            height=10
        )

        tree.heading("#0", text="File")
        tree.heading("pages", text="Pages")
        tree.heading("Selected Pages", text="Selected Pages")

        tree.column("#0", width=420)
        tree.column("pages", width=80, anchor="center")
        tree.column("Selected Pages", width=120, anchor="center")

        tree.pack(pady=10, fill="x")

        range_var = tk.StringVar(value="Select a file to edit pages")

        def add_pdf(f):
            try:
                with open(f, "rb") as fh:
                    reader = PdfReader(fh)
                    pages = len(reader.pages)
            except Exception:
                pages = "?"

            files.append(f)
            page_ranges.append("all")
            tree.insert("", "end", text=f.split("/")[-1], values=(pages, "all"))

        def add_files():
            for f in filedialog.askopenfilenames(
                filetypes=[("PDF Files", "*.pdf")]
            ):
                add_pdf(f)

        def on_drop(data):
            for f in self.tk.splitlist(data):
                if f.lower().endswith(".pdf"):
                    add_pdf(f)

        tree.drop_target_register(DND_FILES)
        tree.dnd_bind("<<Drop>>", lambda e: on_drop(e.data))

        def on_select(event):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            entry.state(["!disabled"])
            range_var.set(page_ranges[idx])

        def apply_range():
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            page_ranges[idx] = range_var.get().strip()
            tree.set(sel[0], "Selected Pages", page_ranges[idx])

        def remove_file():
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            tree.delete(sel[0])
            files.pop(idx)
            page_ranges.pop(idx)
            entry.state(["disabled"])
            range_var.set("Select a file to edit pages")

        def merge():
            if not files:
                messagebox.showerror("Error", "No PDFs Selected")
                return

            out = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )
            if not out:
                return

            writer = PdfWriter()
            for f, r in zip(files, page_ranges):
                reader = PdfReader(f)
                if r == "all":
                    for p in reader.pages:
                        writer.add_page(p)
                else:
                    for i in parse_ranges(r, len(reader.pages)):
                        writer.add_page(reader.pages[i])

            with open(out, "wb") as fh:
                writer.write(fh)

            messagebox.showinfo("Success", "PDFs merged")

        tree.bind("<<TreeviewSelect>>", on_select)

        btns = ttk.Frame(frame)
        btns.pack()

        ttk.Button(btns, text="Add PDFs", command=add_files).grid(row=0, column=0, padx=5)
        ttk.Button(btns, text="Remove", command=remove_file).grid(row=0, column=1, padx=5)

        range_frame = ttk.Frame(frame)
        range_frame.pack(pady=10)

        ttk.Label(range_frame, text="Pages:").grid(row=0, column=0)
        entry = ttk.Entry(range_frame, textvariable=range_var, width=30)
        entry.grid(row=0, column=1)
        entry.state(["disabled"])

        ttk.Button(range_frame, text="Apply", command=apply_range).grid(row=0, column=2, padx=5)
        ttk.Button(frame, text="Merge PDFs", command=merge).pack(pady=10)

        return frame

    # ---------------- EXTRACT TAB ----------------
    def _extract_tab(self, parent):
        frame = ttk.Frame(parent)

        pdf = tk.StringVar()
        rng = tk.StringVar()

        ttk.Entry(frame, textvariable=pdf, width=60).pack()
        ttk.Button(frame, text="Browse",
                   command=lambda: pdf.set(filedialog.askopenfilename())).pack()

        ttk.Entry(frame, textvariable=rng, width=30).pack()
        ttk.Button(frame, text="Extract",
                   command=lambda: extract_pages(pdf.get(), rng.get(),
                                                  filedialog.asksaveasfilename())).pack()

        return frame

    # ---------------- TEXT TAB ----------------
    def _text_tab(self, parent):
        frame = ttk.Frame(parent)

        pdf = tk.StringVar()
        rng = tk.StringVar()

        ttk.Entry(frame, textvariable=pdf, width=60).pack()
        ttk.Button(frame, text="Browse",
                   command=lambda: pdf.set(filedialog.askopenfilename())).pack()

        ttk.Entry(frame, textvariable=rng, width=30).pack()
        ttk.Button(frame, text="Extract Text",
                   command=lambda: extract_text(pdf.get(), rng.get(),
                                                 filedialog.asksaveasfilename())).pack()

        return frame

    # ---------------- IMAGES TAB ----------------
    def _img_tab(self, parent):
        frame = ttk.Frame(parent)
        images = []

        def pick():
            nonlocal images
            images = filedialog.askopenfilenames(
                filetypes=[("Images", "*.png *.jpg *.jpeg")]
            )

        ttk.Button(frame, text="Select Images", command=pick).pack()
        ttk.Button(frame, text="Convert",
                   command=lambda: images_to_pdf(
                       filedialog.asksaveasfilename(defaultextension=".pdf"), images
                   )).pack()

        return frame


if __name__ == "__main__":
    PDFApp().mainloop()
