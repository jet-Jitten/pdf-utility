import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter
from pdf_tools.extract import parse_ranges, extract_pages
from pdf_tools.text_extract import extract_text
from pdf_tools.img_to_pdf import images_to_pdf

import pandas as pd
import pdfplumber
from docx import Document


from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        self.configure(bg="#1f1f1f")
        self.title("PDF Utility")
        self.geometry("750x750")
        

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        ctk.set_widget_scaling(1.1)

        self.attributes("-alpha", 1)

        self._style_treeview()
        self._create_tabs()

    def _update_preview(self, file, preview_var):
        try:
            reader = PdfReader(file)
            text_preview = ""

            try:
                text = reader.pages[0].extract_text()
                if text:
                    text_preview = text[:200].replace("\n", " ")
            except:
                text_preview = ""

            preview_var.set(
                f"{file.split('/')[-1]}\nPages: {len(reader.pages)}\n{text_preview}"
            )
        except:
            preview_var.set("Preview unavailable")

    # ---------- TREEVIEW DARK STYLE ----------
    def _style_treeview(self):
        style = ttk.Style()
        style.theme_use("default")

        style.configure("Treeview",
            background="#2b2b2b",
            foreground="white",
            fieldbackground="#2b2b2b",
            rowheight=28
        )

        style.configure("Treeview.Heading",
            background="#1f1f1f",
            foreground="white"
        )

        style.map("Treeview",
            background=[("selected", "#1f6aa5")]
        )

    # ---------- TABS ----------
    def _create_tabs(self):
        tabview = ctk.CTkTabview(self)
        tabview.pack(expand=True, fill="both", padx=0.5, pady=0.5)

        tabview.add("Merge")
        tabview.add("Extract Pages")
        tabview.add("Extract Text")
        tabview.add("Images")
        tabview.add("PDF to Excel")
        tabview.add("PDF to Word")
        tabview.configure(fg_color="#1f1f1f")

        self._merge_tab(tabview.tab("Merge"))
        self._extract_tab(tabview.tab("Extract Pages"))
        self._text_tab(tabview.tab("Extract Text"))
        self._img_tab(tabview.tab("Images"))
        self._pdf_to_excel_tab(tabview.tab("PDF to Excel"))
        self._pdf_to_word_tab(tabview.tab("PDF to Word"))

    # ---------- MERGE ----------
    def _merge_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
        
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(frame, text="Merge PDF Files", font=("Segoe UI", 18, "bold")).pack(pady=10)

        files, ranges = [], []
        drag = {"item": None}

        tree = ttk.Treeview(frame, columns=("pages", "selected"), show="tree headings", height=10)
        tree.heading("#0", text="File")
        tree.heading("pages", text="Pages")
        tree.heading("selected", text="Selected")
        tree.column("#0", width=420)
        tree.column("pages", width=80, anchor="center")
        tree.column("selected", width=120, anchor="center")
        tree.pack(fill="x", pady=10, padx=10)

        preview_var = tk.StringVar(value="Drop or select a file")

        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)

        range_var = tk.StringVar(value="Select a file")

        def drop_files(event):
            dropped_files = self.tk.splitlist(event.data)
            for f in dropped_files:
                if f.lower().endswith(".pdf"):
                    files.append(f)
                    ranges.append("all")
                    try:
                        pages = len(PdfReader(f).pages)
                    except:
                        pages = "?"
                    tree.insert("", "end", text=f.split("/")[-1], values=(pages, "all"))
                    self._update_preview(f, preview_var)
                    preview_box.delete("1.0", "end")
                    preview_box.insert("end", preview_var.get())
        tree.drop_target_register(DND_FILES)
        tree.dnd_bind('<<Drop>>', drop_files)

        def add_files():
            for f in filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")]):
                files.append(f)
                ranges.append("all")
                try:
                    pages = len(PdfReader(f).pages)
                except:
                    pages = "?"
                tree.insert("", "end", text=f.split("/")[-1], values=(pages, "all"))

        def remove():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                tree.delete(sel[0])
                files.pop(i)
                ranges.pop(i)

        def move_up():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                if i > 0:
                    files[i], files[i-1] = files[i-1], files[i]
                    ranges[i], ranges[i-1] = ranges[i-1], ranges[i]
                    tree.move(sel[0], "", i-1)

        def move_down():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                if i < len(files)-1:
                    files[i], files[i+1] = files[i+1], files[i]
                    ranges[i], ranges[i+1] = ranges[i+1], ranges[i]
                    tree.move(sel[0], "", i+1)

        def apply_range():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                ranges[i] = range_var.get()
                tree.set(sel[0], "selected", ranges[i])

        def on_select(e):
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                entry.configure(state="normal")
                range_var.set(ranges[i])

                try:
                    reader = PdfReader(files[i])
                    preview_var.set(f"{files[i].split('/')[-1]} | Pages: {len(reader.pages)}")
                except:
                    preview_var.set("Preview unavailable")
                self._update_preview(files[i], preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        def merge():
            if not files:
                messagebox.showerror("Error", "No PDFs selected")
                return

            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            if not out:
                return

            writer = PdfWriter()
            for f, r in zip(files, ranges):
                reader = PdfReader(f)
                if r == "all":
                    for p in reader.pages:
                        writer.add_page(p)
                else:
                    for i in parse_ranges(r, len(reader.pages)):
                        writer.add_page(reader.pages[i])

            with open(out, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Success", "PDF merged successfully")

        def drag_start(e):
            drag["item"] = tree.identify_row(e.y)

        def drag_motion(e):
            item = drag["item"]
            target = tree.identify_row(e.y)
            if item and target and item != target:
                i, j = tree.index(item), tree.index(target)
                files[i], files[j] = files[j], files[i]
                ranges[i], ranges[j] = ranges[j], ranges[i]
                tree.move(item, "", j)

        tree.bind("<<TreeviewSelect>>", on_select)
        tree.bind("<ButtonPress-1>", drag_start)
        tree.bind("<B1-Motion>", drag_motion)

        controls = ctk.CTkFrame(frame)
        controls.pack(pady=10)

        ctk.CTkButton(controls, text="Add Files", width=120, command=add_files).grid(row=0, column=0, padx=5)
        ctk.CTkButton(controls, text="Move Up", width=120, command=move_up).grid(row=0, column=1, padx=5)
        ctk.CTkButton(controls, text="Move Down", width=120, command=move_down).grid(row=0, column=2, padx=5)
        ctk.CTkButton(controls, text="Remove", width=120, command=remove).grid(row=0, column=3, padx=5)

        range_frame = ctk.CTkFrame(frame)
        range_frame.pack(pady=10)

        ctk.CTkLabel(range_frame, text="Pages (e.g. 1-3,5):").grid(row=0, column=0, padx=5)
        entry = ctk.CTkEntry(range_frame, textvariable=range_var, width=200)
        entry.grid(row=0, column=1, padx=10)
        entry.configure(state="disabled")

        ctk.CTkButton(range_frame, text="Apply", width=100, command=apply_range).grid(row=0, column=2, padx=5)

        ctk.CTkButton(frame, text="Merge & Save", width=200, command=merge).pack(pady=15)

    # ---------- EXTRACT ----------
    def _extract_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
       
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        

        ctk.CTkLabel(frame, text="Extract Pages", font=("Segoe UI", 18, "bold")).pack(pady=10)

        path = tk.StringVar()
        rng = tk.StringVar()
        info = tk.StringVar(value="No file selected")

        ctk.CTkEntry(frame, textvariable=path).pack(fill="x", padx=20, pady=5)

        preview_var = tk.StringVar(value="Drop or select a file")
        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)

        def browse():
            f = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if f:
                path.set(f)
                info.set(f"Total Pages: {len(PdfReader(f).pages)}")
            
            self._update_preview(f, preview_var)
            preview_box.delete("1.0", "end")
            preview_box.insert("end", preview_var.get())

        ctk.CTkButton(frame, text="Browse", command=browse).pack(pady=5)
        ctk.CTkLabel(frame, textvariable=info).pack()
        ctk.CTkLabel(
            frame,
            text="Enter pages like: 1-3,5,8",
            text_color="gray"
        ).pack()

        ctk.CTkEntry(frame, textvariable=rng).pack(pady=10)

        def drop(event):
            f = self.tk.splitlist(event.data)[0]
            if f.lower().endswith(".pdf"):
                path.set(f)
                self._update_preview(f, preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', drop)
        

        def run():
            if not path.get() or not rng.get():
                messagebox.showerror("Error", "Select file and pages")
                return
            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            extract_pages(path.get(), rng.get(), out)
            messagebox.showinfo("Success", "Pages extracted")

        ctk.CTkButton(frame, text="Extract Pages", command=run).pack(pady=10)

    # ---------- TEXT ----------
    def _text_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
        
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(frame, text="Extract Text", font=("Segoe UI", 18, "bold")).pack(pady=10)

        path = tk.StringVar()
        rng = tk.StringVar()
        info = tk.StringVar(value="No file selected")

        ctk.CTkEntry(frame, textvariable=path).pack(fill="x", padx=20, pady=5)

        preview_var = tk.StringVar(value="Drop or select a file")
        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)

        def browse():
            f = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if f:
                path.set(f)
                info.set(f"Total Pages: {len(PdfReader(f).pages)}")

            self._update_preview(f, preview_var)
            preview_box.delete("1.0", "end")
            preview_box.insert("end", preview_var.get())


        ctk.CTkButton(frame, text="Browse", command=browse).pack(pady=5)
        ctk.CTkLabel(frame, textvariable=info).pack()
        ctk.CTkLabel(
        frame,
        text="Enter pages like: 1-3,5,8",
        text_color="gray"
        ).pack()

        ctk.CTkEntry(frame, textvariable=rng).pack(pady=10)

        def drop(event):
            f = self.tk.splitlist(event.data)[0]
            if f.lower().endswith(".pdf"):
                path.set(f)
                self._update_preview(f, preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', drop)

        def run():
            if not path.get() or not rng.get():
                messagebox.showerror("Error", "Select file and pages")
                return
            out = filedialog.asksaveasfilename(defaultextension=".txt")
            extract_text(path.get(), rng.get(), out)
            messagebox.showinfo("Success", "Text extracted")

        ctk.CTkButton(frame, text="Extract Text", command=run).pack(pady=10)

    # ---------- IMAGE ----------
    def _img_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(frame, text="Images to PDF", font=("Segoe UI", 18, "bold")).pack(pady=10)

        images = []
        drag = {"item": None}

        tree = ttk.Treeview(frame, columns=("name",), show="tree headings", height=10)
        tree.heading("#0", text="File")
        tree.column("#0", width=600)
        tree.pack(fill="x", pady=10, padx=10)

        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)

        # ---------- DRAG & DROP ----------
        def drop_files(event):
            dropped_files = self.tk.splitlist(event.data)
            for f in dropped_files:
                if f.lower().endswith((".png", ".jpg", ".jpeg")):
                    images.append(f)
                    tree.insert("", "end", text=f.split("/")[-1])

                    preview_box.delete("1.0", "end")
                    preview_box.insert("end", f.split("/")[-1])

        tree.drop_target_register(DND_FILES)
        tree.dnd_bind('<<Drop>>', drop_files)

        # ---------- ADD ----------
        def add():
            for f in filedialog.askopenfilenames(filetypes=[("Images", "*.png *.jpg *.jpeg")]):
                images.append(f)
                tree.insert("", "end", text=f.split("/")[-1])

        # ---------- REMOVE ----------
        def remove():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                tree.delete(sel[0])
                images.pop(i)

        # ---------- MOVE ----------
        def move_up():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                if i > 0:
                    images[i], images[i-1] = images[i-1], images[i]
                    tree.move(sel[0], "", i-1)

        def move_down():
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                if i < len(images)-1:
                    images[i], images[i+1] = images[i+1], images[i]
                    tree.move(sel[0], "", i+1)

        # ---------- SELECT ----------
        def on_select(e):
            sel = tree.selection()
            if sel:
                i = tree.index(sel[0])
                preview_box.delete("1.0", "end")
                preview_box.insert("end", images[i].split("/")[-1])

        # ---------- CONVERT ----------
        def convert():
            if not images:
                messagebox.showerror("Error", "No images selected")
                return

            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            if not out:
                return

            images_to_pdf(out, images)
            messagebox.showinfo("Success", "PDF created")

        # ---------- DRAG REORDER ----------
        def drag_start(e):
            drag["item"] = tree.identify_row(e.y)

        def drag_motion(e):
            item = drag["item"]
            target = tree.identify_row(e.y)
            if item and target and item != target:
                i, j = tree.index(item), tree.index(target)
                images[i], images[j] = images[j], images[i]
                tree.move(item, "", j)

        tree.bind("<<TreeviewSelect>>", on_select)
        tree.bind("<ButtonPress-1>", drag_start)
        tree.bind("<B1-Motion>", drag_motion)

        # ---------- CONTROLS ----------
        controls = ctk.CTkFrame(frame)
        controls.pack(pady=10)

        ctk.CTkButton(controls, text="Add Images", width=140, command=add).grid(row=0, column=0, padx=5)
        ctk.CTkButton(controls, text="Remove", width=140, command=remove).grid(row=0, column=3, padx=5)
        ctk.CTkButton(controls, text="Move Up", width=120, command=move_up).grid(row=0, column=1, padx=5)
        ctk.CTkButton(controls, text="Move Down", width=120, command=move_down).grid(row=0, column=2, padx=5)

        ctk.CTkButton(frame, text="Create PDF", width=200, command=convert).pack(pady=15)

    # ---------- PDF TO EXCEL (OFFICE SAFE) ----------
    def _pdf_to_excel_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(frame, text="PDF to Excel", font=("Segoe UI", 18, "bold")).pack(pady=10)

        path = tk.StringVar()

        ctk.CTkEntry(frame, textvariable=path).pack(fill="x", padx=20, pady=5)

        preview_var = tk.StringVar(value="Drop or select a file")
        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)

        def browse():
            f = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if f:
                path.set(f)
            
            self._update_preview(f, preview_var)
            preview_box.delete("1.0", "end")
            preview_box.insert("end", preview_var.get())

        ctk.CTkButton(frame, text="Browse PDF", command=browse).pack(pady=5)

        def drop(event):
            f = self.tk.splitlist(event.data)[0]
            if f.lower().endswith(".pdf"):
                path.set(f)
                self._update_preview(f, preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', drop)

        def convert():
            if not path.get():
                messagebox.showerror("Error", "Select a PDF file")
                return

            out = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if not out:
                return

            try:
                all_tables = []

                with pdfplumber.open(path.get()) as pdf:
                    for page_num, page in enumerate(pdf.pages, start=1):
                        tables = page.extract_tables()
                        

                        if not tables:
                            text = page.extract_text()
                            if text:
                                tables = [[line.split() for line in text.split("\n")]]

                        for table in tables:
                            df = pd.DataFrame(table)

                            # Clean empty columns
                            df = df.dropna(axis=1, how='all')

                            # Fill missing values
                            df = df.fillna("")

                            # Optional: auto header detection
                            if len(df) > 1:
                                df.columns = df.iloc[0]
                                df = df[1:]
                            df.insert(0, "Page", page_num)
                            all_tables.append(df)

                if not all_tables:
                    messagebox.showwarning("No Tables", "No tables found in PDF")
                    return

                with pd.ExcelWriter(out) as writer:
                    for i, df in enumerate(all_tables):
                        df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

                messagebox.showinfo("Success", f"{len(all_tables)} tables exported")

            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(frame, text="Convert to Excel", command=convert).pack(pady=15)
    # ---------- PDF TO WORD ----------
    def _pdf_to_word_tab(self, parent):
        frame = ctk.CTkScrollableFrame(parent)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(frame, text="PDF to Word", font=("Segoe UI", 18, "bold")).pack(pady=10)

        path = tk.StringVar()

        ctk.CTkEntry(frame, textvariable=path).pack(fill="x", padx=20, pady=5)

        preview_var = tk.StringVar(value="Drop or select a file")
        preview_box = ctk.CTkTextbox(frame, height=80)
        preview_box.pack(fill="x", padx=10, pady=5)


        def browse():
            f = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if f:
                path.set(f)

                self._update_preview(f, preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        ctk.CTkButton(frame, text="Browse PDF", command=browse).pack(pady=5)

        def drop(event):
            f = self.tk.splitlist(event.data)[0]
            if f.lower().endswith(".pdf"):
                path.set(f)
                self._update_preview(f, preview_var)
                preview_box.delete("1.0", "end")
                preview_box.insert("end", preview_var.get())

        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', drop)

        def convert():
            if not path.get():
                messagebox.showerror("Error", "Select a PDF file")
                return

            out = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")]
            )

            if not out:
                return

            try:
                doc = Document()

                with pdfplumber.open(path.get()) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            doc.add_paragraph(text)

                doc.save(out)

                messagebox.showinfo("Success", "Converted to Word successfully")

            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(frame, text="Convert to Word", command=convert).pack(pady=15)


if __name__ == "__main__":
    app = PDFApp()
    app.mainloop()