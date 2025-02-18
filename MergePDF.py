import os
import fitz
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image as PILImage, ImageTk, Image

class MergePDF:
    def __init__(self, root):
        self.root = root
        self.root.title("MergePDF")
        self.root.geometry("1024x768")
        
        self.style = ttk.Style()
        self.style.configure('Modern.TButton', padding=10, font=('Helvetica', 10))
        self.style.configure('Delete.TButton', padding=5, background='#ff4444')
        
        self.file_list = []
        self.page_images = []
        self.current_pages = []
        self.thumbnail_width = 500
        self.thumbnail_height = 700

        self.main_container = ttk.Frame(root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.create_header()

        self.create_drag_drop_zone()

        self.create_scrollable_content()

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop_files)

    def create_header(self):
        header_frame = ttk.Frame(self.main_container)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        title_label = ttk.Label(
            header_frame, 
            text="MergePDF", 
            font=('Helvetica', 24, 'bold')
        )
        title_label.pack(side=tk.LEFT)

        buttons_frame = ttk.Frame(header_frame)
        buttons_frame.pack(side=tk.RIGHT)

        self.clear_btn = ttk.Button(
            buttons_frame,
            text="Clear All",
            style='Modern.TButton',
            command=self.clear_list
        )
        self.clear_btn.pack(side=tk.RIGHT, padx=5)

    def create_drag_drop_zone(self):
        self.drop_frame = ttk.Frame(
            self.main_container,
            style='Drop.TFrame'
        )
        self.drop_frame.pack(fill=tk.X, pady=(0, 20))

        drop_label = ttk.Label(
            self.drop_frame,
            text="Drag and drop PDF files here",
            font=('Helvetica', 12),
            padding=20
        )
        drop_label.pack(pady=20)

    def create_scrollable_content(self):
        content_frame = ttk.Frame(self.main_container)
        content_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(content_frame)
        self.scrollbar = ttk.Scrollbar(
            content_frame,
            orient="vertical",
            command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def display_pdf(self, pdf_path):
        pdf_document = fitz.open(pdf_path)
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))

            img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((self.thumbnail_width, self.thumbnail_height), Image.Resampling.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)

            self.page_images.append(img_tk)

            page_frame = ttk.Frame(self.scrollable_frame)
            page_frame.pack(pady=10, padx=10, fill=tk.X)

            image_container = ttk.Frame(page_frame)
            image_container.pack(pady=5)

            page_label = ttk.Label(
                image_container,
                text=f"Page {page_num+1}",
                font=('Helvetica', 10, 'bold')
            )
            page_label.pack(pady=(0, 5))

            thumbnail_label = ttk.Label(image_container, image=img_tk)
            thumbnail_label.image = img_tk
            thumbnail_label.pack()

            page_index = len(self.current_pages)
            remove_btn = ttk.Button(
                image_container,
                text="×",
                style='Delete.TButton',
                command=lambda idx=page_index: self.remove_page(idx),
                width=3
            )
            remove_btn.place(relx=1.0, rely=0, x=-5, y=5, anchor="ne")

            page_info = {
                'frame': page_frame,
                'container': image_container,
                'label': thumbnail_label,
                'button': remove_btn,
                'image': img_tk,
                'source_pdf': pdf_path,
                'page_number': page_num
            }
            self.current_pages.append(page_info)

        pdf_document.close()

    def remove_page(self, page_index):
        if 0 <= page_index < len(self.current_pages):
            page_info = self.current_pages[page_index]
            page_info['frame'].destroy()
            self.current_pages.pop(page_index)
            
            for i, page in enumerate(self.current_pages):
                page['button'].configure(
                    command=lambda idx=i: self.remove_page(idx)
                )

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                if file.endswith('.pdf'):
                    self.handle_pdf(file)
                else:
                    self.convert_image_to_pdf(file)

    def handle_pdf(self, pdf_path):
        pdf_document = fitz.open(pdf_path)
        page_count = pdf_document.page_count
        pdf_document.close()

        if page_count == 1:
            self.file_list.append(pdf_path)
            self.display_pdf(pdf_path)
        else:
            answer = messagebox.askyesno(
                "Import PDF",
                f"This PDF has {page_count} pages. Do you want to import the entire PDF?\n\n"
                f"Click 'Yes' to import all pages, or 'No' to specify a range.\n\n"
                f"File: {os.path.basename(pdf_path)}",
                icon='question'
            )

            if answer:
                self.file_list.append(pdf_path)
                self.display_pdf(pdf_path)
            else:
                self.ask_page_range(pdf_path, page_count)

    def ask_page_range(self, pdf_path, total_pages):
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Page Range")
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        content = ttk.Frame(dialog, padding="20")
        content.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            content,
            text=f"Enter page range (1-{total_pages}):",
            font=('Helvetica', 12, 'bold')
        ).pack(pady=(0, 20))

        start_frame = ttk.Frame(content)
        start_frame.pack(fill=tk.X, pady=5)
        ttk.Label(start_frame, text="Start Page:").pack(side=tk.LEFT)
        start_page_entry = ttk.Entry(start_frame)
        start_page_entry.pack(side=tk.RIGHT, expand=True)

        end_frame = ttk.Frame(content)
        end_frame.pack(fill=tk.X, pady=5)
        ttk.Label(end_frame, text="End Page:").pack(side=tk.LEFT)
        end_page_entry = ttk.Entry(end_frame)
        end_page_entry.pack(side=tk.RIGHT, expand=True)

        def on_confirm():
            try:
                start_page = int(start_page_entry.get())
                end_page = int(end_page_entry.get())
                if start_page > end_page or start_page < 1 or end_page > total_pages:
                    raise ValueError("Invalid page range")

                self.import_pdf_pages(pdf_path, start_page, end_page)
                dialog.destroy()
            except ValueError:
                messagebox.showerror(
                    "Error",
                    f"Please enter valid page numbers (1-{total_pages}).",
                    parent=dialog
                )

        button_frame = ttk.Frame(content)
        button_frame.pack(pady=20)
        ttk.Button(
            button_frame,
            text="Confirm",
            style='Modern.TButton',
            command=on_confirm
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Cancel",
            style='Modern.TButton',
            command=dialog.destroy
        ).pack(side=tk.LEFT, padx=5)

    def import_pdf_pages(self, pdf_path, start_page, end_page):
        self.display_pdf_range(pdf_path, start_page - 1, end_page - 1)

    def display_pdf_range(self, pdf_path, start_page, end_page):
        pdf_document = fitz.open(pdf_path)
        
        for page_num in range(start_page, end_page + 1):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))

            img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((self.thumbnail_width, self.thumbnail_height), Image.Resampling.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)

            self.page_images.append(img_tk)

            page_frame = ttk.Frame(self.scrollable_frame)
            page_frame.pack(pady=10, padx=10, fill=tk.X)

            image_container = ttk.Frame(page_frame)
            image_container.pack(pady=5)

            page_label = ttk.Label(
                image_container,
                text=f"Page {page_num+1}",
                font=('Helvetica', 10, 'bold')
            )
            page_label.pack(pady=(0, 5))

            thumbnail_label = ttk.Label(image_container, image=img_tk)
            thumbnail_label.image = img_tk
            thumbnail_label.pack()

            page_index = len(self.current_pages)
            remove_btn = ttk.Button(
                image_container,
                text="×",
                style='Delete.TButton',
                command=lambda idx=page_index: self.remove_page(idx),
                width=3
            )
            remove_btn.place(relx=1.0, rely=0, x=-5, y=5, anchor="ne")

            page_info = {
                'frame': page_frame,
                'container': image_container,
                'label': thumbnail_label,
                'button': remove_btn,
                'image': img_tk,
                'source_pdf': pdf_path,
                'page_number': page_num
            }
            self.current_pages.append(page_info)

        pdf_document.close()

    def convert_image_to_pdf(self, image_path):
        img = PILImage.open(image_path)
        pdf_path = image_path.rsplit('.', 1)[0] + ".pdf"
        img.save(pdf_path, "PDF", resolution=100.0)
        self.file_list.append(pdf_path)
        self.display_pdf(pdf_path)

    def clear_list(self):
        if self.current_pages:
            if messagebox.askyesno(
                "Clear All",
                "Are you sure you want to clear all pages?",
                icon='warning'
            ):
                self.file_list.clear()
                self.page_images.clear()
                self.current_pages.clear()
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.configure(bg='white')
    app = MergePDF(root)
    root.mainloop()
