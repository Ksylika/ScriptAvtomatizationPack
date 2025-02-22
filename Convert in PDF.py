import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from docx import Document
from fpdf import FPDF
import PyPDF2
import threading

class PDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Конвертер и объединитель PDF")
        master.geometry("600x400")
        master.configure(bg="#f5f5f5")

        # Создание интерфейса
        self.create_widgets()

        # Переменные для хранения путей
        self.folder_path = ""
        self.output_file_path = ""

    def create_widgets(self):
        # Основной фрейм
        self.main_frame = tk.Frame(self.master, bg="#f5f5f5", padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Метка для ввода папки с файлами
        self.folder_label = tk.Label(self.main_frame, text="Папка с файлами:", bg="#f5f5f5", font=("Helvetica", 12))
        self.folder_label.grid(row=0, column=0, sticky="w", pady=(0, 5))

        # Поле ввода для папки с файлами
        self.folder_entry = tk.Entry(self.main_frame, width=50, font=("Helvetica", 12))
        self.folder_entry.grid(row=1, column=0, sticky="w", pady=(0, 5))

        # Кнопка для выбора папки с файлами
        self.select_folder_button = tk.Button(self.main_frame, text="Обзор...", command=self.select_folder, bg="#4CAF50", fg="white", font=("Helvetica", 12))
        self.select_folder_button.grid(row=1, column=1, padx=(10, 0), pady=(0, 10))

        # Метка для ввода выходного файла PDF
        self.output_label = tk.Label(self.main_frame, text="Выходной PDF файл:", bg="#f5f5f5", font=("Helvetica", 12))
        self.output_label.grid(row=2, column=0, sticky="w", pady=(0, 5))

        # Поле ввода для выходного файла PDF
        self.output_entry = tk.Entry(self.main_frame, width=50, font=("Helvetica", 12))
        self.output_entry.grid(row=3, column=0, sticky="w", pady=(0, 5))

        # Кнопка для выбора выходного файла PDF
        self.select_output_button = tk.Button(self.main_frame, text="Обзор...", command=self.select_output_file, bg="#4CAF50", fg="white", font=("Helvetica", 12))
        self.select_output_button.grid(row=3, column=1, padx=(10, 0), pady=(0, 10))

        # Кнопка для начала конвертации
        self.convert_button = tk.Button(self.main_frame, text="Конвертировать и объединить в PDF", command=self.start_conversion, bg="#2196F3", fg="white", font=("Helvetica", 12))
        self.convert_button.grid(row=4, column=0, columnspan=2, pady=(20, 10))

        # Прогрессбар
        self.progress = tk.ttk.Progressbar(self.main_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=2, pady=10)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, self.folder_path)
            messagebox.showinfo("Выбранная папка", self.folder_path)

    def select_output_file(self):
        self.output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                              filetypes=[("PDF files", "*.pdf")])
        if self.output_file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, self.output_file_path)
            messagebox.showinfo("Выбранный выходной файл", self.output_file_path)

    def start_conversion(self):
        self.folder_path = self.folder_entry.get()
        self.output_file_path = self.output_entry.get()

        if not self.folder_path or not self.output_file_path:
            messagebox.showwarning("Ошибка", "Пожалуйста, выберите папку с файлами и выходной файл PDF.")
            return
        
        # Запуск конвертации в отдельном потоке
        threading.Thread(target=self.convert_and_merge).start()

    def convert_and_merge(self):
        pdf_files = []
        files = os.listdir(self.folder_path)
        total_files = len(files)
        self.update_progress(0, total_files)

        for i, filename in enumerate(files):
            file_path = os.path.join(self.folder_path, filename)
            if filename.endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff')):
                pdf_file = self.convert_image_to_pdf(file_path)
                pdf_files.append(pdf_file)
            elif filename.endswith('.docx'):
                pdf_file = self.convert_docx_to_pdf(file_path)
                pdf_files.append(pdf_file)
            elif filename.endswith('.txt'):
                pdf_file = self.convert_txt_to_pdf(file_path)
                pdf_files.append(pdf_file)
            elif filename.endswith('.html'):
                pdf_file = self.convert_html_to_pdf(file_path)
                pdf_files.append(pdf_file)

            self.update_progress(i + 1, total_files)

        if pdf_files:
            self.merge_pdfs(pdf_files, self.output_file_path)
            self.cleanup_files(pdf_files)
            messagebox.showinfo("Готово", "Конвертация и объединение завершено!")
        else:
            messagebox.showwarning("Ошибка", "Нет файлов для конвертации.")

    def convert_image_to_pdf(self, image_path):
        image = Image.open(image_path)
        pdf_path = image_path.rsplit('.', 1)[0] + '.pdf'
        image.convert("RGB").save(pdf_path, "PDF")
        return pdf_path

    def convert_docx_to_pdf(self, docx_path):
        pdf_path = docx_path.rsplit('.', 1)[0] + '.pdf'
        doc = Document(docx_path)
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        for para in doc.paragraphs:
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, para.text)

        pdf.output(pdf_path)
        return pdf_path

    def convert_txt_to_pdf(self, txt_path):
        pdf_path = txt_path.rsplit('.', 1)[0] + '.pdf'
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        with open(txt_path, 'r', encoding='utf-8') as file:
            for line in file:
                pdf.multi_cell(0, 10, line)

        pdf.output(pdf_path)
        return pdf_path

    def convert_html_to_pdf(self, html_path):
        try:
            from weasyprint import HTML
        except ImportError:
            messagebox.showerror("Ошибка", "WeasyPrint не установлен. Установите его командой 'pip install weasyprint'")
            return None

        pdf_path = html_path.rsplit('.', 1)[0] + '.pdf'
        HTML(html_path).write_pdf(pdf_path)
        return pdf_path

    def merge_pdfs(self, pdf_files, output_file):
        pdf_writer = PyPDF2.PdfWriter()

        for pdf in pdf_files:
            pdf_reader = PyPDF2.PdfReader(pdf)
            for page in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page])

        with open(output_file, 'wb') as out:
            pdf_writer.write(out)

    def cleanup_files(self, pdf_files):
        for pdf in pdf_files:
            os.remove(pdf)

    def update_progress(self, current, total):
        self.progress["value"] = (current / total) * 100
        self.master.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()