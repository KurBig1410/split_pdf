import tkinter as tk
from tkinter import filedialog, messagebox
from split_pdf import split_pdf  # Импортируем функцию разбиения PDF

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter")
        self.root.geometry("500x300")

        # Файл PDF
        tk.Label(root, text="Выберите PDF файл:").pack(pady=5)
        self.pdf_entry = tk.Entry(root, width=50)
        self.pdf_entry.pack(pady=5)
        tk.Button(root, text="Выбрать PDF", command=self.select_pdf).pack(pady=5)

        # Файл Excel
        tk.Label(root, text="Выберите Excel файл:").pack(pady=5)
        self.excel_entry = tk.Entry(root, width=50)
        self.excel_entry.pack(pady=5)
        tk.Button(root, text="Выбрать Excel", command=self.select_excel).pack(pady=5)

        # Папка для сохранения
        tk.Label(root, text="Выберите папку для сохранения:").pack(pady=5)
        self.folder_entry = tk.Entry(root, width=50)
        self.folder_entry.pack(pady=5)
        tk.Button(root, text="Выбрать папку", command=self.select_folder).pack(pady=5)

        # Кнопка "Начать"
        tk.Button(root, text="НАЧАТЬ", command=self.start_split, bg="green", fg="white").pack(pady=10)

    def select_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, file_path)

    def select_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder_path)

    def start_split(self):
        pdf_path = self.pdf_entry.get()
        excel_path = self.excel_entry.get()
        output_folder = self.folder_entry.get()

        if not (pdf_path and excel_path and output_folder):
            messagebox.showerror("Ошибка", "Заполните все поля перед началом!")
            return

        try:
            split_pdf(pdf_path, excel_path, output_folder)
            messagebox.showinfo("Готово", "✅ Разделение PDF завершено!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ Ошибка при разбиении PDF: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()