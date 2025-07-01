import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from generator import generate_documents

class App:
    def __init__(self, master):
        self.master = master
        master.title("DOCX Generator")

        self.root_dir = tk.StringVar(value=".")
        self.main_file = tk.StringVar()
        self.template_file = tk.StringVar()
        self.output_dir = tk.StringVar(value="output_docs")
        self.common_column = tk.StringVar(value="id")
        self.file_name_column = tk.StringVar(value="id")
        self.stop_flag = False
        row = 0
        tk.Label(master, text="Папка з таблицями:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.root_dir, width=40).grid(row=row, column=1)
        tk.Button(master, text="...", command=self.select_root_dir).grid(row=row, column=2)
        row += 1

        tk.Label(master, text="Основний Excel-файл:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.main_file, width=40).grid(row=row, column=1)
        tk.Button(master, text="...", command=self.select_main_file).grid(row=row, column=2)
        row += 1

        tk.Label(master, text="Шаблон DOCX:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.template_file, width=40).grid(row=row, column=1)
        tk.Button(master, text="...", command=self.select_template_file).grid(row=row, column=2)
        row += 1

        tk.Label(master, text="Папка збереження:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.output_dir, width=40).grid(row=row, column=1)
        tk.Button(master, text="...", command=self.select_output_dir).grid(row=row, column=2)
        row += 1

        tk.Label(master, text="Назва спільного стовпця:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.common_column).grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        tk.Label(master, text="Стовпець для імені файлу:").grid(row=row, column=0, sticky="e")
        tk.Entry(master, textvariable=self.file_name_column).grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        tk.Button(master, text="Старт", command=self.generate).grid(row=row, column=0, columnspan=3, pady=10)
        tk.Button(master,text = "Стоп", command =self.stop_generation).grid(row=row, column=1, columnspan=3, pady=2)
        row += 1

        self.log = scrolledtext.ScrolledText(master, width=60, height=15, state='disabled', font=("Consolas", 10))
        self.log.grid(row=row, column=0, columnspan=3, pady=5, sticky="we")

    def log_write(self, text, end="\n"):
        self.log['state'] = 'normal'
        self.log.insert('end', text + end)
        self.log.see('end')
        self.log['state'] = 'disabled'
        self.master.update_idletasks()

    def select_root_dir(self):
        dirname = filedialog.askdirectory(title="Оберіть папку з Excel файлами")
        if dirname:
            self.root_dir.set(dirname)

    def select_main_file(self):
        filename = filedialog.askopenfilename(title="Оберіть основний Excel-файл", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.main_file.set(filename)
            self.root_dir.set(os.path.dirname(filename))

    def select_template_file(self):
        filename = filedialog.askopenfilename(title="Оберіть шаблон DOCX", filetypes=[("DOCX files", "*.docx")])
        if filename:
            self.template_file.set(filename)

    def select_output_dir(self):
        dirname = filedialog.askdirectory(title="Оберіть папку для збереження DOCX")
        if dirname:
            self.output_dir.set(dirname)

    def generate(self):
        import threading
        threading.Thread(target=self._generate_thread).start()

    def stop_generation(self):
        self.stop_flag = True
        self.log_write("⛔ Зупинка процесу запрошена...")

    def _generate_thread(self):
        self.stop_flag = False
        generate_documents(
            root_dir=self.root_dir.get(),
            main_path=self.main_file.get(),
            template_path=self.template_file.get(),
            output_dir=self.output_dir.get(),
            common_column=self.common_column.get(),
            file_name_column=self.file_name_column.get(),
            log_callback=self.log_write,
            stop_flag=lambda: self.stop_flag
        )
