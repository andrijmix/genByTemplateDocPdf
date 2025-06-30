import pandas as pd
import os
import glob
from datetime import datetime
from docxtpl import DocxTemplate
import jinja2
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def format_date(val):
    if pd.isnull(val):
        return '—'
    if isinstance(val, datetime):
        if val.time() == datetime.min.time():
            return val.strftime('%d.%m.%Y')
        elif val.second == 0:
            return val.strftime('%d.%m.%Y %H:%M')
        else:
            return val.strftime('%d.%m.%Y %H:%M:%S')
    return str(val)

def floatformat(val, precision=2):
    try:
        precision = int(precision)
        return f"{float(val):.{precision}f}".replace('.', ',')
    except Exception:
        return val

jinja_env = jinja2.Environment()
jinja_env.filters['floatformat'] = floatformat

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
        row += 1

        # Лог вікно
        self.log = scrolledtext.ScrolledText(master, width=60, height=15, state='disabled', font=("Consolas", 10))
        self.log.grid(row=row, column=0, columnspan=3, pady=5, sticky="we")

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

    def log_write(self, text, end="\n"):
        self.log['state'] = 'normal'
        self.log.insert('end', text + end)
        self.log.see('end')
        self.log['state'] = 'disabled'
        self.master.update_idletasks()

    def generate(self):
        import threading
        threading.Thread(target=self._generate_thread).start()

    def _generate_thread(self):
        try:
            self.log_write("=== Старт генерації DOCX ===")
            root_dir = self.root_dir.get()
            main_path = self.main_file.get()
            template_path = self.template_file.get()
            output_dir = self.output_dir.get()
            common_column = self.common_column.get()
            file_name_column = self.file_name_column.get()

            if not all([os.path.exists(main_path), os.path.exists(template_path), os.path.isdir(root_dir)]):
                self.log_write("❌ Помилка: Перевірте всі шляхи до файлів!")
                return

            os.makedirs(output_dir, exist_ok=True)

            main_df = pd.read_excel(main_path)
            main_df.columns = main_df.columns.str.strip()
            all_xlsx = glob.glob(os.path.join(root_dir, "*.xlsx"))
            other_xlsx = [f for f in all_xlsx if os.path.abspath(f) != os.path.abspath(main_path)]
            other_tables = {}
            for fname in other_xlsx:
                name = os.path.splitext(os.path.basename(fname))[0].lower()
                df = pd.read_excel(fname)
                df.columns = df.columns.str.strip()
                other_tables[name] = df
            self.log_write(f"Знайдено {len(main_df)} записів. Запущено потоки...")

            def generate_docx(borrower):
                context = {}
                for col in main_df.columns:
                    val = borrower[col]
                    context[f"{col}_credit"] = val if pd.notnull(val) else "—"
                for tablename, df in other_tables.items():
                    if common_column not in df.columns:
                        continue
                    filtered = df[df[common_column] == borrower[common_column]]
                    rows = []
                    for _, row in filtered.iterrows():
                        row_dict = {}
                        for col in df.columns:
                            val = row[col]
                            if isinstance(val, datetime):
                                row_dict[col] = format_date(val)
                            else:
                                row_dict[col] = val if pd.notnull(val) else "—"
                        rows.append(row_dict)
                    context[f"{tablename}_table"] = rows
                tpl = DocxTemplate(template_path)
                tpl.render(context, jinja_env)
                safe_name = str(borrower.get(file_name_column, borrower[common_column])).replace(" ", "_")
                docx_filename = os.path.join(output_dir, f"doc_{safe_name}.docx")
                tpl.save(docx_filename)
                return docx_filename

            created_docx_files = []
            with ThreadPoolExecutor() as executor:
                futures = [executor.submit(generate_docx, borrower) for _, borrower in main_df.iterrows()]
                for i, future in enumerate(as_completed(futures), 1):
                    fname = future.result()
                    created_docx_files.append(fname)
                    self.log_write(f"[{i}/{len(futures)}] Згенеровано: {os.path.basename(fname)}")
            self.log_write(f"\n✅ Успішно створено {len(created_docx_files)} DOCX документів у {output_dir}\n")
        except Exception as e:
            self.log_write("❌ ПОМИЛКА: " + str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
