import os
import glob
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from concurrent.futures import ThreadPoolExecutor, as_completed
from utils import format_date, floatformat
import jinja2

jinja_env = jinja2.Environment()
jinja_env.filters['floatformat'] = floatformat

def generate_documents(root_dir, main_path, template_path, output_dir,
                       common_column, file_name_column, log_callback):
    try:
        log_callback("=== Старт генерації DOCX ===")

        if not all([os.path.exists(main_path), os.path.exists(template_path), os.path.isdir(root_dir)]):
            log_callback("❌ Помилка: Перевірте всі шляхи до файлів!")
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

        log_callback(f"Знайдено {len(main_df)} записів. Запущено потоки...")

        def generate_docx(borrower):
            context = {}
            for col in main_df.columns:
                val = borrower[col]
                key = f"{col}_credit"
                context[key] = format_date(val) if isinstance(val, datetime) else (val if pd.notnull(val) else "—")

            for tablename, df in other_tables.items():
                if common_column not in df.columns:
                    continue
                filtered = df[df[common_column] == borrower[common_column]]
                rows = []
                for _, row in filtered.iterrows():
                    row_dict = {}
                    for col in df.columns:
                        val = row[col]
                        row_dict[col] = format_date(val) if isinstance(val, datetime) else (val if pd.notnull(val) else "—")
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
                log_callback(f"[{i}/{len(futures)}] Згенеровано: {os.path.basename(fname)}")

        log_callback(f"\n✅ Успішно створено {len(created_docx_files)} DOCX документів у {output_dir}\n")

    except Exception as e:
        log_callback("❌ ПОМИЛКА: " + str(e))
