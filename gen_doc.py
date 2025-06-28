import pandas as pd
import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import yaml
import jinja2
import glob

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

# === 1. Читання параметрів з config.yaml ===
with open("config.yaml", encoding="utf-8") as f:
    config = yaml.safe_load(f)

credits_path = config.get("credits_path", "credits.xlsx")
common_column = config.get("common_column", "id")
template_path = config.get("template_path", "template.docx")
output_dir = config.get("output_dir", "output_docs")
save_format = config.get("save_format", "both").lower()
file_name_column = config.get("file_name_column", "id")

save_docx = save_format in ("docx", "both")
save_pdf = save_format in ("pdf", "both")

os.makedirs(output_dir, exist_ok=True)
pdf_output_dir = os.path.join(output_dir, "pdfs")
if save_pdf:
    os.makedirs(pdf_output_dir, exist_ok=True)

# === 2. Читання credits.xlsx ===
credits_df = pd.read_excel(credits_path)
credits_df.columns = credits_df.columns.str.strip()

root_dir = os.path.dirname(os.path.abspath(credits_path))
all_xlsx = glob.glob(os.path.join(root_dir, "*.xlsx"))
other_xlsx = [f for f in all_xlsx if os.path.abspath(f) != os.path.abspath(credits_path)]

# === 3. Завантажуємо всі інші таблиці (phones, payments, address...) ===
other_tables = {}
for fname in other_xlsx:
    name = os.path.splitext(os.path.basename(fname))[0].lower()
    df = pd.read_excel(fname)
    df.columns = df.columns.str.strip()
    other_tables[name] = df

# === 4. Генерація DOCX ===
created_docx_files = []

for idx, borrower in credits_df.iterrows():
    context = {}

    # Додаємо всі поля з credits_df (можна форматувати тут як хочеш)
    for col in credits_df.columns:
        val = borrower[col]
        context[f"{col}_credit"] = val if pd.notnull(val) else "—"

    # Для кожної додаткової таблиці робимо підбір по ключу (одна-many, може бути 0..N рядків!)
    for tablename, df in other_tables.items():
        if common_column not in df.columns:
            continue  # Пропускаємо таблиці без спільного ключа

        filtered = df[df[common_column] == borrower[common_column]]

        # Формуємо список словників для шаблону
        rows = []
        for _, row in filtered.iterrows():
            row_dict = {}
            for col in df.columns:
                val = row[col]
                # Автоформатування: дати і числа
                if isinstance(val, datetime):
                    row_dict[col] = format_date(val)
                else:
                    row_dict[col] = val if pd.notnull(val) else "—"
            rows.append(row_dict)
        context[f"{tablename}_table"] = rows

    # Завантаження шаблону
    tpl = DocxTemplate(template_path)
    tpl.render(context, jinja_env)

    # Формування імені файлу
    safe_name = str(borrower.get(file_name_column, borrower[common_column])).replace(" ", "_")
    docx_filename = os.path.join(output_dir, f"doc_{safe_name}.docx")

    tpl.save(docx_filename)
    created_docx_files.append(docx_filename)

print(f"✅ Успішно створено {len(created_docx_files)} DOCX документів.")

if save_pdf:
    print("\n📄 Починаємо конвертацію DOCX в PDF...")
    try:
        convert(output_dir, pdf_output_dir)
        print(f"✅ Успішно конвертовано всі документи у папку: {pdf_output_dir}")
    except Exception as e:
        print(f"⚠️ Помилка при масовій конвертації у PDF: {e}")

print("\n🏁 Роботу завершено!")
