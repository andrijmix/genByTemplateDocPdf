import pandas as pd
import os
import glob
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import jinja2

try:
    import yaml
    config_path = "config.yaml"
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            config = yaml.safe_load(f)
    else:
        config = None
except ImportError:
    config = None

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

# ==== 1. Отримання налаштувань ====
if config:
    root_dir = config.get("data_folder", ".")
    main_name = config.get("main_file", "main.xlsx")
    template_path = config.get("template_path", "template.docx")
    output_dir = config.get("output_dir", "output_docs")
    save_format = config.get("save_format", "both").lower()
    common_column = config.get("common_column", "id")
    file_name_column = config.get("file_name_column", "id")
else:
    print("Конфіг файл не знайдено! Введіть дані вручну:")
    root_dir = input("Вкажіть папку з таблицями (де main.xlsx): ").strip() or "."
    main_name = input("Вкажіть ім'я основної таблиці (main.xlsx): ").strip() or "main.xlsx"
    template_path = input("Вкажіть шлях до шаблону Word (template.docx): ").strip() or "template.docx"
    output_dir = input("Куди зберігати документи (output_docs): ").strip() or "output_docs"
    save_format = input("Формат збереження (docx/pdf/both): ").strip().lower() or "both"
    common_column = input("Назва спільного стовпця (id): ").strip() or "id"
    file_name_column = input("Назва стовпця для імені файлу (id): ").strip() or "id"

save_docx = save_format in ("docx", "both")
save_pdf = save_format in ("pdf", "both")

os.makedirs(output_dir, exist_ok=True)
pdf_output_dir = os.path.join(output_dir, "pdfs")
if save_pdf:
    os.makedirs(pdf_output_dir, exist_ok=True)

# ==== 2. Знаходимо main.xlsx і всі додаткові таблиці ====
main_path = os.path.join(root_dir, main_name)
if not os.path.exists(main_path):
    raise FileNotFoundError(f"Основний файл не знайдено: {main_path}")

main_df = pd.read_excel(main_path)
main_df.columns = main_df.columns.str.strip()

all_xlsx = glob.glob(os.path.join(root_dir, "*.xlsx"))
other_xlsx = [f for f in all_xlsx if os.path.abspath(f) != os.path.abspath(main_path)]

# ==== 3. Завантажуємо всі інші таблиці (phones, payments, address...) ====
other_tables = {}
for fname in other_xlsx:
    name = os.path.splitext(os.path.basename(fname))[0].lower()
    df = pd.read_excel(fname)
    df.columns = df.columns.str.strip()
    other_tables[name] = df

# ==== 4. Генерація DOCX ====
created_docx_files = []

for idx, borrower in main_df.iterrows():
    context = {}
    # Додаємо всі поля з main_df (можна форматувати тут як хочеш)
    for col in main_df.columns:
        val = borrower[col]
        context[f"{col}_credit"] = val if pd.notnull(val) else "—"

    # Для кожної додаткової таблиці робимо підбір по ключу (one-to-many)
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
