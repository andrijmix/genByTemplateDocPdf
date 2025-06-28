import pandas as pd
import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import yaml
import jinja2

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

def format_number(val):
    if pd.isnull(val):
        return '—'
    if isinstance(val, (int, float)):
        return "{:,.2f}".format(val).replace(",", " ").replace(".", ",")
    return str(val)

# ==== floatformat filter for template ====
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
payments_path = config.get("payments_path", "payments.xlsx")
template_path = config.get("template_path", "template.docx")
output_dir = config.get("output_dir", "output_docs")
save_format = config.get("save_format", "both").lower()
common_column = config.get("common_column", "id")
file_name_column = config.get("file_name_column", "id")

# === 2. Налаштування форматів збереження ===
save_docx = save_format in ("docx", "both")
save_pdf = save_format in ("pdf", "both")

# === 3. Створення папок ===
os.makedirs(output_dir, exist_ok=True)
pdf_output_dir = os.path.join(output_dir, "pdfs")
if save_pdf:
    os.makedirs(pdf_output_dir, exist_ok=True)

# === 4. Читання Excel-файлів ===
credits_df = pd.read_excel(credits_path)
payments_df = pd.read_excel(payments_path)

# Очищення назв стовпців від пробілів
credits_df.columns = credits_df.columns.str.strip()
payments_df.columns = payments_df.columns.str.strip()

# Перевірка наявності спільного стовпця
if common_column not in credits_df.columns or common_column not in payments_df.columns:
    raise ValueError(f"Спільного стовпця '{common_column}' немає в одній із таблиць. Перевірте дані.")

# Злиття даних
merged_df = pd.merge(payments_df, credits_df, on=common_column, how="left", suffixes=('_payment', '_credit'))

# Групування
grouped = merged_df.groupby(common_column)

# === 5. Генерація всіх DOCX ===
created_docx_files = []

for borrower_id, group in grouped:
    borrower_info = group.iloc[0]

    # Створення контексту
    context = {}

    # Дані з кредитів
    for col in credits_df.columns:
        val = borrower_info[col]
        if isinstance(val, (int, float)):
            context[f"{col}_credit"] = val  # не форматуй тут, залиш сире число для фільтра!
        elif isinstance(val, datetime):
            context[f"{col}_credit"] = format_date(val)
        else:
            context[f"{col}_credit"] = val if pd.notnull(val) else "—"

    # Дані з платежів
    payment_rows = []
    payment_columns = [col for col in payments_df.columns if col != common_column]
    for _, row in group.iterrows():
        row_data = {}
        for col in payment_columns:
            val = row[col]

            if isinstance(val, (int, float)):
                row_data[f"{col}_payment"] = val  # залишаємо число, не форматуй!
            elif isinstance(val, datetime):
                row_data[f"{col}_payment"] = format_date(val)
            elif pd.isnull(val):
                row_data[f"{col}_payment"] = "—"
            else:
                row_data[f"{col}_payment"] = val

        payment_rows.append(row_data)

    context["payments_table"] = payment_rows

    # Завантаження шаблону
    tpl = DocxTemplate(template_path)
    tpl.render(context, jinja_env)

    # Формування імені файлу
    safe_name = str(borrower_info.get(file_name_column, borrower_id)).replace(" ", "_")
    docx_filename = os.path.join(output_dir, f"doc_{safe_name}.docx")

    # Збереження DOCX
    tpl.save(docx_filename)
    created_docx_files.append(docx_filename)

print(f"✅ Успішно створено {len(created_docx_files)} DOCX документів.")

# === 6. Масова конвертація у PDF ===
if save_pdf:
    print("\n📄 Починаємо конвертацію DOCX в PDF...")
    try:
        # convert(папка_з_DOCX, папка_куди_зберегти_PDF)
        convert(output_dir, pdf_output_dir)
        print(f"✅ Успішно конвертовано всі документи у папку: {pdf_output_dir}")
    except Exception as e:
        print(f"⚠️ Помилка при масовій конвертації у PDF: {e}")

print("\n🏁 Роботу завершено!")
