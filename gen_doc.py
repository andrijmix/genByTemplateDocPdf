import pandas as pd
import os
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert

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

# === 1. Введення параметрів від користувача ===
credits_path = input("Введіть шлях до файлу кредитів (credits.xlsx): ").strip() or "credits.xlsx"
payments_path = input("Введіть шлях до файлу платежів (payments.xlsx): ").strip() or "payments.xlsx"
template_path = input("Введіть шлях до шаблону Word (template.docx): ").strip() or "template.docx"
output_dir = input("Введіть шлях для збереження результату (output_docs): ").strip() or "output_docs"

save_format = input("У якому форматі зберігати документи? (docx/pdf/both): ").strip().lower() or "both"
common_column = input("Введіть назву спільного стовпця для об'єднання таблиць (id): ").strip()
file_name_column = input("Введіть назву стовпця для імені файлу (КД або ПІБ): ").strip()

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
            context[f"{col}_credit"] = format_number(val)
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
                val = format_number(val)
            elif isinstance(val, datetime):
                val = format_date(val)
            elif pd.isnull(val):
                val = "—"

            row_data[f"{col}_payment"] = val

        payment_rows.append(row_data)

    context["payments_table"] = payment_rows

    # Завантаження шаблону
    tpl = DocxTemplate(template_path)
    tpl.render(context)

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
