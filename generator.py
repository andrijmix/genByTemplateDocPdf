import glob
import multiprocessing
import os
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime

import jinja2
import pandas as pd
from docxtpl import DocxTemplate

from utils import format_date, floatformat, is_date_string


# Отримуємо оптимальну кількість процесів (половина від доступних ядер)
def get_optimal_workers():
    """Повертає оптимальну кількість робочих процесів (половина від CPU ядер)"""
    cpu_count = multiprocessing.cpu_count()
    # Використовуємо половину ядер, але мінімум 2 і максимум 8 для стабільності
    optimal_workers = max(2, min(90, cpu_count // 2))
    return optimal_workers


def process_single_document(args):
    """
    Функція для обробки одного документа в окремому процесі.
    Має бути на верхньому рівні модуля для pickle серіалізації.
    """
    try:
        (row_data, template_path, output_dir, common_column,
         file_name_column, other_tables, main_columns) = args

        index, borrower_dict = row_data

        # Створюємо Jinja2 середовище в кожному процесі
        jinja_env = jinja2.Environment()

        # Визначаємо фільтри прямо тут для уникнення проблем з імпортом
        def dateonly_filter(val):
            """Тільки дата без часу"""
            if pd.isnull(val):
                return '—'

            try:
                # Спочатку спробуємо конвертувати все в pandas datetime
                parsed_date = None

                # Якщо це вже datetime об'єкт
                if hasattr(val, 'strftime'):
                    parsed_date = val
                # Якщо це рядок
                elif isinstance(val, str):
                    if is_date_string(val):
                        parsed_date = pd.to_datetime(val, errors='coerce')
                    else:
                        return str(val)  # Не схоже на дату
                # Якщо це число (timestamp)
                elif isinstance(val, (int, float)):
                    parsed_date = pd.to_datetime(val, unit='s', errors='coerce')
                else:
                    # Спробуємо конвертувати будь-що інше
                    parsed_date = pd.to_datetime(val, errors='coerce')

                # Якщо успішно розпарсили дату
                if parsed_date is not None and pd.notna(parsed_date):
                    # Форматуємо ТІЛЬКИ дату без часу
                    if hasattr(parsed_date, 'date'):
                        return parsed_date.date().strftime('%d.%m.%Y')
                    else:
                        return parsed_date.strftime('%d.%m.%Y')

                return str(val)

            except Exception as e:
                return str(val)

        def datetime_full_filter(val):
            """Дата з часом"""
            if pd.isnull(val):
                return '—'

            try:
                # Pandas Timestamp або будь-який об'єкт з методом strftime
                if hasattr(val, 'strftime'):
                    # Перевіряємо чи є час
                    if hasattr(val, 'time') and val.time() != datetime.min.time():
                        return val.strftime('%d.%m.%Y %H:%M:%S')
                    else:
                        return val.strftime('%d.%m.%Y')

                # Рядок дати
                if isinstance(val, str):
                    if is_date_string(val):
                        parsed_date = pd.to_datetime(val, errors='coerce')
                        if pd.notna(parsed_date):
                            if parsed_date.time() != datetime.min.time():
                                return parsed_date.strftime('%d.%m.%Y %H:%M:%S')
                            else:
                                return parsed_date.strftime('%d.%m.%Y')
                    return str(val)

                return str(val)

            except Exception as e:
                return str(val)

        def datetime_full_no_sec_filter(val):
            """Дата з часом"""
            if pd.isnull(val):
                return '—'

            try:
                # Pandas Timestamp або будь-який об'єкт з методом strftime
                if hasattr(val, 'strftime'):
                    # Перевіряємо чи є час
                    if hasattr(val, 'time') and val.time() != datetime.min.time():
                        return val.strftime('%d.%m.%Y %H:%M')
                    else:
                        return val.strftime('%d.%m.%Y')

                # Рядок дати
                if isinstance(val, str):
                    if is_date_string(val):
                        parsed_date = pd.to_datetime(val, errors='coerce')
                        if pd.notna(parsed_date):
                            if parsed_date.time() != datetime.min.time():
                                return parsed_date.strftime('%d.%m.%Y %H:%M')
                            else:
                                return parsed_date.strftime('%d.%m.%Y')
                    return str(val)

                return str(val)

            except Exception as e:
                return str(val)

        def number_thousands_filter(val):
            """Число з тисячними розділювачами"""
            try:
                if isinstance(val, str):
                    val = val.replace(' ', '').replace(',', '.')
                num = float(val)
                formatted = f"{num:.2f}".replace('.', ',')
                parts = formatted.split(',')
                integer_part = parts[0]
                decimal_part = parts[1] if len(parts) > 1 else ''
                integer_with_spaces = ''
                for i, digit in enumerate(reversed(integer_part)):
                    if i > 0 and i % 3 == 0:
                        integer_with_spaces = ' ' + integer_with_spaces
                    integer_with_spaces = digit + integer_with_spaces
                if decimal_part:
                    return integer_with_spaces + ',' + decimal_part
                else:
                    return integer_with_spaces
            except Exception:
                return val

        def currency_uah_filter(val):
            """Гривні"""
            try:
                formatted = number_thousands_filter(val)
                return f"{formatted} ₴"
            except Exception:
                return val

        def currency_usd_filter(val):
            """Долари"""
            try:
                formatted = number_thousands_filter(val)
                return f"{formatted} $"
            except Exception:
                return val

        # Реєструємо фільтри
        jinja_env.filters['floatformat'] = floatformat
        jinja_env.filters['dateonly'] = dateonly_filter
        jinja_env.filters['datetime_full'] = datetime_full_filter
        jinja_env.filters['datetime_full_no_sec'] = datetime_full_no_sec_filter
        jinja_env.filters['number_thousands'] = number_thousands_filter
        jinja_env.filters['currency_uah'] = currency_uah_filter
        jinja_env.filters['currency_usd'] = currency_usd_filter

        # Додаткові фільтри для дат (синоніми)
        jinja_env.filters['date'] = dateonly_filter
        jinja_env.filters['dateformat'] = dateonly_filter

        # Підготовка контексту для шаблону
        context = {}

        # Додаємо дані з основної таблиці
        for col in main_columns:
            val = borrower_dict.get(col)
            key = f"{col}_credit"

            # Спеціальна обробка для дат
            if val is not None and not pd.isnull(val):
                # Якщо це datetime об'єкт, зберігаємо його як є
                if isinstance(val, (pd.Timestamp, datetime)):
                    context[key] = val
                # Якщо це рядок дати
                elif isinstance(val, str) and val != "NaT" and is_date_string(val):
                    try:
                        parsed_date = pd.to_datetime(val, errors='coerce')
                        if pd.notna(parsed_date):
                            context[key] = parsed_date
                        else:
                            context[key] = val
                    except:
                        context[key] = val
                else:
                    context[key] = val
            else:
                context[key] = "—"

        # Додаємо дані з додаткових таблиць
        for tablename, df_dict in other_tables.items():
            # Відновлюємо DataFrame з словника
            df = pd.DataFrame(df_dict['data'])
            df.columns = df_dict['columns']

            if common_column not in df.columns:
                continue

            # Фільтруємо по спільному стовпцю
            borrower_id = borrower_dict.get(common_column)
            filtered = df[df[common_column] == borrower_id]

            rows = []
            for _, row in filtered.iterrows():
                row_dict = {}
                for col in df.columns:
                    val = row[col]

                    # Спеціальна обробка для дат
                    if val is not None and not pd.isnull(val):
                        # Якщо це datetime об'єкт, зберігаємо його як є
                        if isinstance(val, (pd.Timestamp, datetime)):
                            row_dict[col] = val
                        # Якщо це рядок дати
                        elif isinstance(val, str) and val != "NaT" and is_date_string(val):
                            try:
                                parsed_date = pd.to_datetime(val, errors='coerce')
                                if pd.notna(parsed_date):
                                    row_dict[col] = parsed_date
                                else:
                                    row_dict[col] = val
                            except:
                                row_dict[col] = val
                        else:
                            row_dict[col] = val
                    else:
                        row_dict[col] = "—"
                rows.append(row_dict)
            context[f"{tablename}_table"] = rows

        # Генерація документа
        tpl = DocxTemplate(template_path)
        tpl.render(context, jinja_env)

        # Створення імені файлу
        safe_name = str(borrower_dict.get(file_name_column, borrower_dict.get(common_column, f"doc_{index}"))).replace(
            " ", "_")
        # Видаляємо небезпечні символи з імені файлу
        safe_name = "".join(c for c in safe_name if c.isalnum() or c in ('-', '_', '.'))

        docx_filename = os.path.join(output_dir, f"doc_{safe_name}.docx")
        tpl.save(docx_filename)

        return {"success": True, "filename": docx_filename, "index": index}

    except Exception as e:
        return {"success": False, "error": str(e), "index": index}


def generate_documents(root_dir, main_path, template_path, output_dir,
                       common_column, file_name_column, log_callback, stop_flag):
    try:
        # Визначаємо кількість робочих процесів
        max_workers = get_optimal_workers()
        log_callback(f"=== Старт генерації DOCX ===")
        log_callback(f"💻 Використовуємо {max_workers} процесів з {multiprocessing.cpu_count()} доступних ядер")

        if not all([os.path.exists(main_path), os.path.exists(template_path), os.path.isdir(root_dir)]):
            log_callback("❌ Помилка: Перевірте всі шляхи до файлів!")
            return

        os.makedirs(output_dir, exist_ok=True)

        # Читаємо основну таблицю
        log_callback("📖 Читання основної таблиці...")
        # Читаємо з автоматичним парсингом дат
        main_df = pd.read_excel(main_path, parse_dates=True)
        main_df.columns = main_df.columns.str.strip()

        # Читаємо додаткові таблиці
        log_callback("📖 Читання додаткових таблиць...")
        all_xlsx = glob.glob(os.path.join(root_dir, "*.xlsx"))
        other_xlsx = [f for f in all_xlsx if os.path.abspath(f) != os.path.abspath(main_path)]
        other_tables = {}

        for fname in other_xlsx:
            if stop_flag():
                log_callback("⛔ Генерацію зупинено користувачем.")
                return

            name = os.path.splitext(os.path.basename(fname))[0].lower()
            # Читаємо з автоматичним парсингом дат
            df = pd.read_excel(fname, parse_dates=True)
            df.columns = df.columns.str.strip()

            # Конвертуємо DataFrame в формат, який можна серіалізувати
            other_tables[name] = {
                'data': df.to_dict('records'),
                'columns': df.columns.tolist()
            }
            log_callback(f"✓ Завантажено таблицю: {name} ({len(df)} записів)")

        log_callback(f"📊 Знайдено {len(main_df)} записів для обробки")
        log_callback(f"🚀 Запуск {max_workers} паралельних процесів...")

        # Підготовка даних для паралельної обробки
        main_columns = main_df.columns.tolist()

        # Конвертуємо рядки в словники для серіалізації
        tasks = []
        for i, (_, row) in enumerate(main_df.iterrows()):
            if stop_flag():
                break

            row_dict = row.to_dict()
            # Конвертуємо datetime об'єкти в рядки для серіалізації, але зберігаємо інформацію про тип
            for key, val in row_dict.items():
                if isinstance(val, (pd.Timestamp, datetime)):
                    row_dict[key] = val.isoformat()
                elif pd.isna(val):
                    row_dict[key] = None

            task_args = (
                (i, row_dict),
                template_path,
                output_dir,
                common_column,
                file_name_column,
                other_tables,
                main_columns
            )
            tasks.append(task_args)

        if stop_flag():
            log_callback("⛔ Генерацію зупинено користувачем.")
            return

        # Паралельна обробка з ProcessPoolExecutor
        created_docx_files = []
        failed_files = []
        start_time = time.time()

        # Використовуємо ProcessPoolExecutor для справжньої багатопроцесорності
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            # Запускаємо всі завдання
            future_to_index = {
                executor.submit(process_single_document, task): task[0][0]
                for task in tasks
            }

            # Обробляємо результати по мірі завершення
            completed_count = 0
            active_processes = len(future_to_index)

            log_callback(f"⚡ Активно обробляється {active_processes} завдань паралельно...")

            for future in as_completed(future_to_index):
                if stop_flag():
                    log_callback("⛔ Зупинка всіх процесів...")
                    # Скасовуємо всі незавершені завдання
                    for f in future_to_index:
                        f.cancel()
                    executor.shutdown(wait=False)
                    break

                try:
                    result = future.result(timeout=30)  # Таймаут 30 секунд на документ
                    completed_count += 1

                    if result["success"]:
                        created_docx_files.append(result["filename"])
                        elapsed = time.time() - start_time
                        speed = completed_count / elapsed if elapsed > 0 else 0
                        log_callback(
                            f"✅ [{completed_count}/{len(tasks)}] {os.path.basename(result['filename'])} | {speed:.1f} док/сек")
                    else:
                        error_msg = f"Рядок {result['index']}: {result['error']}"
                        failed_files.append(error_msg)
                        log_callback(f"❌ [{completed_count}/{len(tasks)}] {error_msg}")

                except Exception as e:
                    failed_files.append(f"Критична помилка процесу: {str(e)}")
                    log_callback(f"❌ Критична помилка процесу: {str(e)}")

        # Підсумки
        total_time = time.time() - start_time
        if not stop_flag():
            log_callback(f"\n🎉 Генерацію завершено за {total_time:.1f} секунд!")
            log_callback(f"✅ Успішно створено: {len(created_docx_files)} документів")
            log_callback(f"⚡ Середня швидкість: {len(created_docx_files) / total_time:.1f} документів/секунду")
            if failed_files:
                log_callback(f"❌ Помилок: {len(failed_files)}")
                for error in failed_files[:3]:  # Показуємо перші 3 помилки
                    log_callback(f"   • {error}")
                if len(failed_files) > 3:
                    log_callback(f"   • ... та ще {len(failed_files) - 3} помилок")
            log_callback(f"📁 Папка збереження: {output_dir}")
        else:
            log_callback(
                f"\n⛔ Генерацію зупинено за {total_time:.1f} сек. Створено: {len(created_docx_files)} документів")

    except Exception as e:
        log_callback(f"❌ КРИТИЧНА ПОМИЛКА: {str(e)}")
        import traceback
        log_callback(f"Деталі помилки: {traceback.format_exc()}")