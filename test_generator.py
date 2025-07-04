# test_generator.py - Простий модуль для тестування генерації документів
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
from datetime import datetime


class DocumentTestRunner:
    """Клас для запуску тестів генерації документів"""

    def __init__(self, test_folder="test", log_callback=None):
        self.test_folder = Path(test_folder)
        self.log_callback = log_callback or print

    def log(self, message):
        """Логування повідомлень"""
        self.log_callback(message)

    def run_tests(self):
        """Запускає тест: генерує документ і порівнює з еталоном"""
        self.log("🧪 Початок тестування системи...")

        # Перевірка наявності тестових файлів
        if not self._check_test_files():
            return False

        # Створюємо main.xlsx якщо його немає
        main_file = self.test_folder / "main.xlsx"
        if not main_file.exists():
            self._create_test_main_xlsx()

        # Генеруємо тестовий документ
        generated_file = self._generate_test_document()
        if not generated_file:
            return False

        # Зберігаємо згенерований файл в папку test для перевірки
        test_generated = self.test_folder / "generated_test.docx"
        import shutil
        shutil.copy(generated_file, test_generated)
        self.log(f"💾 Згенерований файл збережено: {test_generated}")

        # Порівняння з еталонним файлом
        success = self._compare_documents(generated_file)

        # Очищення тимчасового файлу
        try:
            os.unlink(generated_file)
        except:
            pass

        if success:
            self.log("✅ Всі тести пройшли успішно! Система готова до роботи.")
        else:
            self.log("❌ Тест провалений! Перевірте файли test/generated_test.docx та test/reference.docx")

        return success

    def _check_test_files(self):
        """Перевіряє наявність необхідних тестових файлів"""
        required_files = [
            self.test_folder / "template.docx",
            self.test_folder / "reference.docx"
        ]

        for file_path in required_files:
            if not file_path.exists():
                self.log(f"❌ Тестовий файл не знайдено: {file_path}")
                return False

        self.log("✓ Всі тестові файли знайдено")
        return True

    def _create_test_main_xlsx(self):
        """Створює тестовий main.xlsx файл з еталонними даними"""
        try:
            self.log("📝 Створюємо тестовий main.xlsx...")

            # Еталонні дані з reference.docx
            test_data = {
                'НАЗВА_СУДУ': ["Святошинський районний суд м. Києва"],
                'ЄДРПОУ_СУДУ': ["02896733"],
                'ІПН_БОРЖНИКА': ["3663108343"],
                'ДАТА_НАРОДЖЕННЯ_БОРЖНИКА': [datetime(2000, 4, 16)],
                'ДАТА_ЗАРАХУВАННЯ_ВІД': [datetime(2021, 12, 15, 14, 51)],
                'СУМА_ЗАРАХУВАННЯ': [16500.00],
                'id': [1]
            }

            df = pd.DataFrame(test_data)

            # Зберігаємо в Excel з правильним форматуванням
            main_file = self.test_folder / "main.xlsx"
            with pd.ExcelWriter(main_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

                # Форматуємо текстові поля
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # ЄДРПОУ та ІПН як текст
                for row in range(2, len(df) + 2):
                    worksheet[f'B{row}'].number_format = '@'  # ЄДРПОУ_СУДУ
                    worksheet[f'C{row}'].number_format = '@'  # ІПН_БОРЖНИКА

            self.log("✅ Створено test/main.xlsx")

        except Exception as e:
            self.log(f"⚠️ Помилка створення main.xlsx: {e}")

    def _generate_test_document(self):
        """Генерує тестовий документ використовуючи логіку основної програми"""
        try:
            self.log("📝 Генерація тестового документу...")

            template_file = self.test_folder / "template.docx"
            main_file = self.test_folder / "main.xlsx"

            # Читаємо дані
            from generator import smart_read_excel
            df = smart_read_excel(str(main_file), self.log)

            if len(df) == 0:
                self.log("❌ Тестовий Excel файл порожній")
                return None

            # Беремо перший рядок
            test_row = df.iloc[0].to_dict()

            # Використовуємо функцію з основної програми
            from generator import process_single_document

            # Підготовка даних
            other_tables = {}
            main_columns = df.columns.tolist()

            # Конвертуємо datetime для серіалізації
            for key, val in test_row.items():
                if isinstance(val, (pd.Timestamp, datetime)):
                    test_row[key] = val.isoformat()
                elif pd.isna(val):
                    test_row[key] = None

            # Створюємо тимчасову папку
            temp_dir = tempfile.mkdtemp()

            args = (
                (0, test_row),
                str(template_file),
                temp_dir,
                "id",
                "id",
                other_tables,
                main_columns
            )

            # Генеруємо
            result = process_single_document(args)

            if result["success"]:
                self.log("✓ Тестовий документ згенеровано")
                return result["filename"]
            else:
                self.log(f"❌ Помилка генерації: {result['error']}")
                return None

        except Exception as e:
            self.log(f"❌ Помилка генерації: {e}")
            import traceback
            self.log(f"Деталі: {traceback.format_exc()}")
            return None

    def _compare_documents(self, generated_file):
        """Порівнює згенерований документ з еталонним"""
        reference_file = self.test_folder / "reference.docx"

        self.log("🔍 Порівняння з еталонним документом...")

        try:
            # Витягуємо текст з обох документів
            generated_text = self._extract_text_from_docx(generated_file)
            reference_text = self._extract_text_from_docx(str(reference_file))

            # Нормалізуємо тексти
            generated_normalized = self._normalize_text(generated_text)
            reference_normalized = self._normalize_text(reference_text)

            # Перевіряємо ключові елементи з reference.docx
            key_elements = [
                "Святошинський районний суд м. Києва",
                "02896733",
                "3663108343",
                "16.04.2000",
                "15.12.2021 14:51",
                "16 500,00"
            ]

            missing_elements = []
            for element in key_elements:
                if element not in generated_normalized:
                    missing_elements.append(element)

            if not missing_elements:
                self.log("✅ Всі ключові елементи присутні в документі!")
                self.log("🎯 Документ згенеровано правильно!")
                return True
            else:
                self.log("❌ Відсутні ключові елементи:")
                for element in missing_elements:
                    self.log(f"   • {element}")

                # Зберігаємо для аналізу
                self._save_debug_texts(reference_normalized, generated_normalized)
                return False

        except Exception as e:
            self.log(f"❌ Помилка порівняння: {e}")
            return False

    def _save_debug_texts(self, reference, generated):
        """Зберігає тексти для детального аналізу"""
        debug_folder = Path("debug")
        debug_folder.mkdir(exist_ok=True)

        try:
            with open(debug_folder / "reference_text.txt", "w", encoding="utf-8") as f:
                f.write(reference)
            with open(debug_folder / "generated_text.txt", "w", encoding="utf-8") as f:
                f.write(generated)
            self.log(f"💾 Тексти збережено в папку debug/ для аналізу")
        except Exception as e:
            self.log(f"⚠️ Не вдалося зберегти debug файли: {e}")

    def _extract_text_from_docx(self, file_path):
        """Витягує текст з DOCX файлу"""
        text_content = []

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            try:
                xml_content = docx_zip.read('word/document.xml')
                root = ET.fromstring(xml_content)

                # Пошук всіх текстових елементів
                for elem in root.iter():
                    if elem.tag.endswith('}t'):
                        if elem.text:
                            text_content.append(elem.text)

            except Exception as e:
                self.log(f"⚠️ Помилка читання XML: {e}")
                return ""

        return ' '.join(text_content)

    def _normalize_text(self, text):
        """Нормалізує текст для порівняння"""
        if not text:
            return ""

        # Видаляємо зайві пробіли
        normalized = ' '.join(text.split())

        # Заміняємо спеціальні символи
        normalized = normalized.replace('\u00a0', ' ')  # Нерозривний пробіл
        normalized = normalized.replace('\u2010', '-')  # Дефіси
        normalized = normalized.replace('\u2011', '-')
        normalized = normalized.replace('\u2012', '-')
        normalized = normalized.replace('\u2013', '-')
        normalized = normalized.replace('\u2014', '-')

        return normalized.strip()


def run_integration_test(log_callback=None):
    """Функція для запуску тесту"""
    test_runner = DocumentTestRunner(log_callback=log_callback)
    return test_runner.run_tests()


if __name__ == "__main__":
    success = run_integration_test()
    if success:
        print("🎉 Тести пройшли успішно!")
    else:
        print("💥 Тести провалені!")
        exit(1)