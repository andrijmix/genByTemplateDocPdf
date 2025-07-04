# build_exe.py - Скрипт для створення EXE файлу
import subprocess
import sys
import os


def install_requirements():
    """Встановлює необхідні пакети"""
    print("📦 Встановлення залежностей...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])


def build_exe():
    """Створює EXE файл з оптимізацією для багатопоточності"""
    print("🔨 Створення EXE файлу...")

    # Команда для PyInstaller з оптимізаціями для багатопроцесорності
    cmd = [
        "pyinstaller",
        "--onefile",  # Один файл
        "--windowed",  # Без консолі (тільки GUI)
        "--name=DocxGenerator",  # Ім'я файлу
        "--icon=icon.ico",  # Іконка (якщо є)
        "--add-data=requirements.txt;.",  # Додаємо requirements.txt
        "--hidden-import=pandas",  # Явно вказуємо модулі
        "--hidden-import=docxtpl",
        "--hidden-import=jinja2",
        "--hidden-import=openpyxl",
        "--hidden-import=concurrent.futures",
        "--hidden-import=multiprocessing",
        "--hidden-import=multiprocessing.spawn",  # Для ProcessPoolExecutor
        "--hidden-import=pickle",  # Для серіалізації між процесами
        "--hidden-import=utils",  # Наш модуль utils
        "--optimize=2",  # Максимальна оптимізація
        "--strip",  # Видаляємо зайві символи
        "--noupx",  # Відключаємо UPX (може конфліктувати з багатопроцесорністю)
        "main.py"
    ]

    # Видаляємо іконку з команди, якщо файл не існує
    if not os.path.exists("icon.ico"):
        cmd.remove("--icon=icon.ico")

    try:
        subprocess.check_call(cmd)
        print("✅ EXE файл успішно створено!")
        print("📁 Знайдіть файл DocxGenerator.exe в папці dist/")

        # Створюємо довідкову інформацію
        create_readme()

    except subprocess.CalledProcessError as e:
        print(f"❌ Помилка при створенні EXE: {e}")
        return False

    return True


def create_readme():
    """Створює README файл для користувача"""
    readme_content = """
# DOCX Generator - Інструкція по використанню

## Системні вимоги
- Windows 7/8/10/11 (64-bit)
- Мінімум 4 ГБ RAM
- Для оптимальної роботи: багатоядерний процесор

## Як використовувати

1. **Підготуйте файли:**
   - main.xlsx - основна таблиця (кожен рядок = один документ)
   - Додаткові .xlsx файли - таблиці з деталями
   - template.docx - Word шаблон з Jinja2 змінними

2. **Запустіть DocxGenerator.exe**

3. **Налаштуйте параметри:**
   - Оберіть папку з Excel файлами
   - Вкажіть основний Excel файл (main.xlsx)
   - Оберіть Word шаблон (template.docx)
   - Вкажіть папку для збереження
   - Налаштуйте назви стовпців

4. **Натисніть "Старт"**

## Оптимізація продуктивності

Програма автоматично:
- Використовує 50% ядер процесора для максимальної швидкості
- Показує прогрес обробки в реальному часі
- Дозволяє зупинити процес кнопкою "Стоп"

## Приклад змінних у Word шаблоні

```
{{ name_credit }} - значення зі стовпця "name" основної таблиці
{{ transactions_table }} - таблиця з файлу "transactions.xlsx"
{{ amount_credit|floatformat:2 }} - число з 2 знаками після коми
```

## Підтримка
Якщо виникли проблеми - перевірте:
1. Всі Excel файли відкриваються в Excel
2. Word шаблон містить правильні змінні Jinja2
3. Назви стовпців співпадають у всіх файлах
"""

    with open("README_USER.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)

    print("📄 Створено файл README_USER.txt з інструкціями")


if __name__ == "__main__":
    print("🚀 Початок створення EXE файлу...")
    print("⚡ Налаштовано багатопоточність для максимальної швидкості")

    try:
        install_requirements()
        if build_exe():
            print("\n🎉 Готово! Ваш EXE файл готовий до використання.")
            print("💡 Програма автоматично використає половину ядер процесора")
        else:
            print("\n❌ Виникли проблеми при створенні EXE")
    except KeyboardInterrupt:
        print("\n⛔ Процес перерваний користувачем")
    except Exception as e:
        print(f"\n❌ Неочікувана помилка: {e}")