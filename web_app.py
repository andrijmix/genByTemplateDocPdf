# web_app.py
import os
from flask import Flask, render_template, request, send_file, redirect, url_for, flash,abort
from werkzeug.utils import secure_filename
import shutil
import tempfile
from generator import generate_documents
import threading
import glob
import shutil
import time

app = Flask(__name__)
app.secret_key = "some_secret_key"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def clean_old_temp_dirs(base_folder="uploads", minutes=30):
    """Видаляє всі тимчасові папки, старші за вказаний час (default: 30 хвилин)"""
    now = time.time()
    for dir in glob.glob(os.path.join(base_folder, "*")):
        if os.path.isdir(dir):
            mtime = os.path.getmtime(dir)
            if now - mtime > minutes * 60:
                try:
                    shutil.rmtree(dir)
                except Exception as e:
                    print(f"Не вдалося видалити {dir}: {e}")

@app.route("/", methods=["GET", "POST"])
def index():
    clean_old_temp_dirs()
    if request.method == "POST":
        # Читаємо поля форми
        main_file = request.files.get("main_file")
        template_file = request.files.get("template_file")
        root_zip = request.files.get("root_zip")
        output_dir = tempfile.mkdtemp(dir=UPLOAD_FOLDER)
        common_column = request.form.get("common_column", "id")
        file_name_column = request.form.get("file_name_column", "id")
        logs = []

        # Валідація
        if not main_file or not template_file or not root_zip:
            flash("Завантаж всі потрібні файли!", "danger")
            return redirect(request.url)

        # Зберігаємо файли
        main_path = os.path.join(output_dir, secure_filename(main_file.filename))
        template_path = os.path.join(output_dir, secure_filename(template_file.filename))
        root_dir = os.path.join(output_dir, "tables")
        os.makedirs(root_dir, exist_ok=True)

        main_file.save(main_path)
        template_file.save(template_path)

        # Розпаковуємо додаткові таблиці
        import zipfile
        with zipfile.ZipFile(root_zip, "r") as zip_ref:
            zip_ref.extractall(root_dir)

        # Запуск генерації
        def log_callback(msg):
            logs.append(msg)

        def stop_flag():
            return False

        output_docs_dir = os.path.join(output_dir, "docs")
        os.makedirs(output_docs_dir, exist_ok=True)

        generate_documents(
            root_dir=root_dir,
            main_path=main_path,
            template_path=template_path,
            output_dir=output_docs_dir,
            common_column=common_column,
            file_name_column=file_name_column,
            log_callback=log_callback,
            stop_flag=stop_flag,
        )

        # Пакуємо результати
        result_zip = os.path.join(output_dir, "results.zip")
        shutil.make_archive(os.path.splitext(result_zip)[0], 'zip', output_docs_dir)

        # Повертаємо на сторінку з лінком
        return render_template(
            "index.html",
            logs="\n".join(logs),
            download_link=url_for("download_file", temp_path=result_zip)
        )

    return render_template("index.html", logs=None, download_link=None)


@app.route('/download/')
def download_file():
    temp_path = request.args.get("temp_path")
    if not temp_path or not os.path.exists(temp_path):
        abort(404)

    # Відправляємо файл і одразу ставимо на видалення у фоновому потоці
    def remove_file_delayed(path):
        import time
        time.sleep(5)  # Дати Flask завершити відправку
        # Видаляємо всю тимчасову папку (можеш залишити тільки файл, якщо хочеш)
        dir_path = os.path.dirname(path)
        import shutil
        try:
            shutil.rmtree(dir_path)
        except Exception as e:
            print(f"Не вдалося видалити {dir_path}: {e}")

    threading.Thread(target=remove_file_delayed, args=(temp_path,)).start()
    return send_file(temp_path, as_attachment=True)

@app.route("/faq")
def faq():
    faqs = [
        {
            "q": "Які файли потрібні для генерації документів?",
            "a": """
            <ul>
              <li><b>main.xlsx</b> – основна таблиця (кожен рядок = один документ)</li>
              <li>Інші .xlsx – додаткові таблиці (деталі, списки)</li>
              <li><b>template.docx</b> – Word-шаблон із змінними Jinja2</li>
              <li><i>config.yaml</i> – (необовʼязково) для швидкого запуску</li>
            </ul>
            """
        },
        {
            "q": "Як підготувати шаблон Word?",
            "a": """
            <ul>
                <li>Використовуй змінні у форматі <code>{{ name_credit }}</code>, <code>{{ amount_credit|currency_uah }}</code> тощо.</li>
                <li>Для таблиць – <code>{{ transactions_table }}</code> або блок <code>{% for item in transactions_table %}</code>...</li>
                <li>Доступні фільтри для дати, суми, валюти.</li>
            </ul>
            <pre>
    Документ для: {{ name_credit }}
    Сума: {{ amount_credit|currency_uah }}
    Дата народження: {{ birth_date_credit|dateonly }}
            </pre>
            """
        },
        {
            "q": "Яка структура основного Excel-файлу?",
            "a": """
            <p>Кожен рядок – це окремий документ.<br>
            Стовпці мають співпадати із змінними у шаблоні.</p>
            <pre>
    | id | name         | amount   | birth_date  |
    |----|--------------|----------|-------------|
    | 1  | Іван Петров  | 50000.50 | 1985-03-15  |
            </pre>
            """
        },
        {
            "q": "Які є фільтри для форматування у шаблоні?",
            "a": """
            <table class="table table-sm table-bordered">
            <tr><th>Фільтр</th><th>Опис</th><th>Приклад</th></tr>
            <tr><td>|dateonly</td><td>Тільки дата</td><td>15.03.2024</td></tr>
            <tr><td>|datetime_full</td><td>Дата+час</td><td>15.03.2024 14:30:45</td></tr>
            <tr><td>|number_thousands</td><td>З пробілами</td><td>1 234 567,89</td></tr>
            <tr><td>|currency_uah</td><td>Гривні</td><td>50 000,00 ₴</td></tr>
            <tr><td>|currency_usd</td><td>Долари</td><td>1 234,50 $</td></tr>
            <tr><td>|floatformat:2</td><td>Кількість знаків</td><td>1234,56</td></tr>
            </table>
            """,
            "is_open": True
        },
        {
            "q": "Що робити, якщо виникає помилка або документ не створюється?",
            "a": """
            <ul>
                <li>Перевір чи всі файли мають правильний формат і структуру</li>
                <li>Переконайся, що назви змінних у шаблоні співпадають зі стовпцями в Excel</li>
                <li>Дивись лог виконання — він підкаже, де саме помилка</li>
                <li>Спробуй протестувати шаблон на малих даних</li>
            </ul>
            """
        },

    ]

    return render_template("faq.html", faqs=faqs)

if __name__ == '__main__':
    app.run(debug=True, port=8080)
