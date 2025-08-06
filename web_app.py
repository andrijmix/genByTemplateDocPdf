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

LOGS_FOLDER = "logs"
os.makedirs(LOGS_FOLDER, exist_ok=True)

sessions = {}  # session_id: {"log": log_path, "result": result_zip}
stop_flags = {}  # session_id: threading.Event
def background_generate(session_id, root_dir, main_path, template_path, output_dir, common_column, file_name_column):
    log_path = os.path.join(LOGS_FOLDER, f"{session_id}.log")

    def log_callback(msg):
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    def stop_flag():
        return stop_flags[session_id].is_set()

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
    sessions[session_id]["result"] = result_zip
    del stop_flags[session_id]


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
@app.route("/logs")
def get_logs():
    session_id = request.args.get("session_id")
    log_path = f"logs/{session_id}.log"
    if os.path.exists(log_path):
        with open(log_path, encoding='utf-8') as f:
            return f.read()
    return ""
@app.route("/", methods=["GET", "POST"])
def index():
    clean_old_temp_dirs()
    if request.method == "POST":
        import random
        session_id = str(int(time.time() * 1000)) + str(random.randint(100,999))
        output_dir = tempfile.mkdtemp(dir=UPLOAD_FOLDER)
        main_file = request.files.get("main_file")
        template_file = request.files.get("template_file")
        root_zip = request.files.get("root_zip")
        common_column = request.form.get("common_column", "id")
        file_name_column = request.form.get("file_name_column", "id")

        # Зберігаємо файли
        main_path = os.path.join(output_dir, secure_filename(main_file.filename))
        template_path = os.path.join(output_dir, secure_filename(template_file.filename))
        root_dir = os.path.join(output_dir, "tables")
        os.makedirs(root_dir, exist_ok=True)
        main_file.save(main_path)
        template_file.save(template_path)

        import zipfile
        with zipfile.ZipFile(root_zip, "r") as zip_ref:
            zip_ref.extractall(root_dir)

        # Сесія
        sessions[session_id] = {"log": os.path.join(LOGS_FOLDER, f"{session_id}.log"), "result": None}
        stop_flags[session_id] = threading.Event()
        # Генерація у фоні
        t = threading.Thread(target=background_generate, args=(session_id, root_dir, main_path, template_path, output_dir, common_column, file_name_column))
        t.start()

        return redirect(url_for("progress", session_id=session_id))
    return render_template("index.html", logs=None, download_link=None)
@app.route("/logs/<session_id>")
def logs(session_id):
    log_path = sessions.get(session_id, {}).get("log")
    if not log_path or not os.path.exists(log_path):
        return ""
    with open(log_path, encoding="utf-8") as f:
        return f.read()
@app.route("/stop/<session_id>", methods=["POST"])
def stop(session_id):
    if session_id in stop_flags:
        stop_flags[session_id].set()
        return "OK"
    return "Not found", 404


@app.route("/result/<session_id>")
def result(session_id):
    result_zip = sessions.get(session_id, {}).get("result")
    if not result_zip or not os.path.exists(result_zip):
        return "Not ready", 404
    return send_file(result_zip, as_attachment=True)
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
@app.route("/progress/<session_id>")
def progress(session_id):
    return render_template("progress.html", session_id=session_id)
if __name__ == '__main__':
    app.run(debug=True, port=8080)
