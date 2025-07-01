import sys
import os
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
                             QFileDialog, QGridLayout, QMessageBox, QProgressBar,
                             QGroupBox, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon
from generator import generate_documents


class LoggerThread(QThread):
    """Потік для обробки логування без блокування UI"""
    log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.messages = []
        self.running = True

    def add_message(self, message):
        self.messages.append(message)
        self.log_signal.emit(message)

    def stop(self):
        self.running = False


class GeneratorThread(QThread):
    """Потік для генерації документів"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, root_dir, main_file, template_file, output_dir,
                 common_column, file_name_column):
        super().__init__()
        self.root_dir = root_dir
        self.main_file = main_file
        self.template_file = template_file
        self.output_dir = output_dir
        self.common_column = common_column
        self.file_name_column = file_name_column
        self.stop_flag = False

    def run(self):
        try:
            generate_documents(
                root_dir=self.root_dir,
                main_path=self.main_file,
                template_path=self.template_file,
                output_dir=self.output_dir,
                common_column=self.common_column,
                file_name_column=self.file_name_column,
                log_callback=self.log_message,
                stop_flag=lambda: self.stop_flag
            )
        except Exception as e:
            self.log_message(f"❌ Критична помилка: {str(e)}")
        finally:
            self.finished_signal.emit()

    def log_message(self, message):
        self.log_signal.emit(message)

    def stop_generation(self):
        self.stop_flag = True


class ModernButton(QPushButton):
    """Проста стилізована кнопка"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 6px 12px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)


class ModernLineEdit(QLineEdit):
    """Просте поле вводу"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QLineEdit {
                border: 1px solid #cccccc;
                border-radius: 3px;
                padding: 6px 8px;
                font-size: 12px;
                background-color: white;
                min-height: 16px;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
            QLineEdit:hover {
                border: 1px solid #0078d4;
            }
        """)


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.generator_thread = None
        self.logger_thread = LoggerThread()
        self.logger_thread.log_signal.connect(self.log_write)
        self.logger_thread.start()

        self.init_ui()
        self.apply_modern_style()

    def init_ui(self):
        # Основне вікно
        self.setWindowTitle("DOCX Generator")
        self.setGeometry(100, 100, 900, 800)
        self.setMinimumSize(550, 450)

        # Центральний віджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # Заголовок
        title_label = QLabel("DOCX Generator")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #000000;
                margin-bottom: 20px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Група налаштувань файлів
        files_group = QGroupBox("Налаштування файлів")
        files_group.setStyleSheet("""
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                background-color: white;
            }
        """)
        files_layout = QGridLayout(files_group)
        files_layout.setSpacing(15)

        # Поля вводу
        self.root_dir = ModernLineEdit()
        self.root_dir.setPlaceholderText("Папка з Excel файлами...")
        self.main_file = ModernLineEdit()
        self.main_file.setPlaceholderText("Основний Excel файл...")
        self.template_file = ModernLineEdit()
        self.template_file.setPlaceholderText("Word шаблон...")
        self.output_dir = ModernLineEdit()
        self.output_dir.setText("output_docs")
        self.output_dir.setPlaceholderText("Папка збереження...")

        # Кнопки вибору файлів
        btn_root = ModernButton("Обрати")
        btn_root.clicked.connect(self.select_root_dir)

        btn_main = ModernButton("Обрати")
        btn_main.clicked.connect(self.select_main_file)

        btn_template = ModernButton("Обрати")
        btn_template.clicked.connect(self.select_template_file)

        btn_output = ModernButton("Обрати")
        btn_output.clicked.connect(self.select_output_dir)

        # Розміщення полів і кнопок
        files_layout.addWidget(QLabel("Папка з таблицями:"), 0, 0)
        files_layout.addWidget(self.root_dir, 0, 1)
        files_layout.addWidget(btn_root, 0, 2)

        files_layout.addWidget(QLabel("Основний Excel-файл:"), 1, 0)
        files_layout.addWidget(self.main_file, 1, 1)
        files_layout.addWidget(btn_main, 1, 2)

        files_layout.addWidget(QLabel("Шаблон DOCX:"), 2, 0)
        files_layout.addWidget(self.template_file, 2, 1)
        files_layout.addWidget(btn_template, 2, 2)

        files_layout.addWidget(QLabel("Папка збереження:"), 3, 0)
        files_layout.addWidget(self.output_dir, 3, 1)
        files_layout.addWidget(btn_output, 3, 2)

        main_layout.addWidget(files_group)

        # Група налаштувань стовпців
        columns_group = QGroupBox("Налаштування стовпців")
        columns_group.setStyleSheet(files_group.styleSheet())
        columns_layout = QGridLayout(columns_group)
        columns_layout.setSpacing(15)

        self.common_column = ModernLineEdit()
        self.common_column.setText("id")
        self.common_column.setPlaceholderText("Спільний стовпець...")

        self.file_name_column = ModernLineEdit()
        self.file_name_column.setText("id")
        self.file_name_column.setPlaceholderText("Стовпець для імені файлу...")

        columns_layout.addWidget(QLabel("Спільний стовпець:"), 0, 0)
        columns_layout.addWidget(self.common_column, 0, 1, 1, 2)

        columns_layout.addWidget(QLabel("Стовпець для імені файлу:"), 1, 0)
        columns_layout.addWidget(self.file_name_column, 1, 1, 1, 2)

        main_layout.addWidget(columns_group)

        # Кнопки управління
        control_layout = QHBoxLayout()
        control_layout.setSpacing(15)

        self.start_btn = ModernButton("Початок генерації")
        self.start_btn.setStyleSheet(self.start_btn.styleSheet() + """
            QPushButton {
                background-color: #4CAF50;
                font-size: 14px;
                font-weight: bold;
                min-height: 35px;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        self.start_btn.clicked.connect(self.generate)

        self.stop_btn = ModernButton("Зупинити")
        self.stop_btn.setStyleSheet(self.stop_btn.styleSheet() + """
            QPushButton {
                background-color: #f44336;
                min-height: 35px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:pressed {
                background-color: #c1171b;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_generation)
        self.stop_btn.setEnabled(False)

        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.stop_btn)
        control_layout.addStretch()

        main_layout.addLayout(control_layout)

        # Лог
        log_group = QGroupBox("Журнал виконання")
        log_group.setStyleSheet(files_group.styleSheet())
        log_layout = QVBoxLayout(log_group)

        self.log = QTextEdit()
        self.log.setStyleSheet("""
            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                background-color: #ffffff;
                color: #000000;
            }
        """)
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(200)

        log_layout.addWidget(self.log)
        main_layout.addWidget(log_group)

        # Прогрес бар (поки що прихований)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
                font-weight: bold;
                background-color: #f5f5f5;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 3px;
            }
        """)
        main_layout.addWidget(self.progress_bar)

    def apply_modern_style(self):
        """Застосовує простий чистий стиль"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
            QLabel {
                color: #000000;
                font-size: 12px;
                font-weight: normal;
            }
        """)

    def log_write(self, text):
        """Додає текст до логу"""
        self.log.append(text)
        self.log.ensureCursorVisible()

    def select_root_dir(self):
        dirname = QFileDialog.getExistingDirectory(self, "Оберіть папку з Excel файлами")
        if dirname:
            self.root_dir.setText(dirname)

    def select_main_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Оберіть основний Excel-файл", "", "Excel files (*.xlsx)")
        if filename:
            self.main_file.setText(filename)
            self.root_dir.setText(os.path.dirname(filename))

    def select_template_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Оберіть шаблон DOCX", "", "DOCX files (*.docx)")
        if filename:
            self.template_file.setText(filename)

    def select_output_dir(self):
        dirname = QFileDialog.getExistingDirectory(self, "Оберіть папку для збереження DOCX")
        if dirname:
            self.output_dir.setText(dirname)

    def generate(self):
        """Запускає генерацію документів"""
        # Перевірка заповнення полів
        if not all([self.root_dir.text(), self.main_file.text(),
                    self.template_file.text(), self.output_dir.text()]):
            QMessageBox.warning(self, "Увага",
                                "Будь ласка, заповніть всі обов'язкові поля!")
            return

        # Перевірка існування файлів
        if not os.path.exists(self.main_file.text()):
            QMessageBox.critical(self, "Помилка",
                                 "Основний Excel файл не знайдено!")
            return

        if not os.path.exists(self.template_file.text()):
            QMessageBox.critical(self, "Помилка",
                                 "Файл шаблону не знайдено!")
            return

        # Запуск генерації в окремому потоці
        self.generator_thread = GeneratorThread(
            self.root_dir.text(),
            self.main_file.text(),
            self.template_file.text(),
            self.output_dir.text(),
            self.common_column.text() or "id",
            self.file_name_column.text() or "id"
        )

        self.generator_thread.log_signal.connect(self.log_write)
        self.generator_thread.finished_signal.connect(self.generation_finished)

        # Блокуємо кнопку старту та активуємо стоп
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        # Очищаємо лог
        self.log.clear()

        # Запускаємо потік
        self.generator_thread.start()

    def stop_generation(self):
        """Зупиняє генерацію документів"""
        if self.generator_thread:
            self.generator_thread.stop_generation()
            self.log_write("⛔ Запит на зупинку надіслано...")

    def generation_finished(self):
        """Викликається після завершення генерації"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        # Показуємо повідомлення про завершення
        QMessageBox.information(self, "Готово",
                                "Генерація документів завершена!\n"
                                "Перевірте папку збереження.")

    def closeEvent(self, event):
        """Обробка закриття програми"""
        if self.generator_thread and self.generator_thread.isRunning():
            reply = QMessageBox.question(self, "Закрити програму",
                                         "Генерація ще виконується. Зупинити та закрити?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.generator_thread.stop_generation()
                self.generator_thread.wait(3000)  # Чекаємо до 3 секунд
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    """Головна функція для запуску програми з Windows Forms UI"""
    import multiprocessing

    # Ініціалізація багатопроцесорності
    multiprocessing.freeze_support()
    if hasattr(multiprocessing, 'set_start_method'):
        try:
            multiprocessing.set_start_method('spawn', force=True)
        except RuntimeError:
            pass

    # Створення додатку
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Сучасний стиль

    # Встановлення іконки (якщо є)
    try:
        app.setWindowIcon(QIcon('icon.ico'))
    except:
        pass

    # Налаштування шрифтів
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Створення вікна
    window = App()
    window.show()

    # Запуск головного циклу
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()