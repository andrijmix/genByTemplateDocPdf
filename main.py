# main.py
import tkinter as tk
import multiprocessing
from ui import App


def main():
    """Головна функція з правильною ініціалізацією багатопроцесорності"""
    # Необхідно для правильної роботи багатопроцесорності в EXE
    multiprocessing.freeze_support()

    # Встановлюємо метод запуску для Windows (важливо для ProcessPoolExecutor)
    if hasattr(multiprocessing, 'set_start_method'):
        try:
            # Використовуємо 'spawn' для кращої ізоляції процесів
            multiprocessing.set_start_method('spawn', force=True)
        except RuntimeError:
            pass  # Метод вже встановлено

    # Створюємо та запускаємо GUI
    root = tk.Tk()

    # Налаштування вікна
    root.geometry("600x500")
    root.resizable(True, True)

    # Встановлюємо заголовок з інформацією про потужність
    cpu_count = multiprocessing.cpu_count()
    optimal_workers = max(2, min(8, cpu_count // 2))
    root.title(f"DOCX Generator - {optimal_workers}/{cpu_count} потоків")

    # Створюємо додаток
    app = App(root)

    # Додаємо обробку закриття
    def on_closing():
        if hasattr(app, 'stop_flag'):
            app.stop_flag = True
        root.quit()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # Запускаємо основний цикл
    try:
        root.mainloop()
    except KeyboardInterrupt:
        on_closing()


if __name__ == "__main__":
    main()