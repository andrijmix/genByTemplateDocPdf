# main.py
import sys
import multiprocessing

# Імпортуємо новий UI
from ui import main as winforms_main


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

    # Запускаємо нову Windows Forms UI
    try:
        winforms_main()
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print(f"Помилка запуску: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()