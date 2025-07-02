import pandas as pd
from datetime import datetime
import re


def is_date_string(val):
    """
    Перевіряє, чи схожий рядок на дату.
    Повертає True тільки для рядків, які дійсно схожі на дати.
    """
    if not isinstance(val, str) or len(val.strip()) == 0:
        return False

    val = val.strip()

    # Якщо рядок містить тільки цифри та букви без типових роздільників дат - це не дата
    if re.match(r'^[A-Za-z0-9]+$', val) and not re.search(r'[\-\.\/:\ ]', val):
        return False

    # Якщо рядок виглядає як код (багато букв та цифр без роздільників) - не дата
    if len(val) > 6 and re.match(r'^[A-Z0-9]+$', val):
        return False

    # Якщо містить більше ніж 4 літери підряд - ймовірно не дата
    if re.search(r'[A-Za-z]{5,}', val):
        return False

    # Типові паттерни дат
    date_patterns = [
        r'^\d{1,2}[\.\-\/]\d{1,2}[\.\-\/]\d{2,4}$',  # DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
        r'^\d{4}[\.\-\/]\d{1,2}[\.\-\/]\d{1,2}$',  # YYYY.MM.DD, YYYY/MM/DD, YYYY-MM-DD
        r'^\d{1,2}[\.\-\/]\d{1,2}[\.\-\/]\d{2,4}\s+\d{1,2}:\d{2}',  # DD.MM.YYYY HH:MM
        r'^\d{4}\-\d{2}\-\d{2}T\d{2}:\d{2}:\d{2}',  # ISO format
        r'^\d{1,2}\s+(січ|лют|бер|кві|тра|чер|лип|сер|вер|жов|лис|гру)',  # Українські місяці
        r'^\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)',  # Англійські місяці (скорочені)
        r'^\d{1,2}\s+(january|february|march|april|may|june|july|august|september|october|november|december)',
        # Повні назви
    ]

    for pattern in date_patterns:
        if re.match(pattern, val, re.IGNORECASE):
            return True

    return False


def format_date(val):
    """Форматує дату в читабельний вигляд (тільки дата)"""
    if pd.isnull(val):
        return '—'
    if isinstance(val, datetime):
        return val.strftime('%d.%m.%Y')
    return str(val)


def format_datetime(val):
    """Форматує дату з часом в читабельний вигляд"""
    if pd.isnull(val):
        return '—'
    if isinstance(val, datetime):
        if val.time() == datetime.min.time():
            # Якщо час 00:00:00, показуємо тільки дату
            return val.strftime('%d.%m.%Y')
        elif val.second == 0:
            # Якщо секунди = 0, показуємо без секунд
            return val.strftime('%d.%m.%Y %H:%M')
        else:
            # Повний формат з секундами
            return val.strftime('%d.%m.%Y %H:%M:%S')
    return str(val)


def dateformat(val, format_type="date"):
    """
    Універсальний фільтр для форматування дат.
    format_type: "date" (тільки дата) або "datetime" (дата + час)

    Використання в шаблоні:
    {{ my_date|dateformat:"date" }} - тільки дата
    {{ my_date|dateformat:"datetime" }} - дата + час
    """
    if pd.isnull(val):
        return '—'

    # Спробуємо конвертувати в datetime якщо це не так
    if isinstance(val, str):
        try:
            if is_date_string(val):
                val = pd.to_datetime(val, errors='coerce')
                if pd.isna(val):
                    return str(val)
            else:
                return str(val)
        except:
            return str(val)

    if isinstance(val, datetime):
        if format_type.lower() == "datetime":
            return format_datetime(val)
        else:
            return format_date(val)

    return str(val)


def floatformat(val, precision=2):
    """Форматує число з заданою кількістю знаків після коми"""
    try:
        precision = int(precision)
        return f"{float(val):.{precision}f}".replace('.', ',')
    except Exception:
        return val


def numberformat(val, thousands_sep=True, decimal_places=2):
    """
    Форматує число з розділювачами тисяч та десятковими знаками.

    Використання в шаблоні:
    {{ my_number|numberformat }} - з пробілами для тисяч, 2 знаки після коми
    {{ my_number|numberformat:False:0 }} - без пробілів, без десяткових
    {{ my_number|numberformat:True:3 }} - з пробілами, 3 знаки після коми
    """
    try:
        # Конвертуємо в число
        if isinstance(val, str):
            # Очищаємо рядок від можливих розділювачів
            val = val.replace(' ', '').replace(',', '.')

        num = float(val)
        decimal_places = int(decimal_places)

        # Форматуємо число
        if decimal_places == 0:
            formatted = f"{num:.0f}"
        else:
            formatted = f"{num:.{decimal_places}f}"

        # Заміняємо крапку на кому для десяткових
        formatted = formatted.replace('.', ',')

        # Додаємо розділювачі тисяч (пробіли)
        if thousands_sep and str(thousands_sep).lower() != "false":
            parts = formatted.split(',')
            integer_part = parts[0]
            decimal_part = parts[1] if len(parts) > 1 else ''

            # Додаємо пробіли для тисяч (справа наліво)
            integer_with_spaces = ''
            for i, digit in enumerate(reversed(integer_part)):
                if i > 0 and i % 3 == 0:
                    integer_with_spaces = ' ' + integer_with_spaces
                integer_with_spaces = digit + integer_with_spaces

            if decimal_part:
                formatted = integer_with_spaces + ',' + decimal_part
            else:
                formatted = integer_with_spaces

        return formatted

    except Exception:
        return val


def currencyformat(val, currency="₴", decimal_places=2):
    """
    Форматує число як валюту.

    Використання в шаблоні:
    {{ amount|currencyformat }} - з гривнею
    {{ amount|currencyformat:"$" }} - з доларом
    {{ amount|currencyformat:"€":0 }} - з євро, без копійок
    """
    try:
        formatted_number = numberformat(val, True, decimal_places)
        return f"{formatted_number} {currency}"
    except Exception:
        return val