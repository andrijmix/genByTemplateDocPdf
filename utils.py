import pandas as pd
from datetime import datetime

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

def floatformat(val, precision=2):
    try:
        precision = int(precision)
        return f"{float(val):.{precision}f}".replace('.', ',')
    except Exception:
        return val
