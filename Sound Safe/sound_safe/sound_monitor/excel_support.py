# excel_support.py

_OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, Alignment, numbers
    _OPENPYXL_AVAILABLE = True
except Exception:
    _OPENPYXL_AVAILABLE = False
    Workbook = None
    get_column_letter = None
    Font = None
    Alignment = None
    numbers = None
