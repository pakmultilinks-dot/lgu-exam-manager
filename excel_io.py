"""Excel container selection; academic interpretation lives in the schema mapping."""
from pathlib import Path
SUPPORTED_EXTENSIONS=('.xlsx','.xlsm','.xls','.xlsb')
FILE_TYPES=[('Excel workbooks','*.xlsx *.xlsm *.xls *.xlsb')]

def excel_engine(path):
    suffix=Path(path).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:raise ValueError('Select an Excel workbook (.xlsx, .xlsm, .xls or .xlsb).')
    return 'openpyxl' if suffix in ('.xlsx','.xlsm') else 'calamine'
