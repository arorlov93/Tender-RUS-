"""Повторная вёрстка готовых xlsx (одинаковая высота строк, перенос, печать по ширине)."""
import sys, openpyxl
from tidy import tidy
for f in sys.argv[1:]:
    wb=openpyxl.load_workbook(f); tidy(wb); wb.save(f); print("ok",f)
