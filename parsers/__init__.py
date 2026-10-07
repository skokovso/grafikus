# parsers/__init__.py
from .base import read_excel_or_csv, detect_department, get_dept_abbr, get_dept_order
from .standard import parse_standard_graph, parse_responsible_graph
from .oar import parse_oar_graph
from .surgery import parse_surgery_pdf, parse_surgery_txt
from .multisheet import parse_excel_with_months
from .administration import parse_administration_graph  # <-- ДОБАВЛЕНО