"""Cong cu tao danh sach tai khoan (Nen tang AI cong vu).

core    - quy tac chuan hoa dung chung + Che do 1 (Ho ten / Email / SDT)
mode2   - Che do 2 (them Don vi / To chuc), dung lai core
gui     - giao dien Tkinter
web     - giao dien web (FastAPI)
"""
import os

PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(PACKAGE_DIR, "templates")
TEMPLATE_MODE1 = os.path.join(TEMPLATE_DIR, "cls_template_users.xlsx")
TEMPLATE_MODE2 = os.path.join(TEMPLATE_DIR, "TemplateV2.xlsx")
