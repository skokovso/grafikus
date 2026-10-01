# main.py
"""Точка входа в приложение Сводный график."""
import sys
import os

# Добавляем корневую папку проекта в sys.path для корректных импортов
# (важно при запуске из IDE или через PyInstaller)
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

if base_path not in sys.path:
    sys.path.insert(0, base_path)

# Настройка CustomTkinter
import customtkinter as ctk
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

from ui.app import App

if __name__ == '__main__':
    app = App()
    app.mainloop()