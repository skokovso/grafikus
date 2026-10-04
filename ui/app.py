# ui/app.py
"""Графический интерфейс приложения (CustomTkinter)."""
import os
import sys
import datetime
import ctypes
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from core.config import MONTHS_RU, MONTHS_RU_NOMINATIVE, MONTHS_RU_UPPER
from core.schedule_builder import build_master_schedule

def get_base_path():
    """Возвращает путь к папке, где находится исполняемый файл (или скрипт)"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load_spravochnik(filename):
    base_path = get_base_path()
    file_path = os.path.join(base_path, filename)
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f.readlines() if line.strip()]
    except FileNotFoundError:
        return []

def save_spravochnik(filename, data):
    base_path = get_base_path()
    file_path = os.path.join(base_path, filename)
    with open(file_path, 'w', encoding='utf-8') as f:
        for item in data:
            f.write(item + '\n')

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Иконка
        try:
            icon_path = os.path.join(get_base_path(), "icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
                try:
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SvodnyGrafik")
                except:
                    pass
        except Exception as e:
            print(f"⚠️ Ошибка установки иконки: {e}")

        self.title("Сводный график дежурств")
        self.geometry("700x750")
        self.minsize(650, 700)

        self.folder_path = ctk.StringVar(value="")
        self.rukovoditel = ctk.StringVar()
        self.ploshadka = ctk.StringVar()
        self.month_var = ctk.StringVar()
        self.year_var = ctk.StringVar()
        self.last_success_message = ""

        self.rukovoditeli_list = load_spravochnik('spravochnik_rukovoditeli.txt')
        self.ploshadki_list = load_spravochnik('spravochnik_ploshadki.txt')

        if not self.rukovoditeli_list:
            self.rukovoditeli_list = [
                'Главный врач ГОБУЗ МОКМЦ|Тарбаев Е.Ю.',
                'Зам. гл. врача по медицинской части ГОБУЗ МОКМЦ|Гредягин С.С.',
            ]
            save_spravochnik('spravochnik_rukovoditeli.txt', self.rukovoditeli_list)
        if not self.ploshadki_list:
            self.ploshadki_list = ['Ломоносова, д.18', 'Володарского, д.18', 'Перинатальный центр']
            save_spravochnik('spravochnik_ploshadki.txt', self.ploshadki_list)

        now = datetime.datetime.now()
        default_year = now.year
        if now.month == 12 and now.day > 15:
            default_year += 1
        self.year_var.set(str(default_year))

        self.create_widgets()

    def create_widgets(self):
        self.label_title = ctk.CTkLabel(self, text="📊 Сводный график дежурств", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_title.pack(pady=(20, 15))

        # Папка
        self.frame_folder = ctk.CTkFrame(self)
        self.frame_folder.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(self.frame_folder, text="📁 Папка с графиками:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=5, pady=(5, 0))
        folder_row = ctk.CTkFrame(self.frame_folder, fg_color="transparent")
        folder_row.pack(fill="x", padx=5, pady=5)
        self.entry_folder = ctk.CTkEntry(folder_row, textvariable=self.folder_path, placeholder_text="Выберите папку...", width=500)
        self.entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.btn_folder = ctk.CTkButton(folder_row, text="Обзор...", command=self.select_folder, width=80)
        self.btn_folder.pack(side="right")

        # Период
        self.frame_date = ctk.CTkFrame(self)
        self.frame_date.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(self.frame_date, text="📅 Период:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=5, pady=(5, 0))
        date_row = ctk.CTkFrame(self.frame_date, fg_color="transparent")
        date_row.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(date_row, text="Месяц:").pack(side="left", padx=(0, 10))
        self.combo_month = ctk.CTkComboBox(date_row, values=list(MONTHS_RU_NOMINATIVE.values()), width=150)
        self.combo_month.pack(side="left", padx=(0, 20))
        self.combo_month.set(MONTHS_RU_NOMINATIVE.get(datetime.datetime.now().month, "Октябрь"))
        ctk.CTkLabel(date_row, text="Год:").pack(side="left", padx=(0, 10))
        self.combo_year = ctk.CTkComboBox(date_row, values=[str(y) for y in range(2024, 2031)], width=100)
        self.combo_year.pack(side="left")
        self.combo_year.set(str(datetime.datetime.now().year))

        # Руководитель
        self.frame_ruk = ctk.CTkFrame(self)
        self.frame_ruk.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(self.frame_ruk, text="👤 Руководитель:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=5, pady=(5, 0))
        self.combo_ruk = ctk.CTkComboBox(self.frame_ruk, values=self.rukovoditeli_list, width=500)
        self.combo_ruk.pack(padx=5, pady=5)
        self.combo_ruk.set(self.rukovoditeli_list[0] if self.rukovoditeli_list else "")

        # Площадка
        self.frame_plo = ctk.CTkFrame(self)
        self.frame_plo.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(self.frame_plo, text="🏥 Площадка:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=5, pady=(5, 0))
        self.combo_plo = ctk.CTkComboBox(self.frame_plo, values=self.ploshadki_list, width=500)
        self.combo_plo.pack(padx=5, pady=5)
        self.combo_plo.set(self.ploshadki_list[0] if self.ploshadki_list else "")

        # Файлы
        self.frame_files = ctk.CTkFrame(self)
        self.frame_files.pack(fill="both", expand=True, padx=20, pady=5)
        ctk.CTkLabel(self.frame_files, text=" Файлы в папке:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=5, pady=(5, 0))
        self.files_listbox = ctk.CTkTextbox(self.frame_files, height=80, font=ctk.CTkFont(size=11))
        self.files_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkButton(self.frame_files, text="🔄 Обновить список", command=self.refresh_files_list, width=150).pack(pady=5)

        # Кнопка запуска
        self.btn_run = ctk.CTkButton(self, text=" Составить сводный график", font=ctk.CTkFont(size=14, weight="bold"),
                                     fg_color="#4CAF50", hover_color="#388E3C", height=45, command=self.run)
        self.btn_run.pack(pady=20, padx=20, fill="x")

        # Статус
        self.status_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.status_frame.pack(fill="x", padx=20, pady=(0, 15))
        self.status_textbox = ctk.CTkTextbox(self.status_frame, height=80, font=ctk.CTkFont(size=13), wrap="word")
        self.status_textbox.pack(fill="x", pady=5)
        self.status_textbox.insert("0.0", "✅ Готов к работе")
        self.status_textbox.configure(state="disabled")
        self.status_textbox.bind("<Button-1>", self.on_status_click)
        ctk.CTkButton(self.status_frame, text="🗑️ Очистить", command=self.clear_status, width=80, height=28, fg_color="#607D8B").pack(pady=(0, 5))

    def set_status(self, message, status_type="info", details=None):
        colors = {"info": "#4FC3F7", "success": "#66BB6A", "error": "#EF5350", "warning": "#FFA726"}
        color = colors.get(status_type, "#FFFFFF")
        self.status_textbox.configure(state="normal")
        self.status_textbox.delete("0.0", "end")
        self.status_textbox.insert("0.0", message + "\n", status_type)
        self.status_textbox.tag_config(status_type, foreground=color)
        if details:
            self.status_textbox.insert("end", "\n" + details, "details")
            self.status_textbox.tag_config("details", foreground="#B0BEC5")
        self.status_textbox.configure(state="disabled")
        if status_type == "success" and "успешно создан" in message:
            self.last_success_message = message
            self.status_textbox.configure(cursor="hand2")
        else:
            self.status_textbox.configure(cursor="arrow")
        self.update()

    def clear_status(self):
        self.status_textbox.configure(state="normal")
        self.status_textbox.delete("0.0", "end")
        self.status_textbox.insert("0.0", "✅ Готов к работе")
        self.status_textbox.configure(state="disabled")
        self.status_textbox.configure(cursor="arrow")
        self.last_success_message = ""

    def show_excluded_files(self, files_with_errors):
        if not files_with_errors: return
        count = len(files_with_errors)
        files_list = "\n".join([f"  • {fname} — {error}" for fname, error in files_with_errors[:10]])
        if count > 10: files_list += f"\n  ... и ещё {count - 10} файлов"
        self.set_status(f"⚠️ {count} файлов исключены (не соответствуют месяцу)", "warning", f"Исключённые файлы:\n{files_list}")

    def refresh_files_list(self):
        folder = self.folder_path.get().strip()
        self.files_listbox.delete("0.0", "end")
        if not folder or not os.path.exists(folder):
            self.files_listbox.insert("0.0", "⚠️ Папка не выбрана или не существует")
            return
        try:
            graph_files = sorted([f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls', '.csv', '.pdf', '.txt'))])
            if graph_files:
                for f in graph_files: self.files_listbox.insert("end", f + "\n")
                self.set_status(f"✅ Найдено {len(graph_files)} файлов", "success")
            else:
                self.files_listbox.insert("0.0", "⚠️ Нет файлов графиков")
        except Exception as e:
            self.files_listbox.insert("0.0", f"❌ Ошибка: {str(e)}")

    def select_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с графиками")
        if folder:
            self.folder_path.set(folder)
            self.refresh_files_list()

    def on_status_click(self, event):
        text = self.status_textbox.get("0.0", "end").strip()
        if "успешно создан" in text or "❌ Файл" in text:
            self.open_output_folder()

    def open_output_folder(self):
        folder = self.folder_path.get().strip()
        if not folder: return
        output_folder = Path(folder) / "Сводный график"
        if output_folder.exists():
            try:
                os.startfile(str(output_folder))
            except:
                pass

    def run(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.exists(folder):
            self.set_status("❌ Ошибка: выберите папку с графиками!", "error")
            return
        
        ruk_text = self.combo_ruk.get().strip()
        plo_text = self.combo_plo.get().strip()
        if not ruk_text or not plo_text:
            self.set_status("❌ Ошибка: выберите руководителя и площадку!", "error")
            return

        month_str = self.combo_month.get().strip()
        month_names = {v: k for k, v in MONTHS_RU_NOMINATIVE.items()}
        month_num = month_names.get(month_str)
        if not month_num:
            self.set_status(f" Неверный месяц: {month_str}", "error")
            return

        try:
            year_num = int(self.combo_year.get().strip())
        except:
            self.set_status("❌ Неверный формат года!", "error")
            return

        self.set_status("⏳ Обработка... Пожалуйста, подождите", "info")
        self.btn_run.configure(state="disabled", text="⏳ Обработка...")
        self.update()

        try:
            success, files_with_errors, error_message = build_master_schedule(
                folder, ruk_text, plo_text, month_num, year_num
            )
            if success:
                self.set_status("✅ Сводный график успешно создан! Файлы в папке 'Сводный график'", "success")
                if files_with_errors: self.show_excluded_files(files_with_errors)
            else:
                self.set_status(error_message if error_message else "❌ Ошибка при создании графика", "error")
        except Exception as e:
            self.set_status(f"❌ Критическая ошибка: {str(e)}", "error")
            import traceback
            print(traceback.format_exc())
        finally:
            self.btn_run.configure(state="normal", text="🚀 Составить сводный график")
            self.update()