# main.py - СФМ ВОЛП: Оценка качества (ПОЛНОСТЬЮ РАБОЧАЯ ВЕРСИЯ)
# Запуск: python main.py
# Зависимости: pip install numpy pandas matplotlib openpyxl

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json
import os
import base64
from io import BytesIO
from datetime import datetime


# ================== МАТЕМАТИЧЕСКОЕ ЯДРО ==================
def calculate_weights(m, n):
    """Расчёт нормированных весовых коэффициентов по формуле (7) из ТЗ"""
    total = m * n
    if total == 0:
        return np.array([])
    return np.array([2 * (total - l + 1) / (total * (total + 1)) for l in range(1, total + 1)])


def extract_numeric_rows(data):
    """Извлекает только числовые колонки (начиная с 4-й: индекс 4+)"""
    numeric_data = []
    for row in data:
        if len(row) < 4:
            raise ValueError("Строка содержит менее 4 колонок")
        numeric_part = row[4:]
        numeric_data.append(numeric_part)
    return numeric_data


def calculate_matrices(base, real):
    """Расчёт матриц разностей по формулам из ТЗ"""
    base_num = extract_numeric_rows(base)
    real_num = extract_numeric_rows(real)

    # ✅ КОНВЕРТИРУЕМ в numpy массивы
    base_arr = np.array(base_num, dtype=float)
    real_arr = np.array(real_num, dtype=float)

    if base_arr.shape != real_arr.shape:
        raise ValueError(f"Размеры не совпадают: база {base_arr.shape} vs реальная {real_arr.shape}")
    if base_arr.size == 0:
        raise ValueError("Пустая матрица данных")

    BQ1 = base_arr.copy()
    BQ2 = base_arr.copy()
    RQ1 = real_arr.copy()
    RQ2 = real_arr.copy()

    BQ1[:, ::2] = 0  # Обнуляем Min для BQ1
    BQ2[:, 1::2] = 0  # Обнуляем Max для BQ2
    RQ1[:, ::2] = 0
    RQ2[:, 1::2] = 0

    CQ1 = BQ1 - RQ1
    CQ2 = RQ2 - BQ2

    if (CQ1 < 0).any() or (CQ2 < 0).any():
        return None, CQ1, CQ2, BQ1, BQ2, RQ1, RQ2

    CQ = CQ1 + CQ2
    return CQ, CQ1, CQ2, BQ1, BQ2, RQ1, RQ2


def calculate_Q(CQ):
    """Расчёт комплексного показателя качества по формуле (9) из ТЗ"""
    if CQ is None or CQ.size == 0:
        return 0.0
    m, n = CQ.shape
    weights = calculate_weights(m, n)
    return np.sum(CQ.flatten() * weights)


def calculate_weighted_table(CQ):
    """Расчёт таблицы взвешенных значений (Таблица 6 из ТЗ)"""
    if CQ is None or CQ.size == 0:
        return None, None
    m, n = CQ.shape
    weights = calculate_weights(m, n)
    weighted = CQ * weights.reshape(m, n)
    return weighted, weights.reshape(m, n)


# ================== GUI: РЕДАКТОР ТАБЛИЦЫ ==================
class TableEditor:
    def __init__(self, parent, headers=None):
        self.frame = tk.Frame(parent)
        self.frame.pack(fill="both", expand=True)
        self.headers = headers or ["№", "Функция", "СП", "ТС",
                                   "P1_Min", "P1_Max", "P2_Min", "P2_Max",
                                   "P3_Min", "P3_Max", "P4_Min", "P4_Max"]
        self.rows_count = tk.IntVar(value=5)
        self.params_count = tk.IntVar(value=4)
        self.entries = []
        self.header_labels = []
        self._create_controls()

    def _create_controls(self):
        ctrl = tk.Frame(self.frame)
        ctrl.pack(fill="x", pady=5)

        tk.Label(ctrl, text="Строк:").pack(side="left")
        tk.Spinbox(ctrl, from_=1, to=50, textvariable=self.rows_count, width=4,
                   command=self._rebuild).pack(side="left")

        tk.Label(ctrl, text="Параметров (пар Min/Max):").pack(side="left")
        tk.Spinbox(ctrl, from_=1, to=10, textvariable=self.params_count, width=4,
                   command=self._rebuild).pack(side="left")

        tk.Button(ctrl, text="Обновить", command=self._rebuild).pack(side="left", padx=5)
        tk.Button(ctrl, text="Очистить", command=self._clear).pack(side="left")
        tk.Button(ctrl, text="➕ Добавить строку", command=self._add_row, bg="#4CAF50", fg="white").pack(side="left",
                                                                                                        padx=5)
        tk.Button(ctrl, text="🗑️ Удалить строку", command=self._delete_row, bg="#f44336", fg="white").pack(side="left",
                                                                                                           padx=5)

        self.canvas = tk.Canvas(self.frame)
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._rebuild()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _rebuild(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.entries = []
        self.header_labels = []

        rows = self.rows_count.get()
        cols = 4 + self.params_count.get() * 2

        for j, header in enumerate(self.headers[:cols]):
            lbl = tk.Label(self.scrollable_frame, text=header, font=("Arial", 9, "bold"),
                           bg="#e0e0e0", relief="ridge", padx=2, pady=2)
            lbl.grid(row=0, column=j, sticky="ew")
            self.header_labels.append(lbl)

        for i in range(rows):
            row_entries = []
            for j in range(cols):
                e = tk.Entry(self.scrollable_frame, width=10 if j >= 4 else 12, font=("Arial", 9))
                e.grid(row=i + 1, column=j, padx=1, pady=1)
                if j == 0:
                    e.insert(0, str(i + 1))
                row_entries.append(e)
            self.entries.append(row_entries)

    def _add_row(self):
        self.rows_count.set(self.rows_count.get() + 1)
        self._rebuild()

    def _delete_row(self):
        if self.rows_count.get() > 1:
            self.rows_count.set(self.rows_count.get() - 1)
            self._rebuild()

    def _clear(self):
        for row in self.entries:
            for j, e in enumerate(row):
                if j >= 4:
                    e.delete(0, tk.END)

    def get_full_data(self):
        data = []
        for i, row in enumerate(self.entries):
            row_data = []
            for j, e in enumerate(row):
                val = e.get().strip()
                if j < 4:
                    row_data.append(val if val else f"Item_{i + 1}_{j + 1}")
                else:
                    if val == "":
                        row_data.append(0.0)
                    else:
                        try:
                            row_data.append(float(val))
                        except ValueError:
                            row_data.append(0.0)
            data.append(row_data)
        return data

    def get_numeric_data(self):
        full = self.get_full_data()
        return [row[4:] for row in full]

    def set_data(self, data):
        """Установка данных в таблицу с обработкой NaN и преобразованием значений"""
        if not data:
            return

        cleaned_data = []
        for row in data:
            cleaned_row = []
            for val in row:
                if pd.isna(val):
                    cleaned_row.append("")
                elif isinstance(val, (float, int)):
                    if isinstance(val, float):
                        if val == int(val):
                            cleaned_row.append(str(int(val)))
                        else:
                            cleaned_row.append(f"{val:.3f}".rstrip('0').rstrip('.'))
                    else:
                        cleaned_row.append(str(val))
                else:
                    cleaned_row.append(str(val))
            cleaned_data.append(cleaned_row)

        data = cleaned_data

        if not data:
            return

        rows = len(data)
        cols = len(data[0]) if data else 4
        num_params = max(0, (cols - 4) // 2)

        self.params_count.set(num_params)
        self.rows_count.set(rows)
        self._rebuild()

        for i in range(min(rows, len(self.entries))):
            for j in range(min(cols, len(self.entries[i]))):
                val = data[i][j]
                self.entries[i][j].delete(0, tk.END)
                self.entries[i][j].insert(0, val if val else "")


# ================== ОСНОВНОЕ ПРИЛОЖЕНИЕ ==================
class App:
    DEFAULT_HEADERS = ["№", "Функция", "СП", "ТС",
                       "P1_Затухание_Min", "P1_Затухание_Max",
                       "P2_Скорость_Min", "P2_Скорость_Max",
                       "P3_Отказы_Min", "P3_Отказы_Max",
                       "P4_Срок_службы_Min", "P4_Срок_службы_Max"]

    PARAM_NAMES = ["P1_Затухание", "P2_Скорость", "P3_Отказы", "P4_Срок_службы"]
    PARAM_UNITS = ["дБ", "Гбит/с", "отказов", "лет"]

    def __init__(self, root):
        self.root = root
        self.root.title("📡 СФМ ВОЛП: Оценка качества")
        self.root.geometry("1600x1000")

        self.base = None
        self.models = {}
        self.results = []
        self.detailed_results = []
        self.selected_model = None
        self.editing_editor = None
        self.graph_canvas = None
        self.matrix_canvas = None

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook.Tab', padding=[12, 8], font=('Arial', 10))

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_data = tk.Frame(self.nb)
        self.tab_calc = tk.Frame(self.nb)
        self.tab_graphs = tk.Frame(self.nb)
        self.tab_matrices = tk.Frame(self.nb)
        self.tab_help = tk.Frame(self.nb)

        self.nb.add(self.tab_data, text="📥 Данные")
        self.nb.add(self.tab_calc, text="⚙ Расчёт")
        self.nb.add(self.tab_graphs, text="📊 Графики")
        self.nb.add(self.tab_matrices, text="🔢 Матрицы")
        self.nb.add(self.tab_help, text="❓ Справка")

        self.init_data_tab()
        self.init_calc_tab()
        self.init_graphs_tab()
        self.init_matrices_tab()
        self.init_help_tab()

        self.status = tk.Label(root, text="Готов к работе", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    # ===== ВКЛАДКА: ДАННЫЕ =====
    def init_data_tab(self):
        toolbar = tk.Frame(self.tab_data)
        toolbar.pack(fill="x", pady=5)
        tk.Button(toolbar, text="📂 Загрузить Excel", command=self.load_excel, bg="#4CAF50", fg="white").pack(
            side="left", padx=3)
        tk.Button(toolbar, text="💾 Сохранить проект", command=self.save_project).pack(side="left", padx=3)
        tk.Button(toolbar, text="📤 Загрузить проект", command=self.load_project).pack(side="left", padx=3)
        tk.Button(toolbar, text="🧹 Очистить всё", command=self.clear_all, bg="#f44336", fg="white").pack(side="right",
                                                                                                         padx=3)

        tk.Label(self.tab_data, text="🎯 Базовая (идеальная) модель", font=("Arial", 12, "bold")).pack(pady=(10, 5))
        self.base_editor = TableEditor(self.tab_data, headers=self.DEFAULT_HEADERS)

        btn_frame = tk.Frame(self.tab_data)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="✅ Сохранить базу", command=self.save_base, bg="#2196F3", fg="white").pack(
            side="left", padx=5)
        tk.Button(btn_frame, text="📋 Копировать из базы", command=self.copy_base_to_model).pack(side="left", padx=5)

        tk.Label(self.tab_data, text="📚 Загруженные модели:", font=("Arial", 11, "bold")).pack(pady=(15, 5))
        self.models_listbox = tk.Listbox(self.tab_data, width=80, height=8, font=("Arial", 10))
        self.models_listbox.pack(fill="x", padx=10, pady=5)
        self.models_listbox.bind('<<ListboxSelect>>', self.on_model_select)

        self.editing_frame = tk.Frame(self.tab_data)
        self.editing_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def save_base(self):
        try:
            self.base = self.base_editor.get_full_data()
            self.status.config(text=f"✅ База сохранена: {len(self.base)} строк")
            messagebox.showinfo("Успех", "Базовая модель сохранена в памяти")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить базу:\n{str(e)}")

    def copy_base_to_model(self):
        if not self.base:
            messagebox.showwarning("Внимание", "Сначала сохраните базовую модель")
            return
        name = simpledialog.askstring("Новая модель", "Название модели:", initialvalue=f"Model_{len(self.models) + 1}")
        if not name or name in self.models:
            if name in self.models:
                messagebox.showwarning("Внимание", f"Модель '{name}' уже существует!")
            return
        self.models[name] = [row[:] for row in self.base]
        self._update_models_list()
        self._update_matrix_combo()
        self.status.config(text=f"✅ Добавлена модель: {name}")

    def on_model_select(self, event):
        sel = self.models_listbox.curselection()
        if not sel:
            return

        listbox_text = self.models_listbox.get(sel[0])
        model_name = listbox_text.split(" (")[0]

        if messagebox.askyesno("Редактирование модели",
                               f"Отредактировать модель\n«{model_name}»?\n\n"
                               f"Данных строк: {len(self.models.get(model_name, []))}"):

            self.selected_model = model_name

            for widget in self.editing_frame.winfo_children():
                widget.destroy()

            tk.Label(self.editing_frame, text=f"✏️ Редактирование модели: {model_name}",
                     font=("Arial", 12, "bold"), fg="#2196F3").pack(pady=(0, 5))

            self.editing_editor = TableEditor(self.editing_frame, headers=self.DEFAULT_HEADERS)

            if model_name in self.models and self.models[model_name]:
                self.editing_editor.set_data(self.models[model_name])
                self.status.config(text=f"📝 Редактирование: {model_name}")
            else:
                messagebox.showwarning("Внимание", f"Нет данных для модели {model_name}")
                self.cancel_editing()
                return

            btns = tk.Frame(self.editing_frame)
            btns.pack(pady=10)
            tk.Button(btns, text="💾 Сохранить изменения",
                      command=lambda: self.save_edited_model(model_name),
                      bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=8)
            tk.Button(btns, text="❌ Отмена", command=self.cancel_editing,
                      bg="#f44336", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=8)
        else:
            self.status.config(text=f"ℹ️ Выбрана модель: {model_name} (редактирование отменено)")

    def save_edited_model(self, name):
        if self.editing_editor:
            self.models[name] = self.editing_editor.get_full_data()
            self._update_models_list()
            self._update_matrix_combo()
            messagebox.showinfo("✅ Сохранено",
                                f"Изменения в модели «{name}» успешно сохранены!\n"
                                f"Строк: {len(self.models[name])}")
            self.status.config(text=f"✅ Модель {name} обновлена")
            self.cancel_editing()

    def cancel_editing(self):
        for widget in self.editing_frame.winfo_children():
            widget.destroy()
        self.editing_editor = None
        self.selected_model = None
        self.status.config(text="Редактирование отменено")

    def _update_models_list(self):
        self.models_listbox.delete(0, tk.END)
        for name in sorted(self.models.keys()):
            rows = len(self.models[name])
            self.models_listbox.insert(tk.END, f"{name} ({rows} строк)")

    def _update_matrix_combo(self):
        """Обновление списка моделей в комбобоксе матриц"""
        model_names = sorted(self.models.keys())
        self.matrix_model_combo['values'] = model_names
        if model_names:
            self.matrix_model_combo.current(0)
        print(f"📋 Обновлён список матриц: {model_names}")

    def load_excel(self):
        file = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file:
            return

        try:
            xls = pd.ExcelFile(file)
            sheet_names = xls.sheet_names

            if "Base" not in sheet_names:
                messagebox.showerror("Ошибка",
                                     "Файл должен содержать лист 'Base'!\n"
                                     f"Доступные листы: {', '.join(sheet_names)}")
                return

            base_df = pd.read_excel(xls, sheet_name="Base")
            self.base = base_df.values.tolist()
            self.base_editor.set_data(self.base)

            self.models = {}
            for sheet in sheet_names:
                if sheet != "Base":
                    model_df = pd.read_excel(xls, sheet_name=sheet)
                    self.models[sheet] = model_df.values.tolist()

            self._update_models_list()
            self._update_matrix_combo()

            self.status.config(text=f"✅ Загружено: {len(self.models)} моделей из {file.split('/')[-1]}")

            print(f"📊 Загружено листов: {sheet_names}")
            print(f"📊 Размер базы: {len(self.base)} строк")
            for name, data in self.models.items():
                print(f"📊 Модель {name}: {len(data)} строк")

            messagebox.showinfo("Успех",
                                f"Файл загружен!\n📊 Листов: {len(sheet_names)}\n📚 Моделей: {len(self.models)}\n\n"
                                f"📌 ШАГ 2: Перейдите на вкладку '⚙ Расчёт'\n"
                                f"📌 ШАГ 3: Нажмите '🚀 Рассчитать все модели'\n"
                                f"📌 ШАГ 4: Перейдите на '📊 Графики' или '🔢 Матрицы'")

        except Exception as e:
            messagebox.showerror("Ошибка загрузки", f"Не удалось загрузить файл:\n{str(e)}")
            print(f"❌ Ошибка: {e}")
            import traceback
            traceback.print_exc()

    def save_project(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file:
            return
        try:
            data = {
                "base": self.base,
                "models": self.models,
                "headers": self.DEFAULT_HEADERS,
                "saved_at": datetime.now().isoformat()
            }
            with open(file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.status.config(text=f"💾 Проект сохранён: {os.path.basename(file)}")
            messagebox.showinfo("Успех", "Проект сохранён")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить проект:\n{str(e)}")

    def load_project(self):
        file = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file:
            return
        try:
            with open(file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.base = data.get("base")
            self.models = data.get("models", {})
            if "headers" in data:
                self.DEFAULT_HEADERS = data["headers"]
            if self.base:
                self.base_editor.set_data(self.base)
            self._update_models_list()
            self._update_matrix_combo()
            self.status.config(text=f"📤 Проект загружен: {os.path.basename(file)}")
            messagebox.showinfo("Успех", f"Загружено моделей: {len(self.models)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить проект:\n{str(e)}")

    def clear_all(self):
        if messagebox.askyesno("Подтверждение", "Очистить все данные?\n\nЭто действие нельзя отменить!"):
            self.base = None
            self.models = {}
            self.results = []
            self.detailed_results = []
            self.selected_model = None

            for w in self.editing_frame.winfo_children():
                w.destroy()

            self.base_editor._rebuild()
            self.editing_editor = None
            self._update_models_list()
            self._update_matrix_combo()

            if self.graph_canvas:
                self.graph_canvas.get_tk_widget().pack_forget()
                self.graph_canvas = None

            if self.matrix_canvas:
                self.matrix_canvas.get_tk_widget().pack_forget()
                self.matrix_canvas = None

            self.status.config(text="🧹 Все данные очищены")

    # ===== ВКЛАДКА: РАСЧЁТ =====
    def init_calc_tab(self):
        toolbar = tk.Frame(self.tab_calc)
        toolbar.pack(fill="x", pady=5)
        tk.Button(toolbar, text="🚀 Рассчитать все модели", command=self.calculate_all,
                  bg="#4CAF50", fg="white", font=("Arial", 11, "bold")).pack(side="left", padx=5)
        tk.Button(toolbar, text="🔄 Сбросить", command=self.reset_results).pack(side="left", padx=5)
        tk.Button(toolbar, text="📄 Экспорт отчёта (HTML)", command=self.export_full_report,
                  bg="#FF9800", fg="white").pack(side="right", padx=5)

        self.text = tk.Text(self.tab_calc, font=("Consolas", 10), wrap=tk.WORD)
        self.text.pack(fill="both", expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(self.text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.text.yview)

    def calculate_all(self):
        if not self.base:
            messagebox.showerror("Ошибка", "Сначала загрузите или создайте базовую модель!")
            return
        if not self.models:
            messagebox.showerror("Ошибка", "Нет моделей для расчёта!")
            return

        self.results = []
        self.detailed_results = []
        self.text.delete(1.0, tk.END)
        self.text.insert(tk.END, "🔍 НАЧАЛО РАСЧЁТА\n" + "=" * 60 + "\n\n")

        for name, model in self.models.items():
            try:
                result = calculate_matrices(self.base, model)
                CQ, CQ1, CQ2, BQ1, BQ2, RQ1, RQ2 = result

                if CQ is None:
                    self.text.insert(tk.END, f"❌ {name}: НЕ СООТВЕТСТВУЕТ (∆Q < 0 в матрице)\n")
                    self.detailed_results.append({
                        'name': name, 'status': 'FAIL', 'Q_ke': 0, 'CQ': None,
                        'CQ1': CQ1, 'CQ2': CQ2, 'BQ1': BQ1, 'BQ2': BQ2,
                        'RQ1': RQ1, 'RQ2': RQ2, 'weighted': None, 'weights': None
                    })
                    continue

                Q_ke = calculate_Q(CQ)
                weighted, weights = calculate_weighted_table(CQ)

                self.results.append((name, Q_ke, CQ, CQ1, CQ2))
                self.detailed_results.append({
                    'name': name, 'status': 'PASS', 'Q_ke': Q_ke, 'CQ': CQ,
                    'CQ1': CQ1, 'CQ2': CQ2, 'BQ1': BQ1, 'BQ2': BQ2,
                    'RQ1': RQ1, 'RQ2': RQ2, 'weighted': weighted, 'weights': weights
                })

                self.text.insert(tk.END, f"✅ {name}:\n")
                self.text.insert(tk.END, f"   Qкэ = {Q_ke:.6f}\n")
                self.text.insert(tk.END, f"   Размер матрицы: {CQ.shape[0]}×{CQ.shape[1]}\n")
                self.text.insert(tk.END, f"   Сумма ∆Q: {np.sum(CQ):.4f}\n\n")

            except Exception as e:
                self.text.insert(tk.END, f"❌ {name}: ОШИБКА — {str(e)}\n\n")
                import traceback
                traceback.print_exc()

        if self.results:
            self.results.sort(key=lambda x: x[1], reverse=True)
            self.detailed_results.sort(key=lambda x: x['Q_ke'], reverse=True)

            self.text.insert(tk.END, "\n" + "🏆 РЕЙТИНГ КАЧЕСТВА".center(60, "=") + "\n")
            for i, (name, Q_ke, CQ, _, _) in enumerate(self.results, 1):
                medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "•"
                self.text.insert(tk.END, f"{medal} {i}. {name:25s} → Qкэ = {Q_ke:.6f}\n")

            self.status.config(text=f"✅ Расчёт завершён: {len(self.results)} моделей обработано")
            messagebox.showinfo("Расчёт завершён",
                                f"✅ Обработано: {len(self.results)} моделей\n"
                                f"❌ Не прошли: {len(self.detailed_results) - len(self.results)}\n\n"
                                f"📊 Теперь перейдите на вкладку 'Графики' или 'Матрицы'!")
        else:
            self.text.insert(tk.END, "\n⚠ Нет успешных расчётов для формирования рейтинга")
            self.status.config(text="⚠ Расчёт завершён с ошибками")

    def reset_results(self):
        self.results = []
        self.detailed_results = []
        self.text.delete(1.0, tk.END)
        self.text.insert(tk.END, "🔄 Результаты сброшены. Нажмите 'Рассчитать' для нового расчёта.")
        self.status.config(text="🔄 Результаты сброшены")

    # ===== ВКЛАДКА: ГРАФИКИ =====
    def init_graphs_tab(self):
        toolbar = tk.Frame(self.tab_graphs)
        toolbar.pack(fill="x", pady=5)

        tk.Button(toolbar, text="📊 Все графики", command=self.show_all_graphs,
                  bg="#2196F3", fg="white", font=("Arial", 11, "bold")).pack(side="left", padx=3)
        tk.Button(toolbar, text="🏆 Рейтинг", command=self.show_rating_chart).pack(side="left", padx=3)
        tk.Button(toolbar, text="📈 По параметрам", command=self.show_params_chart).pack(side="left", padx=3)
        tk.Button(toolbar, text="🎯 Радары", command=self.show_radar_charts).pack(side="left", padx=3)
        tk.Button(toolbar, text="📉 По компонентам", command=self.show_components_chart).pack(side="left", padx=3)
        tk.Button(toolbar, text="🔄 Обновить", command=self.show_all_graphs).pack(side="right", padx=3)

        self.graph_frame = tk.Frame(self.tab_graphs)
        self.graph_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.graph_canvas = None

    def show_all_graphs(self):
        if not self.results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт на вкладке '⚙ Расчёт'!\n\n"
                                        "1. Перейдите на вкладку 'Расчёт'\n"
                                        "2. Нажмите '🚀 Рассчитать все модели'\n"
                                        "3. Вернитесь на 'Графики'")
            return

        if self.graph_canvas:
            try:
                self.graph_canvas.get_tk_widget().pack_forget()
            except:
                pass
            self.graph_canvas = None
            plt.close('all')

        try:
            fig = plt.figure(figsize=(18, 12))
            fig.suptitle("📡 СФМ ВОЛП: Детальная визуализация результатов оценки качества",
                         fontsize=16, fontweight="bold", y=0.98)

            ax1 = fig.add_subplot(2, 3, 1)
            self._plot_rating_chart(ax1)

            ax2 = fig.add_subplot(2, 3, 2)
            self._plot_params_chart(ax2)

            ax3 = fig.add_subplot(2, 3, 3, polar=True)
            self._plot_radar_chart(ax3, self.detailed_results[0] if self.detailed_results else None)

            ax4 = fig.add_subplot(2, 3, 4)
            self._plot_components_chart(ax4)

            ax5 = fig.add_subplot(2, 3, 5)
            self._plot_weights_distribution(ax5)

            ax6 = fig.add_subplot(2, 3, 6)
            self._plot_minmax_comparison(ax6)

            plt.tight_layout(rect=[0, 0, 1, 0.96])

            self.graph_canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.graph_canvas.draw()
            self.graph_canvas.get_tk_widget().pack(fill="both", expand=True)
            self.status.config(text="📊 Все графики построены")

        except Exception as e:
            messagebox.showerror("Ошибка графиков", f"Не удалось построить графики:\n{str(e)}")
            import traceback
            traceback.print_exc()

    def show_rating_chart(self):
        if not self.results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт")
            return

        if self.graph_canvas:
            self.graph_canvas.get_tk_widget().pack_forget()
            plt.close('all')

        fig, ax = plt.subplots(figsize=(12, 6))
        self._plot_rating_chart(ax)
        plt.tight_layout()

        self.graph_canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        self.graph_canvas.draw()
        self.graph_canvas.get_tk_widget().pack(fill="both", expand=True)

    def show_params_chart(self):
        if not self.results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт")
            return

        if self.graph_canvas:
            self.graph_canvas.get_tk_widget().pack_forget()
            plt.close('all')

        fig, ax = plt.subplots(figsize=(14, 7))
        self._plot_params_chart(ax)
        plt.tight_layout()

        self.graph_canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        self.graph_canvas.draw()
        self.graph_canvas.get_tk_widget().pack(fill="both", expand=True)

    def show_radar_charts(self):
        if not self.detailed_results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт")
            return

        if self.graph_canvas:
            self.graph_canvas.get_tk_widget().pack_forget()
            plt.close('all')

        passed = [r for r in self.detailed_results if r['status'] == 'PASS']
        n_models = min(len(passed), 6)

        fig = plt.figure(figsize=(16, 10))
        fig.suptitle("🎯 Радары разностей для всех моделей", fontsize=14, fontweight="bold")

        for i, result in enumerate(passed[:n_models]):
            ax = fig.add_subplot(2, 3, i + 1, polar=True)
            self._plot_radar_chart(ax, result)

        plt.tight_layout(rect=[0, 0, 1, 0.95])

        self.graph_canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        self.graph_canvas.draw()
        self.graph_canvas.get_tk_widget().pack(fill="both", expand=True)

    def show_components_chart(self):
        if not self.results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт")
            return

        if self.graph_canvas:
            self.graph_canvas.get_tk_widget().pack_forget()
            plt.close('all')

        fig, ax = plt.subplots(figsize=(12, 8))
        self._plot_components_chart(ax)
        plt.tight_layout()

        self.graph_canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        self.graph_canvas.draw()
        self.graph_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _plot_rating_chart(self, ax):
        names = [r[0] for r in self.results]
        values = [r[1] for r in self.results]

        colors = ['#FFD700' if i == 0 else '#C0C0C0' if i == 1 else '#CD7F32' if i == 2 else '#4CAF50'
                  for i in range(len(values))]
        bars = ax.bar(names, values, color=colors, edgecolor='black', linewidth=1.5, alpha=0.8)

        ax.set_xlabel("Модель", fontsize=11, fontweight='bold')
        ax.set_ylabel("Комплексный показатель качества (Qкэ)", fontsize=11, fontweight='bold')
        ax.set_title("🏆 Рейтинг качества моделей ВОЛП", fontsize=12, fontweight='bold')
        ax.tick_params(axis='x', rotation=45)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height, f'{height:.4f}',
                    ha='center', va='bottom', fontsize=9, fontweight='bold')

    def _plot_params_chart(self, ax):
        if not self.base or not self.results:
            return

        param_indices = [0, 2, 4, 6]
        x = np.arange(len(self.PARAM_NAMES))
        width = 0.15

        # ✅ КОНВЕРТИРУЕМ в numpy массив
        base_numeric = np.array(extract_numeric_rows(self.base), dtype=float)
        base_means = []
        for i, idx in enumerate(param_indices):
            if idx < base_numeric.shape[1]:
                min_vals = base_numeric[:, idx]
                max_vals = base_numeric[:, idx + 1]
                base_means.append(np.mean((min_vals + max_vals) / 2))
            else:
                base_means.append(0)

        model_data = []
        for name, Q_ke, CQ, _, _ in self.results[:5]:
            model = self.models.get(name)
            if model:
                # ✅ КОНВЕРТИРУЕМ в numpy массив
                model_numeric = np.array(extract_numeric_rows(model), dtype=float)
                means = []
                for idx in param_indices:
                    if idx < model_numeric.shape[1]:
                        min_vals = model_numeric[:, idx]
                        max_vals = model_numeric[:, idx + 1]
                        means.append(np.mean((min_vals + max_vals) / 2))
                    else:
                        means.append(0)
                model_data.append((name, means))

        for i, (name, means) in enumerate(model_data):
            offset = (i - len(model_data) / 2 + 0.5) * width
            ax.bar(x + offset, means, width, label=name, alpha=0.8, edgecolor='black')

        ax.plot(x, base_means, 'k--', linewidth=2, marker='o', markersize=8, label='🎯 БАЗА (идеал)')

        ax.set_xlabel("Параметр качества", fontsize=11, fontweight='bold')
        ax.set_ylabel("Среднее значение", fontsize=11, fontweight='bold')
        ax.set_title("📈 Сравнение параметров моделей с базой", fontsize=12, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([f"{p}\n({u})" for p, u in zip(self.PARAM_NAMES, self.PARAM_UNITS)], fontsize=9)
        ax.legend(loc='upper right', fontsize=8)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

    def _plot_radar_chart(self, ax, result):
        if result is None or result['CQ'] is None:
            ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center', transform=ax.transAxes)
            ax.set_title("❌ Нет данных")
            return

        CQ = result['CQ']
        flat_CQ = CQ.flatten()
        n = len(flat_CQ)
        angles = np.linspace(0, 2 * np.pi, n, endpoint=False)
        flat_CQ_closed = np.concatenate((flat_CQ, [flat_CQ[0]]))
        angles_closed = np.concatenate((angles, [angles[0]]))

        color = '#2196F3' if result['status'] == 'PASS' else '#f44336'
        ax.plot(angles_closed, flat_CQ_closed, 'o-', linewidth=2, color=color,
                label=f"{result['name'][:15]} (Q={result['Q_ke']:.3f})")
        ax.fill(angles_closed, flat_CQ_closed, alpha=0.25, color=color)

        ax.set_title(f"🎯 {result['name'][:20]}", size=10, fontweight='bold', pad=15)
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=8)
        ax.set_xticks(angles)
        ax.set_xticklabels([f'C{j + 1}' for j in range(n)], fontsize=7)

    def _plot_components_chart(self, ax):
        if not self.base or not self.results:
            return

        ts_names = [row[3] if len(row) > 3 else f"ТС_{i + 1}" for i, row in enumerate(self.base)]

        if self.detailed_results and self.detailed_results[0]['weighted'] is not None:
            best = self.detailed_results[0]
            weighted = best['weighted']
            row_contributions = np.sum(weighted, axis=1)

            y = np.arange(len(ts_names))
            colors = ['#4CAF50' if v >= 0 else '#f44336' for v in row_contributions]

            bars = ax.barh(y, row_contributions, color=colors, edgecolor='black', alpha=0.8)
            ax.set_yticks(y)
            ax.set_yticklabels(ts_names, fontsize=9)
            ax.set_xlabel("Вклад в Qкэ (∑ ql × ∆Ql по строке)", fontsize=11, fontweight='bold')
            ax.set_title(f"📉 Вклад компонентов в качество\n(модель: {best['name'][:25]})",
                         fontsize=12, fontweight='bold')
            ax.grid(axis='x', alpha=0.3, linestyle='--')
            ax.set_axisbelow(True)

            for i, v in enumerate(row_contributions):
                ax.text(v + (0.01 if v >= 0 else -0.01), i, f'{v:.4f}',
                        va='center', fontsize=8, fontweight='bold')

    def _plot_weights_distribution(self, ax):
        if not self.base:
            return

        # ✅ КОНВЕРТИРУЕМ в numpy массив
        base_numeric = np.array(extract_numeric_rows(self.base), dtype=float)
        m, n = base_numeric.shape
        weights = calculate_weights(m, n)
        weights_matrix = weights.reshape(m, n)

        im = ax.imshow(weights_matrix, cmap='YlOrRd', aspect='auto')
        ax.set_xlabel("Столбец (параметр)", fontsize=11, fontweight='bold')
        ax.set_ylabel("Строка (функция/ТС)", fontsize=11, fontweight='bold')
        ax.set_title("⚖️ Распределение весовых коэффициентов (ql)", fontsize=12, fontweight='bold')

        for i in range(m):
            for j in range(n):
                ax.text(j, i, f'{weights_matrix[i, j]:.3f}', ha='center', va='center',
                        fontsize=8, color='black')

        plt.colorbar(im, ax=ax, label='Значение ql')

    def _plot_minmax_comparison(self, ax):
        if not self.detailed_results:
            return

        passed = [r for r in self.detailed_results if r['status'] == 'PASS'][:5]

        min_sums = []
        max_sums = []
        names = []

        for result in passed:
            CQ1 = result['CQ1']
            CQ2 = result['CQ2']
            min_sums.append(np.sum(CQ1))
            max_sums.append(np.sum(CQ2))
            names.append(result['name'][:15])

        x = np.arange(len(names))
        width = 0.35

        bars1 = ax.bar(x - width / 2, min_sums, width, label='∆Q Min (CQ1)', color='#2196F3', alpha=0.8)
        bars2 = ax.bar(x + width / 2, max_sums, width, label='∆Q Max (CQ2)', color='#4CAF50', alpha=0.8)

        ax.set_xlabel("Модель", fontsize=11, fontweight='bold')
        ax.set_ylabel("Сумма отклонений", fontsize=11, fontweight='bold')
        ax.set_title("📊 Сравнение отклонений Min и Max параметров", fontsize=12, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(names, rotation=45, ha='right')
        ax.legend()
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

    # ===== ВКЛАДКА: МАТРИЦЫ =====
    def init_matrices_tab(self):
        toolbar = tk.Frame(self.tab_matrices)
        toolbar.pack(fill="x", pady=5)

        tk.Label(toolbar, text="Выберите модель:", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        self.matrix_model_var = tk.StringVar()
        self.matrix_model_combo = ttk.Combobox(toolbar, textvariable=self.matrix_model_var, width=30, state="readonly")
        self.matrix_model_combo.pack(side="left", padx=5)

        tk.Button(toolbar, text="🔢 Показать матрицы", command=self.show_matrices,
                  bg="#2196F3", fg="white").pack(side="left", padx=5)
        tk.Button(toolbar, text="📊 Тепловая карта", command=self.show_heatmaps,
                  bg="#FF9800", fg="white").pack(side="left", padx=5)

        self.matrices_frame = tk.Frame(self.tab_matrices)
        self.matrices_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.matrix_canvas = None

    def show_matrices(self):
        model_name = self.matrix_model_var.get()
        if not model_name:
            messagebox.showinfo("Инфо", "Выберите модель из списка\n\n"
                                        "Если список пустой:\n"
                                        "1. Загрузите Excel файл\n"
                                        "2. Нажмите 'Рассчитать все модели'")
            return

        result = None
        for r in self.detailed_results:
            if r['name'] == model_name:
                result = r
                break

        if not result:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт на вкладке '⚙ Расчёт'!")
            return

        for w in self.matrices_frame.winfo_children():
            w.destroy()

        text = tk.Text(self.matrices_frame, font=("Consolas", 9))
        text.pack(fill="both", expand=True)

        text.insert(tk.END, "=" * 80 + "\n")
        text.insert(tk.END, f"📊 МАТРИЦЫ ДЛЯ МОДЕЛИ: {model_name}\n")
        text.insert(tk.END, "=" * 80 + "\n\n")

        text.insert(tk.END, f"Статус: {'✅ СООТВЕТСТВУЕТ' if result['status'] == 'PASS' else '❌ НЕ СООТВЕТСТВУЕТ'}\n")
        text.insert(tk.END, f"Qкэ = {result['Q_ke']:.6f}\n\n")

        if result['CQ'] is not None:
            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, "📋 Матрица разностей CQ (∆Q):\n")
            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, self._matrix_to_text(result['CQ']) + "\n\n")

            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, "⚖️ Матрица весовых коэффициентов (ql):\n")
            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, self._matrix_to_text(result['weights']) + "\n\n")

            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, "📈 Взвешенная матрица (ql × ∆Ql):\n")
            text.insert(tk.END, "─" * 80 + "\n")
            text.insert(tk.END, self._matrix_to_text(result['weighted']) + "\n\n")

        text.config(state=tk.DISABLED)

    def show_heatmaps(self):
        model_name = self.matrix_model_var.get()
        if not model_name:
            messagebox.showinfo("Инфо", "Выберите модель из списка")
            return

        result = None
        for r in self.detailed_results:
            if r['name'] == model_name:
                result = r
                break

        if not result or result['CQ'] is None:
            messagebox.showinfo("Инфо", "Нет данных для отображения\n\nСначала выполните расчёт!")
            return

        if self.matrix_canvas:
            self.matrix_canvas.get_tk_widget().pack_forget()
            plt.close('all')

        fig, axes = plt.subplots(1, 3, figsize=(18, 5))
        fig.suptitle(f"🔢 Тепловые карты матриц: {model_name}", fontsize=14, fontweight="bold")

        im1 = axes[0].imshow(result['CQ'], cmap='RdYlGn', aspect='auto')
        axes[0].set_title("📋 Матрица разностей CQ", fontsize=12, fontweight='bold')
        axes[0].set_xlabel("Столбец")
        axes[0].set_ylabel("Строка")
        plt.colorbar(im1, ax=axes[0])

        im2 = axes[1].imshow(result['weights'], cmap='YlOrRd', aspect='auto')
        axes[1].set_title("⚖️ Весовые коэффициенты ql", fontsize=12, fontweight='bold')
        axes[1].set_xlabel("Столбец")
        axes[1].set_ylabel("Строка")
        plt.colorbar(im2, ax=axes[1])

        im3 = axes[2].imshow(result['weighted'], cmap='RdYlGn', aspect='auto')
        axes[2].set_title("📈 Взвешенная матрица (ql × ∆Ql)", fontsize=12, fontweight='bold')
        axes[2].set_xlabel("Столбец")
        axes[2].set_ylabel("Строка")
        plt.colorbar(im3, ax=axes[2])

        plt.tight_layout(rect=[0, 0, 1, 0.92])

        self.matrix_canvas = FigureCanvasTkAgg(fig, master=self.matrices_frame)
        self.matrix_canvas.draw()
        self.matrix_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _matrix_to_text(self, matrix):
        if matrix is None:
            return "Данные отсутствуют"

        lines = []
        for i in range(matrix.shape[0]):
            row = " | ".join([f"{matrix[i, j]:8.4f}" for j in range(matrix.shape[1])])
            lines.append(f"Строка {i + 1}: {row}")
        return "\n".join(lines)

    # ===== ЭКСПОРТ ОТЧЁТА =====
    def export_full_report(self):
        if not self.detailed_results:
            messagebox.showinfo("Инфо", "Сначала выполните расчёт на вкладке 'Расчёт'")
            return

        file = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        if not file:
            return

        try:
            fig_bar, fig_radar, fig_params, fig_components = self._generate_report_figures()
            bar_img = self._fig_to_base64(fig_bar)
            radar_img = self._fig_to_base64(fig_radar)
            params_img = self._fig_to_base64(fig_params)
            components_img = self._fig_to_base64(fig_components)
            html_content = self._generate_html_report(bar_img, radar_img, params_img, components_img)

            with open(file, "w", encoding="utf-8") as f:
                f.write(html_content)

            messagebox.showinfo("Успех", f"Полный отчёт сохранён: {file}\n\n"
                                         f"📌 Откройте файл в браузере для просмотра.\n"
                                         f"📌 Для сохранения в PDF: Ctrl+P → Сохранить как PDF")
            self.status.config(text=f"📄 Отчёт экспортирован: {os.path.basename(file)}")

            plt.close(fig_bar)
            plt.close(fig_radar)
            plt.close(fig_params)
            plt.close(fig_components)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчёт:\n{str(e)}")
            import traceback
            traceback.print_exc()

    def _fig_to_base64(self, fig):
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode('utf-8')
        return img_base64

    def _generate_report_figures(self):
        passed_results = [r for r in self.detailed_results if r['status'] == 'PASS']

        fig_bar, ax1 = plt.subplots(figsize=(10, 6))
        self._plot_rating_chart(ax1)
        plt.tight_layout()

        fig_radar, ax2 = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
        if passed_results:
            self._plot_radar_chart(ax2, passed_results[0])
        plt.tight_layout()

        fig_params, ax3 = plt.subplots(figsize=(12, 6))
        self._plot_params_chart(ax3)
        plt.tight_layout()

        fig_components, ax4 = plt.subplots(figsize=(10, 6))
        self._plot_components_chart(ax4)
        plt.tight_layout()

        return fig_bar, fig_radar, fig_params, fig_components

    def _generate_html_report(self, bar_img, radar_img, params_img, components_img):
        rating_rows = ""
        for i, result in enumerate(self.detailed_results, 1):
            if result['status'] == 'PASS':
                medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "•"
                rating_rows += f"""
                <tr>
                    <td style="text-align:center; font-size:18px;">{medal}</td>
                    <td style="text-align:center; font-weight:bold;">{i}</td>
                    <td>{result['name']}</td>
                    <td style="text-align:center; color:#4CAF50; font-weight:bold;">{result['Q_ke']:.6f}</td>
                    <td style="text-align:center;"><span style="color:#4CAF50; font-weight:bold;">✅ СООТВЕТСТВУЕТ</span></td>
                </tr>"""
            else:
                rating_rows += f"""
                <tr>
                    <td style="text-align:center; font-size:18px;">❌</td>
                    <td style="text-align:center;">-</td>
                    <td>{result['name']}</td>
                    <td style="text-align:center; color:#f44336;">0.000000</td>
                    <td style="text-align:center;"><span style="color:#f44336; font-weight:bold;">НЕ СООТВЕТСТВУЕТ</span></td>
                </tr>"""

        html = f"""
        <!DOCTYPE html>
        <html lang="ru">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Отчёт по оценке качества СФМ ВОЛП</title>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; max-width: 1400px; margin: 0 auto; padding: 20px; background-color: #f5f5f5; }}
                .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; text-align: center; margin-bottom: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
                .header h1 {{ margin: 0; font-size: 28px; }}
                .section {{ background: white; padding: 25px; margin-bottom: 25px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
                .section h2 {{ color: #667eea; border-bottom: 2px solid #667eea; padding-bottom: 10px; margin-top: 0; }}
                table {{ width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px; }}
                th, td {{ border: 1px solid #ddd; padding: 10px; text-align: center; }}
                th {{ background-color: #667eea; color: white; font-weight: bold; }}
                .graphs {{ display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; }}
                .graph-container {{ background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex: 1; min-width: 450px; }}
                .graph-container img {{ width: 100%; height: auto; }}
                .footer {{ text-align: center; margin-top: 40px; padding: 20px; color: #666; font-size: 12px; }}
                .print-btn {{ background: #667eea; color: white; border: none; padding: 12px 25px; font-size: 14px; border-radius: 5px; cursor: pointer; margin: 20px 0; }}
                @media print {{ .print-btn {{ display: none; }} body {{ background: white; }} .section {{ box-shadow: none; border: 1px solid #ddd; }} }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>📡 ОТЧЁТ ПО ОЦЕНКЕ КАЧЕСТВА СФМ ВОЛП</h1>
                <p>Компьютерное структурно-функциональное моделирование волоконно-оптических линий передачи</p>
                <p><strong>Дата формирования:</strong> {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}</p>
            </div>
            <button class="print-btn" onclick="window.print()">🖨️ Печать / Сохранить в PDF</button>
            <div class="section">
                <h2>🏆 РЕЙТИНГ КАЧЕСТВА МОДЕЛЕЙ</h2>
                <table>
                    <thead>
                        <tr>
                            <th style="width:50px;">🏅</th>
                            <th style="width:50px;">№</th>
                            <th>Наименование модели</th>
                            <th style="width:120px;">Qкэ</th>
                            <th style="width:150px;">Статус</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rating_rows}
                    </tbody>
                </table>
            </div>
            <div class="section">
                <h2>📊 ГРАФИЧЕСКАЯ ВИЗУАЛИЗАЦИЯ</h2>
                <div class="graphs">
                    <div class="graph-container">
                        <h3 style="text-align:center; color:#667eea;">Рейтинг моделей (Qкэ)</h3>
                        <img src="data:image/png;base64,{bar_img}" alt="Рейтинг моделей">
                    </div>
                    <div class="graph-container">
                        <h3 style="text-align:center; color:#764ba2;">Радар разностей (лучшая модель)</h3>
                        <img src="data:image/png;base64,{radar_img}" alt="Радар разностей">
                    </div>
                    <div class="graph-container">
                        <h3 style="text-align:center; color:#2196F3;">Сравнение параметров с базой</h3>
                        <img src="data:image/png;base64,{params_img}" alt="Параметры">
                    </div>
                    <div class="graph-container">
                        <h3 style="text-align:center; color:#4CAF50;">Вклад компонентов в качество</h3>
                        <img src="data:image/png;base64,{components_img}" alt="Компоненты">
                    </div>
                </div>
            </div>
            <div class="footer">
                <p>Отчёт сформирован автоматически программой "СФМ ВОЛП: Оценка качества"</p>
                <p>Методика расчёта соответствует ТЗ Этап 2 (формулы 2-10)</p>
            </div>
        </body>
        </html>
        """
        return html

    # ===== ВКЛАДКА: СПРАВКА =====
    def init_help_tab(self):
        help_text = tk.Text(self.tab_help, wrap=tk.WORD, font=("Arial", 11), bg="#f5f5f5")
        help_text.pack(fill="both", expand=True, padx=15, pady=15)

        scrollbar = ttk.Scrollbar(help_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        help_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=help_text.yview)

        content = """
📘 СПРАВКА: СФМ ВОЛП — Оценка качества (ТЗ Этап 2)

═══════════════════════════════════════════════════════════

📋 ПОРЯДОК РАБОТЫ:
───────────────────────────────────────────────────────────
1. 📥 Загрузите Excel файл (кнопка "Загрузить Excel")
2. ⚙ Перейдите на вкладку "Расчёт"
3. 🚀 Нажмите "Рассчитать все модели"
4. 📊 Перейдите на "Графики" → "Все графики"
5. 🔢 Перейдите на "Матрицы" → выберите модель → "Показать"

⚠️ ВАЖНО:
───────────────────────────────────────────────────────────
• Список моделей в "Матрицы" заполняется ПОСЛЕ загрузки Excel!
• Графики строятся ТОЛЬКО после расчёта!
• Если список пустой — сначала нажмите "Рассчитать все модели"

📊 ГРАФИКИ (6 типов):
───────────────────────────────────────────────────────────
1. 🏆 Рейтинг моделей — бар-чарт Qкэ
2. 📈 Параметры — сравнение P1-P4 с базой
3. 🎯 Радары — матрица разностей для каждой модели
4. 📉 Компоненты — вклад ТС в качество
5. ⚖️ Веса — распределение ql
6. 📊 Min/Max — отклонения CQ1 и CQ2

🔢 МАТРИЦЫ:
───────────────────────────────────────────────────────────
• CQ — матрица разностей (∆Q)
• ql — весовые коэффициенты
• ql × ∆Ql — взвешенная матрица
• Тепловые карты для визуализации

📁 ФОРМАТЫ ФАЙЛОВ:
───────────────────────────────────────────────────────────
• Excel (.xlsx): лист "Base" + листы моделей
• JSON (.json): полный проект
• HTML (.html): отчёт с графиками

═══════════════════════════════════════════════════════════
Версия: 3.2 (ПОЛНОСТЬЮ РАБОЧАЯ) | 2026
        """

        help_text.insert(tk.END, content)
        help_text.config(state=tk.DISABLED)


# ================== ЗАПУСК ==================
if __name__ == "__main__":
    root = tk.Tk()

    try:
        root.iconbitmap(default='')
    except:
        pass

    app = App(root)


    def on_closing():
        if messagebox.askokcancel("Выход", "Сохранить проект перед выходом?"):
            app.save_project()
        root.destroy()


    root.protocol("WM_DELETE_WINDOW", on_closing)

    print("🚀 Запуск приложения СФМ ВОЛП...")
    print("💡 Загрузите файл VOLP_7_Models_Test.xlsx для начала работы")
    print("📊 6 типов графиков для детальной визуализации")
    print("🔢 Вкладка 'Матрицы' для просмотра всех расчётных матриц")
    print("\n⚠️ ВАЖНО:")
    print("1. После загрузки Excel нажмите '⚙ Расчёт' → '🚀 Рассчитать все модели'")
    print("2. Только после расчёта будут доступны графики и матрицы!")

    root.mainloop()