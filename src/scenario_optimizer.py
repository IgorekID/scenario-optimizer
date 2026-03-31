import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import os


class ScenarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Оптимизатор сценариев реагирования")
        self.root.geometry("1200x900")

        self.df = pd.DataFrame(columns=[
            'Name', 'Cost', 'Time', 'Personnel', 'Risk',
            'Complexity', 'Reliability'
        ])

        # ===== СЦЕНАРИИ =====
        tk.Label(root, text="Сценарии реагирования", font=("Arial", 14)).pack()
        table_frame = tk.Frame(root, height=200)
        table_frame.pack(fill="x")
        columns = ('No', 'Name', 'Cost', 'Time', 'Personnel', 'Risk', 'Complexity', 'Reliability')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=6)
        headings = [
            "№", "Название", "Стоимость (руб)", "Время (ч)",
            "Персонал", "Снижение риска", "Сложность", "Надежность"
        ]
        for col, head in zip(columns, headings):
            self.tree.heading(col, text=head)
        self.tree.pack(fill='x')

        frame_buttons = tk.Frame(root)
        frame_buttons.pack(pady=5)
        tk.Button(frame_buttons, text="Добавить", command=self.add_scenario).pack(side='left', padx=5)
        tk.Button(frame_buttons, text="Удалить", command=self.delete_scenario).pack(side='left', padx=5)
        tk.Button(frame_buttons, text="Загрузить шаблон", command=self.load_template).pack(side='left', padx=5)
        tk.Button(frame_buttons, text="Загрузить Excel", command=self.load_from_excel).pack(side='left', padx=5)

        # ===== ОГРАНИЧЕНИЯ =====
        tk.Label(root, text="Ограничения ресурсов", font=("Arial", 12)).pack()
        frame_constraints = tk.Frame(root)
        frame_constraints.pack()
        tk.Label(frame_constraints, text="Макс. бюджет").pack(side='left')
        self.max_budget = tk.Entry(frame_constraints, width=10)
        self.max_budget.pack(side='left', padx=5)
        tk.Label(frame_constraints, text="Макс. время").pack(side='left')
        self.max_time = tk.Entry(frame_constraints, width=10)
        self.max_time.pack(side='left', padx=5)
        tk.Label(frame_constraints, text="Макс. персонал").pack(side='left')
        self.max_personnel = tk.Entry(frame_constraints, width=10)
        self.max_personnel.pack(side='left', padx=5)

        # ===== ВЕСА =====
        tk.Label(root, text="Веса критериев", font=("Arial", 12)).pack()
        frame_weights = tk.Frame(root)
        frame_weights.pack()
        self.w_risk = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal',
                               label="Снижение риска")
        self.w_risk.set(0.4)
        self.w_risk.pack(side='left', padx=10)
        self.w_cost = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal', label="Стоимость")
        self.w_cost.set(0.3)
        self.w_cost.pack(side='left', padx=10)
        self.w_time = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal', label="Время реакции")
        self.w_time.set(0.3)
        self.w_time.pack(side='left', padx=10)

        # ===== КНОПКА РАСЧЕТА =====
        tk.Button(root, text="Найти оптимальный сценарий", command=self.calculate).pack(pady=10)

        # ===== ЭКСПОРТ =====
        frame_export = tk.Frame(root)
        frame_export.pack()
        tk.Button(frame_export, text="Экспорт Excel", command=self.export_excel).pack(side='left', padx=10)
        tk.Button(frame_export, text="Экспорт PNG", command=self.export_png).pack(side='left', padx=10)
        tk.Button(frame_export, text="Экспорт PDF", command=self.export_pdf).pack(side='left', padx=10)

        # ===== РЕЗУЛЬТАТ (теперь в отдельной строке сверху) =====
        self.result_frame = tk.Frame(root, bg='lightblue', height=60)
        self.result_frame.pack(fill='x', pady=5)
        self.result_label = tk.Label(self.result_frame, text="", font=("Arial", 10),
                                     bg='lightblue', justify="left")
        self.result_label.pack(side='left', padx=10, pady=5)

        # ===== ГРАФИКИ =====
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        self.tab_radar = tk.Frame(self.notebook)
        self.tab_bar = tk.Frame(self.notebook)
        self.tab_resources = tk.Frame(self.notebook)
        self.notebook.add(self.tab_radar, text="Радарная диаграмма")
        self.notebook.add(self.tab_bar, text="Рейтинг сценариев")
        self.notebook.add(self.tab_resources, text="Использование ресурсов")

    def update_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, row in self.df.iterrows():
            self.tree.insert('', 'end', values=(
                idx + 1,
                row['Name'],
                row['Cost'],
                row['Time'],
                row['Personnel'],
                row['Risk'],
                row['Complexity'],
                row['Reliability']
            ))

    def add_scenario(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить сценарий")
        entries = {}
        fields = [
            ("Название", "Name"),
            ("Стоимость (руб)", "Cost"),
            ("Время (ч)", "Time"),
            ("Персонал", "Personnel"),
            ("Снижение риска (0-1)", "Risk"),
            ("Сложность (0-1)", "Complexity"),
            ("Надежность (0-1)", "Reliability")
        ]
        for label_text, key in fields:
            tk.Label(dialog, text=label_text).pack()
            entry = tk.Entry(dialog)
            entry.pack()
            entries[key] = entry

        def save():
            try:
                new_row = {}
                for k in entries:
                    if k == "Name":
                        new_row[k] = entries[k].get()
                    else:
                        new_row[k] = float(entries[k].get())
                self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
                self.update_tree()
                dialog.destroy()
            except:
                messagebox.showerror("Ошибка", "Неверные данные")

        tk.Button(dialog, text="Сохранить", command=save).pack(pady=10)

    def delete_scenario(self):
        selected = self.tree.selection()
        if selected:
            idx = int(self.tree.item(selected)['values'][0]) - 1
            self.df = self.df.drop(idx).reset_index(drop=True)
            self.update_tree()

    def load_template(self):
        data = [
            {'Name': 'Активация резервной системы', 'Cost': 900000, 'Time': 4, 'Personnel': 3,
             'Risk': 0.8, 'Complexity': 0.6, 'Reliability': 0.9},
            {'Name': 'Анализ угрозы', 'Cost': 500000, 'Time': 8, 'Personnel': 2,
             'Risk': 0.5, 'Complexity': 0.3, 'Reliability': 0.6}
        ]
        self.df = pd.DataFrame(data)
        self.update_tree()

    def load_from_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            df_new = pd.read_excel(file)
            self.df = pd.concat([self.df, df_new], ignore_index=True)
            self.update_tree()

    def calculate(self):
        try:
            max_budget = float(self.max_budget.get())
            max_time = float(self.max_time.get())
            max_personnel = float(self.max_personnel.get())
        except:
            messagebox.showerror("Ошибка", "Введите ограничения ресурсов")
            return

        if self.df.empty:
            messagebox.showwarning("Нет данных", "Добавьте хотя бы один сценарий")
            return

        df = self.df.copy()
        eps = 1e-6

        # Нормализация критериев
        df['v_cost'] = (df['Cost'].max() - df['Cost']) / (df['Cost'].max() - df['Cost'].min() + eps)
        df['v_time'] = (df['Time'].max() - df['Time']) / (df['Time'].max() - df['Time'].min() + eps)
        df['v_risk'] = (df['Risk'] - df['Risk'].min()) / (df['Risk'].max() - df['Risk'].min() + eps)

        # ===== РАБОТА С ВЕСАМИ =====
        w_risk = self.w_risk.get()
        w_cost = self.w_cost.get()
        w_time = self.w_time.get()

        total_w = w_risk + w_cost + w_time
        if total_w > 1.0:
            messagebox.showwarning(
                "Внимание",
                f"Сумма весов = {total_w:.2f} > 1.0\n"
                "Веса будут автоматически нормализованы к сумме 1.0"
            )

        if total_w > 0:
            w_risk /= total_w
            w_cost /= total_w
            w_time /= total_w
        else:
            w_risk = w_cost = w_time = 1.0 / 3

        # Расчёт рейтинга R
        df['R'] = w_risk * df['v_risk'] + w_cost * df['v_cost'] + w_time * df['v_time']

        # Фильтрация допустимых сценариев
        feasible = df[
            (df['Cost'] <= max_budget) &
            (df['Time'] <= max_time) &
            (df['Personnel'] <= max_personnel)
            ]

        if feasible.empty:
            self.result_label.config(text="Нет допустимых сценариев в рамках ограничений")
            return

        # Выбор оптимального
        optimal = feasible.loc[feasible['R'].idxmax()]

        # ТЕПЕРЬ ИНФОРМАЦИЯ В ОДНУ СТРОКУ
        text = f"Рекомендуемый сценарий: {optimal['Name']} | Стоимость: {optimal['Cost']} руб | Снижение риска: {optimal['Risk'] * 100:.1f}% | Бюджет: {optimal['Cost']}/{max_budget} | Время: {optimal['Time']}/{max_time} | Персонал: {optimal['Personnel']}/{max_personnel} | Рейтинг R: {optimal['R']:.3f}"
        self.result_label.config(text=text)

        # Построение графиков
        self.plot_radar(df)
        self.plot_bar(df, optimal['Name'])
        self.plot_resources(df, max_budget, max_time, max_personnel)

    def plot_radar(self, df):
        for w in self.tab_radar.winfo_children():
            w.destroy()

        categories = ['Снижение риска', 'Эффективность стоимости', 'Эффективность времени', 'Сложность', 'Надежность']
        angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False)
        angles = np.concatenate((angles, [angles[0]]))

        fig = plt.Figure(figsize=(10, 8))  # Немного увеличим размер
        ax = fig.add_subplot(111, polar=True)

        for _, row in df.iterrows():
            values = [
                row['Risk'],
                1 - row['v_cost'],
                1 - row['v_time'],
                row['Complexity'],
                row['Reliability']
            ]
            values.append(values[0])
            ax.plot(angles, values, linewidth=2, label=row['Name'])
            ax.fill(angles, values, alpha=0.1)

        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories)

        # Сдвигаем легенду дальше вправо
        ax.legend(loc='center left', bbox_to_anchor=(1.3, 0.5), fontsize=10)

        # Увеличиваем правый отступ для легенды
        fig.subplots_adjust(right=0.75)

        # Добавляем заголовок
        ax.set_title('Радарная диаграмма сценариев', pad=20, fontsize=14)

        self.fig_radar = fig
        canvas = FigureCanvasTkAgg(fig, master=self.tab_radar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    def plot_bar(self, df, optimal_name):
        for w in self.tab_bar.winfo_children():
            w.destroy()

        fig = plt.Figure(figsize=(8, 7))
        ax = fig.add_subplot(111)
        colors = ['green' if n == optimal_name else 'gray' for n in df['Name']]
        ax.bar(df['Name'], df['R'], color=colors)
        ax.set_ylabel("Рейтинг R")

        # Горизонтальные подписи
        ax.set_xticklabels(df['Name'], rotation=0, ha='center')

        # Увеличиваем нижний отступ для горизонтальных подписей
        fig.subplots_adjust(bottom=0.2)

        self.fig_bar = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_bar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot_resources(self, df, max_budget, max_time, max_personnel):
        for w in self.tab_resources.winfo_children():
            w.destroy()

        # Увеличиваем высоту графика и поднимаем его выше
        fig = plt.Figure(figsize=(8, 6))
        ax = fig.add_subplot(111)
        x = np.arange(len(df))

        ax.bar(x, df['Cost'] / max_budget, label='Бюджет')
        ax.bar(x, df['Time'] / max_time, bottom=df['Cost'] / max_budget, label='Время')
        ax.bar(x, df['Personnel'] / max_personnel,
               bottom=(df['Cost'] / max_budget + df['Time'] / max_time),
               label='Персонал')

        ax.set_xticks(x)
        # Горизонтальные подписи
        ax.set_xticklabels(df['Name'], rotation=0, ha='center')
        ax.axhline(1, color='red', linestyle='--', label='Лимит')
        ax.legend()

        # Увеличиваем нижний отступ для горизонтальных подписей
        fig.subplots_adjust(bottom=0.2)

        # Добавляем заголовок для ясности
        ax.set_title('Использование ресурсов по сценариям', fontsize=12, pad=20)

        self.fig_resources = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_resources)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot_resources(self, df, max_budget, max_time, max_personnel):
        for w in self.tab_resources.winfo_children():
            w.destroy()

        # Увеличиваем ширину графика для лучшего размещения подписей
        fig = plt.Figure(figsize=(10, 6))
        ax = fig.add_subplot(111)
        x = np.arange(len(df))

        ax.bar(x, df['Cost'] / max_budget, label='Бюджет')
        ax.bar(x, df['Time'] / max_time, bottom=df['Cost'] / max_budget, label='Время')
        ax.bar(x, df['Personnel'] / max_personnel,
               bottom=(df['Cost'] / max_budget + df['Time'] / max_time),
               label='Персонал')

        ax.set_xticks(x)
        # Горизонтальные подписи с переносом длинных названий
        labels = [name.replace(' ', '\n') for name in df['Name']]  # Заменяем пробелы на перенос строки
        ax.set_xticklabels(labels, rotation=0, ha='center')
        ax.axhline(1, color='red', linestyle='--', label='Лимит')
        ax.legend()

        # Увеличиваем нижний отступ для многострочных подписей
        fig.subplots_adjust(bottom=0.25)

        # Добавляем заголовок для ясности
        ax.set_title('Использование ресурсов по сценариям', fontsize=12, pad=20)

        # Добавляем подписи осей для наглядности
        ax.set_ylabel('Относительное использование ресурсов')

        self.fig_resources = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_resources)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def export_excel(self):
        if self.df.empty:
            messagebox.showwarning("Нет данных", "Нечего экспортировать")
            return
        file = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if file:
            self.df.to_excel(file, index=False)

    def export_png(self):
        folder = filedialog.askdirectory()
        if folder:
            if hasattr(self, "fig_radar"):
                self.fig_radar.savefig(os.path.join(folder, "radar.png"))
            if hasattr(self, "fig_bar"):
                self.fig_bar.savefig(os.path.join(folder, "rating.png"))
            if hasattr(self, "fig_resources"):
                self.fig_resources.savefig(os.path.join(folder, "resources.png"))

    def export_pdf(self):
        file = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not file:
            return
        with PdfPages(file) as pdf:
            if hasattr(self, "fig_radar"):
                pdf.savefig(self.fig_radar)
            if hasattr(self, "fig_bar"):
                pdf.savefig(self.fig_bar)
            if hasattr(self, "fig_resources"):
                pdf.savefig(self.fig_resources)


if __name__ == "__main__":
    root = tk.Tk()
    app = ScenarioApp(root)
    root.mainloop()