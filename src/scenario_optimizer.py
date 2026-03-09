import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import os
import sys
class ScenarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Оптимизатор сценариев реагирования")
        self.root.geometry("1200x850")
        self.df = pd.DataFrame(columns=[
            'Name', 'Cost', 'Time', 'Personnel',
            'Risk', 'Complexity', 'Reliability'
        ])
        # ===== Сценарии =====
        tk.Label(root, text="Сценарии реагирования", font=("Arial", 14)).pack()
        columns = (
            'No', 'Name', 'Cost', 'Time',
            'Personnel', 'Risk', 'Complexity', 'Reliability'
        )
        self.tree = ttk.Treeview(root, columns=columns, show='headings')
        headings = [
            "№", "Название", "Стоимость (руб)", "Время (ч)",
            "Персонал", "Снижение риска",
            "Сложность", "Надежность"
        ]
        for col, head in zip(columns, headings):
            self.tree.heading(col, text=head)
        self.tree.pack(fill='both', expand=True)
        frame_buttons = tk.Frame(root)
        frame_buttons.pack()
        tk.Button(frame_buttons, text="Добавить", command=self.add_scenario).pack(side='left')
        tk.Button(frame_buttons, text="Удалить", command=self.delete_scenario).pack(side='left')
        tk.Button(frame_buttons, text="Загрузить шаблон", command=self.load_template).pack(side='left')
        tk.Button(frame_buttons, text="Загрузить Excel", command=self.load_from_excel).pack(side='left')
        # ===== Ограничения =====
        tk.Label(root, text="Ограничения ресурсов", font=("Arial", 12)).pack()
        frame_constraints = tk.Frame(root)
        frame_constraints.pack()
        tk.Label(frame_constraints, text="Макс. бюджет").pack(side='left')
        self.max_budget = tk.Entry(frame_constraints, width=10)
        self.max_budget.pack(side='left')
        tk.Label(frame_constraints, text="Макс. время").pack(side='left')
        self.max_time = tk.Entry(frame_constraints, width=10)
        self.max_time.pack(side='left')
        tk.Label(frame_constraints, text="Макс. персонал").pack(side='left')
        self.max_personnel = tk.Entry(frame_constraints, width=10)
        self.max_personnel.pack(side='left')
        # ===== Веса =====
        tk.Label(root, text="Веса критериев").pack()
        frame_weights = tk.Frame(root)
        frame_weights.pack()
        self.w_risk = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1,
                               orient='horizontal', label="Риск")
        self.w_risk.set(0.4)
        self.w_risk.pack(side='left')
        self.w_cost = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1,
                               orient='horizontal', label="Стоимость")
        self.w_cost.set(0.3)
        self.w_cost.pack(side='left')
        self.w_time = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1,
                               orient='horizontal', label="Время")
        self.w_time.set(0.3)
        self.w_time.pack(side='left')
        tk.Button(root, text="Найти оптимальный сценарий",
                  command=self.calculate).pack(pady=5)
        # ===== Результат =====
        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack()
        # ===== Графики =====
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        self.tab_radar = tk.Frame(self.notebook)
        self.tab_bar = tk.Frame(self.notebook)
        self.tab_resources = tk.Frame(self.notebook)
        self.notebook.add(self.tab_radar, text="Радарная диаграмма")
        self.notebook.add(self.tab_bar, text="Рейтинг сценариев")
        self.notebook.add(self.tab_resources, text="Использование ресурсов")
        # ===== Экспорт =====
        frame_export = tk.Frame(root)
        frame_export.pack()
        tk.Button(frame_export, text="Экспорт Excel",
                  command=self.export_excel).pack(side='left')
        tk.Button(frame_export, text="Экспорт PNG",
                  command=self.export_png).pack(side='left')
        tk.Button(frame_export, text="Экспорт PDF",
                  command=self.export_pdf).pack(side='left')
    # =====================================================
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
    # =====================================================
    def add_scenario(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить сценарий")
        entries = {}
        fields = [
            ("Название", "Name"),
            ("Стоимость", "Cost"),
            ("Время", "Time"),
            ("Персонал", "Personnel"),
            ("Снижение риска", "Risk"),
            ("Сложность", "Complexity"),
            ("Надежность", "Reliability")
        ]
        for label, key in fields:
            tk.Label(dialog, text=label).pack()
            entry = tk.Entry(dialog)
            entry.pack()
            entries[key] = entry
        def save():
            try:
                new_row = {k: float(entries[k].get())
                           if k != 'Name' else entries[k].get()
                           for k in entries}
                self.df = pd.concat(
                    [self.df, pd.DataFrame([new_row])],
                    ignore_index=True
                )
                self.update_tree()
                dialog.destroy()
            except:
                messagebox.showerror("Ошибка", "Неверные данные")
        tk.Button(dialog, text="Сохранить", command=save).pack()
    # =====================================================
    def delete_scenario(self):
        selected = self.tree.selection()
        if selected:
            idx = int(self.tree.item(selected)['values'][0]) - 1
            self.df = self.df.drop(idx).reset_index(drop=True)
            self.update_tree()
    # =====================================================
    def load_template(self):
        data = [
            {
                'Name': 'Активация резервной системы',
                'Cost': 900000,
                'Time': 4,
                'Personnel': 3,
                'Risk': 0.8,
                'Complexity': 0.6,
                'Reliability': 0.9
            },
            {
                'Name': 'Анализ угрозы',
                'Cost': 500000,
                'Time': 8,
                'Personnel': 2,
                'Risk': 0.5,
                'Complexity': 0.3,
                'Reliability': 0.6
            }
        ]
        self.df = pd.DataFrame(data)
        self.update_tree()
    # =====================================================
    def load_from_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            df = pd.read_excel(file)
            self.df = pd.concat([self.df, df], ignore_index=True)
            self.update_tree()
    # =====================================================
    def calculate(self):
        try:
            max_budget = float(self.max_budget.get())
            max_time = float(self.max_time.get())
            max_personnel = float(self.max_personnel.get())
        except:
            messagebox.showerror("Ошибка", "Введите ограничения")
            return
        if self.df.empty:
            messagebox.showwarning("Нет данных", "Добавьте сценарии")
            return
        df = self.df.copy()
        eps = 1e-6
        df['v_cost'] = (df['Cost'].max() - df['Cost']) / (df['Cost'].max() - df['Cost'].min() + eps)
        df['v_time'] = (df['Time'].max() - df['Time']) / (df['Time'].max() - df['Time'].min() + eps)
        df['v_risk'] = (df['Risk'] - df['Risk'].min()) / (df['Risk'].max() - df['Risk'].min() + eps)
        w_risk = self.w_risk.get()
        w_cost = self.w_cost.get()
        w_time = self.w_time.get()
        df['R'] = w_risk * df['v_risk'] + w_cost * df['v_cost'] + w_time * df['v_time']
        feasible = df[
            (df['Cost'] <= max_budget) &
            (df['Time'] <= max_time) &
            (df['Personnel'] <= max_personnel)
        ]
        if feasible.empty:
            self.result_label.config(text="Нет допустимых сценариев")
            return
        optimal = feasible.loc[feasible['R'].idxmax()]
        text = f"""
Рекомендуемый сценарий: {optimal['Name']}
Стоимость: {optimal['Cost']} руб
Снижение риска: {optimal['Risk']*100:.1f}%
Использование ресурсов
Бюджет: {optimal['Cost']} / {max_budget}
Время: {optimal['Time']} / {max_time}
Персонал: {optimal['Personnel']} / {max_personnel}
"""
        self.result_label.config(text=text)
        self.plot_radar(df)
        self.plot_bar(df, optimal['Name'])
        self.plot_resources(df, max_budget, max_time, max_personnel)
    # =====================================================
    def plot_radar(self, df):
        for w in self.tab_radar.winfo_children():
            w.destroy()
        categories = [
            'Снижение риска',
            'Стоимость',
            'Время',
            'Сложность',
            'Надежность'
        ]
        angles = np.linspace(0, 2*np.pi, len(categories), endpoint=False)
        angles = np.concatenate((angles, [angles[0]]))
        fig = plt.Figure(figsize=(6,6))
        ax = fig.add_subplot(111, polar=True)
        for _, row in df.iterrows():
            values = [
                row['Risk'],
                1-row['v_cost'],
                1-row['v_time'],
                row['Complexity'],
                row['Reliability']
            ]
            values.append(values[0])
            ax.plot(angles, values, label=row['Name'])
            ax.fill(angles, values, alpha=0.1)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        self.fig_radar = fig
        canvas = FigureCanvasTkAgg(fig, master=self.tab_radar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    # =====================================================
    def plot_bar(self, df, optimal):
        for w in self.tab_bar.winfo_children():
            w.destroy()
        fig = plt.Figure(figsize=(6,6))
        ax = fig.add_subplot(111)
        colors = ['green' if n == optimal else 'gray' for n in df['Name']]
        ax.bar(df['Name'], df['R'], color=colors)
        ax.set_ylabel("Рейтинг")
        self.fig_bar = fig
        canvas = FigureCanvasTkAgg(fig, master=self.tab_bar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    # =====================================================
    def plot_resources(self, df, max_budget, max_time, max_personnel):
        for w in self.tab_resources.winfo_children():
            w.destroy()
        fig = plt.Figure(figsize=(6,6))
        ax = fig.add_subplot(111)
        x = np.arange(len(df))
        ax.bar(x, df['Cost']/max_budget, label='Бюджет')
        ax.bar(x, df['Time']/max_time, bottom=df['Cost']/max_budget, label='Время')
        ax.bar(
            x,
            df['Personnel']/max_personnel,
            bottom=(df['Cost']/max_budget + df['Time']/max_time),
            label='Персонал'
        )
        ax.set_xticks(x)
        ax.set_xticklabels(df['Name'])
        ax.legend()
        self.fig_resources = fig
        canvas = FigureCanvasTkAgg(fig, master=self.tab_resources)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    # =====================================================
    def export_excel(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if file:
            self.df.to_excel(file, index=False)
    # =====================================================
    def export_png(self):
        folder = filedialog.askdirectory()
        if folder:
            self.fig_radar.savefig(os.path.join(folder, "radar.png"))
            self.fig_bar.savefig(os.path.join(folder, "rating.png"))
            self.fig_resources.savefig(os.path.join(folder, "resources.png"))
    # =====================================================
    def export_pdf(self):
        file = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not file:
            return
        with PdfPages(file) as pdf:
            pdf.savefig(self.fig_radar)
            pdf.savefig(self.fig_bar)
            pdf.savefig(self.fig_resources)
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ScenarioApp(root)
    root.mainloop()