import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import sys

# To make it executable on Windows and Linux, use pyinstaller:
# Install pyinstaller: pip install pyinstaller
# Then run: pyinstaller --onefile --windowed scenario_optimizer.py
# This will create a dist folder with the executable.

class ScenarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Оптимизатор сценариев")
        self.root.geometry("1200x800")

        # DataFrame for scenarios
        self.df = pd.DataFrame(columns=['Name', 'Cost', 'Time', 'Personnel', 'Risk'])

        # Scenarios section
        tk.Label(root, text="Сценарии реагирования").pack()
        self.tree = ttk.Treeview(root, columns=('No', 'Name', 'Cost', 'Time', 'Personnel', 'Risk'), show='headings')
        self.tree.heading('No', text='№')
        self.tree.heading('Name', text='Название')
        self.tree.heading('Cost', text='Стоимость (руб)')
        self.tree.heading('Time', text='Время (часы)')
        self.tree.heading('Personnel', text='Персонал (чел)')
        self.tree.heading('Risk', text='Снижение риска (0-1)')
        self.tree.pack(fill='both', expand=True)

        # Buttons for scenarios
        frame_buttons = tk.Frame(root)
        frame_buttons.pack()
        tk.Button(frame_buttons, text="Добавить сценарий", command=self.add_scenario).pack(side='left')
        tk.Button(frame_buttons, text="Удалить сценарий", command=self.delete_scenario).pack(side='left')
        tk.Button(frame_buttons, text="Загрузить шаблон", command=self.load_template).pack(side='left')

        # Constraints section
        tk.Label(root, text="Ограничения по ресурсам").pack()
        frame_constraints = tk.Frame(root)
        frame_constraints.pack()
        tk.Label(frame_constraints, text="Макс. бюджет (руб):").pack(side='left')
        self.max_budget = tk.Entry(frame_constraints)
        self.max_budget.pack(side='left')
        tk.Label(frame_constraints, text="Макс. время (часы):").pack(side='left')
        self.max_time = tk.Entry(frame_constraints)
        self.max_time.pack(side='left')
        tk.Label(frame_constraints, text="Макс. персонал (чел):").pack(side='left')
        self.max_personnel = tk.Entry(frame_constraints)
        self.max_personnel.pack(side='left')

        # Weights section
        tk.Label(root, text="Веса критериев").pack()
        frame_weights = tk.Frame(root)
        frame_weights.pack()
        tk.Label(frame_weights, text="Снижение риска:").pack(side='left')
        self.w_risk = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal')
        self.w_risk.set(0.5)
        self.w_risk.pack(side='left')
        tk.Label(frame_weights, text="Стоимость:").pack(side='left')
        self.w_cost = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal')
        self.w_cost.set(0.3)
        self.w_cost.pack(side='left')
        tk.Label(frame_weights, text="Время:").pack(side='left')
        self.w_time = tk.Scale(frame_weights, from_=0, to=1, resolution=0.1, orient='horizontal')
        self.w_time.set(0.2)
        self.w_time.pack(side='left')

        # Calculate button
        tk.Button(root, text="Найти оптимальный сценарий", command=self.calculate).pack()

        # Result section
        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack()

        # Graphs notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        self.tab_radar = tk.Frame(self.notebook)
        self.notebook.add(self.tab_radar, text='Радарная диаграмма')

        self.tab_bar = tk.Frame(self.notebook)
        self.notebook.add(self.tab_bar, text='Столбчатая диаграмма рейтинга')

        self.tab_resources = tk.Frame(self.notebook)
        self.notebook.add(self.tab_resources, text='Использование ресурсов')

        # Export buttons
        frame_export = tk.Frame(root)
        frame_export.pack()
        tk.Button(frame_export, text="Экспорт данных в Excel", command=self.export_excel).pack(side='left')
        tk.Button(frame_export, text="Экспорт графиков в PNG", command=self.export_png).pack(side='left')

    def update_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, row in self.df.iterrows():
            self.tree.insert('', 'end', values=(idx+1, row['Name'], row['Cost'], row['Time'], row['Personnel'], row['Risk']))

    def add_scenario(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить сценарий")

        tk.Label(dialog, text="Название:").pack()
        name_entry = tk.Entry(dialog)
        name_entry.pack()

        tk.Label(dialog, text="Стоимость (руб):").pack()
        cost_entry = tk.Entry(dialog)
        cost_entry.pack()

        tk.Label(dialog, text="Время (часы):").pack()
        time_entry = tk.Entry(dialog)
        time_entry.pack()

        tk.Label(dialog, text="Персонал (чел):").pack()
        personnel_entry = tk.Entry(dialog)
        personnel_entry.pack()

        tk.Label(dialog, text="Снижение риска (0-1):").pack()
        risk_entry = tk.Entry(dialog)
        risk_entry.pack()

        def save():
            try:
                new_row = {
                    'Name': name_entry.get(),
                    'Cost': float(cost_entry.get()),
                    'Time': float(time_entry.get()),
                    'Personnel': float(personnel_entry.get()),
                    'Risk': float(risk_entry.get())
                }
                self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
                self.update_tree()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Ошибка", "Неверные значения ввода")

        tk.Button(dialog, text="Сохранить", command=save).pack()

    def delete_scenario(self):
        selected = self.tree.selection()
        if selected:
            idx = int(self.tree.item(selected)['values'][0]) - 1
            self.df = self.df.drop(idx).reset_index(drop=True)
            self.update_tree()
        else:
            messagebox.showwarning("Предупреждение", "Сценарий не выбран")

    def load_template(self):
        # Load example data from TZ
        data = [
            {'Name': 'Активация резервной системы', 'Cost': 900000, 'Time': 4, 'Personnel': 3, 'Risk': 0.8},
            {'Name': 'Анализ угрозы', 'Cost': 500000, 'Time': 8, 'Personnel': 2, 'Risk': 0.5}
        ]
        self.df = pd.DataFrame(data)
        self.update_tree()
        self.max_budget.delete(0, tk.END)
        self.max_budget.insert(0, '800000')
        self.max_time.delete(0, tk.END)
        self.max_time.insert(0, '6')
        self.max_personnel.delete(0, tk.END)
        self.max_personnel.insert(0, '3')

    def calculate(self):
        try:
            max_budget = float(self.max_budget.get())
            max_time = float(self.max_time.get())
            max_personnel = float(self.max_personnel.get())
            w_risk = self.w_risk.get()
            w_cost = self.w_cost.get()
            w_time = self.w_time.get()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверные значения ограничений")
            return

        if self.df.empty:
            messagebox.showwarning("Предупреждение", "Нет добавленных сценариев")
            return

        df = self.df.astype({'Cost': 'float', 'Time': 'float', 'Personnel': 'float', 'Risk': 'float'})

        # Filter feasible scenarios
        feasible = df[(df['Cost'] <= max_budget) & (df['Time'] <= max_time) & (df['Personnel'] <= max_personnel)]

        if feasible.empty:
            self.result_label.config(text="Нет допустимых сценариев")
            return

        # Normalize across all for graphs, but for selection use feasible
        eps = 1e-6
        min_cost = df['Cost'].min()
        max_cost = df['Cost'].max()
        min_time = df['Time'].min()
        max_time = df['Time'].max()
        min_risk = df['Risk'].min()
        max_risk = df['Risk'].max()

        df['v_cost'] = (max_cost - df['Cost']) / (max_cost - min_cost + eps)
        df['v_time'] = (max_time - df['Time']) / (max_time - min_time + eps)
        df['v_risk'] = (df['Risk'] - min_risk) / (max_risk - min_risk + eps)
        df['R'] = w_risk * df['v_risk'] + w_cost * df['v_cost'] + w_time * df['v_time']

        # Recalculate for feasible
        feasible = df[(df['Cost'] <= max_budget) & (df['Time'] <= max_time) & (df['Personnel'] <= max_personnel)]
        optimal_idx = feasible['R'].idxmax()
        optimal = feasible.loc[optimal_idx]

        result_text = f"Рекомендуемый сценарий: {optimal['Name']}\n"
        result_text += f"Общие затраты: {optimal['Cost']} руб\n"
        result_text += f"Риск снижен на: {optimal['Risk'] * 100}%\n"
        result_text += f"Ресурсы: Бюджет {optimal['Cost']}/{max_budget}, Время {optimal['Time']}/{max_time}, Персонал {optimal['Personnel']}/{max_personnel}"
        self.result_label.config(text=result_text)

        # Plot graphs
        self.plot_radar(df)
        self.plot_bar(df, optimal['Name'])
        self.plot_resources(df, max_budget, max_time, max_personnel)

    def plot_radar(self, df):
        for widget in self.tab_radar.winfo_children():
            widget.destroy()

        categories = ['Снижение риска', 'Эффективность стоимости', 'Эффективность времени']
        N = len(categories)
        angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
        angles += angles[:1]

        fig = plt.Figure(figsize=(6, 6))
        ax = fig.add_subplot(111, polar=True)
        for idx, row in df.iterrows():
            values = [row['v_risk'], row['v_cost'], row['v_time']]
            values += values[:1]
            ax.plot(angles, values, linewidth=1, linestyle='solid', label=row['Name'])
            ax.fill(angles, values, alpha=0.1)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories)
        ax.set_rlabel_position(30)
        ax.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
        self.fig_radar = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_radar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot_bar(self, df, optimal_name):
        for widget in self.tab_bar.winfo_children():
            widget.destroy()

        fig = plt.Figure(figsize=(6, 6))
        ax = fig.add_subplot(111)
        colors = ['green' if name == optimal_name else 'gray' for name in df['Name']]
        ax.bar(df['Name'], df['R'], color=colors)
        ax.set_xlabel('Сценарии')
        ax.set_ylabel('Рейтинг')
        self.fig_bar = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_bar)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot_resources(self, df, max_budget, max_time, max_personnel):
        for widget in self.tab_resources.winfo_children():
            widget.destroy()

        resources = ['Бюджет', 'Время', 'Персонал']
        norm_values = pd.DataFrame({
            'Бюджет': df['Cost'] / max_budget,
            'Время': df['Time'] / max_time,
            'Персонал': df['Personnel'] / max_personnel
        })

        fig = plt.Figure(figsize=(6, 6))
        ax = fig.add_subplot(111)
        x = np.arange(len(resources))
        width = 0.15
        for i, (name, row) in enumerate(norm_values.iterrows()):
            offset = width * (i - len(df)/2)
            ax.bar(x + offset, row, width, label=df.loc[i, 'Name'])

        ax.axhline(1, color='red', linestyle='--', label='Лимит')
        ax.set_xticks(x)
        ax.set_xticklabels(resources)
        ax.set_ylabel('Нормированное использование')
        ax.legend()
        self.fig_resources = fig

        canvas = FigureCanvasTkAgg(fig, master=self.tab_resources)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def export_excel(self):
        if self.df.empty:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        file = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if file:
            self.df.to_excel(file, index=False)

    def export_png(self):
        folder = filedialog.askdirectory()
        if folder:
            if hasattr(self, 'fig_radar'):
                self.fig_radar.savefig(os.path.join(folder, 'radar.png'))
            if hasattr(self, 'fig_bar'):
                self.fig_bar.savefig(os.path.join(folder, 'bar.png'))
            if hasattr(self, 'fig_resources'):
                self.fig_resources.savefig(os.path.join(folder, 'resources.png'))

if __name__ == "__main__":
    # For Windows, to avoid issues with multiprocessing in pyinstaller
    if sys.platform.startswith('win'):
        import multiprocessing
        multiprocessing.freeze_support()
    root = tk.Tk()
    app = ScenarioApp(root)
    root.mainloop()