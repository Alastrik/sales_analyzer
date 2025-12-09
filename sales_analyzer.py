import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
from itertools import combinations
import os
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import chardet
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle

def detect_encoding(filepath):
    with open(filepath, 'rb') as f:
        raw_data = f.read(10000)
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        if encoding is None:
            encoding = 'utf-8'
        return encoding

class SalesAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор продаж")
        self.root.geometry("600x450")
        self.df = None
        self.create_widgets()


    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Выберите файл данных для анализа продаж").pack(pady=(0, 15))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Выбрать файл", command=self.load_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Анализировать", command=self.analyze, state=tk.DISABLED).pack(side=tk.LEFT, padx=5)
        self.analyze_btn = self.root.nametowidget(btn_frame.winfo_children()[1])

        ttk.Button(frame, text="Показать графики", command=self.show_charts, state=tk.DISABLED).pack(pady=5)
        self.chart_btn = self.root.nametowidget(frame.winfo_children()[-1])
        ttk.Button(frame, text="Сохранить отчёт", command=self.save_report, state=tk.DISABLED).pack(pady=5)
        self.save_btn = self.root.nametowidget(frame.winfo_children()[-1])

        self.file_label = ttk.Label(frame, text="Файл не выбран", foreground="gray")
        self.file_label.pack(pady=(20, 0))

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл данных",
            filetypes=[
                ("CSV файлы", "*.csv"),
                ("Текстовые файлы", "*.txt"),
                ("Excel файлы", "*.xlsx"),
                ("Все файлы", "*.*")
            ]
        )
        if filepath:
            self.filepath = filepath
            self.file_label.config(text=f"Выбран: {os.path.basename(filepath)}")
            self.analyze_btn.config(state=tk.NORMAL)
            self.chart_btn.config(state=tk.DISABLED)
            self.save_btn.config(state=tk.DISABLED)
            self.df = None

    def analyze(self):
        if not self.filepath:
            messagebox.showwarning("Ошибка", "Сначала выберите файл!")
            return

        try:
            ext = os.path.splitext(self.filepath)[1].lower()
            encoding = detect_encoding(self.filepath)
            if ext == ".csv":
                self.df = pd.read_csv(self.filepath, encoding=encoding)
            elif ext == ".txt":
                with open(self.filepath, 'r', encoding=encoding, errors='replace') as f:
                    sample = f.read(1024)
                sep = ',' if ',' in sample else ('\t' if '\t' in sample else ';')
                self.df = pd.read_csv(self.filepath, sep=sep, encoding=encoding, errors='replace')
            elif ext == ".xlsx":
                self.df = pd.read_excel(self.filepath)
            else:
                raise ValueError("Неподдерживаемый формат файла")

            has_sales = {'Date', 'Total'}.issubset(self.df.columns)
            has_prices = {'Date', 'Price', 'Product'}.issubset(self.df.columns)

            if not (has_sales or has_prices):
                messagebox.showerror(
                    "Ошибка структуры",
                    "Файл должен содержать либо:\n"
                    "  • Date, Total                → для анализа ПРОДАЖ\n"
                    "  • Date, Price, Product       → для анализа ЦЕН\n\n"
                    f"Фактические колонки: {', '.join(self.df.columns)}"
                )
                return

            self.analysis_mode = 'sales' if has_sales else 'prices'

            self.df['Date'] = pd.to_datetime(self.df['Date'])
            self.df['Year'] = self.df['Date'].dt.year
            self.df['Month'] = self.df['Date'].dt.to_period('M')

            self.forecast = 0
            self.basket_rules = []

            if self.analysis_mode == 'sales':
                last_year = self.df['Year'].max()
                df_last = self.df[self.df['Year'] == last_year]
                monthly = df_last.groupby('Month')['Total'].sum().sort_index()

                if len(monthly) == 0:
                    self.forecast = 0
                elif len(monthly) <= 3:
                    self.forecast = round(monthly.mean())
                else:
                    self.forecast = round(monthly.tail(3).mean())

                if {'OrderID', 'Product'}.issubset(self.df.columns):
                    basket = self.df.groupby(['OrderID', 'Product'])['Total'].count().unstack().fillna(0)
                    basket = basket.applymap(lambda x: 1 if x > 0 else 0)
                    self.basket_rules = self.get_frequent_pairs(basket)

            else:
                sample_product = self.df['Product'].iloc[0]
                product_data = self.df[self.df['Product'] == sample_product].sort_values('Date')
                monthly_price = product_data.groupby('Month')['Price'].mean().sort_index()

                if len(monthly_price) == 0:
                    self.forecast = 0
                elif len(monthly_price) <= 3:
                    self.forecast = round(monthly_price.mean())
                else:
                    self.forecast = round(monthly_price.tail(3).mean())

            msg = "✅ Анализ завершён!\n\n"
            if self.analysis_mode == 'sales':
                msg += f"🔹 Прогноз продаж на следующий месяц: {self.forecast:,.0f} руб.\n"
                if self.basket_rules:
                    msg += "🔹 Часто покупают вместе:\n"
                    for pair, freq in self.basket_rules[:3]:
                        msg += f"   {' + '.join(pair)} — {freq} раз(а)\n"
                else:
                    msg += "🔹 Совместные покупки не обнаружены.\n"
            else:
                product_name = self.df['Product'].iloc[0]
                msg += f"🔹 Прогноз цены на «{product_name}»: {self.forecast:,.0f} руб.\n"
                msg += "🔹 Анализ корзины недоступен (режим «Цены»)."

            messagebox.showinfo("Результат", msg)
            self.chart_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Ошибка анализа", f"Не удалось проанализировать файл:\n{str(e)}")

    def get_frequent_pairs(self, basket):
        from itertools import combinations
        from collections import Counter

        pairs = []
        for order in basket.index:
            products = basket.columns[basket.loc[order] == 1]
            for pair in combinations(products, 2):
                pairs.append(tuple(sorted(pair)))

        pair_counts = Counter(pairs)
        return [(pair, count) for pair, count in pair_counts.most_common() if count >= 2]

    def show_charts(self):
        if self.df is None:
            return

        has_sales = 'Total' in self.df.columns
        has_prices = 'Price' in self.df.columns and 'Product' in self.df.columns

        fig, axes = plt.subplots(1, 2, figsize=(12, 5))

        if has_sales:
            self.df['Month'] = self.df['Date'].dt.to_period('M')
            monthly = self.df.groupby('Month')['Total'].sum()
            monthly.plot(kind='line', marker='o', ax=axes[0], color='purple')
            axes[0].set_title('Продажи по месяцам')
            axes[0].set_ylabel('Рубли')
            if hasattr(self, 'forecast'):
                axes[0].axhline(self.forecast, color='red', linestyle='--', label=f'Прогноз: {self.forecast:,.0f}')
                axes[0].legend()
        else:
            axes[0].text(0.5, 0.5, 'Нет данных\nо продажах', ha='center', va='center')
            axes[0].set_title('Продажи')

        if has_prices:
            top_product = self.df['Product'].iloc[0]
            price_data = self.df[self.df['Product'] == top_product].sort_values('Date')
            price_data.set_index('Date')['Price'].plot(kind='line', marker='o', ax=axes[1], color='green')
            axes[1].set_title(f'Цена: {top_product}')
            axes[1].set_ylabel('Рубли')
            axes[1].tick_params(axis='x', rotation=45)
        else:
            axes[1].text(0.5, 0.5, 'Нет данных\nо ценах', ha='center', va='center')
            axes[1].set_title('Цены')

        plt.tight_layout()
        plt.show()

    def save_report(self):

        date_style = NamedStyle(name='datetime', number_format='YYYY-MM-DD')

        output_path = os.path.splitext(self.filepath)[0] + "_sales_report.xlsx"

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_to_save = self.df.copy()

            if 'Date' in df_to_save.columns:
                df_to_save['Date'] = pd.to_datetime(df_to_save['Date']).dt.strftime('%Y-%m-%d')

            df_to_save.to_excel(writer, sheet_name='Данные', index=False)
            self._auto_adjust_columns(writer, 'Данные', df_to_save)

            mode_name = "продаж" if self.analysis_mode == 'sales' else "цены"
            forecast_df = pd.DataFrame({
                'Прогноз': [f"Прогноз {mode_name} на следующий месяц"],
                'Значение': [f"{self.forecast:,.0f} руб."]
            })
            forecast_df.to_excel(writer, sheet_name='Прогноз', index=False)
            self._auto_adjust_columns(writer, 'Прогноз', forecast_df)

            if self.basket_rules:
                basket_df = pd.DataFrame(self.basket_rules, columns=['Товары', 'Частота'])
                basket_df['Товары'] = basket_df['Товары'].apply(lambda x: ' + '.join(x))
                basket_df.to_excel(writer, sheet_name='Корзина', index=False)
                self._auto_adjust_columns(writer, 'Корзина', basket_df)

        messagebox.showinfo("Сохранено", f"Отчёт готов!\n{output_path}")

    def _auto_adjust_columns(self, writer, sheet_name, dataframe):
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(dataframe.columns, 1):
            max_length = max(
                len(str(col)),
                dataframe[col].astype(str).map(len).max() if not dataframe.empty else 0
            )
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[get_column_letter(idx)].width = adjusted_width

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesAnalyzerApp(root)
    root.mainloop()