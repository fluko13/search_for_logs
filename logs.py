import re
from collections import defaultdict
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ---------- Логика анализа (исправленная) ----------
def analyze_serials_from_file(file_path: str) -> dict:
    """
    Возвращает словарь:
    ключ = серийный номер (строка)
    значение = (verdict, status_text)
        verdict: "OK" если был Done, иначе "NOT OK"
        status_text: подробный статус
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        text = f.read()
    if not text.strip():
        return {}

    # Строгий поиск: границы слова, ровно 7 цифр, подчёркивание, ровно 7 цифр
    pattern = re.compile(r'PLC# (\b\d{7}_\d{7}\b)')
    matches = list(pattern.finditer(text))
    if not matches:
        return {}

    stats = defaultdict(lambda: {'has_done': False, 'had_failure': False})

    for i, match in enumerate(matches):
        serial = match.group(1)
        start_pos = match.end()
        end_pos = matches[i+1].start() if i+1 < len(matches) else len(text)
        chunk = text[start_pos:end_pos]

        if re.search(r'\bDone\b', chunk):
            stats[serial]['has_done'] = True
        else:
            stats[serial]['had_failure'] = True

    result = {}
    for serial, data in stats.items():
        # Определяем подробный статус
        if not data['has_done']:
            detail_status = "провалил тестирование"
            verdict = "NOT OK"
        elif data['had_failure']:
            detail_status = "тест пройден не с первого раза"
            verdict = "OK"
        else:
            detail_status = "тест пройден"
            verdict = "OK"
        result[serial] = (verdict, detail_status)

    return result

def save_to_csv(data: dict, output_file: str):
    """Сохраняет три столбца: Серийный номер, Вердикт, Результат"""
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(['Серийный номер', 'Вердикт', 'Результат тестирования'])
        for serial, (verdict, status) in data.items():
            writer.writerow([serial, verdict, status])

# ---------- GUI ----------
class SerialAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор серийных номеров")
        self.root.geometry("850x650")
        self.root.resizable(True, True)

        self.data = {}          # {serial: (verdict, detail_status)}
        self.current_file = None

        # Верхняя панель
        top_frame = tk.Frame(root)
        top_frame.pack(pady=10)

        btn_open = tk.Button(top_frame, text="Выбрать файл с логами", command=self.load_file, width=25)
        btn_open.pack(side=tk.LEFT, padx=5)

        btn_save = tk.Button(top_frame, text="Сохранить в CSV", command=self.save_csv, width=20)
        btn_save.pack(side=tk.LEFT, padx=5)

        # Область вывода результатов (текстовое поле)
        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=20)
        self.text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Нижняя панель поиска
        search_frame = tk.Frame(root)
        search_frame.pack(pady=10, fill=tk.X, padx=10)

        lbl_search = tk.Label(search_frame, text="Серийный номер:")
        lbl_search.pack(side=tk.LEFT)

        self.entry_search = tk.Entry(search_frame, width=30)
        self.entry_search.pack(side=tk.LEFT, padx=5)

        btn_search = tk.Button(search_frame, text="Найти", command=self.search_serial)
        btn_search.pack(side=tk.LEFT, padx=5)

        self.search_result = tk.Label(search_frame, text="", fg="blue")
        self.search_result.pack(side=tk.LEFT, padx=10)

        # Статусная строка
        self.status_var = tk.StringVar()
        self.status_var.set("Готов. Выберите файл с логами.")
        status_bar = tk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите текстовый файл с результатами тестирования",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if not file_path:
            return
        try:
            self.data = analyze_serials_from_file(file_path)
            if not self.data:
                messagebox.showwarning("Пустой результат", "В файле не найдено серийных номеров формата PLC# 1234567_8901234.")
                self.text_area.delete(1.0, tk.END)
                self.status_var.set("Файл не содержит корректных записей.")
                return
            self.current_file = file_path
            self.display_results()
            self.status_var.set(f"Загружен: {os.path.basename(file_path)} | Найдено серийников: {len(self.data)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать файл:\n{str(e)}")
            self.status_var.set("Ошибка загрузки файла")

    def display_results(self):
        """Выводит таблицу: Номер | Вердикт | Результат"""
        self.text_area.delete(1.0, tk.END)
        if not self.data:
            self.text_area.insert(tk.END, "Нет данных для отображения.")
            return
        # Заголовки
        self.text_area.insert(tk.END, f"{'Серийный номер':<20} {'Вердикт':<10} {'Результат тестирования':<35}\n")
        self.text_area.insert(tk.END, "-" * 65 + "\n")
        for serial, (verdict, status) in self.data.items():
            self.text_area.insert(tk.END, f"{serial:<20} {verdict:<10} {status:<35}\n")

    def save_csv(self):
        if not self.data:
            messagebox.showwarning("Нет данных", "Сначала загрузите файл с логами.")
            return
        script_dir = os.path.dirname(os.path.abspath(__file__))
        default_name = "results.csv"
        default_path = os.path.join(script_dir, default_name)
        file_path = filedialog.asksaveasfilename(
            title="Сохранить как CSV",
            defaultextension=".csv",
            initialfile=default_name,
            initialdir=script_dir,
            filetypes=[("CSV файлы", "*.csv")]
        )
        if not file_path:
            return
        try:
            save_to_csv(self.data, file_path)
            messagebox.showinfo("Успех", f"Файл сохранён:\n{file_path}")
            self.status_var.set(f"Сохранён CSV: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def search_serial(self):
        serial = self.entry_search.get().strip()
        if not serial:
            self.search_result.config(text="Введите серийный номер", fg="red")
            return
        if not self.data:
            self.search_result.config(text="Нет загруженных данных", fg="red")
            return
        if serial in self.data:
            verdict, status = self.data[serial]
            result_text = f"Вердикт: {verdict}, Результат: {status}"
            self.search_result.config(text=result_text, fg="green")
        else:
            self.search_result.config(text="Не найден", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialAnalyzerApp(root)
    root.mainloop()