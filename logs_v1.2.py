import re
from collections import defaultdict
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import Workbook, load_workbook

# ---------- Логика анализа ----------
def analyze_serials_from_file(file_path: str) -> dict:
    """
    Возвращает словарь {серийный номер: (verdict, status_text)}.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        text = f.read()
    if not text.strip():
        return {}

    # Находим все позиции ключевого слова PCL_FW#
    keyword = "PLC_FW#"
    positions = []
    start = 0
    while True:
        pos = text.find(keyword, start)
        if pos == -1:
            break
        positions.append(pos)
        start = pos + len(keyword)

    if not positions:
        return {}

    stats = defaultdict(lambda: {'has_done': False, 'had_failure': False})
    serial_pattern = re.compile(r'^\s*(\b\d{7}_\d{7}\b)', re.MULTILINE)

    for i, pos in enumerate(positions):
        fragment_start = pos + len(keyword)
        fragment_end = positions[i+1] if i+1 < len(positions) else len(text)
        fragment = text[fragment_start:fragment_end]

        match = serial_pattern.search(fragment)
        if not match:
            continue
        serial = match.group(1)

        if re.search(r'\bDone\b', fragment):
            stats[serial]['has_done'] = True
        else:
            stats[serial]['had_failure'] = True

    result = {}
    for serial, data in stats.items():
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


# ---------- Определение имени листа ----------
def get_sheet_name(serial: str) -> str:
    prefix = serial.split('_')[0]
    if len(prefix) != 7 or not prefix.isdigit():
        return "Ошибка"
    if not prefix.startswith('075'):
        return "Ошибка"
    base = prefix[3:6]
    flag = prefix[6]
    if flag == '0':
        return f"БТП75-{base}"
    elif flag == '1':
        return f"БТП75-{base}А"
    else:
        return "Ошибка"


# ---------- Сохранение в Excel (добавление + обновление статуса) ----------
def save_to_excel_append(data: dict, output_file: str):
    if not data:
        return

    if os.path.exists(output_file):
        wb = load_workbook(output_file)
    else:
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        wb.create_sheet("Все данные")
        wb["Все данные"].append(["Серийный номер", "Вердикт", "Результат тестирования"])

    # Убедимся, что лист "Все данные" существует
    if "Все данные" not in wb.sheetnames:
        ws_all = wb.create_sheet("Все данные")
        ws_all.append(["Серийный номер", "Вердикт", "Результат тестирования"])
    else:
        ws_all = wb["Все данные"]

    # Соберём все существующие номера из листа "Все данные" для быстрого поиска
    existing_serials = {}
    for row in ws_all.iter_rows(min_row=2, values_only=True):
        if row[0]:
            existing_serials[row[0]] = row[1]  # serial -> verdict

    # Обрабатываем каждый серийный номер из нового лога
    for serial, (verdict, status) in data.items():
        sheet_name = get_sheet_name(serial)
        # Проверяем, существует ли номер в файле
        if serial in existing_serials:
            # Обновляем статус, если было NOT OK, а стало OK
            old_verdict = existing_serials[serial]
            if old_verdict == "NOT OK" and verdict == "OK":
                # Найти строку в "Все данные" и обновить
                for row_idx, row in enumerate(ws_all.iter_rows(min_row=2, values_only=False), start=2):
                    if row[0].value == serial:
                        ws_all.cell(row=row_idx, column=2, value=verdict)
                        ws_all.cell(row=row_idx, column=3, value=status)
                        break
                # Обновить в специализированном листе
                if sheet_name in wb.sheetnames:
                    ws_spec = wb[sheet_name]
                    for r_idx, r in enumerate(ws_spec.iter_rows(min_row=2, values_only=False), start=2):
                        if r[0].value == serial:
                            ws_spec.cell(row=r_idx, column=2, value=verdict)
                            ws_spec.cell(row=r_idx, column=3, value=status)
                            break
                # Обновляем existing_serials для возможных следующих итераций (хотя в одной сессии дублей не будет)
                existing_serials[serial] = verdict
        else:
            # Новый номер – добавляем в "Все данные"
            ws_all.append([serial, verdict, status])
            existing_serials[serial] = verdict
            # Добавляем в специализированный лист
            if sheet_name not in wb.sheetnames:
                ws_spec = wb.create_sheet(sheet_name)
                ws_spec.append(["Серийный номер", "Вердикт", "Результат тестирования"])
            else:
                ws_spec = wb[sheet_name]
            ws_spec.append([serial, verdict, status])

    # Автоширина столбцов
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for col in ['A', 'B', 'C']:
            max_len = 0
            for cell in ws[col]:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col].width = max_len + 2

    wb.save(output_file)


# ---------- GUI ----------
class SerialAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор серийных номеров")
        self.root.geometry("850x650")
        self.root.resizable(True, True)

        self.data = {}
        self.current_file = None

        top_frame = tk.Frame(root)
        top_frame.pack(pady=10)

        btn_open = tk.Button(top_frame, text="Выбрать файл с логами", command=self.load_file, width=25)
        btn_open.pack(side=tk.LEFT, padx=5)

        btn_save = tk.Button(top_frame, text="Сохранить в Excel", command=self.save_excel, width=20)
        btn_save.pack(side=tk.LEFT, padx=5)

        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=20)
        self.text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

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
                messagebox.showwarning("Пустой результат", "В файле не найдено серийных номеров формата PCL_FW# 1234567_8901234.")
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
        self.text_area.delete(1.0, tk.END)
        if not self.data:
            self.text_area.insert(tk.END, "Нет данных для отображения.")
            return
        self.text_area.insert(tk.END, f"{'Серийный номер':<20} {'Вердикт':<10} {'Результат тестирования':<35}\n")
        self.text_area.insert(tk.END, "-" * 65 + "\n")
        for serial, (verdict, status) in self.data.items():
            self.text_area.insert(tk.END, f"{serial:<20} {verdict:<10} {status:<35}\n")

    def save_excel(self):
        if not self.data:
            messagebox.showwarning("Нет данных", "Сначала загрузите файл с логами.")
            return
        script_dir = os.path.dirname(os.path.abspath(__file__))
        default_name = "results.xlsx"
        default_path = os.path.join(script_dir, default_name)
        file_path = filedialog.asksaveasfilename(
            title="Сохранить как Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            initialdir=script_dir,
            filetypes=[("Excel файлы", "*.xlsx")]
        )
        if not file_path:
            return
        try:
            save_to_excel_append(self.data, file_path)
            messagebox.showinfo("Успех", f"Файл сохранён:\n{file_path}")
            self.status_var.set(f"Сохранён Excel: {os.path.basename(file_path)}")
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
            self.search_result.config(text="Не найден в текущем логе", fg="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = SerialAnalyzerApp(root)
    root.mainloop()