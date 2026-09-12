#!/usr/bin/env python3
"""
Script: script_with_save_prompt.py
Same as script.py but before opening the Save-As dialog shows an informational message:
"Теперь выбери куда мне сохранить результат"
"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime

# Настройки
ALLOWED_AMOUNTS = {400, 400000, 1300000, 2650000, 3550000}
MIN_CONTRACT_YEAR = 2022  # год контракта не может быть ниже этого

# ------------------------ Утилиты ------------------------

def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9а-яё\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def name_matches(target_text, name):
    # Проверяем, что все слова из name присутствуют в target_text (приближённое совпадение)
    n = normalize_text(name)
    t = normalize_text(target_text)
    if not n:
        return False
    words = [w for w in n.split() if w]
    return all(w in t for w in words)

def try_parse_date(val):
    """
    Надёжно парсим дату, избегая лишних предупреждений pandas.
    Возвращаем pd.Timestamp или None.
    """
    if pd.isna(val) or str(val).strip() == '':
        return None

    s = str(val).strip()

    # 1) Попробовать популярные явные форматы (ISO и dd.mm.yyyy варианты)
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%Y-%m-%d',
        '%d.%m.%Y %H:%M:%S',
        '%d.%m.%Y %H:%M',
        '%d.%m.%Y',
        '%d/%m/%Y',
        '%d/%m/%Y %H:%M:%S',
        '%Y.%m.%d',
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return pd.Timestamp(dt)
        except Exception:
            pass

    # 2) Универсальный парсер (dayfirst=False чтобы не получать предупреждение для ISO)
    try:
        dt = pd.to_datetime(s, dayfirst=False, errors='coerce')
        if not pd.isna(dt):
            yr = getattr(dt, "year", None)
            if yr and 1900 <= yr <= 2100:
                return dt
    except Exception:
        pass

    # 3) Регекс для dd.mm.yyyy и подобного, затем парсинг с dayfirst=True (как запасной вариант)
    m2 = re.search(r'(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})', s)
    if m2:
        try:
            dt = pd.to_datetime(m2.group(0), dayfirst=True, errors='coerce')
            if not pd.isna(dt):
                return dt
        except:
            pass

    # 4) Также можно найти год явно в тексте (например "2024"), тогда вернуть 1 января этого года
    m_year = re.search(r'\b(20\d{2})\b', s)
    if m_year:
        try:
            y = int(m_year.group(1))
            if 1900 <= y <= 2100:
                return pd.Timestamp(year=y, month=1, day=1)
        except:
            pass

    return None

def parse_int_like(val):
    if pd.isna(val) or str(val).strip() == '':
        return None
    try:
        if isinstance(val, (int, float)) and not pd.isna(val):
            return int(val)
    except:
        pass
    s = str(val)
    s_digits = re.sub(r'[^0-9]', '', s)
    if not s_digits:
        return None
    try:
        return int(s_digits)
    except:
        return None

# ------------------------ Поиск даты/суммы ------------------------

def find_date_candidates_in_row_by_columns(row, df_columns, predicate=lambda dt: True):
    dates = []
    for col in df_columns:
        try:
            val = row.get(col)
        except Exception:
            val = None
        dt = try_parse_date(val)
        if dt is not None and predicate(dt):
            dates.append(dt)
    return dates

def find_date_in_row(row, df_columns):
    def is_contract_year(dt):
        return dt.year >= MIN_CONTRACT_YEAR

    # 1) колонки с заголовками, связанные с контрактом
    header_cols = []
    for col in df_columns:
        if col is None:
            continue
        cl = str(col).lower()
        if 'контрак' in cl or 'заключ' in cl or ('дата' in cl and ('контрак' in cl or 'заключ' in cl)):
            header_cols.append(col)
    if header_cols:
        dates = find_date_candidates_in_row_by_columns(row, header_cols, predicate=is_contract_year)
        if dates:
            return min(dates)

    # 2) по всем ячейкам строки
    all_dates = []
    for col in df_columns:
        try:
            val = row.get(col)
        except Exception:
            val = None
        dt = try_parse_date(val)
        if dt is not None and is_contract_year(dt):
            all_dates.append(dt)
    if all_dates:
        return min(all_dates)

    # 3) запасные позиции I(8), H(7), L(11)
    fallback_indices = [8, 7, 11]
    cols_list = list(df_columns)
    for idx in fallback_indices:
        if 0 <= idx < len(cols_list):
            col = cols_list[idx]
            try:
                val = row.get(col)
            except Exception:
                val = None
            dt = try_parse_date(val)
            if dt is not None and is_contract_year(dt):
                return dt

    return None

def find_amount_in_row(row, df_columns):
    # сначала поиск по явным заголовкам (размер/сумма/выплат)
    for col in df_columns:
        if col is None:
            continue
        label = str(col).lower()
        if any(k in label for k in ['размер', 'выплат', 'сумма', 'объем']):
            val = row.get(col)
            iv = parse_int_like(val)
            if iv is not None and iv in ALLOWED_AMOUNTS:
                return iv
    # иначе по всем ячейкам
    for col in df_columns:
        val = row.get(col)
        iv = parse_int_like(val)
        if iv is not None and iv in ALLOWED_AMOUNTS:
            return iv
    return None

# ------------------------ Основная обработка ------------------------

def process_source_file(src_path, out_path=None, parent_root=None):
    src_df = pd.read_excel(src_path, dtype=str)
    cols = list(src_df.columns)

    if len(cols) >= 1:
        fio_col = cols[0]
    else:
        raise RuntimeError("В исходном файле нет колонок.")

    if len(cols) >= 5:
        files_col = cols[4]
    else:
        files_col = None
        for c in cols:
            sample = src_df[c].astype(str).fillna('').head(30).to_list()
            if any('\\' in str(x) or '/' in str(x) for x in sample):
                files_col = c
                break
        if files_col is None:
            raise RuntimeError("Не найден столбец с путями к файлам (E). Убедитесь, что в исходном файле есть колонка с путями.")

    result_df = src_df.copy()

    # 4-й столбец -> "Год контракта"
    target_year_col = "Год контракта"
    if len(result_df.columns) >= 4:
        cols_list = list(result_df.columns)
        cols_list[3] = target_year_col
        result_df.columns = cols_list
    else:
        while len(result_df.columns) < 4:
            result_df[f'_pad_{len(result_df.columns)+1}'] = ''
        cols_list = list(result_df.columns)
        cols_list[3] = target_year_col
        result_df.columns = cols_list

    # 6-й столбец -> "Сумма выплаты"
    amount_col_name = "Сумма выплаты"
    if len(result_df.columns) >= 6:
        cols_list = list(result_df.columns)
        cols_list[5] = amount_col_name
        result_df.columns = cols_list
    else:
        while len(result_df.columns) < 6:
            result_df[f'_pad_{len(result_df.columns)+1}'] = ''
        cols_list = list(result_df.columns)
        cols_list[5] = amount_col_name
        result_df.columns = cols_list

    total = len(src_df)

    # --- окно прогресса ---
    stop_flag = {'stop': False}
    if parent_root is None:
        root = tk.Tk()
        root.withdraw()
    else:
        root = parent_root

    progress_win = tk.Toplevel(root)
    progress_win.title("Прогресс обработки")
    progress_win.resizable(False, False)
    tk.Label(progress_win, text="Обработка записей:").grid(row=0, column=0, columnspan=2, padx=10, pady=(10,0))

    progress_var = tk.IntVar(value=0)
    pb = ttk.Progressbar(progress_win, orient="horizontal", length=420, mode="determinate", maximum=total, variable=progress_var)
    pb.grid(row=1, column=0, columnspan=2, padx=10, pady=(5,0))

    lbl_status = tk.Label(progress_win, text=f"Обработано 0 из {total}")
    lbl_status.grid(row=2, column=0, columnspan=2, padx=10, pady=(5,0))

    def on_cancel():
        if messagebox.askyesno("Отмена", "Прервать обработку? Частичный результат будет сохранён с суффиксом _partial."):
            stop_flag['stop'] = True
            btn_cancel.config(state="disabled")

    btn_cancel = tk.Button(progress_win, text="Отменить", command=on_cancel, width=12)
    btn_cancel.grid(row=3, column=0, padx=10, pady=10, sticky="e")

    btn_hide = tk.Button(progress_win, text="Скрыть", command=progress_win.withdraw, width=12)
    btn_hide.grid(row=3, column=1, padx=10, pady=10, sticky="w")

    try:
        progress_win.update_idletasks()
        if parent_root:
            progress_win.geometry("+%d+%d" % (parent_root.winfo_rootx() + 50, parent_root.winfo_rooty() + 50))
    except:
        pass

    cancelled = False
    processed = 0

    # --- основной цикл ---
    for idx, row in src_df.iterrows():
        if stop_flag['stop']:
            cancelled = True
            break

        fio = row.get(fio_col, '')
        files_cell = row.get(files_col, '')
        if pd.isna(files_cell) or str(files_cell).strip() == '':
            result_df.at[idx, target_year_col] = ""
            result_df.at[idx, amount_col_name] = "пусто"
        else:
            file_paths = [p.strip() for p in str(files_cell).split(';') if p.strip()]
            found_date = None
            found_amount = None
            for fp in file_paths:
                fp_clean = fp.strip().strip('"').strip("'")
                if not os.path.isabs(fp_clean):
                    base_dir = os.path.dirname(src_path)
                    candidate = os.path.join(base_dir, fp_clean)
                    if os.path.exists(candidate):
                        fp_clean = candidate
                if not os.path.exists(fp_clean):
                    continue
                try:
                    sheets = pd.read_excel(fp_clean, sheet_name=None, dtype=str)
                except Exception:
                    try:
                        sheets = {'sheet1': pd.read_csv(fp_clean, dtype=str)}
                    except:
                        sheets = {}
                for sheet_name, df in sheets.items():
                    if df is None or df.shape[0] == 0:
                        continue
                    df_cols = list(df.columns)
                    combined = df.fillna('').astype(str).agg(' '.join, axis=1)
                    matched = []
                    for ridx, txt in combined.items():
                        if name_matches(txt, fio):
                            matched.append(ridx)
                    if not matched:
                        continue
                    for ridx in matched:
                        row_target = df.loc[ridx]
                        if found_date is None:
                            dt = find_date_in_row(row_target, df_cols)
                            if dt is not None:
                                found_date = dt
                        if found_amount is None:
                            amt = find_amount_in_row(row_target, df_cols)
                            if amt is not None:
                                found_amount = amt
                        if found_date is not None and found_amount is not None:
                            break
                    if found_date is not None and found_amount is not None:
                        break
                if found_date is not None and found_amount is not None:
                    break

            if found_date is not None:
                result_df.at[idx, target_year_col] = str(found_date.year)
            else:
                result_df.at[idx, target_year_col] = ""

            if found_amount is not None:
                result_df.at[idx, amount_col_name] = int(found_amount)
            else:
                result_df.at[idx, amount_col_name] = "пусто"

        processed += 1
        progress_var.set(processed)
        lbl_status.config(text=f"Обработано {processed} из {total}")
        try:
            progress_win.update_idletasks()
            progress_win.update()
        except:
            pass

    try:
        progress_win.destroy()
    except:
        pass

    # Подготовка имени выходного файла
    base = os.path.splitext(os.path.basename(src_path))[0]
    if out_path:
        out_file = out_path
        root_name, ext = os.path.splitext(out_file)
        if ext.lower() not in ['.xlsx', '.xlsm', '.xls']:
            out_file = out_file + '.xlsx'
        if os.path.exists(out_file):
            if not messagebox.askyesno("Подтвердите перезапись", f"Файл {os.path.basename(out_file)} уже существует. Перезаписать?"):
                return None
    else:
        if cancelled:
            default_name = f"результат_{base}_partial.xlsx"
        else:
            default_name = f"результат_{base}.xlsx"
        out_file = os.path.join(os.path.dirname(src_path), default_name)
        i = 1
        while os.path.exists(out_file):
            if cancelled:
                out_file = os.path.join(os.path.dirname(src_path), f"результат_{base}_partial_{i}.xlsx")
            else:
                out_file = os.path.join(os.path.dirname(src_path), f"результат_{base}_{i}.xlsx")
            i += 1

    result_df.to_excel(out_file, index=False)
    if cancelled:
        messagebox.showinfo("Обработка прервана", f"Обработка прервана пользователем. Частичный результат сохранён:\n{out_file}")
    return out_file

# ------------------------ GUI: выбор и запуск ------------------------

def ask_save_path(default_base):
    # Показываем информационное окно перед выбором места сохранения
    try:
        messagebox.showinfo("Сохранение", "Теперь выбери куда мне сохранить результат")
    except Exception:
        pass

    opts = {
        'defaultextension': '.xlsx',
        'filetypes': [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        'initialfile': default_base,
        'title': 'Сохранить результат как'
    }
    path = filedialog.asksaveasfilename(**opts)
    if not path:
        return None
    return path

def main():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Начнем", "Здравствуй хозяин! Позволь мне за тебя найти информацию о том, когда твои бойцы заключили контракты и сколько им заплатили. Будь добр, покажи, в каком файле записаны твои бойцы.")
    global src_path
    src_path = filedialog.askopenfilename(title="Выберите исходный файл", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    if not src_path:
        print("Файл не выбран. Выход.")
        return

    base = os.path.splitext(os.path.basename(src_path))[0]
    default_name = f"результат_{base}.xlsx"

    save_path = ask_save_path(default_name)
    final_out = None
    while True:
        out = process_source_file(src_path, out_path=save_path, parent_root=root)
        if out is None:
            retry = messagebox.askyesno("Сохранение", "Вы отменили перезапись или сохранение. Хотите выбрать другой путь для сохранения?")
            if retry:
                save_path = ask_save_path(default_name)
                if save_path is None:
                    save_path = None
                    continue
                else:
                    continue
            else:
                # Сохранить по умолчанию в папке исходного файла
                save_path = None
                continue
        else:
            final_out = out
            break

    if final_out:
        messagebox.showinfo("Готово", f"Готово. Я постарался найти для тебя всю нужную информацию и сохранить там, где ты меня попросил. Если я хорошо поработал, можешь положить конфетку на системник")
    else:
        messagebox.showinfo("Отмена", "Результат не был сохранён.")

if __name__ == "__main__":
    main()
