# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from pathlib import Path
import numpy as np
import re
import calendar
from datetime import date, datetime
from openpyxl.styles import PatternFill, Font, Alignment

# ===============================
# НАСТРОЙКИ
# ===============================
APP_VERSION = "v2.9.0"
MAIN_SHEET_NAME = "Сводная мотивация"
OKLAD_SHEET_NAME = "Оклады"           # <-- подтвержденное имя листа
LOCAL_OKLADS_FILE = Path("оклады.xlsx")
USE_VLOOKUP_FALLBACK_DEFAULT = False  # можно переключить в UI

# ===============================
# УТИЛИТЫ
# ===============================
def normalize_fio(s: str):
    """Мягкая очистка ФИО: убираем лишние пробелы."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return s
    return " ".join(str(s).split()).strip()

def read_main_workbook(uploaded_file):
    """Открываем Excel без разрушения оформления (сохраняем формулы/стили)."""
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        if MAIN_SHEET_NAME not in wb.sheetnames:
            st.error("Нет листа «Сводная мотивация».")
            return None
        return wb
    except Exception as e:
        st.error(f"Ошибка загрузки Excel: {e}")
        return None

def ws_header_map(ws):
    """Карта заголовков первой строки → {имя: индекс_колонки}."""
    hmap = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v is not None and str(v).strip() != "":
            name = str(v).strip()
            if name not in hmap:
                hmap[name] = col
    return hmap

def get_last_row_by_fio(ws, fio_col_idx=2):
    """Последняя фактическая строка — по непустым ФИО в колонке B."""
    for r in range(ws.max_row, 1, -1):
        v = ws.cell(row=r, column=fio_col_idx).value
        if v not in (None, ""):
            return r
    return 1

def clean_fio(ws, fio_col_idx, last_row):
    """Нормализуем ФИО по фактическим строкам (оформление не трогаем)."""
    for r in range(2, last_row + 1):
        ws.cell(row=r, column=fio_col_idx).value = normalize_fio(
            ws.cell(row=r, column=fio_col_idx).value
        )

def delete_rows_partial(ws, fio_col_idx, patterns):
    """
    Удаляем строки по частичным совпадениям:
    строка удаляется, если ФИО содержит все части шаблона (без регистра).
    """
    filters = []
    for p in patterns:
        parts = [x.strip().lower() for x in p.split() if x.strip()]
        if parts:
            filters.append(parts)

    rows_to_delete = []
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=fio_col_idx).value
        if not val:
            continue
        fio_norm = normalize_fio(val).lower()
        for parts in filters:
            if all(part in fio_norm for part in parts):
                rows_to_delete.append(r)
                break

    for r in reversed(rows_to_delete):
        ws.delete_rows(r, 1)
    return len(rows_to_delete)

def set_zero(ws, col_idx_list, last_row):
    """Ставит 0 в указанных колонках по фактическим строкам."""
    changed = 0
    for col in col_idx_list:
        if col is None:
            continue
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=col).value = 0
            changed += 1
    return changed

def read_local_oklads():
    """Читаем локальный «оклады.xlsx» (первый лист файла)."""
    if not LOCAL_OKLADS_FILE.exists():
        st.error("❗ Файл «оклады.xlsx» не найден рядом с приложением.")
        return None
    try:
        # engine='openpyxl' — безопасно для локальной среды
        return pd.read_excel(LOCAL_OKLADS_FILE, engine="openpyxl")
    except Exception as e:
        st.error(f"Ошибка чтения «оклады.xlsx»: {e}")
        return None

def ensure_oklad_sheet(df_ok):
    """Упорядочиваем столбцы: A=ФИО, B=Сумма, остальное — справа."""
    if df_ok is None or df_ok.empty:
        return df_ok
    cols = df_ok.columns.tolist()
    fio = "ФИО" if "ФИО" in cols else cols[0]
    sm  = "Сумма" if "Сумма" in cols else (cols[1] if len(cols) > 1 else cols[0])
    return df_ok[[fio, sm] + [c for c in cols if c not in (fio, sm)]].copy()

def add_oklad_sheet(wb, df_ok):
    """Создаём/пересоздаём лист «Оклады» значениями из DataFrame."""
    if OKLAD_SHEET_NAME in wb.sheetnames:
        del wb[OKLAD_SHEET_NAME]
    ws = wb.create_sheet(OKLAD_SHEET_NAME)
    # Заголовки
    for c_idx, col in enumerate(df_ok.columns, 1):
        ws.cell(row=1, column=c_idx).value = col
    # Данные
    for r_idx, (_, row) in enumerate(df_ok.iterrows(), 2):
        for c_idx, col in enumerate(df_ok.columns, 1):
            ws.cell(row=r_idx, column=c_idx).value = row[col]

def remove_sheet_protection(ws):
    """Снимаем защиту листа, если стояла."""
    try:
        ws.protection.sheet = False
        ws.protection.enable = False
        ws.protection.password = None
    except Exception:
        pass

def finalize_sheets(wb, df_ok_sheet):
    """
    Удаляем все листы, кроме «Сводная мотивация», и добавляем лист «Оклады».
    Возвращаем (успех, [удалённые_листы], добавлен_оклады: bool).
    Если включена защита структуры книги — Excel не даст удалить/добавить листы.
    """
    deleted = []
    try:
        for name in list(wb.sheetnames):
            if name != MAIN_SHEET_NAME:
                deleted.append(name)
                del wb[name]
    except Exception as e:
        st.error("❗ Не удалось удалить лишние листы. Возможно, включена защита структуры книги.\n"
                 f"Техническая деталь: {e}")
        return False, deleted, False

    added_oklads = False
    try:
        if df_ok_sheet is not None and not df_ok_sheet.empty:
            add_oklad_sheet(wb, df_ok_sheet)
            added_oklads = True
        return True, deleted, added_oklads
    except Exception as e:
        st.error("❗ Не удалось добавить лист «Оклады». Возможно, структура книги под защитой.\n"
                 f"Техническая деталь: {e}")
        return False, deleted, added_oklads

def preview_df(ws, last_row):
    """
    Формируем DataFrame‑предпросмотр (ничего в книге не меняем).
    Уникализируем заголовки, чтобы Streamlit/pyarrow не падали из-за дублей.
    """
    rows = []
    for r in range(1, last_row + 1):
        rows.append([ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])
    if not rows:
        return pd.DataFrame()
    header = rows[0]
    data = rows[1:]

    cols = []
    seen = {}
    for h in header:
        nm = "Unnamed" if h in (None, "", "None") else str(h).strip()
        if nm in seen:
            seen[nm] += 1
            nm = f"{nm}_{seen[nm]}"
        else:
            seen[nm] = 1
        cols.append(nm)

    try:
        return pd.DataFrame(data, columns=cols)
    except Exception:
        return pd.DataFrame(data)

def build_formula_xlookup(i: int) -> str:
    """
    Английский синтаксис с запятыми, одинарные кавычки вокруг имени листа.
    Пустое значение по ошибке → "".
    """
    return (
        f"=IFERROR(XLOOKUP($B{i},'{OKLAD_SHEET_NAME}'!$A:$A,'{OKLAD_SHEET_NAME}'!$B:$B,\"\"),\"\")"
    )

def build_formula_vlookup(i: int) -> str:
    return (
        f"=IFERROR(VLOOKUP($B{i},'{OKLAD_SHEET_NAME}'!$A:$B,2,FALSE),\"\")"
    )

def append_suffix_to_filename(original_name: str, suffix: str) -> str:
    """
    Вставляет текст `suffix` перед расширением файла.
    Пример: 'март 2026.xlsx' + ' — processed' -> 'март 2026 — processed.xlsx'
    """
    if "." not in original_name:
        return original_name + suffix
    base, ext = original_name.rsplit(".", 1)
    return f"{base}{suffix}.{ext}"

def add_log_sheet(wb, changes_log, meta: dict):
    """Создаёт/пересоздаёт лист «Лог» с отчётом об обработке."""
    name = "Лог"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)

    # Заголовок
    ws["A1"].value = "Отчёт об обработке файла"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:C1")

    # Метаданные
    rows_meta = [
        ("Дата и время", meta.get("timestamp", "")),
        ("Версия приложения", meta.get("app_version", "")),
        ("Исходный файл", meta.get("filename", "")),
        ("Месяц/Год", f"{meta.get('month_name', '')} {meta.get('year', '')}"),
        ("Рабочие дни (Пн–Пт, минус праздники РФ)", meta.get("workdays", "")),
        ("Режим формул", "VLOOKUP" if meta.get("use_vlookup") else "XLOOKUP"),
    ]
    start_row = 3
    for i, (k, v) in enumerate(rows_meta, start=start_row):
        ws.cell(row=i, column=1).value = k
        ws.cell(row=i, column=2).value = v
        ws.cell(row=i, column=1).font = Font(bold=True)
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="left")
        ws.cell(row=i, column=2).alignment = Alignment(horizontal="left")

    # Список изменений
    body_row = start_row + len(rows_meta) + 1
    ws.cell(row=body_row, column=1).value = "Изменения"
    ws.cell(row=body_row, column=1).font = Font(bold=True)
    for j, line in enumerate(changes_log, start=body_row + 1):
        ws.cell(row=j, column=1).value = f"• {line}"

    # Немного ширины
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 22

# ===============================
# UI
# ===============================
st.set_page_config(page_title=f"Обработчик мотивации {APP_VERSION}", layout="wide")
st.title(f"📊 Обработчик мотивации — {APP_VERSION}")

with st.sidebar:
    st.header("⚙️ Настройки")
    default_bad = (
        "Хамидуллин Руслан\n"
        "Селюх Артем\n"
        "Орлов Дмитрий\n"
        "Чиканов Алексей\n"
        "Захаров Никита"
    )
    bad_text = st.text_area("ФИО для удаления (каждое с новой строки):", default_bad, height=150)
    remove_bad = st.checkbox("Удалять строки по списку ФИО", value=False)

    use_vlookup_ui = st.checkbox("VLOOKUP вместо XLOOKUP", value=USE_VLOOKUP_FALLBACK_DEFAULT)

uploaded_file = st.file_uploader("Загрузите файл Excel (.xlsx) с листом «Сводная мотивация»", type=["xlsx"])

# ======================================================================
# ОСНОВНАЯ ОБРАБОТКА
# ======================================================================
if uploaded_file:
    # --- Определяем месяц и год из названия файла (например: «март 2026») ---
    filename_lower = uploaded_file.name.lower()
    original_filename = uploaded_file.name  # важно не менять это имя

    months = {
        "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "май": 5, "мая": 5,
        "июн": 6, "июл": 7, "август": 8, "сентябр": 9, "октябр": 10,
        "ноябр": 11, "декабр": 12
    }
    month = None
    month_key_found = None
    for key, value in months.items():
        if key in filename_lower:
            month = value
            month_key_found = key
            break

    year_match = re.search(r"20\d{2}", filename_lower)
    year = int(year_match.group(0)) if year_match else date.today().year

    if not month:
        st.error("❗ Не удалось определить месяц из названия файла. Укажите месяц в имени файла (например, «март 2026»).")
        st.stop()

    # --- Считаем рабочие дни (Пн‑Пт, минус ключевые праздники РФ) ---
    holidays = []
    if month == 1:
        holidays += [date(year, 1, d) for d in range(1, 8)]  # 1–7 января
    if month == 2:
        holidays.append(date(year, 2, 23))
    if month == 3:
        holidays.append(date(year, 3, 8))
    if month == 5:
        holidays += [date(year, 5, 1), date(year, 5, 9)]
    if month == 6:
        holidays.append(date(year, 6, 12))
    if month == 11:
        holidays.append(date(year, 11, 4))

    _, days_in_month = calendar.monthrange(year, month)
    workdays = sum(
        1
        for d in range(1, days_in_month + 1)
        if (date(year, month, d).weekday() < 5) and (date(year, month, d) not in holidays)
    )

    # Для метаданных/журнала
    month_name_ru = {
        1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май",
        6: "июнь", 7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь",
        11: "ноябрь", 12: "декабрь"
    }.get(month, "")

    # --- Работа с книгой ---
    wb = read_main_workbook(uploaded_file)
    if wb is None:
        st.stop()

    ws = wb[MAIN_SHEET_NAME]
    remove_sheet_protection(ws)
    hmap = ws_header_map(ws)

    FIO_COL = 2                               # B
    AC_COL  = hmap.get("Заработная плата инженера", 29)  # по названию или AC (29)

    last_row = get_last_row_by_fio(ws, FIO_COL)

    st.subheader("Предпросмотр (до изменений)")
    st.dataframe(preview_df(ws, last_row), use_container_width=True)

    # Оклады (локальный файл)
    df_ok = read_local_oklads()
    if df_ok is None:
        st.stop()
    df_ok_sheet = ensure_oklad_sheet(df_ok)

    st.subheader("Лист «Оклады» (будет создан/обновлён)")
    st.dataframe(df_ok_sheet, use_container_width=True)

    # --- Кнопка сборки ---
    if st.button("Сформировать processed.xlsx"):
        changes_log = []
        meta = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "app_version": APP_VERSION,
            "filename": original_filename,
            "month_name": month_name_ru,
            "year": year,
            "workdays": workdays,
            "use_vlookup": use_vlookup_ui,
        }

        # 1) Чистка ФИО
        try:
            before_last_row = last_row
            clean_fio(ws, FIO_COL, last_row)
            processed_rows = max(last_row - 1, 0)
            changes_log.append(f"Нормализованы ФИО (обработано строк: {processed_rows}).")
        except Exception as e:
            st.error(f"Ошибка при нормализации ФИО: {e}")
            changes_log.append(f"❗ Ошибка при нормализации ФИО: {e}")

        # 2) Удаление строк по ФИО (опционально)
        try:
            if remove_bad:
                patterns = [x.strip() for x in bad_text.splitlines() if x.strip()]
                removed = delete_rows_partial(ws, FIO_COL, patterns)
                changes_log.append(f"Удалено строк по списку ФИО: {removed}.")
                last_row = get_last_row_by_fio(ws, FIO_COL)
            else:
                changes_log.append("Удаление строк по списку ФИО пропущено (выключено).")
        except Exception as e:
            st.error(f"Ошибка при удалении строк по ФИО: {e}")
            changes_log.append(f"❗ Ошибка при удалении строк по ФИО: {e}")

        # 3) Жёсткое обнуление T (20) и U (21)
        try:
            zeroed_cells = set_zero(ws, [20, 21], last_row)
            changes_log.append(f"Обнулены столбцы T и U (установлено нулей: {zeroed_cells}).")
        except Exception as e:
            st.error(f"Ошибка при обнулении T/U: {e}")
            changes_log.append(f"❗ Ошибка при обнулении T/U: {e}")

        # 4) Заполнение V (22) — рабочих дней месяца, если 0/пусто
        V_COL = 22
        try:
            days_filled = 0
            for r in range(2, last_row + 1):
                v = ws.cell(row=r, column=V_COL).value
                if v in (0, None, "", "0"):
                    ws.cell(row=r, column=V_COL).value = workdays
                    days_filled += 1
            changes_log.append(f"Подставлены рабочие дни ({workdays}) в {days_filled} строках.")
        except Exception as e:
            st.error(f"Ошибка при заполнении рабочих дней в V: {e}")
            changes_log.append(f"❗ Ошибка при заполнении рабочих дней в V: {e}")

        # 5) Очистка AC и вставка формул XLOOKUP/VLOOKUP
        try:
            if AC_COL is None:
                AC_COL = 29  # AC по позиции
            # Очистка
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=AC_COL).value = None

            # Запись формул
            use_vlookup = use_vlookup_ui
            formulas_written = 0
            for r in range(2, last_row + 1):
                formula = build_formula_vlookup(r) if use_vlookup else build_formula_xlookup(r)
                ws.cell(row=r, column=AC_COL).value = formula
                formulas_written += 1
            changes_log.append(
                f"Перезаписаны формулы в столбце AC ({'VLOOKUP' if use_vlookup else 'XLOOKUP'}), строк: {formulas_written}."
            )
        except Exception as e:
            st.error(f"Ошибка при вставке формул в AC: {e}")
            changes_log.append(f"❗ Ошибка при вставке формул в AC: {e}")

        # 6) Окраска столбца R (18) в жёлтый, если значение < 8.00
        try:
            R_COL = 18
            YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            painted = 0
            for r in range(2, last_row + 1):
                cell = ws.cell(row=r, column=R_COL)

                # Пробуем получить вычисленное значение формулы (кэш Excel)
                val = getattr(cell, "internal_value", None)
                # Если нет — используем обычное значение
                if val is None:
                    val = cell.value

                # Приведение к числу
                num = None
                try:
                    if isinstance(val, str):
                        num = float(val.replace(",", ".").strip())
                    else:
                        num = float(val)
                except Exception:
                    num = None

                if num is not None and num < 8:
                    cell.fill = YELLOW
                    painted += 1
            changes_log.append(f"Окрашено ячеек в столбце R (значение < 8.00): {painted}.")
        except Exception as e:
            st.error(f"Ошибка при окрашивании R: {e}")
            changes_log.append(f"❗ Ошибка при окрашивании R: {e}")

        # 7) Структура книги: оставляем «Сводная мотивация», добавляем «Оклады»
        ok_finalize = False
        deleted_sheets = []
        oklads_added = False
        try:
            ok_finalize, deleted_sheets, oklads_added = finalize_sheets(wb, df_ok_sheet)
            if ok_finalize:
                if deleted_sheets:
                    changes_log.append(f"Удалены листы: {', '.join(deleted_sheets)}.")
                else:
                    changes_log.append("Лишние листы для удаления не найдены.")
                changes_log.append("Добавлен лист «Оклады»." if oklads_added else "Лист «Оклады» не добавлен (нет данных).")
            else:
                changes_log.append("❗ Не удалось финализировать структуру книги (см. ошибки выше).")
        except Exception as e:
            st.error(f"Ошибка при финализации структуры книги: {e}")
            changes_log.append(f"❗ Ошибка при финализации структуры книги: {e}")

        # 8) Автофильтр на заголовки A1:AM1
        try:
            ws = wb[MAIN_SHEET_NAME]
            ws.auto_filter.ref = "A1:AM1"
            changes_log.append("Установлен автофильтр диапазона A1:AM1 на листе «Сводная мотивация».")
        except Exception as e:
            changes_log.append(f"Предупреждение: автофильтр не установлен (техническая деталь: {e}).")

        # 9) Лист «Лог»
        try:
            add_log_sheet(wb, changes_log, meta)
            changes_log.append("Создан лист «Лог» с отчётом об обработке.")
        except Exception as e:
            st.error(f"Ошибка при создании листа «Лог»: {e}")
            changes_log.append(f"❗ Ошибка при создании листа «Лог»: {e}")

        # 10) Сохранение
        try:
            out = BytesIO()
            wb.save(out)

            st.success("Готово! Рабочие дни заполнены, T/U обнулены, R<8 окрашен, формулы добавлены, фильтр включён, журнал сформирован.")
            # Имя для скачивания: оригинальное + ' — processed' перед расширением
            download_name = append_suffix_to_filename(original_filename, " — processed")
            st.download_button(
                label=f"Скачать {download_name}",
                data=out.getvalue(),
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ошибка сохранения файла: {e}")
            changes_log.append(f"❗ Ошибка сохранения файла: {e}")

        # 11) Вывод отчёта в UI
        st.subheader("Отчёт о внесённых изменениях")
        for line in changes_log:
            # Маркерная строка
            st.write("• " + line)