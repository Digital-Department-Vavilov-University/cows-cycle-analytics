#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from datetime import date


# Срок стельности в днях
GESTATION_DAYS = 285

# Столбцы с датами
DATE_COLUMNS = {'Дата отела', 'Дата посл. осеменения', 'Дата планового отела'}

# Столбцы для шаблона
TEMPLATE_COLUMNS = [
    'Инвентарный номер',
    'Дата рождения',
    'Возраст, полных лет',
    'Номер текущей лактации',
    'Дата отела',
    'Дата посл. осеменения',
    'Дата планового отела',
]

MONTH_NAMES_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

#  Главное окно
root = tk.Tk()
root.title("Лакталькулятор")

screen_width  = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width  = int(screen_width  * 0.92)
window_height = int(screen_height * 0.92)
offset_x      = int((screen_width  - window_width)  / 2)
offset_y      = int((screen_height - window_height) / 2)
root.geometry(f"{window_width}x{window_height}+{offset_x}+{offset_y}")

# Главный блок
main = ttk.Frame(root)
main.pack(expand=True, fill='both', padx=5, pady=5)

# Панель кнопок
frame_controls = ttk.LabelFrame(main, text="Панель управления", padding=(5, 5))
frame_controls.grid(row=0, column=0, sticky="ew", pady=(0, 4))

# Таблица со всеми коровами
frame_table = ttk.LabelFrame(main, text="Данные о животных", padding=(5, 5))
frame_table.grid(row=1, column=0, sticky="nsew", pady=(0, 4))

# Календарь
frame_detail = ttk.LabelFrame(main, text="Детали по месяцу", padding=(5, 5))
frame_detail.grid(row=2, column=0, sticky="nsew")

# Фикс габариты
main.grid_rowconfigure(1, weight=3)
main.grid_rowconfigure(2, weight=2)
main.grid_columnconfigure(0, weight=1)

# Гибкие габариты
frame_detail.grid_columnconfigure(0, weight=0)
frame_detail.grid_columnconfigure(1, weight=1)
frame_detail.grid_rowconfigure(0, weight=1)

# Датафреймы
original_df  = None   # исходные данные после загрузки (не меняются)
displayed_df = None   # то, что сейчас видно в таблице (с фильтром и сортировкой)
detail_df    = None   # коровы выбранного месяца в нижней панели

# Виджеты таблиц
tree_widget     = None
detail_cow_tree = None

# Сортировка главной таблицы
current_sort_column = None
current_sort_order  = None

# Сортировка таблицы коров по месяцу
detail_sort_column = None
detail_sort_order  = None

# Фильтр скрытия неосем
hide_uninseminated = False

# Календарь_buttons
month_buttons = {} # кнопки месяцев
current_month_key = None # выбранный месяц
cal_year = date.today().year # год, отображаемый в календаре

# Уведомления 5s кд
def _notify(msg, color='green', ms=5000):
    error_label.config(text=msg, foreground=color, font=('Arial', 9, 'bold'))
    root.after(ms, lambda: error_label.config(text="", font=('Arial', 9)))


# Сортировка по dd.mm.yyyy
def _sort_series(df, col_name, ascending):
    if col_name in DATE_COLUMNS:
        key = pd.to_datetime(df[col_name], dayfirst=True, errors='coerce')
        na_mask   = key.isna()
        valid     = df[~na_mask]
        invalid   = df[na_mask]
        key_valid = key[~na_mask]
        sorted_valid = valid.loc[key_valid.sort_values(ascending=ascending).index]
        return pd.concat([sorted_valid, invalid], ignore_index=True)
    return df.sort_values(col_name, ascending=ascending,
                          na_position='last').reset_index(drop=True)

# Шаблон excel & csv
def download_template():
    fmt = _ask_template_format()
    if not fmt:
        return

    if fmt == 'xlsx':
        fp = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Шаблон лактации коров",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not fp:
            return
        try:
            df_tpl = pd.DataFrame(columns=TEMPLATE_COLUMNS)
            with pd.ExcelWriter(fp, engine='openpyxl') as writer:
                df_tpl.to_excel(writer, index=False, sheet_name='Лактация')
                ws = writer.sheets['Лактация']
                for i, col in enumerate(TEMPLATE_COLUMNS, 1):
                    ws.column_dimensions[
                        ws.cell(row=1, column=i).column_letter
                    ].width = max(len(col) + 4, 18)
            _notify(f'Шаблон Excel сохранён: {fp}', 'green')
        except Exception as e:
            _notify(f'Ошибка сохранения шаблона: {e}', 'red')
    else:
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile="Шаблон лактации коров",
            filetypes=[("CSV", "*.csv")]
        )
        if not fp:
            return
        try:
            pd.DataFrame(columns=TEMPLATE_COLUMNS).to_csv(
                fp, index=False, sep=';', encoding='utf-8-sig')
            _notify(f'Шаблон CSV сохранён: {fp}', 'green')
        except Exception as e:
            _notify(f'Ошибка сохранения шаблона: {e}', 'red')


def _ask_template_format():
    choice = tk.StringVar(value='')
    dlg = tk.Toplevel(root)
    dlg.title("Формат шаблона")
    dlg.resizable(False, False)
    dlg.grab_set()

    dlg.update_idletasks()
    rx = root.winfo_x() + root.winfo_width()  // 2 - 210
    ry = root.winfo_y() + root.winfo_height() // 2 - 80
    dlg.geometry(f"420x160+{rx}+{ry}")

    ttk.Label(dlg, text="Выберите формат шаблона:",
              font=('Arial', 11)).pack(pady=(18, 8))

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack()

    def pick(fmt):
        choice.set(fmt)
        dlg.destroy()

    ttk.Button(btn_frame, text="📊 Excel (.xlsx)", width=16,
               command=lambda: pick('xlsx')).grid(row=0, column=0, padx=10, pady=4)
    ttk.Button(btn_frame, text="📄 CSV (.csv)", width=16,
               command=lambda: pick('csv')).grid(row=0, column=1, padx=10, pady=4)
    ttk.Button(dlg, text="Отмена", command=dlg.destroy).pack(pady=(6, 0))

    root.wait_window(dlg)
    return choice.get() or None

# Расчёт даты планового отёла
def calculate_periods(df):
    PLANNED_COL = 'Дата планового отела'
    INS_COL     = 'Дата посл. осеменения'

    drop_cols = ['Статус', 'Срок стельности (дней)', 'Сервис-период (дней)']
    df.drop(columns=[c for c in drop_cols if c in df.columns], inplace=True)

    if PLANNED_COL not in df.columns:
        df[PLANNED_COL] = ''

    planned = pd.to_datetime(df[PLANNED_COL], dayfirst=True, errors='coerce')
    ins     = pd.to_datetime(df[INS_COL],     dayfirst=True, errors='coerce')

    mask_fill = planned.isna() & ins.notna()
    planned.loc[mask_fill] = ins[mask_fill] + pd.Timedelta(days=GESTATION_DAYS)
    df[PLANNED_COL] = planned.dt.strftime('%d.%m.%Y').fillna('')

    df[INS_COL] = df[INS_COL].replace('', pd.NA).fillna('Не осеменялась')
    return df

# Импорт
def load_file():
    global original_df, displayed_df, tree_widget, hide_uninseminated
    fp = filedialog.askopenfilename(title="Выберите CSV", filetypes=[("CSV", "*.csv")])
    if not fp:
        return
    try:
        df = pd.read_csv(fp, sep=";", encoding="utf-8", header=0)
        if 'Дата посл. осеменения' not in df.columns:
            _notify('Ошибка: нет столбца «Дата посл. осеменения»', 'red')
            return
        calculate_periods(df)
        original_df        = df.copy()
        displayed_df       = original_df.copy()
        hide_uninseminated = False
        toggle_btn.config(text="🙈 Скрыть неосеменённых")
        if tree_widget:
            tree_widget.destroy()
        create_table(displayed_df)
        build_month_calendar()
        clear_cow_detail()
        _update_stats_label()
        error_label.config(text="")
    except Exception as e:
        _notify(f'Ошибка импорта: {e}', 'red')

# Экспорт 
def _main_export_filename():
    # Имя файла: Лактация_коров_<кол-во>_<годы>
    if displayed_df is None or displayed_df.empty:
        return "Лактация_коров"
    count = len(displayed_df)
    dates = pd.to_datetime(
        displayed_df.get('Дата планового отела', pd.Series(dtype=str)),
        dayfirst=True, errors='coerce').dropna()
    if dates.empty:
        dates = pd.to_datetime(
            displayed_df.get('Дата отела', pd.Series(dtype=str)),
            dayfirst=True, errors='coerce').dropna()
    if not dates.empty:
        yr_min = dates.dt.year.min()
        yr_max = dates.dt.year.max()
        period = f"{yr_min}" if yr_min == yr_max else f"{yr_min}-{yr_max}"
    else:
        period = "период_неизвестен"
    return f"Лактация_коров_{count}_{period}"


def export_file():
    if displayed_df is None or displayed_df.empty:
        _notify('Сначала загрузите файл.', 'red')
        return
    fp = filedialog.asksaveasfilename(
        defaultextension=".csv",
        initialfile=_main_export_filename(),
        filetypes=[("CSV", "*.csv")]
    )
    if not fp:
        return
    try:
        displayed_df.to_csv(fp, index=False, sep=';', encoding='utf-8-sig')
        _notify(f'Сохранено: {fp}', 'green')
    except Exception as e:
        _notify(f'Ошибка экспорта: {e}', 'red')

#  Экспорт выбранного месяца
def export_detail():
    if detail_df is None or detail_df.empty:
        _notify('В этом месяце нет коров', 'red')
        return
    count   = len(detail_df)
    mn_name = (f"{MONTH_NAMES_RU[current_month_key.month-1]}_{current_month_key.year}"
               if current_month_key else "месяц")
    fp = filedialog.asksaveasfilename(
        defaultextension=".csv",
        initialfile=f"Лактация_коров_{mn_name}_{count}",
        filetypes=[("CSV", "*.csv")]
    )
    if not fp:
        return
    try:
        detail_df.to_csv(fp, index=False, sep=';', encoding='utf-8-sig')
        _notify(f'Сохранено: {fp}', 'green')
    except Exception as e:
        _notify(f'Ошибка экспорта: {e}', 'red')

#  Фильтр — неосеменённые коровы
def _get_filtered_df():
    if original_df is None:
        return None
    if hide_uninseminated:
        return original_df[
            original_df['Дата посл. осеменения'] != 'Не осеменялась'
        ].copy()
    return original_df.copy()


def _update_stats_label():
    # Обновляет счётчик «Всего / Скрыто / Показано» в панели управления
    if original_df is None:
        stats_label.config(text="")
        return
    total  = len(original_df)
    shown  = len(_get_filtered_df())
    hidden = total - shown
    if hidden > 0:
        stats_label.config(
            text=f"Всего: {total}  |  Скрыто: {hidden}  |  Показано: {shown}",
            foreground='#8B4513')
    else:
        stats_label.config(
            text=f"Всего: {total}  |  Все показаны",
            foreground='#1a5c1a')


def toggle_uninseminated():
    global hide_uninseminated, displayed_df
    hide_uninseminated = not hide_uninseminated
    toggle_btn.config(
        text="👁 Показать неосеменённых" if hide_uninseminated
        else "🙈 Скрыть неосеменённых")
    base = _get_filtered_df()
    if current_sort_column and current_sort_column in base.columns:
        displayed_df = _sort_series(base, current_sort_column,
                                    ascending=(current_sort_order == 'asc')).copy()
    else:
        displayed_df = base
    create_table(displayed_df)
    if current_sort_column:
        arrow = ' ▲' if current_sort_order == 'asc' else ' ▼'
        tree_widget.heading(current_sort_column,
                            text=f"{current_sort_column}{arrow}")
    _update_stats_label()

#  Сортировка — главная таблица
def sort_by_column(col_name):
    global displayed_df, current_sort_column, current_sort_order
    if displayed_df is None or displayed_df.empty:
        return
    if current_sort_column == col_name:
        current_sort_order = 'desc' if current_sort_order == 'asc' else 'asc'
    else:
        current_sort_column = col_name
        current_sort_order  = 'asc'
    for c in tree_widget['columns']:
        txt = tree_widget.heading(c, option='text').replace(' ▲', '').replace(' ▼', '')
        tree_widget.heading(c, text=txt)
    displayed_df = _sort_series(_get_filtered_df(), col_name,
                                ascending=(current_sort_order == 'asc')).copy()
    create_table(displayed_df)
    arrow = ' ▲' if current_sort_order == 'asc' else ' ▼'
    tree_widget.heading(current_sort_column, text=f"{current_sort_column}{arrow}")


def reset_view():
    # Сбрасывает фильтр и сортировку
    global displayed_df, current_sort_column, current_sort_order, hide_uninseminated
    if original_df is not None:
        hide_uninseminated  = False
        current_sort_column = current_sort_order = None
        displayed_df = original_df.copy()
        toggle_btn.config(text="🙈 Скрыть неосеменённых")
        create_table(displayed_df)
        _update_stats_label()

#  Главная таблица
def create_table(df_data):
    global tree_widget
    cols = list(df_data.columns)
    if tree_widget:
        tree_widget.destroy()
    tree_widget = ttk.Treeview(frame_table, columns=cols, show='headings')
    for i, col in enumerate(cols):
        tree_widget.heading(col, text=col, command=lambda c=col: sort_by_column(c))
        col_width = 200 if i == len(cols) - 1 else 150
        tree_widget.column(col, anchor="center", width=col_width,
                           stretch=(i == len(cols) - 1), minwidth=80)
    scrollbar_v = ttk.Scrollbar(frame_table, orient="vertical",   command=tree_widget.yview)
    scrollbar_h = ttk.Scrollbar(frame_table, orient="horizontal", command=tree_widget.xview)
    tree_widget.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
    tree_widget.grid(row=0, column=0, sticky="nsew")
    scrollbar_v.grid(row=0, column=1, sticky="ns")
    scrollbar_h.grid(row=1, column=0, sticky="ew")
    frame_table.grid_rowconfigure(0, weight=1)
    frame_table.grid_columnconfigure(0, weight=1)
    for row in df_data.itertuples(index=False):
        tree_widget.insert("", "end", values=list(row))

#  Календарь
frame_cal = ttk.LabelFrame(frame_detail, text="Календарь отёлов", padding=(6, 6))
frame_cal.grid(row=0, column=0, sticky="ns", padx=(0, 6))

frame_cow_list = ttk.LabelFrame(frame_detail, text="Коровы — выберите месяц",
                                padding=(5, 5))
frame_cow_list.grid(row=0, column=1, sticky="nsew")
frame_cow_list.grid_rowconfigure(2, weight=1)
frame_cow_list.grid_columnconfigure(0, weight=1)

#  Календарь — логика
def _count_by_month():
    # Сколько коров телится в каждом месяце
    counts = {}
    if original_df is None:
        return counts
    dates = pd.to_datetime(original_df['Дата планового отела'],
                           dayfirst=True, errors='coerce')
    for d in dates.dropna():
        key = d.date().replace(day=1)
        counts[key] = counts.get(key, 0) + 1
    return counts


def _cal_year_range():
    # Возвращает минимальный/максимальный год по данным таблицы
    today_year = date.today().year
    if original_df is None:
        return today_year, today_year
    dates = pd.to_datetime(original_df['Дата планового отела'],
                           dayfirst=True, errors='coerce').dropna()
    if dates.empty:
        return today_year, today_year
    return int(dates.dt.year.min()), int(dates.dt.year.max())


def build_month_calendar():
    # Строит сетку 4×3
    # Навигация
    global month_buttons, cal_year
    for widget in frame_cal.winfo_children():
        widget.destroy()
    month_buttons.clear()

    if original_df is None:
        return

    today  = date.today()
    counts = _count_by_month()
    year_min, year_max = _cal_year_range()
    cal_year = max(year_min, min(year_max, cal_year))

    # Навигационная строка
    nav_frame = ttk.Frame(frame_cal)
    nav_frame.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 6))
    nav_frame.columnconfigure(1, weight=1)

    def go_prev():
        global cal_year
        if cal_year > year_min:
            cal_year -= 1
            build_month_calendar()

    def go_next():
        global cal_year
        if cal_year < year_max:
            cal_year += 1
            build_month_calendar()

    tk.Button(nav_frame, text="◀", width=3, relief="flat",
              font=('Arial', 12, 'bold'),
              fg="#003366" if cal_year > year_min else "#cccccc",
              command=go_prev).grid(row=0, column=0, padx=(0, 4))

    ttk.Label(nav_frame, text=str(cal_year),
              font=('Arial', 13, 'bold'), anchor='center').grid(row=0, column=1)

    tk.Button(nav_frame, text="▶", width=3, relief="flat",
              font=('Arial', 12, 'bold'),
              fg="#003366" if cal_year < year_max else "#cccccc",
              command=go_next).grid(row=0, column=2, padx=(4, 0))

    for month_idx in range(12):
        mk  = date(cal_year, month_idx + 1, 1)
        cnt = counts.get(mk, 0)
        is_current_month = (mk.year == today.year and mk.month == today.month)

        if is_current_month:
            bg, fg = "#42aaff", "#003366"   # голубой — текущий месяц
        elif cnt > 0:
            bg, fg = "#008000", "#1a5c1a"   # зелёный — есть коровы
        else:
            bg, fg = "#4e5754", "#666666"   # серый — нет коров

        row_idx, col_idx = divmod(month_idx, 4)
        btn = tk.Button(
            frame_cal,
            text=f"{MONTH_NAMES_RU[month_idx]}\n🐄  {cnt}",
            width=12, height=3,
            bg=bg, fg=fg,
            activebackground="#ffd700",
            relief="groove",
            font=('Arial', 11, 'bold'),
            command=lambda m=mk: on_month_click(m)
        )
        btn.grid(row=row_idx + 1, column=col_idx, padx=3, pady=3)
        month_buttons[mk] = btn


def on_month_click(month_key):
    today  = date.today()
    counts = _count_by_month()
    for mk, btn in month_buttons.items():
        cnt = counts.get(mk, 0)
        is_current_month = (mk.year == today.year and mk.month == today.month)
        if is_current_month:
            bg, fg = "#d0eaff", "#003366"
        elif cnt > 0:
            bg, fg = "#c8f0c8", "#1a5c1a"
        else:
            bg, fg = "#ebebeb", "#666666"
        btn.config(bg=bg, fg=fg)
    if month_key in month_buttons:
        month_buttons[month_key].config(bg="#ffd700", fg="#000000")
    show_cows_for_month(month_key)

#  Список коров по выбранному месяцу
def _refresh_detail_tree(filtered_df):
    for item in detail_cow_tree.get_children():
        detail_cow_tree.delete(item)
    for row in filtered_df.itertuples(index=False):
        detail_cow_tree.insert("", "end", values=list(row))


def show_cows_for_month(month_key):
    global detail_cow_tree, detail_df, detail_sort_column, detail_sort_order, current_month_key
    detail_sort_column = None
    detail_sort_order  = None
    current_month_key  = month_key

    for widget in frame_cow_list.winfo_children():
        widget.destroy()

    mn = f"{MONTH_NAMES_RU[month_key.month-1]} {month_key.year}"
    frame_cow_list.config(text=f"Коровы — {mn}")

    if original_df is None:
        return

    # Фильтруем коров с плановым отёлом в выбранном месяце
    dates   = pd.to_datetime(original_df['Дата планового отела'],
                             dayfirst=True, errors='coerce')
    mask    = dates.dt.to_period('M') == pd.Period(month_key, freq='M')
    base_df = original_df[mask].copy()
    detail_df = base_df.copy()

    # Строка с количеством и кнопкой экспорта
    info_frame = ttk.Frame(frame_cow_list)
    info_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 2))
    info_frame.columnconfigure(0, weight=1)

    ttk.Label(info_frame,
              text=f"Готовы к отёлу: {len(base_df)} ед.",
              font=('Arial', 11, 'bold'),
              foreground='#1a5c1a').grid(row=0, column=0, sticky="w")

    ttk.Button(info_frame, text="📤 Экспорт месяца",
               command=export_detail).grid(row=0, column=1, padx=(10, 0))

    if base_df.empty:
        ttk.Label(frame_cow_list, text="Нет данных для этого месяца.",
                  foreground='gray').grid(row=1, column=0)
        return

    ttk.Label(frame_cow_list,
              text="Нажмите на заголовок столбца для сортировки",
              foreground='gray', font=('Arial', 9)
              ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 2))

    # Таблица коров
    cols = list(base_df.columns)
    detail_cow_tree = ttk.Treeview(frame_cow_list, columns=cols, show='headings')
    frame_cow_list.grid_rowconfigure(2, weight=1)

    for col in cols:
        detail_cow_tree.heading(col, text=col,
                                command=lambda c=col, b=base_df:
                                    _detail_sort_and_refresh(c, b))
        detail_cow_tree.column(col, anchor="center", width=150, minwidth=80)

    scrollbar_v = ttk.Scrollbar(frame_cow_list, orient="vertical",
                                 command=detail_cow_tree.yview)
    scrollbar_h = ttk.Scrollbar(frame_cow_list, orient="horizontal",
                                 command=detail_cow_tree.xview)
    detail_cow_tree.configure(yscrollcommand=scrollbar_v.set,
                              xscrollcommand=scrollbar_h.set)
    detail_cow_tree.grid(row=2, column=0, sticky="nsew")
    scrollbar_v.grid(row=2, column=1, sticky="ns")
    scrollbar_h.grid(row=3, column=0, sticky="ew")

    _refresh_detail_tree(base_df)


def _detail_sort_and_refresh(col_name, base_df):
    # Сортировка таблицы коров по выбранному столбцу
    global detail_df, detail_sort_column, detail_sort_order
    if detail_sort_column == col_name:
        detail_sort_order = 'desc' if detail_sort_order == 'asc' else 'asc'
    else:
        detail_sort_column = col_name
        detail_sort_order  = 'asc'
    for c in detail_cow_tree['columns']:
        txt = detail_cow_tree.heading(c, option='text').replace(' ▲', '').replace(' ▼', '')
        detail_cow_tree.heading(c, text=txt)
    arrow = ' ▲' if detail_sort_order == 'asc' else ' ▼'
    detail_cow_tree.heading(col_name, text=f"{col_name}{arrow}")
    sorted_df = _sort_series(base_df, col_name, ascending=(detail_sort_order == 'asc'))
    detail_df = sorted_df.copy()
    _refresh_detail_tree(sorted_df)


def clear_cow_detail():
    # Очищает правую панель и показывает подсказку
    global detail_df, current_month_key
    detail_df         = None
    current_month_key = None
    for widget in frame_cow_list.winfo_children():
        widget.destroy()
    frame_cow_list.config(text="Коровы")
    ttk.Label(frame_cow_list,
              text="← Выберите месяц",
              foreground='gray',
              font=('Arial', 11)).grid(row=0, column=0, padx=15, pady=15)
clear_cow_detail()

#  Панель управления
ttk.Button(frame_controls, text="📂 Импортировать CSV",
           command=load_file).grid(row=0, column=0, padx=5, pady=5)
ttk.Button(frame_controls, text="🧩 Экспорт в CSV",
           command=export_file).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_controls, text="⟲ Сбросить",
           command=reset_view).grid(row=0, column=2, padx=5, pady=5)

toggle_btn = ttk.Button(frame_controls, text="🙈 Скрыть неосеменённых",
                        command=toggle_uninseminated)
toggle_btn.grid(row=0, column=3, padx=5, pady=5)

# Счётчик коров
stats_label = ttk.Label(frame_controls, text="", font=('Arial', 9))
stats_label.grid(row=0, column=4, padx=(15, 5), sticky="w")

# Строка уведомлений (растягивается между счётчиком и кнопкой шаблона)
frame_controls.columnconfigure(5, weight=1)
error_label = ttk.Label(frame_controls, text="", font=('Arial', 9), anchor='center')
error_label.grid(row=0, column=5, padx=(5, 5), sticky="ew")

# Кнопка шаблона — крайняя справа
ttk.Button(frame_controls, text="📋 Скачать шаблон",
           command=download_template).grid(row=0, column=6, padx=(5, 8), pady=5)

#  Закрытие окна
def on_closing():
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)

if __name__ == '__main__':
    root.mainloop()


# In[ ]:


q

