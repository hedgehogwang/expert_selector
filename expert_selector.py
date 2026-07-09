import os
import random
import re
import tkinter as tk
from tkinter import END, LEFT, RIGHT, BOTH, EW, W, filedialog, messagebox

import pandas as pd
import ttkbootstrap as tb
from ttkbootstrap.constants import *

EXPECTED_COLUMNS = [
    '序号', '姓名', '性别', '出生年月', '年龄', '学历', '工作单位',
    '职务/职业资格', '职称/职称层级', '手机号码', '申报专业', '业绩成果',
    '推荐部门', '备注'
]
FILTER_LABELS = ['学历', '职称/职称层级', '申报专业', '推荐部门', '所需人数']


def normalize_column_name(name: str) -> str:
    if name is None:
        return ''
    text = str(name)
    text = text.replace('\ufeff', '')
    text = text.replace('\xa0', ' ')
    text = text.replace('／', '/')
    text = text.replace('：', ':')
    text = text.replace('，', ',')
    text = text.replace('；', ';')
    text = re.sub(r'[\s]+', ' ', text).strip()
    return text


def canonical_column_name(name: str) -> str:
    text = normalize_column_name(name)
    text = re.sub(r'[\W_]+', '', text, flags=re.UNICODE)
    return text.lower()


def detect_header_row(file_path: str, max_header_rows: int = 5):
    try:
        preview = pd.read_excel(file_path, dtype=str, header=None, nrows=max_header_rows + 5)
    except Exception:
        return None
    expected_canonicals = {canonical_column_name(c) for c in EXPECTED_COLUMNS}
    for row_index in range(min(max_header_rows, len(preview))):
        candidate = [canonical_column_name(x) for x in preview.iloc[row_index].tolist()]
        if expected_canonicals.issubset(set(candidate)):
            return row_index
    return None


def filter_experts(dataframe: pd.DataFrame, groups, excluded_units=None) -> pd.DataFrame:
    if dataframe is None or dataframe.empty:
        return pd.DataFrame(columns=EXPECTED_COLUMNS)

    available = dataframe.copy()
    excluded_units = excluded_units or []
    if excluded_units:
        available = available[~available['工作单位'].astype(str).isin(excluded_units)]

    selected_groups = []
    selected_indices = set()
    for group in groups or []:
        if not isinstance(group, dict):
            continue

        conditions = {}
        for key in ['学历', '职称/职称层级', '申报专业', '推荐部门']:
            value = str(group.get(key, '')).strip()
            if value:
                conditions[key] = value

        count_value = str(group.get('所需人数', '')).strip()
        if not conditions and not count_value:
            continue

        subset = available.copy()
        for label, value in conditions.items():
            if label == '申报专业':
                subset = subset[subset[label].astype(str).str.contains(value, case=False, na=False)]
            else:
                subset = subset[subset[label].astype(str).str.strip() == value]

        if subset.empty:
            continue

        subset = subset[~subset.index.isin(selected_indices)]
        if subset.empty:
            continue

        if count_value:
            try:
                required = int(count_value)
            except ValueError as exc:
                raise ValueError(f'所需人数必须为整数：{count_value}') from exc
            if required <= 0:
                continue
            if required >= len(subset):
                selected = subset.copy()
            else:
                selected = subset.sample(required, random_state=random.randint(1, 999999))
        else:
            selected = subset.copy()

        selected_groups.append(selected)
        selected_indices.update(selected.index.tolist())

    if selected_groups:
        return pd.concat(selected_groups, ignore_index=True)
    return available.copy()


class ExpertSelectorApp(tb.Window):
    def __init__(self):
        super().__init__(themename='flatly')
        self.title('随机专家抽取器')
        self.geometry('1400x900')
        self.resizable(True, True)
        self.state('zoomed')
        self.bind('<F11>', self.toggle_fullscreen)
        self.bind('<Escape>', self.exit_fullscreen)

        self.data = None
        self.units = []
        self.filter_options = {}
        self.last_filtered_df = None
        self.filter_groups = []

        self.create_widgets()

    def toggle_fullscreen(self, event=None):
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))
        return 'break'

    def exit_fullscreen(self, event=None):
        if self.attributes('-fullscreen'):
            self.attributes('-fullscreen', False)
            return 'break'
        return None

    def create_widgets(self):
        top_frame = tb.Frame(self)
        top_frame.pack(fill=X, padx=12, pady=8)

        btn_open = tb.Button(top_frame, text='选择 Excel 文件', bootstyle='success-outline', command=self.open_file)
        btn_open.grid(row=0, column=0, padx=6, pady=4, sticky=W)

        self.label_file = tb.Label(top_frame, text='未选择文件', bootstyle='secondary')
        self.label_file.grid(row=0, column=1, padx=6, pady=4, sticky=W)

        filter_frame = tb.Labelframe(self, text='筛选条件', padding=10)
        filter_frame.pack(fill=X, padx=12, pady=8)

        self.filter_groups_frame = tb.Frame(filter_frame)
        self.filter_groups_frame.pack(fill=X, expand=True)

        for idx, label in enumerate(FILTER_LABELS):
            tb.Label(self.filter_groups_frame, text=label + '：').grid(row=0, column=idx, sticky=EW, padx=4, pady=2)
            self.filter_groups_frame.columnconfigure(idx, weight=1)
        tb.Label(self.filter_groups_frame, text='操作').grid(row=0, column=len(FILTER_LABELS), sticky=EW, padx=4, pady=2)
        self.filter_groups_frame.columnconfigure(len(FILTER_LABELS), weight=0)

        self.add_filter_group()

        toolbar = tb.Frame(filter_frame)
        toolbar.pack(fill=X, anchor=W, pady=6)
        tb.Button(toolbar, text='+ 新增条件组', bootstyle='success', command=self.add_filter_group).pack(side=LEFT, padx=4)

        exclude_frame = tb.Labelframe(self, text='需排除单位', padding=10)
        exclude_frame.pack(fill=X, padx=12, pady=8)

        tb.Label(exclude_frame, text='输入单位名称关键字：').grid(row=0, column=0, sticky=E, padx=4, pady=4)
        self.combo_exclude = tb.Combobox(exclude_frame, values=[], state='normal', bootstyle='secondary')
        self.combo_exclude.grid(row=0, column=1, sticky=EW, padx=4, pady=4)
        self.combo_exclude.bind('<KeyRelease>', self.on_exclude_input)
        self.combo_exclude.bind('<<ComboboxSelected>>', self.on_suggestion_selected)

        tb.Label(exclude_frame, text='已排除工作单位（用逗号分隔）：').grid(row=1, column=0, sticky=NE, padx=4, pady=4)
        self.text_exclude_list = tb.Text(exclude_frame, height=3)
        self.text_exclude_list.grid(row=1, column=1, sticky=EW, padx=4, pady=4)

        btn_frame = tb.Frame(self)
        btn_frame.pack(fill=X, padx=12, pady=8)

        tb.Button(btn_frame, text='筛选', bootstyle='primary', command=self.apply_filters).pack(side=LEFT, padx=4)
        tb.Button(btn_frame, text='导出', bootstyle='info', command=self.export_results).pack(side=LEFT, padx=4)
        tb.Button(btn_frame, text='清空', bootstyle='outline-secondary', command=self.clear_inputs).pack(side=LEFT, padx=4)

        result_frame = tb.Labelframe(self, text='筛选结果', padding=10)
        result_frame.pack(fill=BOTH, expand=True, padx=12, pady=8)

        self.result_status = tb.Label(result_frame, text='等待筛选', bootstyle='secondary')
        self.result_status.pack(anchor=W, pady=(0, 6))

        self.tree = tb.Treeview(result_frame, columns=EXPECTED_COLUMNS[1:], show='headings', height=20)
        for col in EXPECTED_COLUMNS[1:]:
            heading = '出生日期' if col == '出生年月' else col
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=120, anchor=W)
        self.tree.pack(fill=BOTH, expand=True, side=LEFT)

        vsb = tb.Scrollbar(result_frame, orient='vertical', command=self.tree.yview)
        vsb.pack(side=RIGHT, fill=Y)
        self.tree.configure(yscrollcommand=vsb.set)

    def add_filter_group(self):
        if len(self.filter_groups) >= 5:
            messagebox.showwarning('最多5组', '最多只能添加5组筛选条件。')
            return

        values = []
        for group in self.filter_groups:
            values.append({label: widget.get().strip() for label, widget in group.items()})
        values.append({label: '' for label in FILTER_LABELS})
        self.render_filter_groups(values)

    def remove_filter_group(self, index):
        values = []
        for group in self.filter_groups:
            values.append({label: widget.get().strip() for label, widget in group.items()})
        if 0 <= index < len(values):
            values.pop(index)
        if not values:
            values = [{label: '' for label in FILTER_LABELS}]
        self.render_filter_groups(values)

    def render_filter_groups(self, group_values=None):
        if group_values is None:
            group_values = [{label: '' for label in FILTER_LABELS} for _ in range(max(1, len(self.filter_groups)))]

        for child in list(self.filter_groups_frame.winfo_children()):
            child.destroy()

        for idx, label in enumerate(FILTER_LABELS):
            tb.Label(self.filter_groups_frame, text=label + '：').grid(row=0, column=idx, sticky=EW, padx=4, pady=2)
            self.filter_groups_frame.columnconfigure(idx, weight=1)
        tb.Label(self.filter_groups_frame, text='操作').grid(row=0, column=len(FILTER_LABELS), sticky=EW, padx=4, pady=2)
        self.filter_groups_frame.columnconfigure(len(FILTER_LABELS), weight=0)

        self.filter_groups = []
        for row, values in enumerate(group_values, start=1):
            group_widgets = {}
            for col_idx, label in enumerate(FILTER_LABELS):
                if label == '所需人数':
                    widget = tb.Entry(self.filter_groups_frame, bootstyle='secondary')
                else:
                    widget = tb.Combobox(
                        self.filter_groups_frame,
                        values=self.filter_options.get(label, ['']),
                        state='normal',
                        bootstyle='secondary'
                    )
                widget.grid(row=row, column=col_idx, sticky=EW, padx=4, pady=2)
                widget_value = values.get(label, '')
                if label == '所需人数':
                    widget.delete(0, END)
                    widget.insert(0, widget_value)
                else:
                    widget.set(widget_value)
                group_widgets[label] = widget
                self.filter_groups_frame.columnconfigure(col_idx, weight=1)

            if row > 1:
                btn = tb.Button(
                    self.filter_groups_frame,
                    text='-',
                    bootstyle='danger-outline',
                    width=3,
                    command=lambda r=row - 1: self.remove_filter_group(r)
                )
                btn.grid(row=row, column=len(FILTER_LABELS), sticky=W, padx=4, pady=2)
            else:
                tb.Label(self.filter_groups_frame, text='').grid(row=row, column=len(FILTER_LABELS))
            self.filter_groups.append(group_widgets)

    def on_exclude_input(self, event=None):
        text = self.combo_exclude.get().strip()
        self.update_suggestions(text)

    def on_suggestion_selected(self, event=None):
        unit = self.combo_exclude.get().strip()
        if not unit:
            return
        existing = self.get_excluded_units()
        if unit not in existing:
            existing.append(unit)
            self.text_exclude_list.delete('1.0', END)
            self.text_exclude_list.insert('1.0', '，'.join(existing))
        self.combo_exclude.set('')

    def update_suggestions(self, text):
        text_lower = text.lower()
        if text_lower and self.units:
            matches = [unit for unit in self.units if text_lower in unit.lower()]
            self.combo_exclude.config(values=matches[:50])
        else:
            self.combo_exclude.config(values=[])

    def open_file(self):
        file_path = filedialog.askopenfilename(
            title='选择 Excel 文件',
            filetypes=[('Excel 文件', '*.xlsx')]
        )
        if not file_path:
            return

        header_row = detect_header_row(file_path)
        if header_row is None:
            messagebox.showerror('列缺失', '未能在 Excel 文件前几行中检测到完整的表头，请确保表头在前 5 行内且列名正确。')
            self.data = None
            return

        try:
            self.data = pd.read_excel(file_path, dtype=str, header=header_row)
        except Exception as exc:
            messagebox.showerror('读取失败', f'无法读取文件：{exc}')
            return

        column_map = {}
        expected_map = {canonical_column_name(c): c for c in EXPECTED_COLUMNS}
        for col in self.data.columns:
            canon = canonical_column_name(col)
            if canon in expected_map:
                column_map[col] = expected_map[canon]
            else:
                column_map[col] = normalize_column_name(col)
        self.data = self.data.rename(columns=column_map)

        missing = [c for c in EXPECTED_COLUMNS if c not in self.data.columns]
        if missing:
            messagebox.showerror('列缺失', f'Excel 文件缺少以下列：{", ".join(missing)}')
            self.data = None
            return

        self.label_file.config(text=os.path.basename(file_path))
        self.data = self.data[EXPECTED_COLUMNS].fillna('')
        self.data = self.data.astype(str)
        self.units = sorted({value for value in self.data['工作单位'].dropna().astype(str).tolist() if value})
        self.update_filter_options()
        self.update_suggestions('')
        self.clear_results()
        self.result_status.config(text='已加载 Excel 数据，准备筛选')

    def update_filter_options(self):
        self.filter_options = {}
        for label in FILTER_LABELS:
            if label == '所需人数':
                continue
            options = sorted({value for value in self.data[label].dropna().astype(str).tolist() if value})
            self.filter_options[label] = [''] + options

        for group in self.filter_groups:
            for label, widget in group.items():
                if label == '所需人数':
                    continue
                widget.config(values=self.filter_options.get(label, ['']))
                current_value = widget.get().strip()
                if current_value not in self.filter_options.get(label, []):
                    widget.set('')

    def get_excluded_units(self):
        raw = self.text_exclude_list.get('1.0', END).strip().replace('\n', ' ')
        if not raw:
            return []
        parts = re.split(r'[，,;；]+', raw)
        return [p.strip() for p in parts if p.strip()]

    def format_birth_date(self, value):
        if value is None:
            return ''
        try:
            text = str(value).strip()
            if not text:
                return ''
            if '-' in text:
                parts = text.split('-')
            elif '/' in text:
                parts = text.split('/')
            elif '年' in text and '月' in text:
                return re.sub(r'[^0-9年月]', '', text)
            else:
                parts = re.split(r'[^0-9]+', text)
            if not parts:
                return text
            year = int(parts[0]) if parts and parts[0].isdigit() else None
            month = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else None
            if year is None:
                return text
            if month is None:
                return f'{year}年'
            return f'{year}年{int(month):02d}月'
        except Exception:
            return str(value)

    def apply_filters(self):
        if self.data is None:
            messagebox.showwarning('未选择文件', '请先选择一个 Excel 文件。')
            return

        excluded = self.get_excluded_units()
        groups = []
        for group in self.filter_groups:
            values = {}
            for label, widget in group.items():
                if label == '所需人数':
                    values[label] = widget.get().strip()
                else:
                    values[label] = widget.get().strip()
            if any(values.values()):
                groups.append(values)

        try:
            df = filter_experts(self.data, groups, excluded_units=excluded)
        except ValueError as exc:
            messagebox.showwarning('筛选参数错误', str(exc))
            return

        self.last_filtered_df = df.copy()
        self.show_results(df)
        if df.empty:
            self.result_status.config(text='筛选完成：未找到匹配专家')
            messagebox.showinfo('结果为空', '当前筛选条件下无符合的专家。')
        else:
            self.result_status.config(text=f'筛选完成：共找到 {len(df)} 位专家')

    def show_results(self, df):
        self.clear_results()
        for _, row in df.iterrows():
            values = []
            for col in EXPECTED_COLUMNS[1:]:
                value = row[col]
                if col == '出生年月':
                    value = self.format_birth_date(value)
                values.append(value)
            self.tree.insert('', tk.END, values=values)

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def clear_inputs(self):
        for group in self.filter_groups:
            for widget in group.values():
                widget.set('')
        self.combo_exclude.set('')
        self.text_exclude_list.delete('1.0', END)
        self.clear_results()
        self.result_status.config(text='已清空输入与结果')

    def export_results(self):
        if not hasattr(self, 'last_filtered_df') or self.last_filtered_df is None or self.last_filtered_df.empty:
            messagebox.showwarning('导出失败', '当前没有筛选结果可导出。')
            return

        now = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        default_name = f'筛选结果_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.csv'
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV 文件', '*.csv')],
            initialfile=default_name,
            title='导出筛选结果'
        )
        if not file_path:
            return

        excluded = self.get_excluded_units()
        criteria = []
        for index, group in enumerate(self.filter_groups, start=1):
            conditions = []
            for label in ['学历', '职称/职称层级', '申报专业', '推荐部门']:
                value = group[label].get().strip()
                if value:
                    conditions.append(f'{label}={value}')
            count_value = group['所需人数'].get().strip()
            if count_value:
                conditions.append(f'所需人数={count_value}')
            if conditions:
                criteria.append(f'组{index}: ' + '；'.join(conditions))
            else:
                criteria.append(f'组{index}: 未设置条件')

        try:
            with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                f.write(f'生成时间,{now}\n')
                f.write(f'筛选条件,{" | ".join(criteria)}\n')
                f.write(f'排除单位,{"，".join(excluded) if excluded else "无"}\n')
                f.write('\n')
                self.last_filtered_df.to_csv(f, index=False)
            messagebox.showinfo('导出成功', f'筛选结果已导出到：{file_path}')
        except Exception as exc:
            messagebox.showerror('导出失败', f'无法保存文件：{exc}')


if __name__ == '__main__':
    app = ExpertSelectorApp()
    app.mainloop()
