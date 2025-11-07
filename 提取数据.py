import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import re
import pandas as pd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import threading
import queue
import time
import json
import os.path


class DatFileExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("日志文件数值提取工具")
        self.root.geometry("1050x700")
        self.root.minsize(1000, 600)

        plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
        plt.rcParams["axes.unicode_minus"] = False  # 正确显示负号

        self.selected_files = []
        self.file_checkboxes = {}
        self.extracted_single = {}
        self.extracted_double = {}
        self.extracted_dual = {}
        self.extracted_triple = {}
        self.search_text1 = ""
        self.search_text2 = ""
        self.status_queue = queue.Queue()
        self.process_data = tk.BooleanVar(value=False)
        self.extract_mode = tk.IntVar(value=1)
        self.stop_extraction = threading.Event()

        # 关键词历史记录
        self.keyword_history1 = []
        self.keyword_history2 = []
        self.load_keyword_history()

        self.file_types = {
            "dat": tk.BooleanVar(value=True),
            "txt": tk.BooleanVar(value=False),
            "xlsx": tk.BooleanVar(value=False),
            "log": tk.BooleanVar(value=False),
            "custom": tk.BooleanVar(value=False)
        }
        self.custom_ext = tk.StringVar(value="")

        self.folder_mode = tk.BooleanVar(value=False)
        self.select_all_var = tk.BooleanVar(value=True)
        self.progress_var = tk.DoubleVar(value=0)

        # 添加拖动功能所需的状态变量
        self.drag_data = {'item': None, 'x': 0, 'y': 0}

        self.create_widgets()
        self.root.after(100, self.update_status)

    def load_keyword_history(self):
        """加载关键词历史记录"""
        try:
            if os.path.exists("keyword_history.json"):
                with open("keyword_history.json", "r", encoding="utf-8") as f:
                    history = json.load(f)
                    self.keyword_history1 = history.get("keyword1", [])
                    self.keyword_history2 = history.get("keyword2", [])
        except:
            self.keyword_history1 = []
            self.keyword_history2 = []

    def save_keyword_history(self):
        """保存关键词历史记录"""
        try:
            history = {
                "keyword1": list(set(self.keyword_history1))[:20],  # 去重并保留最多20条
                "keyword2": list(set(self.keyword_history2))[:20]
            }
            with open("keyword_history.json", "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except:
            pass

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        file_selection_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=10)
        file_selection_frame.pack(fill=tk.X, pady=(0, 10))
        file_selection_frame.grid_columnconfigure(0, weight=1)

        mode_frame = ttk.Frame(file_selection_frame)
        mode_frame.grid(row=0, column=0, sticky=tk.W, padx=5)

        ttk.Radiobutton(mode_frame, text="选择文件夹", variable=self.folder_mode, value=True).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="选择文件", variable=self.folder_mode, value=False).pack(side=tk.LEFT, padx=5)

        path_frame = ttk.Frame(file_selection_frame)
        path_frame.grid(row=1, column=0, sticky=tk.W + tk.E, padx=5, pady=5)

        ttk.Label(path_frame, text="文件路径:", padding=(0, 0, 5, 0)).pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(path_frame, width=63)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=False)

        browse_btn = ttk.Button(path_frame, text="浏览", command=self.browse_files)
        browse_btn.pack(side=tk.LEFT, padx=(0, 5))

        file_type_frame = ttk.Frame(file_selection_frame, padding=5)
        file_type_frame.grid(row=2, column=0, sticky=tk.W + tk.E, pady=5)

        ttk.Label(file_type_frame, text="文件类型:", padding=(0, 0, 5, 0)).pack(side=tk.LEFT)

        for ext in ['dat', 'txt', 'xlsx', 'log']:
            cb = ttk.Checkbutton(file_type_frame, text=f'.{ext}', variable=self.file_types[ext])
            cb.pack(side=tk.LEFT, padx=3)

        custom_frame = ttk.Frame(file_type_frame)
        custom_frame.pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(custom_frame, text="自定义:", variable=self.file_types["custom"]).pack(side=tk.LEFT)
        ttk.Entry(custom_frame, textvariable=self.custom_ext, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Label(custom_frame, text="(如: .csv)").pack(side=tk.LEFT)

        extract_settings_frame = ttk.LabelFrame(main_frame, text="提取设置", padding=10)
        extract_settings_frame.pack(fill=tk.X, pady=(0, 10))
        extract_settings_frame.grid_columnconfigure(0, weight=1)

        search_frame = ttk.Frame(extract_settings_frame)
        search_frame.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        keyword1_frame = ttk.Frame(search_frame)
        keyword1_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(keyword1_frame, text="关键词1:").pack(side=tk.LEFT, padx=5)

        # 将关键词1输入框改为下拉选择框
        self.search_entry1 = ttk.Combobox(keyword1_frame, width=23)
        self.search_entry1.pack(side=tk.LEFT, padx=5)
        self.search_entry1["values"] = self.keyword_history1
        if self.keyword_history1:
            self.search_entry1.set(self.keyword_history1[0])

        # 添加快捷按钮
        btn_frame1 = ttk.Frame(keyword1_frame)
        btn_frame1.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame1, text="+", width=2, command=lambda: self.add_keyword_to_history(1)).pack(side=tk.LEFT,
                                                                                                       padx=2)
        ttk.Button(btn_frame1, text="-", width=2, command=lambda: self.remove_keyword_from_history(1)).pack(
            side=tk.LEFT, padx=2)

        keyword2_frame = ttk.Frame(search_frame)
        keyword2_frame.pack(fill=tk.X)

        ttk.Label(keyword2_frame, text="关键词2:").pack(side=tk.LEFT, padx=5)

        # 将关键词2输入框改为下拉选择框
        self.search_entry2 = ttk.Combobox(keyword2_frame, width=23)
        self.search_entry2.pack(side=tk.LEFT, padx=5)
        self.search_entry2["values"] = self.keyword_history2
        if self.keyword_history2:
            self.search_entry2.set(self.keyword_history2[0])

        # 添加快捷按钮
        btn_frame2 = ttk.Frame(keyword2_frame)
        btn_frame2.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame2, text="+", width=2, command=lambda: self.add_keyword_to_history(2)).pack(side=tk.LEFT,
                                                                                                       padx=2)
        ttk.Button(btn_frame2, text="-", width=2, command=lambda: self.remove_keyword_from_history(2)).pack(
            side=tk.LEFT, padx=2)

        mode_frame = ttk.Frame(extract_settings_frame, padding=(10, 0))
        mode_frame.grid(row=0, column=1, sticky=tk.W)

        ttk.Radiobutton(mode_frame, text="提取单值", variable=self.extract_mode, value=1).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="提取双值", variable=self.extract_mode, value=2).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="提取三值", variable=self.extract_mode, value=4).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="双关键词值提取", variable=self.extract_mode, value=3).pack(side=tk.LEFT,
                                                                                                     padx=5)

        btn_frame = ttk.Frame(extract_settings_frame, padding=(10, 0))
        btn_frame.grid(row=0, column=2, sticky=tk.E)

        ttk.Button(btn_frame, text="提取数据", command=self.extract_data, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="停止提取", command=self.stop_extraction_thread, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出Excel", command=self.export_to_excel, width=12).pack(side=tk.LEFT, padx=5)

        process_frame = ttk.Frame(extract_settings_frame, padding=(10, 0))
        process_frame.grid(row=0, column=3, sticky=tk.W)
        ttk.Checkbutton(process_frame, text="数据处理", variable=self.process_data).pack(side=tk.LEFT, padx=5)

        progress_bar = ttk.Progressbar(extract_settings_frame, variable=self.progress_var, length=400)
        progress_bar.grid(row=1, column=0, columnspan=4, sticky=tk.W + tk.E, padx=5, pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        files_tab = ttk.Frame(notebook)
        notebook.add(files_tab, text="文件列表")

        files_tree_frame = ttk.Frame(files_tab)
        files_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        select_all_frame = ttk.Frame(files_tree_frame, padding=(0, 0, 5, 5))
        select_all_frame.pack(anchor=tk.W)

        ttk.Checkbutton(
            select_all_frame,
            text="全选",
            variable=self.select_all_var,
            command=self.toggle_select_all
        ).pack(side=tk.LEFT)

        files_scroll = ttk.Scrollbar(files_tree_frame)
        files_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.files_tree = ttk.Treeview(files_tree_frame, yscrollcommand=files_scroll.set)
        self.files_tree.pack(fill=tk.BOTH, expand=True)
        files_scroll.config(command=self.files_tree.yview)

        self.files_tree["columns"] = ("check", "size", "modified", "order")
        self.files_tree.column("#0", width=300, minwidth=200)
        self.files_tree.column("check", width=50, minwidth=50, anchor=tk.CENTER)
        self.files_tree.column("size", width=100, minwidth=80, anchor=tk.E)
        self.files_tree.column("modified", width=150, minwidth=120)
        self.files_tree.column("order", width=80, minwidth=60, anchor=tk.CENTER)

        self.files_tree.heading("#0", text="文件名")
        self.files_tree.heading("check", text="选择")
        self.files_tree.heading("size", text="大小")
        self.files_tree.heading("modified", text="修改日期")
        self.files_tree.heading("order", text="顺序")

        # 添加拖动功能绑定
        self.files_tree.bind("<ButtonPress-1>", self.on_press)
        self.files_tree.bind("<B1-Motion>", self.on_drag)
        self.files_tree.bind("<ButtonRelease-1>", self.on_release)

        results_tab = ttk.Frame(notebook)
        notebook.add(results_tab, text="提取结果")

        results_tree_frame = ttk.Frame(results_tab)
        results_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        results_scroll = ttk.Scrollbar(results_tree_frame)
        results_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.results_tree = ttk.Treeview(results_tree_frame, yscrollcommand=results_scroll.set)
        self.results_tree.pack(fill=tk.BOTH, expand=True)
        results_scroll.config(command=self.results_tree.yview)

        chart_tab = ttk.Frame(notebook)
        notebook.add(chart_tab, text="统计图表")

        chart_frame = ttk.Frame(chart_tab)
        chart_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.chart = FigureCanvasTkAgg(self.figure, chart_frame)
        self.chart.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def add_keyword_to_history(self, keyword_num):
        """将当前关键词添加到历史记录"""
        if keyword_num == 1:
            keyword = self.search_entry1.get().strip()
            if keyword and keyword not in self.keyword_history1:
                self.keyword_history1.insert(0, keyword)
                self.keyword_history1 = self.keyword_history1[:20]  # 保留最多20条
                self.search_entry1["values"] = self.keyword_history1
                self.save_keyword_history()
        else:
            keyword = self.search_entry2.get().strip()
            if keyword and keyword not in self.keyword_history2:
                self.keyword_history2.insert(0, keyword)
                self.keyword_history2 = self.keyword_history2[:20]  # 保留最多20条
                self.search_entry2["values"] = self.keyword_history2
                self.save_keyword_history()

    def remove_keyword_from_history(self, keyword_num):
        """从历史记录中移除当前关键词"""
        if keyword_num == 1:
            keyword = self.search_entry1.get().strip()
            if keyword in self.keyword_history1:
                self.keyword_history1.remove(keyword)
                self.search_entry1["values"] = self.keyword_history1
                self.save_keyword_history()
        else:
            keyword = self.search_entry2.get().strip()
            if keyword in self.keyword_history2:
                self.keyword_history2.remove(keyword)
                self.search_entry2["values"] = self.keyword_history2
                self.save_keyword_history()

    def update_status(self):
        while not self.status_queue.empty():
            try:
                msg_type, msg = self.status_queue.get()
                if msg_type == "progress":
                    self.progress_var.set(msg)
                elif msg_type == "status":
                    self.status_var.set(msg)
                elif msg_type == "error":
                    self.status_var.set(msg)
                    messagebox.showerror("错误", msg)
                elif msg_type == "chart":
                    self.chart.draw()
            except queue.Empty:
                pass
        self.root.after(100, self.update_status)

    # 新的拖动功能实现
    def on_press(self, event):
        region = self.files_tree.identify_region(event.x, event.y)
        column = self.files_tree.identify_column(event.x)

        # 如果是复选框列，调用原来的点击处理
        if column == "#1":
            self.on_tree_click(event)
            return

        # 记录拖动起始位置
        self.drag_data['item'] = self.files_tree.identify_row(event.y)
        self.drag_data['x'] = event.x
        self.drag_data['y'] = event.y
        self.drag_data['start_y'] = event.y

        if self.drag_data['item']:
            # 高亮显示当前拖动的项目
            self.files_tree.selection_set(self.drag_data['item'])
            self.files_tree.item(self.drag_data['item'], tags=('dragging',))
            self.files_tree.tag_configure('dragging', background='#e1e1e1')

    def on_drag(self, event):
        if not self.drag_data['item']:
            return

        # 计算拖动距离
        dx = abs(event.x - self.drag_data['x'])
        dy = abs(event.y - self.drag_data['y'])

        # 只有拖动距离超过5像素才视为拖动操作
        if dx > 5 or dy > 5:
            # 移动项目到鼠标位置
            self.files_tree.config(cursor="hand2")

            # 获取当前鼠标位置的项目
            target_item = self.files_tree.identify_row(event.y)

            # 如果有目标项目
            if target_item and target_item != self.drag_data['item']:
                # 计算放置位置（目标上方或下方）
                bbox = self.files_tree.bbox(target_item)
                if event.y < bbox[1] + bbox[3] // 2:
                    # 放在目标上方
                    self.files_tree.move(self.drag_data['item'], self.files_tree.parent(target_item),
                                         self.files_tree.index(target_item))
                else:
                    # 放在目标下方
                    next_item = self.files_tree.next(target_item)
                    if next_item:
                        self.files_tree.move(self.drag_data['item'], self.files_tree.parent(target_item),
                                             self.files_tree.index(next_item))
                    else:
                        # 如果目标没有下一个兄弟，放在最后
                        self.files_tree.move(self.drag_data['item'], '', 'end')

                # 更新顺序显示
                self.update_file_order()

                # 更新起始位置
                self.drag_data['y'] = event.y

    def on_release(self, event):
        if not self.drag_data['item']:
            return

        # 恢复光标和项目样式
        self.files_tree.config(cursor="")
        self.files_tree.item(self.drag_data['item'], tags=())
        self.files_tree.tag_configure('dragging', background='')

        # 更新文件顺序
        self.update_selected_files_order()

        # 重置拖动数据
        self.drag_data = {'item': None, 'x': 0, 'y': 0, 'start_y': 0}

    def update_file_order(self):
        """更新文件顺序显示"""
        children = self.files_tree.get_children('')
        for i, item in enumerate(children, 1):
            self.files_tree.set(item, column="order", value=str(i))

    def update_selected_files_order(self):
        """更新选中文件列表的顺序以匹配当前树视图"""
        new_selected_files = []
        for item in self.files_tree.get_children(''):
            if item in self.file_checkboxes:
                file_path = self.file_checkboxes[item]['path']
                new_selected_files.append(file_path)

        self.selected_files = new_selected_files

    def get_selected_filetypes(self):
        exts = []
        for ext, var in self.file_types.items():
            if ext == "custom":
                continue
            if var.get():
                exts.append(ext)

        if self.file_types["custom"].get():
            custom = self.custom_ext.get().strip()
            if custom:
                if not custom.startswith('.'):
                    custom = '.' + custom
                exts.append(custom.lstrip('.'))

        return exts

    def browse_files(self):
        if self.folder_mode.get():
            folder_path = filedialog.askdirectory(title="选择文件夹")
            if not folder_path:
                return

            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, folder_path)

            self.selected_files = self._get_files_from_folder(folder_path)
            self.load_selected_files()
        else:
            selected_exts = self.get_selected_filetypes()
            if not selected_exts:
                messagebox.showwarning("警告", "请至少选择一种文件类型")
                return

            file_types = []
            if len(selected_exts) > 1:
                file_types.append(("选中的文件类型", " ".join(f"*.{ext}" for ext in selected_exts)))

            for ext in selected_exts:
                file_types.append((f"{ext.upper()} 文件", f"*.{ext}"))

            file_types.append(("所有文件", "*.*"))

            selected_files = filedialog.askopenfilenames(
                filetypes=file_types,
                title="选择要处理的文件"
            )

            if selected_files:
                self.selected_files = list(selected_files)
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, f"{len(selected_files)} 个文件已选择")
                self.load_selected_files()

    def _get_files_from_folder(self, folder_path):
        selected_exts = self.get_selected_filetypes()
        if not selected_exts:
            messagebox.showwarning("警告", "请至少选择一种文件类型")
            return []

        files = []
        for root, _, filenames in os.walk(folder_path):
            for filename in filenames:
                ext = os.path.splitext(filename)[1].lower().lstrip('.')
                if ext in selected_exts:
                    files.append(os.path.join(root, filename))

        if not files:
            messagebox.showwarning("警告", f"在文件夹中未找到符合条件的文件 ({', '.join(selected_exts)})")

        return files

    def load_selected_files(self):
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
        self.file_checkboxes.clear()

        valid_files = []
        selected_exts = self.get_selected_filetypes()

        for file_path in self.selected_files:
            ext = os.path.splitext(file_path)[1].lower().lstrip('.')
            if os.path.isfile(file_path) and ext in selected_exts:
                valid_files.append(file_path)

        # 初始按文件名中的数字排序
        self.selected_files = self._sort_files_by_numeric_value(valid_files)

        try:
            for i, file_path in enumerate(self.selected_files):
                file_name = os.path.basename(file_path)
                file_stats = os.stat(file_path)
                size = self.format_size(file_stats.st_size)
                modified = pd.to_datetime(file_stats.st_mtime, unit='s').strftime('%Y-%m-%d %H:%M:%S')

                value = self._extract_numeric_value(file_name)
                order_str = f"{value}" if value is not None else "-"

                check_status = "✓" if self.select_all_var.get() else " "
                item_id = self.files_tree.insert("", tk.END, text=file_name,
                                                 values=(check_status, size, modified, order_str))

                self.file_checkboxes[item_id] = {
                    'path': file_path,
                    'var': tk.BooleanVar(value=self.select_all_var.get())
                }

            # 更新顺序显示
            self.update_file_order()

            self.status_var.set(f"已选择 {len(self.selected_files)} 个文件")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
            self.status_var.set("读取文件时出错")

    def _extract_numeric_value(self, filename):
        try:
            match = re.search(r'(-?\d+\.?\d*)', filename)
            if match:
                return float(match.group(1))
            return None
        except:
            return None

    def _sort_files_by_numeric_value(self, files):
        def get_numeric_value(file_path):
            filename = os.path.basename(file_path)
            value = self._extract_numeric_value(filename)
            return value if value is not None else float('inf')

        return sorted(files, key=get_numeric_value)

    def format_size(self, size_bytes):
        units = ['B', 'KB', 'MB', 'GB', 'TB']
        unit_index = 0
        while size_bytes >= 1024 and unit_index < len(units) - 1:
            size_bytes /= 1024
            unit_index += 1
        return f"{size_bytes:.2f} {units[unit_index]}"

    def on_tree_click(self, event):
        region = self.files_tree.identify_region(event.x, event.y)
        item = self.files_tree.identify_row(event.y)

        if region == "cell" and self.files_tree.identify_column(event.x) == "#1":
            if item in self.file_checkboxes:
                current_state = self.file_checkboxes[item]['var'].get()
                new_state = not current_state
                self.file_checkboxes[item]['var'].set(new_state)

                check_status = "✓" if new_state else " "
                self.files_tree.set(item, column="check", value=check_status)

                all_selected = all(self.file_checkboxes[item]['var'].get()
                                   for item in self.file_checkboxes)
                self.select_all_var.set(all_selected)

                return "break"

    def toggle_select_all(self):
        new_state = self.select_all_var.get()

        for item_id, info in self.file_checkboxes.items():
            info['var'].set(new_state)
            check_status = "✓" if new_state else " "
            self.files_tree.set(item_id, column="check", value=check_status)

    def get_checked_files(self):
        # 按照文件列表中的顺序返回选中的文件
        checked_files = []
        for item in self.files_tree.get_children(''):
            if item in self.file_checkboxes and self.file_checkboxes[item]['var'].get():
                checked_files.append(self.file_checkboxes[item]['path'])
        return checked_files

    def stop_extraction_thread(self):
        self.stop_extraction.set()
        self.status_var.set("正在停止提取...")

    def extract_data(self):
        self.stop_extraction.clear()
        self.search_text1 = self.search_entry1.get().strip()
        self.search_text2 = self.search_entry2.get().strip()

        # 更新关键词历史记录
        if self.search_text1:
            if self.search_text1 not in self.keyword_history1:
                self.keyword_history1.insert(0, self.search_text1)
                self.keyword_history1 = self.keyword_history1[:20]  # 保留最多20条
                self.search_entry1["values"] = self.keyword_history1

        if self.search_text2:
            if self.search_text2 not in self.keyword_history2:
                self.keyword_history2.insert(0, self.search_text2)
                self.keyword_history2 = self.keyword_history2[:20]  # 保留最多20条
                self.search_entry2["values"] = self.keyword_history2

        # 保存历史记录
        self.save_keyword_history()

        if self.extract_mode.get() == 1 and not self.search_text1:
            messagebox.showwarning("警告", "请输入搜索文本1")
            return
        if self.extract_mode.get() == 2 and not self.search_text1:
            messagebox.showwarning("警告", "请输入搜索文本1")
            return
        if self.extract_mode.get() == 3 and (not self.search_text1 or not self.search_text2):
            messagebox.showwarning("警告", "请输入搜索文本1和搜索文本2")
            return
        if self.extract_mode.get() == 4 and not self.search_text1:
            messagebox.showwarning("警告", "请输入搜索文本1")
            return

        checked_files = self.get_checked_files()
        if not checked_files:
            messagebox.showwarning("警告", "请至少选择一个文件")
            return

        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.extracted_single.clear()
        self.extracted_double.clear()
        self.extracted_dual.clear()
        self.extracted_triple.clear()

        if self.extract_mode.get() == 1:
            self.results_tree["columns"] = ("value", "count")
            self.results_tree.column("#0", width=400, minwidth=250)
            self.results_tree.column("value", width=250, minwidth=150, anchor=tk.E)
            self.results_tree.column("count", width=100, minwidth=80, anchor=tk.CENTER)
            self.results_tree.heading("#0", text="文件名")
            self.results_tree.heading("value", text=f"{self.search_text1}值")
            self.results_tree.heading("count", text="值的个数")
        elif self.extract_mode.get() == 2:
            self.results_tree["columns"] = ("value1", "value2", "count")
            self.results_tree.column("#0", width=350, minwidth=200)
            self.results_tree.column("value1", width=200, minwidth=150, anchor=tk.E)
            self.results_tree.column("value2", width=200, minwidth=150, anchor=tk.E)
            self.results_tree.column("count", width=100, minwidth=80, anchor=tk.CENTER)
            self.results_tree.heading("#0", text="文件名")
            self.results_tree.heading("value1", text=f"{self.search_text1}值1")
            self.results_tree.heading("value2", text=f"{self.search_text1}值2")
            self.results_tree.heading("count", text="个数")
        elif self.extract_mode.get() == 3:
            self.results_tree["columns"] = ("value1", "value2", "count")
            self.results_tree.column("#0", width=350, minwidth=200)
            self.results_tree.column("value1", width=200, minwidth=150, anchor=tk.E)
            self.results_tree.column("value2", width=200, minwidth=150, anchor=tk.E)
            self.results_tree.column("count", width=100, minwidth=80, anchor=tk.CENTER)
            self.results_tree.heading("#0", text="文件名")
            self.results_tree.heading("value1", text=f"{self.search_text1}值")
            self.results_tree.heading("value2", text=f"{self.search_text2}值")
            self.results_tree.heading("count", text="个数")
        else:
            self.results_tree["columns"] = ("value1", "value2", "value3", "count")
            self.results_tree.column("#0", width=300, minwidth=150)
            self.results_tree.column("value1", width=150, minwidth=100, anchor=tk.E)
            self.results_tree.column("value2", width=150, minwidth=100, anchor=tk.E)
            self.results_tree.column("value3", width=150, minwidth=100, anchor=tk.E)
            self.results_tree.column("count", width=100, minwidth=80, anchor=tk.CENTER)
            self.results_tree.heading("#0", text="文件名")
            self.results_tree.heading("value1", text=f"{self.search_text1}值1")
            self.results_tree.heading("value2", text=f"{self.search_text1}值2")
            self.results_tree.heading("value3", text=f"{self.search_text1}值3")
            self.results_tree.heading("count", text="个数")

        self.status_var.set("正在提取数据...")
        threading.Thread(target=self._extract_data_thread, args=(checked_files,), daemon=True).start()

    def _extract_data_thread(self, files_to_process):
        total_files = len(files_to_process)
        for i, file_path in enumerate(files_to_process):
            if self.stop_extraction.is_set():
                self.status_queue.put(("status", "提取已停止"))
                return

            try:
                file_name = os.path.basename(file_path)
                content = self.read_file_content(file_path)

                if self.extract_mode.get() == 1:
                    values = self._extract_single_values(content, self.search_text1)
                    self.extracted_single[file_name] = values
                    count = len(values) if values else 0
                    display = ', '.join(map(lambda x: f"{x:.4f}", values)) if values else "未找到"
                    self.results_tree.insert("", tk.END, text=file_name, values=(display, count))
                elif self.extract_mode.get() == 2:
                    values = self._extract_double_values(content, self.search_text1)
                    self.extracted_double[file_name] = values
                    count = len(values) if values else 0
                    if values:
                        self.results_tree.insert("", tk.END, text=file_name,
                                                 values=(f"{values[0][0]:.4f}", f"{values[0][1]:.4f}", count))
                    else:
                        self.results_tree.insert("", tk.END, text=file_name, values=("未找到", "未找到", 0))
                elif self.extract_mode.get() == 3:
                    values = self._extract_dual_values(content)
                    self.extracted_dual[file_name] = values
                    count = len(values) if values else 0
                    if values:
                        self.results_tree.insert("", tk.END, text=file_name,
                                                 values=(f"{values[0][0]:.4f}", f"{values[0][1]:.4f}", count))
                    else:
                        self.results_tree.insert("", tk.END, text=file_name, values=("未找到", "未找到", 0))
                else:
                    values = self._extract_triple_values(content, self.search_text1)
                    self.extracted_triple[file_name] = values
                    count = len(values) if values else 0
                    if values:
                        self.results_tree.insert("", tk.END, text=file_name,
                                                 values=(
                                                     f"{values[0][0]:.4f}", f"{values[0][1]:.4f}",
                                                     f"{values[0][2]:.4f}",
                                                     count))
                    else:
                        self.results_tree.insert("", tk.END, text=file_name, values=("未找到", "未找到", "未找到", 0))

                progress = (i + 1) / total_files * 100
                self.status_queue.put(("progress", progress))
            except Exception as e:
                self.status_queue.put(("error", f"处理文件 {file_name} 时出错: {str(e)}"))

            time.sleep(0.01)

        self._generate_chart()
        self._update_status()

    def _update_status(self):
        mode = self.extract_mode.get()
        if mode == 1:
            valid_files = sum(1 for v in self.extracted_single.values() if v)
            self.status_queue.put(
                ("status", f"完成单文本单值提取，处理 {len(self.extracted_single)} 个文件，成功 {valid_files} 个"))
        elif mode == 2:
            valid_files = sum(1 for v in self.extracted_double.values() if v)
            self.status_queue.put(
                ("status", f"完成单文本双值提取，处理 {len(self.extracted_double)} 个文件，成功 {valid_files} 个"))
        elif mode == 3:
            valid_files = sum(1 for v in self.extracted_dual.values() if v)
            self.status_queue.put(
                ("status", f"完成双文本关联值提取，处理 {len(self.extracted_dual)} 个文件，成功 {valid_files} 个"))
        else:
            valid_files = sum(1 for v in self.extracted_triple.values() if v)
            self.status_queue.put(
                ("status", f"完成单文本三值提取，处理 {len(self.extracted_triple)} 个文件，成功 {valid_files} 个"))

    def read_file_content(self, file_path):
        file_ext = os.path.splitext(file_path)[1].lower().lstrip('.')

        supported_exts = self.get_selected_filetypes()
        if file_ext not in supported_exts:
            raise ValueError(f"不支持的文件类型: {file_ext}")

        if file_ext in ('dat', 'txt', 'log'):
            try:
                # 尝试多种编码格式
                encodings = ['utf-8', 'gbk', 'latin-1', 'cp1252', 'iso-8859-1']
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                            return f.read()
                    except UnicodeDecodeError:
                        continue
                # 如果所有编码都失败，使用二进制模式读取
                with open(file_path, 'rb') as f:
                    return f.read().decode('utf-8', errors='ignore')
            except Exception as e:
                raise ValueError(f"读取文件失败: {str(e)}")
        elif file_ext == 'xlsx':
            try:
                df = pd.read_excel(file_path, engine='openpyxl', dtype=str, nrows=1000)
                return '\n'.join(df.astype(str).values.flatten())
            except Exception as e:
                raise ValueError(f"读取Excel文件失败: {str(e)}")
        else:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
            except Exception as e:
                raise ValueError(f"读取{file_ext}文件失败: {str(e)}")

    # 改进的提取方法 - 处理关键词后无空格的情况
    def _extract_single_values(self, content, keyword):
        # 改进模式：关键词后跟0个或多个空白，然后是一个数字
        pattern = re.compile(
            f"{re.escape(keyword)}[\\s\\t]*(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)"
        )

        values = []
        for match in pattern.finditer(content):
            try:
                values.append(float(match.group(1)))
            except ValueError:
                continue
        return values

    def _extract_double_values(self, content, keyword):
        # 匹配关键词后跟两个数值
        pattern = re.compile(
            f"{re.escape(keyword)}[\\s\\t]*(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)[\\s\\t\\S]+?(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)"
        )

        values = []
        for match in pattern.finditer(content):
            try:
                values.append((float(match.group(1)), float(match.group(2))))
            except ValueError:
                continue
        return values

    def _extract_dual_values(self, content):
        # 匹配两个不同的关键词及其数值
        pattern = re.compile(
            fr'{re.escape(self.search_text1)}[\s\t]*(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?).*?'
            fr'{re.escape(self.search_text2)}[\s\t]*(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)',
            re.DOTALL
        )

        values = []
        for match in pattern.finditer(content):
            try:
                values.append((float(match.group(1)), float(match.group(2))))
            except ValueError:
                continue
        return values

    def _extract_triple_values(self, content, keyword):
        # 匹配关键词后跟三个数值
        pattern = re.compile(
            f"{re.escape(keyword)}[\\s\\t]*(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)[\\s\\t\\S]+?"
            f"(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)[\\s\\t\\S]+?(-?\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)"
        )

        values = []
        for match in pattern.finditer(content):
            try:
                values.append((float(match.group(1)), float(match.group(2)), float(match.group(3))))
            except ValueError:
                continue
        return values

    def _generate_chart(self):
        self.figure.clear()
        mode = self.extract_mode.get()

        if mode == 1:
            valid_data = {k: v for k, v in self.extracted_single.items() if v}
            if not valid_data:
                self.status_queue.put(("status", "没有可用于生成图表的单数值数据"))
                return

            all_values = [val for sublist in valid_data.values() for val in sublist]
            ax = self.figure.add_subplot(111)
            ax.hist(all_values, bins=20)
            ax.set_title(f"{self.search_text1} 后单数值分布")
            ax.set_xlabel("数值")
            ax.set_ylabel("频率")
        elif mode == 2:
            valid_data = {k: v for k, v in self.extracted_double.items() if v}
            if not valid_data:
                self.status_queue.put(("status", "没有可用于生成图表的双数值数据"))
                return

            values1 = [v[0] for sublist in valid_data.values() for v in sublist]
            values2 = [v[1] for sublist in valid_data.values() for v in sublist]

            ax1 = self.figure.add_subplot(211)
            ax1.hist(values1, bins=20, color='blue', alpha=0.7, label='第一个值')
            ax1.set_title(f"{self.search_text1} 后双数值分布")
            ax1.set_ylabel("频率")
            ax1.legend()

            ax2 = self.figure.add_subplot(212)
            ax2.hist(values2, bins=20, color='green', alpha=0.7, label='第二个值')
            ax2.set_xlabel("数值")
            ax2.set_ylabel("频率")
            ax2.legend()

            self.figure.tight_layout()
        elif mode == 3:
            valid_data = {k: v for k, v in self.extracted_dual.items() if v}
            if not valid_data:
                self.status_queue.put(("status", "没有可用于生成图表的双文本关联值数据"))
                return

            all_pairs = [pair for pairs in valid_data.values() for pair in pairs]
            values1 = [pair[0] for pair in all_pairs]
            values2 = [pair[1] for pair in all_pairs]

            ax = self.figure.add_subplot(111)
            ax.scatter(values1, values2, alpha=0.7)
            ax.set_title(f"{self.search_text1} 与 {self.search_text2} 关联关系")
            ax.set_xlabel(self.search_text1)
            ax.set_ylabel(self.search_text2)
            ax.grid(True, linestyle='--', alpha=0.7)

            if len(values1) > 1:
                corr = np.corrcoef(values1, values2)[0, 1]
                ax.text(0.05, 0.95, f"相关系数: {corr:.4f}", transform=ax.transAxes,
                        verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        else:
            valid_data = {k: v for k, v in self.extracted_triple.items() if v}
            if not valid_data:
                self.status_queue.put(("status", "没有可用于生成图表的三数值数据"))
                return

            all_triples = [triple for triples in valid_data.values() for triple in triples]
            values1 = [triple[0] for triple in all_triples]
            values2 = [triple[1] for triple in all_triples]
            values3 = [triple[2] for triple in all_triples]

            # ax1 = self.figure.add_subplot(311)
            # ax1.hist(values1, bins=20, color='blue', alpha=0.7, label='第一个值')
            # ax1.set_title(f"{self.search_text1} 后三数值分布")
            # ax1.set_ylabel("频率")
            # ax1.legend()

            # ax2 = self.figure.add_subplot(312)
            # ax2.hist(values2, bins=20, color='green', alpha=0.7, label='第二个值')
            # ax2.set_ylabel("频率")
            # ax2.legend()

            ax3 = self.figure.add_subplot(111)
            ax3.hist(values3, bins=20, color='red', alpha=0.7, label='值')
            ax3.set_xlabel("数值")
            ax3.set_ylabel("频率")
            ax3.legend()

            self.figure.tight_layout()

        self.figure.tight_layout()
        self.status_queue.put(("chart", "图表已更新"))

    def export_to_excel(self):
        mode = self.extract_mode.get()

        if mode == 1 and not self.extracted_single:
            messagebox.showwarning("警告", "没有提取的单数值数据可导出")
            return
        if mode == 2 and not self.extracted_double:
            messagebox.showwarning("警告", "没有提取的双数值数据可导出")
            return
        if mode == 3 and not self.extracted_dual:
            messagebox.showwarning("警告", "没有提取的双文本关联值数据可导出")
            return
        if mode == 4 and not self.extracted_triple:
            messagebox.showwarning("警告", "没有提取的三数值数据可导出")
            return

        checked_files = [os.path.basename(path) for path in self.get_checked_files()]
        if not checked_files:
            messagebox.showwarning("警告", "请至少选择一个文件")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            if mode == 1:
                filtered_data = {k: v for k, v in self.extracted_single.items() if k in checked_files}
                self._export_single_excel(writer, filtered_data)
            elif mode == 2:
                filtered_data = {k: v for k, v in self.extracted_double.items() if k in checked_files}
                self._export_double_excel(writer, filtered_data)
            elif mode == 3:
                filtered_data = {k: v for k, v in self.extracted_dual.items() if k in checked_files}
                self._export_dual_excel(writer, filtered_data)
            else:
                filtered_data = {k: v for k, v in self.extracted_triple.items() if k in checked_files}
                self._export_triple_excel(writer, filtered_data)

            if self.process_data.get():
                self._export_statistics(writer, checked_files)

        messagebox.showinfo("成功", f"数据已成功导出到 {file_path}")
        self.status_var.set(f"数据已导出到 {file_path}")

    def _export_single_excel(self, writer, data_dict):
        data = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            count = len(values) if values else 0

            if not values:
                data.append({
                    '角度': base_name,
                    '值的个数': count,
                    '序号': 1,
                    f'{self.search_text1}值': '未找到'
                })
                continue

            for idx, val in enumerate(values, 1):
                data.append({
                    '角度': base_name,
                    '个数': idx,
                    f'{self.search_text1}值': val
                })

        pd.DataFrame(data).to_excel(writer, sheet_name='数据预览', index=False)

    def _export_double_excel(self, writer, data_dict):
        data = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            count = len(values) if values else 0

            if not values:
                data.append({
                    '文件名': base_name,
                    '值的组数': count,
                    '序号': 1,
                    f'{self.search_text1}值1': '未找到',
                    f'{self.search_text1}值2': '未找到'
                })
                continue

            for idx, (val1, val2) in enumerate(values, 1):
                data.append({
                    '角度': base_name,
                    '个数': idx,
                    f'{self.search_text1}值1': val1,
                    f'{self.search_text1}值2': val2
                })

        pd.DataFrame(data).to_excel(writer, sheet_name='数值数据', index=False)

    def _export_dual_excel(self, writer, data_dict):
        data = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            count = len(values) if values else 0

            if not values:
                data.append({
                    '文件名': base_name,
                    '匹配行数': count,
                    '序号': 1,
                    f'{self.search_text1}值': '未找到',
                    f'{self.search_text2}值': '未找到'
                })
                continue

            for idx, (val1, val2) in enumerate(values, 1):
                data.append({
                    '角度': base_name,
                    '个数': count,
                    '序号': idx,
                    f'{self.search_text1}值': val1,
                    f'{self.search_text2}值': val2
                })

        pd.DataFrame(data).to_excel(writer, sheet_name='数据预览', index=False)

    def _export_triple_excel(self, writer, data_dict):
        data = []
        for file_name, values in data_dict.items():
            count = len(values) if values else 0

            if not values:
                data.append({
                    '文件名': file_name,
                    '值的组数': count,
                    '序号': 1,
                    f'{self.search_text1}值1': '未找到',
                    f'{self.search_text1}值2': '未找到',
                    f'{self.search_text1}值3': '未找到'
                })
                continue

            for idx, (val1, val2, val3) in enumerate(values, 1):
                data.append({
                    '角度': file_name,
                    '序号': idx,
                    # f'{self.search_text1}值1': val1,
                    # f'{self.search_text1}值2': val2,
                    f'{self.search_text1}值3': val3
                })

        pd.DataFrame(data).to_excel(writer, sheet_name='数值数据', index=False)

    def _export_statistics(self, writer, file_names):
        mode = self.extract_mode.get()
        if mode == 1:
            filtered_data = {k: v for k, v in self.extracted_single.items() if k in file_names}
            self._export_single_statistics(writer, filtered_data)
        elif mode == 2:
            filtered_data = {k: v for k, v in self.extracted_double.items() if k in file_names}
            self._export_double_statistics(writer, filtered_data)
        elif mode == 3:
            filtered_data = {k: v for k, v in self.extracted_dual.items() if k in file_names}
            self._export_dual_statistics(writer, filtered_data)
        else:
            filtered_data = {k: v for k, v in self.extracted_triple.items() if k in file_names}
            self._export_triple_statistics(writer, filtered_data)

    def _export_single_statistics(self, writer, data_dict):
        stats = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            if not values:
                stats.append({
                    '文件名': base_name,
                    '值的个数': 0,
                    '最小值': 'N/A',
                    '最大值': 'N/A',
                    '平均值': 'N/A',
                    '标准差': 'N/A'
                })
                continue

            stats.append({
                '角度': base_name,
                '个数': len(values),
                '最小值': min(values),
                '最大值': max(values),
                '差值': max(values) - min(values),
                '平均值': np.mean(values),
                '标准差': np.std(values) if len(values) > 1 else 0
            })

        pd.DataFrame(stats).to_excel(writer, sheet_name='数据处理', index=False)

    def _export_double_statistics(self, writer, data_dict):
        stats = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            if not values:
                stats.append({
                    '文件名': base_name,
                    '值的组数': 0,
                    '值1最小值': 'N/A', '值1最大值': 'N/A', '值1平均值': 'N/A', '值1标准差': 'N/A',
                    '值2最小值': 'N/A', '值2最大值': 'N/A', '值2平均值': 'N/A', '值2标准差': 'N/A'
                })
                continue

            vals1 = [v[0] for v in values]
            vals2 = [v[1] for v in values]

            stats.append({
                '角度': base_name,
                '个数': len(values),
                '值1最小值': min(vals1), '值1最大值': max(vals1), '值1平均值': np.mean(vals1),
                '值1标准差': np.std(vals1) if len(vals1) > 1 else 0,
                '值2最小值': min(vals2), '值2最大值': max(vals2), '值2平均值': np.mean(vals2),
                '值2标准差': np.std(vals2) if len(vals2) > 1 else 0
            })

        pd.DataFrame(stats).to_excel(writer, sheet_name='双数值统计', index=False)

    def _export_dual_statistics(self, writer, data_dict):
        stats = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            if not values:
                stats.append({
                    '文件名': base_name,
                    '个数': 0,
                    f'{self.search_text1}最小值': 'N/A', f'{self.search_text1}最大值': 'N/A',
                    f'{self.search_text1}平均值': 'N/A', f'{self.search_text1}标准差': 'N/A',
                    f'{self.search_text2}最小值': 'N/A', f'{self.search_text2}最大值': 'N/A',
                    f'{self.search_text2}平均值': 'N/A', f'{self.search_text2}标准差': 'N/A',
                })
                continue

            vals1 = [v[0] for v in values]
            vals2 = [v[1] for v in values]

            corr = np.corrcoef(vals1, vals2)[0, 1] if len(vals1) > 1 else 'N/A'

            stats.append({
                '角度': base_name,
                '个数': len(values),
                f'{self.search_text1}最小值': min(vals1), f'{self.search_text1}最大值': max(vals1),
                f'{self.search_text1}平均值': np.mean(vals1),
                f'{self.search_text1}标准差': np.std(vals1) if len(vals1) > 1 else 0,
                f'{self.search_text2}最小值': min(vals2), f'{self.search_text2}最大值': max(vals2),
                f'{self.search_text2}平均值': np.mean(vals2),
                f'{self.search_text2}标准差': np.std(vals2) if len(vals2) > 1 else 0,
            })

        pd.DataFrame(stats).to_excel(writer, sheet_name='数据统计', index=False)

    def _export_triple_statistics(self, writer, data_dict):
        stats = []
        for file_name, values in data_dict.items():
            base_name = os.path.splitext(file_name)[0]
            if not values:
                stats.append({
                    '文件名': base_name,
                    '个数': 0,
                    f'{self.search_text1}值1最小值': 'N/A', f'{self.search_text1}值1最大值': 'N/A',
                    f'{self.search_text1}值1平均值': 'N/A', f'{self.search_text1}值1标准差': 'N/A',
                    f'{self.search_text1}值2最小值': 'N/A', f'{self.search_text1}值2最大值': 'N/A',
                    f'{self.search_text1}值2平均值': 'N/A', f'{self.search_text1}值2标准差': 'N/A',
                    f'{self.search_text1}值3最小值': 'N/A', f'{self.search_text1}值3最大值': 'N/A',
                    f'{self.search_text1}值3平均值': 'N/A', f'{self.search_text1}值3标准差': 'N/A'
                })
                continue

            vals1 = [v[0] for v in values]
            vals2 = [v[1] for v in values]
            vals3 = [v[2] for v in values]

            stats.append({
                '角度': base_name,
                '个数': len(values),
                # f'值1最小值': min(vals1), f'值1最大值': max(vals1),
                # f'值1平均值': np.mean(vals1),
                # f'值1标准差': np.std(vals1) if len(vals1) > 1 else 0,
                # f'值2最小值': min(vals2), f'值2最大值': max(vals2),
                # f'值2平均值': np.mean(vals2),
                # f'值2标准差': np.std(vals2) if len(vals2) > 1 else 0,
                f'值3最小值': min(vals3), f'值3最大值': max(vals3),
                f'值3平均值': np.mean(vals3),
                f'值3标准差': np.std(vals3) if len(vals3) > 1 else 0
            })

        pd.DataFrame(stats).to_excel(writer, sheet_name='数值统计', index=False)


if __name__ == "__main__":
    root = tk.Tk()
    app = DatFileExtractor(root)
    root.mainloop()