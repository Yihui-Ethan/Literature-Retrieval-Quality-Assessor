import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import os
import glob
import re
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt


# ========== 标题标准化 ==========
def normalize_title(title):
    title = str(title).lower().strip()
    title = re.sub(r'[^\w\s]', '', title)
    title = re.sub(r'\s+', ' ', title)
    return title


# ========== 期刊名标准化 ==========
def normalize_journal(journal):
    journal = str(journal).lower().strip()
    journal = re.sub(r'[^\w\s]', '', journal)
    journal = re.sub(r'\s+', ' ', journal)
    # 常见缩写映射补充
    abbr_map = {
        "nat": "nature",
        "cell": "cell",
        "sci": "science",
        "nejm": "the new england journal of medicine",
        "lancet": "lancet",
        "jama": "jama",
        "nucleic acids res": "nucleic acids research",
        "genome res": "genome research",
        "genome biol": "genome biology",
        "nat genet": "nature genetics",
        "nat med": "nature medicine",
        "nat biotechnol": "nature biotechnology",
        "cell stem cell": "cell stem cell",
        "cell metab": "cell metabolism",
        "nat methods": "nature methods",
        "bmj": "bmj",
        "plos biol": "plos biology",
        "aging cell": "aging cell",
        "j gerontol a biol sci": "journals of gerontology series a-biological sciences and medical sciences",
        "aging res rev": "ageing research reviews",
        "mech ageing dev": "mechanisms of ageing and development",
        "exp gerontol": "experimental gerontology",
        "biogerontology": "biogerontology",
        "front aging": "frontiers in aging",
        "aging dis": "aging and disease",
        "j cell biol": "journal of cell biology",
        "mol cell": "molecular cell",
        "cell rep": "cell reports",
        "proc natl acad sci usa": "proceedings of the national academy of sciences of the united states of america",
        "pnas": "proceedings of the national academy of sciences of the united states of america",
        "bioinformatics": "bioinformatics",
        "bmc bioinformatics": "bmc bioinformatics",
        "brief bioinform": "briefings in bioinformatics",
    }
    if journal in abbr_map:
        return abbr_map[journal]
    return journal


# ========== 多格式读取 ==========
def read_file_safe(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    for enc in ["utf-8-sig", "utf-8", "gbk"]:
        try:
            if ext == ".csv":
                df = pd.read_csv(file_path, encoding=enc)
            elif ext in [".xlsx", ".xls"]:
                df = pd.read_excel(file_path)
            else:
                continue
            return df
        except Exception:
            continue
    return None


# ========== 重叠率计算 ==========
def calc_overlap(x, y):
    common = len(set(x) & set(y))
    avg_total = (len(x) + len(y)) / 2
    return round(common / avg_total * 100, 1)


# ========== 内置生物医学/生信核心期刊分区库（2025中科院） ==========
# 格式：标准化期刊名: (分区, T级)
BUILTIN_JOURNAL_RANK = {
    # ===== T0 顶刊 =====
    "nature": ("1区", "T0"),
    "science": ("1区", "T0"),
    "cell": ("1区", "T0"),
    "the new england journal of medicine": ("1区", "T0"),
    "lancet": ("1区", "T0"),
    "jama": ("1区", "T0"),

    # ===== T1 1区 =====
    "nature genetics": ("1区", "T1"),
    "nature medicine": ("1区", "T1"),
    "nature biotechnology": ("1区", "T1"),
    "nature methods": ("1区", "T1"),
    "cell stem cell": ("1区", "T1"),
    "cell metabolism": ("1区", "T1"),
    "molecular cell": ("1区", "T1"),
    "genome research": ("1区", "T1"),
    "genome biology": ("1区", "T1"),
    "nucleic acids research": ("1区", "T1"),
    "bmj": ("1区", "T1"),
    "plos biology": ("1区", "T1"),
    "proceedings of the national academy of sciences of the united states of america": ("1区", "T1"),
    "journal of cell biology": ("1区", "T1"),
    "cell reports": ("1区", "T1"),
    "briefings in bioinformatics": ("1区", "T1"),
    "ageing research reviews": ("1区", "T1"),
    "aging cell": ("1区", "T1"),
    "aging and disease": ("1区", "T1"),
    "frontiers in aging": ("1区", "T1"),

    # ===== T2 2区 =====
    "bioinformatics": ("2区", "T2"),
    "bmc bioinformatics": ("2区", "T2"),
    "mechanisms of ageing and development": ("2区", "T2"),
    "experimental gerontology": ("2区", "T2"),
    "biogerontology": ("2区", "T2"),
    "journals of gerontology series a-biological sciences and medical sciences": ("2区", "T2"),
    "bmc genomics": ("2区", "T2"),
    "plos one": ("3区", "T3"),
    "scientific reports": ("3区", "T3"),
    "peerj": ("3区", "T3"),
    "frontiers in genetics": ("2区", "T2"),
    "frontiers in bioengineering and biotechnology": ("2区", "T2"),
    "genes": ("3区", "T3"),
    "international journal of molecular sciences": ("2区", "T2"),
    "molecules": ("3区", "T3"),
    "sensors": ("3区", "T3"),
    "computational biology and chemistry": ("3区", "T3"),
    "journal of bioinformatics and computational biology": ("3区", "T3"),
    "evolutionary bioinformatics": ("4区", "T4"),
    "algorithms for molecular biology": ("4区", "T4"),
}

# T级权重分（用于排序）
T_WEIGHT = {"T0": 50, "T1": 20, "T2": 10, "T3": 4, "T4": 1, "未匹配": 0}


class OverlapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文献重叠分析 + 期刊T级自动识别")
        self.root.geometry("1060x860")
        self.root.resizable(True, True)
        self.center_window()

        self.file_list = []
        self.raw_sets = {}  # {文件名: [(原始标题, 原始期刊), ...]}
        self.norm_sets = {}  # {文件名: [标准化标题, ...]}
        self.title_journal_map = {}  # {标准化标题: 原始期刊名}
        self.result_mat = None
        self.all_titles_freq = None
        self.journal_rank = BUILTIN_JOURNAL_RANK.copy()
        self.rank_enabled = True

        self._build_ui()

    def center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build_ui(self):
        # 顶部工具栏
        top_bar = ttk.Frame(self.root, padding=(18, 16))
        top_bar.pack(fill=X)

        ttk.Button(top_bar, text="导入文献", bootstyle=SECONDARY,
                   command=self.select_files).pack(side=LEFT, padx=4)
        ttk.Button(top_bar, text="扫描文件夹", bootstyle=SECONDARY,
                   command=self.select_folder).pack(side=LEFT, padx=4)
        ttk.Button(top_bar, text="导入完整分区表", bootstyle=INFO,
                   command=self.import_rank_table).pack(side=LEFT, padx=4)
        ttk.Button(top_bar, text="清空", bootstyle=SECONDARY,
                   command=self.clear_list).pack(side=LEFT, padx=4)
        ttk.Button(top_bar, text="开始计算", bootstyle=PRIMARY,
                   command=self.calculate).pack(side=RIGHT, padx=4)

        # T级开关
        self.rank_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top_bar, text="启用期刊T级识别", variable=self.rank_var,
                        bootstyle="round-toggle").pack(side=RIGHT, padx=12)

        # 标签页主体
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=18, pady=(0, 12))

        # ===== 标签1：重叠率矩阵 =====
        tab_matrix = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_matrix, text="重叠率矩阵")

        ttk.Label(tab_matrix, text="已选文件", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        file_card = ttk.Frame(tab_matrix, bootstyle=LIGHT, padding=2)
        file_card.pack(fill=X, pady=(6, 14))
        self.listbox = tk.Listbox(file_card, height=4, bd=0, font=("Segoe UI", 10),
                                  fg="#495057", selectbackground="#e3f2fd")
        self.listbox.pack(fill=X, padx=1, pady=1)

        ttk.Label(tab_matrix, text="重叠率矩阵 (%)  双击单元格查看详情",
                  font=("Segoe UI", 10, "bold")).pack(anchor=W)
        table_card = ttk.Frame(tab_matrix, bootstyle=LIGHT, padding=2)
        table_card.pack(fill=BOTH, expand=True, pady=(6, 0))

        self.tree = ttk.Treeview(table_card, show=HEADINGS, bootstyle=INFO)
        v_scroll = ttk.Scrollbar(table_card, orient=VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_card, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side=RIGHT, fill=Y)
        h_scroll.pack(side=BOTTOM, fill=X)
        self.tree.pack(fill=BOTH, expand=True, padx=1, pady=1)
        self.tree.bind("<Double-1>", self.on_cell_double_click)

        # 底部状态栏
        bottom = ttk.Frame(self.root, padding=(18, 12))
        bottom.pack(fill=X, side=BOTTOM)
        self.status_label = ttk.Label(bottom, text="就绪", bootstyle=SECONDARY, font=("Segoe UI", 9))
        self.status_label.pack(side=LEFT)
        ttk.Button(bottom, text="导出矩阵", bootstyle=SECONDARY,
                   command=self.export_matrix).pack(side=RIGHT, padx=4)

        # ===== 标签2：统计与筛选 =====
        tab_stats = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_stats, text="统计&T级筛选")

        # 统计卡片
        stats_frame = ttk.Frame(tab_stats)
        stats_frame.pack(fill=X, pady=(0, 16))

        self.stat_cards = {}
        stats_labels = [
            ("文献集总数", "total_sets"),
            ("总文献数", "total_raw"),
            ("去重后文献数", "unique_papers"),
            ("总重复率", "dup_rate"),
            ("平均两两重叠率", "avg_overlap")
        ]
        for i, (label, key) in enumerate(stats_labels):
            card = ttk.Frame(stats_frame, bootstyle=INFO, padding=12)
            card.grid(row=0, column=i, padx=6, sticky=NSEW)
            ttk.Label(card, text=label, font=("Segoe UI", 9), bootstyle=SECONDARY).pack()
            val_label = ttk.Label(card, text="-", font=("Segoe UI", 14, "bold"))
            val_label.pack(pady=(4, 0))
            self.stat_cards[key] = val_label
        stats_frame.grid_columnconfigure("all", weight=1)

        # 筛选栏
        filter_frame = ttk.Labelframe(tab_stats, text="重复次数 & T级筛选", padding=12)
        filter_frame.pack(fill=X, pady=(0, 14))

        ttk.Label(filter_frame, text="最少出现次数:").grid(row=0, column=0, padx=4)
        self.recur_var = tk.IntVar(value=2)
        ttk.Spinbox(filter_frame, from_=1, to=10, width=5, textvariable=self.recur_var).grid(row=0, column=1, padx=8)
        ttk.Label(filter_frame, text="次").grid(row=0, column=2, padx=4)

        ttk.Label(filter_frame, text="  最低T级:").grid(row=0, column=3, padx=12)
        self.t_filter_var = tk.StringVar(value="全部")
        t_combo = ttk.Combobox(filter_frame, textvariable=self.t_filter_var, width=8, state="readonly")
        t_combo["values"] = ["全部", "T0", "T1及以上", "T2及以上", "T3及以上", "T4及以上"]
        t_combo.grid(row=0, column=4, padx=8)

        ttk.Button(filter_frame, text="筛选", bootstyle=PRIMARY,
                   command=self.filter_recurrence).grid(row=0, column=5, padx=16)
        ttk.Button(filter_frame, text="按重要性排序", bootstyle=WARNING,
                   command=self.sort_by_weight).grid(row=0, column=6, padx=4)
        ttk.Button(filter_frame, text="导出结果", bootstyle=SECONDARY,
                   command=self.export_filtered).grid(row=0, column=7, padx=4)

        # 筛选结果列表
        result_card = ttk.Frame(tab_stats, bootstyle=LIGHT, padding=2)
        result_card.pack(fill=BOTH, expand=True)
        self.filter_tree = ttk.Treeview(result_card, columns=("title", "journal", "tier", "zone", "count", "score"),
                                        show=HEADINGS, bootstyle=INFO)
        self.filter_tree.heading("title", text="文献标题")
        self.filter_tree.heading("journal", text="期刊")
        self.filter_tree.heading("tier", text="T级")
        self.filter_tree.heading("zone", text="分区")
        self.filter_tree.heading("count", text="出现次数")
        self.filter_tree.heading("score", text="重要性得分")

        self.filter_tree.column("title", anchor=W, width=420)
        self.filter_tree.column("journal", anchor=W, width=220)
        self.filter_tree.column("tier", width=60, anchor=CENTER)
        self.filter_tree.column("zone", width=60, anchor=CENTER)
        self.filter_tree.column("count", width=80, anchor=CENTER)
        self.filter_tree.column("score", width=90, anchor=CENTER)

        fv_scroll = ttk.Scrollbar(result_card, orient=VERTICAL, command=self.filter_tree.yview)
        self.filter_tree.configure(yscrollcommand=fv_scroll.set)
        fv_scroll.pack(side=RIGHT, fill=Y)
        self.filter_tree.pack(fill=BOTH, expand=True, padx=1, pady=1)
        self.filtered_df = None

        # ===== 标签3：热力图 =====
        tab_heatmap = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_heatmap, text="热力图")

        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=tab_heatmap)
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        ttk.Button(tab_heatmap, text="保存热力图", bootstyle=SECONDARY,
                   command=self.save_heatmap).pack(pady=10)

    # ========== 文件操作 ==========
    def select_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("数据文件", "*.csv *.xlsx *.xls"), ("CSV文件", "*.csv"), ("Excel文件", "*.xlsx *.xls")]
        )
        if paths:
            self.file_list = list(paths)
            self._refresh_listbox()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.file_list = glob.glob(os.path.join(folder, "*.*"))
            self.file_list = [f for f in self.file_list if f.lower().endswith((".csv", ".xlsx", ".xls"))]
            self._refresh_listbox()

    def import_rank_table(self):
        """导入完整中科院分区表，自动识别期刊名和分区列"""
        path = filedialog.askopenfilename(
            filetypes=[("数据文件", "*.csv *.xlsx"), ("CSV", "*.csv"), ("Excel", "*.xlsx")]
        )
        if not path:
            return
        df = read_file_safe(path)
        if df is None or df.shape[1] < 2:
            messagebox.showerror("错误", "分区表至少需要包含「期刊名」和「分区」两列")
            return

        # 自动识别列
        journal_col = df.columns[0]
        zone_col = None
        for col in df.columns:
            if re.search(r'journal|期刊|刊名|source', col, re.I):
                journal_col = col
            if re.search(r'分区|大区|大类分区|tier|zone', col, re.I):
                zone_col = col

        if zone_col is None:
            messagebox.showerror("错误", "未识别到分区列，请确保列名包含「分区」字样")
            return

        count = 0
        for _, row in df.iterrows():
            j_norm = normalize_journal(row[journal_col])
            zone = str(row[zone_col]).strip()
            # 映射T级
            if "1" in zone:
                tier = "T1"
            elif "2" in zone:
                tier = "T2"
            elif "3" in zone:
                tier = "T3"
            elif "4" in zone:
                tier = "T4"
            else:
                tier = "未匹配"
            self.journal_rank[j_norm] = (zone, tier)
            count += 1

        messagebox.showinfo("成功", f"已导入 {count} 条期刊分区数据\n内置库已合并更新")
        self.status_label.config(text=f"分区库已更新，共 {len(self.journal_rank)} 种期刊")

    def clear_list(self):
        self.file_list = []
        self.listbox.delete(0, tk.END)
        self.status_label.config(text="已清空")

    def _refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for f in self.file_list:
            self.listbox.insert(tk.END, "  " + os.path.basename(f))
        self.status_label.config(text=f"已选择 {len(self.file_list)} 个文件")

    # ========== 核心计算 ==========
    def calculate(self):
        if len(self.file_list) < 2:
            messagebox.showwarning("提示", "请至少选择2个文献文件")
            return

        self.raw_sets = {}
        self.norm_sets = {}
        self.title_journal_map = {}
        self.rank_enabled = self.rank_var.get()

        for f in self.file_list:
            name = os.path.splitext(os.path.basename(f))[0]
            df = read_file_safe(f)
            if df is None or df.shape[1] == 0:
                continue

            # 自动识别标题列和期刊列
            title_col = df.columns[0]
            journal_col = None
            for col in df.columns:
                if re.search(r'title|题目|标题|文献', col, re.I):
                    title_col = col
                if re.search(r'journal|source|期刊|出版物', col, re.I):
                    journal_col = col

            titles = df[title_col].dropna().astype(str).str.strip()
            titles = titles[titles != "标题"].unique().tolist()

            journals = []
            if journal_col is not None:
                journals = df[journal_col].dropna().astype(str).str.strip().tolist()
                if len(journals) != len(titles):
                    journals = ["未知期刊"] * len(titles)
            else:
                journals = ["未知期刊"] * len(titles)

            if len(titles) > 0:
                self.raw_sets[name] = list(zip(titles, journals))
                norm_titles = [normalize_title(t) for t in titles]
                self.norm_sets[name] = norm_titles

                # 建立 标准化标题 -> 原始期刊 的映射
                for t, j in zip(titles, journals):
                    nt = normalize_title(t)
                    if nt not in self.title_journal_map:
                        self.title_journal_map[nt] = j

        names = list(self.norm_sets.keys())
        n = len(names)
        if n < 2:
            messagebox.showerror("错误", "有效文献集少于2个")
            return

        # 计算重叠矩阵
        mat = pd.DataFrame(0.0, index=names, columns=names)
        for i in range(n):
            mat.iloc[i, i] = 100.0
            for j in range(i + 1, n):
                val = calc_overlap(self.norm_sets[names[i]], self.norm_sets[names[j]])
                mat.iloc[i, j] = val
                mat.iloc[j, i] = val

        self.result_mat = mat
        self._show_matrix(mat)
        self._update_stats()
        self._draw_heatmap(mat)
        self.status_label.config(text="计算完成")

    def _show_matrix(self, mat):
        self.tree.delete(*self.tree.get_children())
        cols = ["文献集"] + list(mat.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor=CENTER)

        for idx, (index, row) in enumerate(mat.iterrows()):
            tag = "odd" if idx % 2 == 0 else "even"
            self.tree.insert("", tk.END, values=[index] + row.tolist(), tags=(tag,))
        self.tree.tag_configure("odd", background="#f8f9fa")
        self.tree.tag_configure("even", background="white")

    # ========== 交集/差集详情窗口 ==========
    def on_cell_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        if col_idx == 0:
            return

        row_name = self.tree.item(item, "values")[0]
        col_name = self.tree["columns"][col_idx]
        if row_name == col_name:
            return

        self._show_detail_window(row_name, col_name)

    def _show_detail_window(self, name_a, name_b):
        set_a_raw = {t for t, j in self.raw_sets[name_a]}
        set_b_raw = {t for t, j in self.raw_sets[name_b]}
        set_a_norm = set(self.norm_sets[name_a])
        set_b_norm = set(self.norm_sets[name_b])

        common_norm = set_a_norm & set_b_norm
        common_raw = [t for t in set_a_raw if normalize_title(t) in common_norm]
        a_only = [t for t in set_a_raw if normalize_title(t) not in set_b_norm]
        b_only = [t for t in set_b_raw if normalize_title(t) not in set_a_norm]

        win = ttk.Toplevel(self.root)
        win.title(f"详情: {name_a} vs {name_b}")
        win.geometry("820x600")

        detail_notebook = ttk.Notebook(win)
        detail_notebook.pack(fill=BOTH, expand=True, padx=12, pady=12)

        tabs = [
            (f"共同文献 ({len(common_raw)})", common_raw, True),
            (f"仅在{name_a} ({len(a_only)})", a_only, False),
            (f"仅在{name_b} ({len(b_only)})", b_only, False),
        ]

        for tab_name, data, is_common in tabs:
            tab = ttk.Frame(detail_notebook, padding=8)
            detail_notebook.add(tab, text=tab_name)

            # 带期刊信息的列表
            tree = ttk.Treeview(tab, columns=("title", "journal", "tier"), show=HEADINGS, bootstyle=INFO)
            tree.heading("title", text="标题")
            tree.heading("journal", text="期刊")
            tree.heading("tier", text="T级")
            tree.column("title", anchor=W, width=420)
            tree.column("journal", anchor=W, width=220)
            tree.column("tier", width=60, anchor=CENTER)
            tree.pack(fill=BOTH, expand=True)

            for item in data:
                nt = normalize_title(item)
                j = self.title_journal_map.get(nt, "未知")
                j_norm = normalize_journal(j)
                zone, tier = self.journal_rank.get(j_norm, ("-", "未匹配"))
                tag = ""
                if is_common and self.rank_enabled:
                    if tier == "T0":
                        tag = "t0"
                    elif tier == "T1":
                        tag = "t1"
                    elif tier == "T2":
                        tag = "t2"
                tree.insert("", tk.END, values=(item, j, tier), tags=(tag,))

            tree.tag_configure("t0", background="#ffd7d7", foreground="#b00020", font=("Segoe UI", 9, "bold"))
            tree.tag_configure("t1", background="#ffe3e3", foreground="#c92a2a", font=("Segoe UI", 9, "bold"))
            tree.tag_configure("t2", background="#fff3bf", foreground="#e67700")

        btn_frame = ttk.Frame(win, padding=(12, 0, 12, 12))
        btn_frame.pack(fill=X)

        def export_detail():
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
                initialfile=f"detail_{name_a}_vs_{name_b}.xlsx"
            )
            if path:
                with pd.ExcelWriter(path) as writer:
                    # 共同文献带期刊T级
                    rows = []
                    for t in common_raw:
                        nt = normalize_title(t)
                        j = self.title_journal_map.get(nt, "未知")
                        j_norm = normalize_journal(j)
                        zone, tier = self.journal_rank.get(j_norm, ("-", "未匹配"))
                        rows.append({"标题": t, "期刊": j, "分区": zone, "T级": tier})
                    pd.DataFrame(rows).to_excel(writer, sheet_name="共同文献", index=False)

                    pd.DataFrame({"标题": a_only}).to_excel(writer, sheet_name=f"仅{name_a}", index=False)
                    pd.DataFrame({"标题": b_only}).to_excel(writer, sheet_name=f"仅{name_b}", index=False)
                messagebox.showinfo("成功", "详情已导出")

        ttk.Button(btn_frame, text="导出全部到Excel", bootstyle=PRIMARY, command=export_detail).pack(side=RIGHT)

    # ========== 统计面板 ==========
    def _update_stats(self):
        names = list(self.norm_sets.keys())
        n = len(names)
        total_raw = sum(len(v) for v in self.raw_sets.values())

        all_norm = []
        for v in self.norm_sets.values():
            all_norm.extend(v)
        self.all_titles_freq = pd.Series(all_norm).value_counts()
        unique_count = len(self.all_titles_freq)

        dup_rate = round((1 - unique_count / total_raw) * 100, 1)

        # 平均两两重叠率
        overlaps = []
        for i in range(n):
            for j in range(i + 1, n):
                overlaps.append(self.result_mat.iloc[i, j])
        avg_overlap = round(sum(overlaps) / len(overlaps), 1) if overlaps else 0

        self.stat_cards["total_sets"].config(text=str(n))
        self.stat_cards["total_raw"].config(text=str(total_raw))
        self.stat_cards["unique_papers"].config(text=str(unique_count))
        self.stat_cards["dup_rate"].config(text=f"{dup_rate}%")
        self.stat_cards["avg_overlap"].config(text=f"{avg_overlap}%")

    # ========== 重复+T级筛选 ==========
    def filter_recurrence(self):
        if self.all_titles_freq is None:
            messagebox.showwarning("提示", "请先点击开始计算")
            return

        min_count = self.recur_var.get()
        t_filter = self.t_filter_var.get()
        filtered = self.all_titles_freq[self.all_titles_freq >= min_count].reset_index()
        filtered.columns = ["norm_title", "count"]

        # 匹配回原始标题 + 期刊 + T级
        title_raw_map = {}
        for raw_list in self.raw_sets.values():
            for t, j in raw_list:
                norm = normalize_title(t)
                if norm not in title_raw_map:
                    title_raw_map[norm] = (t, j)

        filtered["title"] = filtered["norm_title"].map(lambda x: title_raw_map.get(x, (x, "未知"))[0])
        filtered["journal"] = filtered["norm_title"].map(lambda x: title_raw_map.get(x, ("", "未知"))[1])

        # 计算分区和T级
        if self.rank_enabled:
            def get_rank(row):
                j = row["journal"]
                j_norm = normalize_journal(j)
                zone, tier = self.journal_rank.get(j_norm, ("-", "未匹配"))
                return pd.Series([zone, tier])

            filtered[["zone", "tier"]] = filtered.apply(get_rank, axis=1)

            # T级筛选
            if t_filter == "T0":
                filtered = filtered[filtered["tier"] == "T0"]
            elif t_filter == "T1及以上":
                filtered = filtered[filtered["tier"].isin(["T0", "T1"])]
            elif t_filter == "T2及以上":
                filtered = filtered[filtered["tier"].isin(["T0", "T1", "T2"])]
            elif t_filter == "T3及以上":
                filtered = filtered[filtered["tier"].isin(["T0", "T1", "T2", "T3"])]
            elif t_filter == "T4及以上":
                filtered = filtered[filtered["tier"].isin(["T0", "T1", "T2", "T3", "T4"])]

            # 计算重要性得分 = T级权重 × 出现次数
            filtered["score"] = filtered["tier"].map(T_WEIGHT) * filtered["count"]
        else:
            filtered["zone"] = "-"
            filtered["tier"] = "-"
            filtered["score"] = filtered["count"]

        self.filtered_df = filtered[["title", "journal", "tier", "zone", "count", "score"]].sort_values("score",
                                                                                                        ascending=False).reset_index(
            drop=True)

        self.filter_tree.delete(*self.filter_tree.get_children())
        for _, row in self.filtered_df.iterrows():
            tag = ""
            if row["tier"] == "T0":
                tag = "t0"
            elif row["tier"] == "T1":
                tag = "t1"
            elif row["tier"] == "T2":
                tag = "t2"
            self.filter_tree.insert("", tk.END,
                                    values=(row["title"], row["journal"], row["tier"], row["zone"], row["count"],
                                            row["score"]), tags=(tag,))

        # T级配色
        self.filter_tree.tag_configure("t0", background="#ffd7d7", foreground="#b00020", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t1", background="#ffe3e3", foreground="#c92a2a", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t2", background="#fff3bf", foreground="#e67700")

        self.status_label.config(text=f"筛选结果: 共 {len(self.filtered_df)} 篇文献")

    def sort_by_weight(self):
        if self.filtered_df is None:
            messagebox.showwarning("提示", "请先执行筛选")
            return
        self.filtered_df = self.filtered_df.sort_values("score", ascending=False).reset_index(drop=True)
        self.filter_tree.delete(*self.filter_tree.get_children())
        for _, row in self.filtered_df.iterrows():
            tag = "t0" if row["tier"] == "T0" else (
                "t1" if row["tier"] == "T1" else ("t2" if row["tier"] == "T2" else ""))
            self.filter_tree.insert("", tk.END,
                                    values=(row["title"], row["journal"], row["tier"], row["zone"], row["count"],
                                            row["score"]), tags=(tag,))
        self.filter_tree.tag_configure("t0", background="#ffd7d7", foreground="#b00020", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t1", background="#ffe3e3", foreground="#c92a2a", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t2", background="#fff3bf", foreground="#e67700")

    def export_filtered(self):
        if self.filtered_df is None:
            messagebox.showwarning("提示", "请先执行筛选")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile="文献筛选结果_带T级.xlsx"
        )
        if path:
            if path.endswith(".csv"):
                self.filtered_df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                self.filtered_df.to_excel(path, index=False)
            messagebox.showinfo("成功", "结果已导出，包含标题、期刊、T级、分区、出现次数")

    # ========== 热力图 ==========
    def _draw_heatmap(self, mat):
        self.ax.clear()
        im = self.ax.imshow(mat.values, cmap="Blues", vmin=0, vmax=100)

        self.ax.set_xticks(range(len(mat.columns)))
        self.ax.set_yticks(range(len(mat.index)))
        self.ax.set_xticklabels(mat.columns, rotation=30, ha="right", fontsize=9)
        self.ax.set_yticklabels(mat.index, fontsize=9)

        for i in range(len(mat.index)):
            for j in range(len(mat.columns)):
                color = "white" if mat.iloc[i, j] > 50 else "black"
                self.ax.text(j, i, f"{mat.iloc[i, j]:.1f}",
                             ha="center", va="center", color=color, fontsize=9)

        self.fig.colorbar(im, ax=self.ax, label="重叠率 (%)")
        self.ax.set_title("文献重叠热力图", fontsize=12, pad=12)
        self.fig.tight_layout()
        self.canvas.draw()

    def save_heatmap(self):
        if self.result_mat is None:
            messagebox.showwarning("提示", "请先计算")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG图片", "*.png"), ("PDF", "*.pdf")],
            initialfile="文献重叠热力图.png"
        )
        if path:
            self.fig.savefig(path, dpi=300, bbox_inches="tight")
            messagebox.showinfo("成功", f"热力图已保存到:\n{path}")

    # ========== 导出矩阵 ==========
    def export_matrix(self):
        if self.result_mat is None:
            messagebox.showwarning("提示", "请先计算")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile="重叠率矩阵.xlsx"
        )
        if path:
            out = self.result_mat.copy()
            out.index.name = "文献集"
            if path.endswith(".csv"):
                out.to_csv(path, encoding="utf-8-sig")
            else:
                out.to_excel(path)
            messagebox.showinfo("成功", "矩阵已导出")


if __name__ == "__main__":
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "Segoe UI"]
    plt.rcParams["axes.unicode_minus"] = False

    root = ttk.Window(themename="cosmo")
    app = OverlapApp(root)
    root.mainloop()
