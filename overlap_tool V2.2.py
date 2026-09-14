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

# ========== Multilingual Text Pack ==========
LANG = {
    "en": {
        "app_title": "Literature Overlap Analyzer + Journal Tier Identification",
        "import_papers": "Import Papers",
        "scan_folder": "Scan Folder",
        "import_rank": "Import Rank Table",
        "clear": "Clear",
        "calculate": "Calculate",
        "enable_tier": "Enable Journal Tier",
        "language": "Language",
        "tab_matrix": "Overlap Matrix",
        "tab_stats": "Statistics & Filter",
        "tab_heatmap": "Heatmap",
        "selected_files": "Selected Files",
        "overlap_matrix_tip": "Overlap Matrix (%)  Double-click cell for details",
        "literature_set": "Literature Set",
        "ready": "Ready",
        "export_matrix": "Export Matrix",
        "total_sets": "Total Sets",
        "total_papers": "Total Papers",
        "unique_papers": "Unique Papers",
        "overall_dup": "Overall Duplication",
        "avg_overlap": "Avg Pairwise Overlap",
        "tier_dist_title": "Tier Distribution per Dataset",
        "set": "Literature Set",
        "total": "Total",
        "t0_cnt": "T0 Count",
        "t0_pct": "T0 Ratio (%)",
        "t1_cnt": "T1 Count",
        "t1_pct": "T1 Ratio (%)",
        "quality_title": "Retrieval Quality per Dataset",
        "unique_rate": "Core Capture (%)",
        "top_tier_rate": "Journal Quality",
        "avg_overlap_col": "Avg Overlap (%)",
        "quality_score": "Quality Score",
        "filter_title": "Occurrence & Tier Filter",
        "min_occur": "Min Occurrences:",
        "times": "times",
        "min_tier": "Min Tier:",
        "filter_btn": "Filter",
        "sort_score": "Sort by Score",
        "export_results": "Export Results",
        "paper_title": "Paper Title",
        "journal_col": "Journal",
        "pubmed_id": "PubMed ID",
        "tier_col": "Tier",
        "quartile_col": "Quartile",
        "occurrences": "Occurrences",
        "importance_score": "Importance Score",
        "save_heatmap": "Save Heatmap",
        "warning": "Warning",
        "error": "Error",
        "success": "Success",
        "select_2_files": "Please select at least 2 literature files",
        "fewer_2_sets": "Fewer than 2 valid literature sets",
        "calc_first": "Please run Calculate first",
        "filter_first": "Please run Filter first",
        "rank_table_error": "Rank table must contain at least Journal and Quartile columns",
        "quartile_not_found": "Quartile column not found",
        "imported_entries": "Imported {count} journal entries\nBuilt-in library updated",
        "rank_updated": "Rank library updated: {num} journals total",
        "cleared": "Cleared",
        "files_selected": "{num} file(s) selected",
        "calc_completed": "Calculation completed",
        "filtered_papers": "Filtered: {num} papers",
        "export_success": "Results exported with title, journal, PMID, tier, quartile, occurrences",
        "heatmap_saved": "Heatmap saved to:\n{path}",
        "matrix_exported": "Matrix exported",
        "detail_title": "Detail: {a} vs {b}",
        "common_papers": "Common Papers ({num})",
        "only_in": "Only in {name} ({num})",
        "title_col": "Title",
        "export_all_excel": "Export All to Excel",
        "detail_exported": "Details exported",
        "default_filter_file": "filtered_papers_with_tier.xlsx",
        "default_matrix_file": "overlap_matrix.xlsx",
        "default_heatmap_file": "literature_overlap_heatmap.png",
        "default_detail_file": "detail_{a}_vs_{b}.xlsx",
        "data_files": "Data Files",
        "csv_files": "CSV Files",
        "excel_files": "Excel Files",
        "png_image": "PNG Image",
        "all_option": "All",
        "t0_option": "T0",
        "t1_option": "T1+",
        "t2_option": "T2+",
        "t3_option": "T3+",
        "t4_option": "T4+",
        "small_sample_note": "* Small sample size, for reference only",
    },
    "zh": {
        "app_title": "文献重叠分析 + 期刊T级自动识别",
        "import_papers": "导入文献",
        "scan_folder": "扫描文件夹",
        "import_rank": "导入分区表",
        "clear": "清空",
        "calculate": "开始计算",
        "enable_tier": "启用期刊T级识别",
        "language": "语言",
        "tab_matrix": "重叠率矩阵",
        "tab_stats": "统计与筛选",
        "tab_heatmap": "热力图",
        "selected_files": "已选文件",
        "overlap_matrix_tip": "重叠率矩阵（%） 双击单元格查看详情",
        "literature_set": "文献集",
        "ready": "就绪",
        "export_matrix": "导出矩阵",
        "total_sets": "文献集总数",
        "total_papers": "总文献数",
        "unique_papers": "去重后文献数",
        "overall_dup": "总重复率",
        "avg_overlap": "平均两两重叠率",
        "tier_dist_title": "各文献集T级分布",
        "set": "文献集名称",
        "total": "总数",
        "t0_cnt": "T0数量",
        "t0_pct": "T0占比(%)",
        "t1_cnt": "T1数量",
        "t1_pct": "T1占比(%)",
        "quality_title": "各文献集检索质量",
        "unique_rate": "核心捕获率(%)",
        "top_tier_rate": "期刊质量分",
        "avg_overlap_col": "平均重叠率(%)",
        "quality_score": "质量得分",
        "filter_title": "出现次数 & T级筛选",
        "min_occur": "最少出现次数：",
        "times": "次",
        "min_tier": "最低T级：",
        "filter_btn": "筛选",
        "sort_score": "按重要性排序",
        "export_results": "导出结果",
        "paper_title": "文献标题",
        "journal_col": "期刊",
        "pubmed_id": "PubMed ID",
        "tier_col": "T级",
        "quartile_col": "分区",
        "occurrences": "出现次数",
        "importance_score": "重要性得分",
        "save_heatmap": "保存热力图",
        "warning": "提示",
        "error": "错误",
        "success": "成功",
        "select_2_files": "请至少选择2个文献文件",
        "fewer_2_sets": "有效文献集少于2个",
        "calc_first": "请先点击开始计算",
        "filter_first": "请先执行筛选",
        "rank_table_error": "分区表至少需要包含「期刊名」和「分区」两列",
        "quartile_not_found": "未识别到分区列",
        "imported_entries": "已导入 {count} 条期刊数据\n内置库已合并更新",
        "rank_updated": "分区库已更新，共 {num} 种期刊",
        "cleared": "已清空",
        "files_selected": "已选择 {num} 个文件",
        "calc_completed": "计算完成",
        "filtered_papers": "筛选结果：共 {num} 篇文献",
        "export_success": "结果已导出，包含标题、期刊、PMID、T级、分区、出现次数",
        "heatmap_saved": "热力图已保存到：\n{path}",
        "matrix_exported": "矩阵已导出",
        "detail_title": "详情：{a} 对比 {b}",
        "common_papers": "共同文献 ({num})",
        "only_in": "仅在{name} ({num})",
        "title_col": "标题",
        "export_all_excel": "导出全部到Excel",
        "detail_exported": "详情已导出",
        "default_filter_file": "文献筛选结果_带T级.xlsx",
        "default_matrix_file": "重叠率矩阵.xlsx",
        "default_heatmap_file": "文献重叠热力图.png",
        "default_detail_file": "详情_{a}_vs_{b}.xlsx",
        "data_files": "数据文件",
        "csv_files": "CSV文件",
        "excel_files": "Excel文件",
        "png_image": "PNG图片",
        "all_option": "全部",
        "t0_option": "T0",
        "t1_option": "T1及以上",
        "t2_option": "T2及以上",
        "t3_option": "T3及以上",
        "t4_option": "T4及以上",
        "small_sample_note": "* 样本量较小，仅供参考",
    }
}


# ========== Title Normalization ==========
def normalize_title(title):
    title = str(title).lower().strip()
    title = re.sub(r'[^\w\s]', '', title)
    title = re.sub(r'\s+', ' ', title)
    return title


# ========== Journal Name Normalization ==========
def normalize_journal(journal):
    journal = str(journal).lower().strip()
    journal = re.sub(r'[^\w\s]', '', journal)
    journal = re.sub(r'\s+', ' ', journal)
    abbr_map = {
        "nat": "nature", "cell": "cell", "sci": "science",
        "nejm": "the new england journal of medicine", "lancet": "lancet", "jama": "jama",
        "nucleic acids res": "nucleic acids research", "genome res": "genome research",
        "genome biol": "genome biology", "nat genet": "nature genetics",
        "nat med": "nature medicine", "nat biotechnol": "nature biotechnology",
        "cell stem cell": "cell stem cell", "cell metab": "cell metabolism",
        "nat methods": "nature methods", "bmj": "bmj", "plos biol": "plos biology",
        "aging cell": "aging cell",
        "j gerontol a biol sci": "journals of gerontology series a-biological sciences and medical sciences",
        "j gerontol": "journals of gerontology series a-biological sciences and medical sciences",
        "aging res rev": "ageing research reviews", "ageing res rev": "ageing research reviews",
        "mech ageing dev": "mechanisms of ageing and development", "exp gerontol": "experimental gerontology",
        "biogerontology": "biogerontology", "front aging": "frontiers in aging",
        "front aging neurosci": "frontiers in aging neuroscience", "aging dis": "aging and disease",
        "j cell biol": "journal of cell biology", "mol cell": "molecular cell",
        "cell rep": "cell reports",
        "proc natl acad sci usa": "proceedings of the national academy of sciences of the united states of america",
        "pnas": "proceedings of the national academy of sciences of the united states of america",
        "bioinformatics": "bioinformatics", "bmc bioinformatics": "bmc bioinformatics",
        "brief bioinform": "briefings in bioinformatics",
        "j bioinform comput biol": "journal of bioinformatics and computational biology",
        "comput biol chem": "computational biology and chemistry", "bmc genomics": "bmc genomics",
        "bmc geriatr": "bmc geriatrics", "elife": "elife", "geroscience": "geroscience",
        "scientific reports": "scientific reports", "plos one": "plos one",
        "ageing research reviews": "ageing research reviews", "biomolecules": "biomolecules",
        "drug discovery today": "drug discovery today", "molecular cell": "molecular cell",
    }
    if journal in abbr_map:
        return abbr_map[journal]
    return journal


# ========== Multi-format File Reader ==========
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


# ========== Overlap Rate Calculation ==========
def calc_overlap(x, y):
    common = len(set(x) & set(y))
    avg_total = (len(x) + len(y)) / 2
    return round(common / avg_total * 100, 1)


# ========== Built-in Journal Rank Database ==========
BUILTIN_JOURNAL_RANK = {
    "nature": ("Q1", "T0"), "science": ("Q1", "T0"), "cell": ("Q1", "T0"),
    "the new england journal of medicine": ("Q1", "T0"), "lancet": ("Q1", "T0"), "jama": ("Q1", "T0"),
    "elife": ("Q1", "T1"),
    "nature genetics": ("Q1", "T1"), "nature medicine": ("Q1", "T1"),
    "nature biotechnology": ("Q1", "T1"), "nature methods": ("Q1", "T1"),
    "cell stem cell": ("Q1", "T1"), "cell metabolism": ("Q1", "T1"),
    "molecular cell": ("Q1", "T1"), "genome research": ("Q1", "T1"),
    "genome biology": ("Q1", "T1"), "nucleic acids research": ("Q1", "T1"),
    "bmj": ("Q1", "T1"), "plos biology": ("Q1", "T1"),
    "proceedings of the national academy of sciences of the united states of america": ("Q1", "T1"),
    "journal of cell biology": ("Q1", "T1"), "cell reports": ("Q1", "T1"),
    "briefings in bioinformatics": ("Q1", "T1"), "ageing research reviews": ("Q1", "T1"),
    "aging cell": ("Q1", "T1"), "aging and disease": ("Q1", "T1"),
    "frontiers in aging": ("Q1", "T1"), "frontiers in aging neuroscience": ("Q1", "T1"),
    "geroscience": ("Q1", "T1"),
    "bioinformatics": ("Q2", "T2"), "bmc bioinformatics": ("Q2", "T2"),
    "mechanisms of ageing and development": ("Q2", "T2"), "experimental gerontology": ("Q2", "T2"),
    "biogerontology": ("Q2", "T2"),
    "journals of gerontology series a-biological sciences and medical sciences": ("Q2", "T2"),
    "bmc genomics": ("Q2", "T2"), "frontiers in genetics": ("Q2", "T2"),
    "frontiers in bioengineering and biotechnology": ("Q2", "T2"),
    "international journal of molecular sciences": ("Q2", "T2"),
    "bmc geriatrics": ("Q2", "T2"), "biomolecules": ("Q2", "T2"),
    "plos one": ("Q3", "T3"), "scientific reports": ("Q3", "T3"), "peerj": ("Q3", "T3"),
    "genes": ("Q3", "T3"), "molecules": ("Q3", "T3"), "sensors": ("Q3", "T3"),
    "computational biology and chemistry": ("Q3", "T3"),
    "journal of bioinformatics and computational biology": ("Q3", "T3"),
    "drug discovery today": ("Q3", "T3"),
    "evolutionary bioinformatics": ("Q4", "T4"), "algorithms for molecular biology": ("Q4", "T4"),
}

T_WEIGHT = {"T0": 50, "T1": 20, "T2": 10, "T3": 4, "T4": 1, "Unmatched": 0}
# 期刊质量分专用权重（百分制）
JOURNAL_QUALITY_WEIGHT = {"T0": 100, "T1": 70, "T2": 40, "T3": 20, "T4": 10, "Unmatched": 5}
# 贝叶斯收缩强度（数值越大，小样本平滑越强）
BAYESIAN_STRENGTH = 5


class OverlapApp:
    def __init__(self, root):
        self.root = root
        self.lang = "zh"
        self.root.geometry("1200x960")
        self.root.resizable(True, True)
        self.center_window()

        self.file_list = []
        self.raw_sets = {}
        self.norm_sets = {}
        self.title_journal_map = {}
        self.title_pmid_map = {}  # 新增：标准化标题 -> PubMed ID
        self.result_mat = None
        self.all_titles_freq = None
        self.journal_rank = BUILTIN_JOURNAL_RANK.copy()
        self.rank_enabled = True
        self.quality_df = None
        self.tier_dist_df = None
        self.filtered_df = None

        self._build_ui()
        self._update_ui_text()

    def tr(self, key):
        return LANG[self.lang].get(key, key)

    def center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build_ui(self):
        # Top toolbar
        self.top_bar = ttk.Frame(self.root, padding=(18, 16))
        self.top_bar.pack(fill=X)

        self.btn_import = ttk.Button(self.top_bar, bootstyle=SECONDARY, command=self.select_files)
        self.btn_import.pack(side=LEFT, padx=4)
        self.btn_scan = ttk.Button(self.top_bar, bootstyle=SECONDARY, command=self.select_folder)
        self.btn_scan.pack(side=LEFT, padx=4)
        self.btn_rank = ttk.Button(self.top_bar, bootstyle=INFO, command=self.import_rank_table)
        self.btn_rank.pack(side=LEFT, padx=4)
        self.btn_clear = ttk.Button(self.top_bar, bootstyle=SECONDARY, command=self.clear_list)
        self.btn_clear.pack(side=LEFT, padx=4)
        self.btn_calc = ttk.Button(self.top_bar, bootstyle=PRIMARY, command=self.calculate)
        self.btn_calc.pack(side=RIGHT, padx=4)

        # Language selector
        self.lang_var = tk.StringVar(value="中文")
        self.lang_combo = ttk.Combobox(self.top_bar, textvariable=self.lang_var, width=10, state="readonly")
        self.lang_combo["values"] = ["English", "中文"]
        self.lang_combo.pack(side=RIGHT, padx=12)
        self.lang_combo.bind("<<ComboboxSelected>>", self._switch_language)

        self.rank_var = tk.BooleanVar(value=True)
        self.chk_rank = ttk.Checkbutton(self.top_bar, variable=self.rank_var, bootstyle="round-toggle")
        self.chk_rank.pack(side=RIGHT, padx=12)

        # Tab panel
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=18, pady=(0, 12))

        # ===== Tab 1: Overlap Matrix =====
        self.tab_matrix = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_matrix, text="")

        self.lbl_selected = ttk.Label(self.tab_matrix, font=("Segoe UI", 10, "bold"))
        self.lbl_selected.pack(anchor=W)
        file_card = ttk.Frame(self.tab_matrix, bootstyle=LIGHT, padding=2)
        file_card.pack(fill=X, pady=(6, 14))
        self.listbox = tk.Listbox(file_card, height=4, bd=0, font=("Segoe UI", 10),
                                  fg="#495057", selectbackground="#e3f2fd")
        self.listbox.pack(fill=X, padx=1, pady=1)

        self.lbl_matrix_tip = ttk.Label(self.tab_matrix, font=("Segoe UI", 10, "bold"))
        self.lbl_matrix_tip.pack(anchor=W)
        table_card = ttk.Frame(self.tab_matrix, bootstyle=LIGHT, padding=2)
        table_card.pack(fill=BOTH, expand=True, pady=(6, 0))

        self.tree = ttk.Treeview(table_card, show="tree headings", bootstyle=INFO)
        v_scroll = ttk.Scrollbar(table_card, orient=VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_card, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side=RIGHT, fill=Y)
        h_scroll.pack(side=BOTTOM, fill=X)
        self.tree.pack(fill=BOTH, expand=True, padx=1, pady=1)
        self.tree.bind("<Double-1>", self.on_cell_double_click)

        # Bottom status bar
        bottom = ttk.Frame(self.root, padding=(18, 12))
        bottom.pack(fill=X, side=BOTTOM)
        self.status_label = ttk.Label(bottom, bootstyle=SECONDARY, font=("Segoe UI", 9))
        self.status_label.pack(side=LEFT)
        self.btn_export_matrix = ttk.Button(bottom, bootstyle=SECONDARY, command=self.export_matrix)
        self.btn_export_matrix.pack(side=RIGHT, padx=4)

        # ===== Tab 2: Statistics & Filter =====
        self.tab_stats = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_stats, text="")

        # Summary cards
        stats_frame = ttk.Frame(self.tab_stats)
        stats_frame.pack(fill=X, pady=(0, 16))

        self.stat_cards = {}
        stats_labels = [
            ("total_sets", "total_sets"),
            ("total_papers", "total_raw"),
            ("unique_papers", "unique_papers"),
            ("overall_dup", "dup_rate"),
            ("avg_overlap", "avg_overlap")
        ]
        self.stat_val_labels = {}
        for i, (label_key, key) in enumerate(stats_labels):
            card = ttk.Frame(stats_frame, bootstyle=INFO, padding=12)
            card.grid(row=0, column=i, padx=6, sticky=NSEW)
            lbl = ttk.Label(card, font=("Segoe UI", 9), bootstyle=SECONDARY)
            lbl.pack()
            val_label = ttk.Label(card, text="-", font=("Segoe UI", 14, "bold"))
            val_label.pack(pady=(4, 0))
            self.stat_cards[key] = val_label
            self.stat_val_labels[label_key] = lbl
        stats_frame.grid_columnconfigure("all", weight=1)

        # Tier Distribution
        self.tier_frame = ttk.Labelframe(self.tab_stats, padding=12)
        self.tier_frame.pack(fill=X, pady=(0, 14))

        self.tier_tree = ttk.Treeview(self.tier_frame, show="tree headings", bootstyle=INFO)
        tv_scroll = ttk.Scrollbar(self.tier_frame, orient=VERTICAL, command=self.tier_tree.yview)
        self.tier_tree.configure(yscrollcommand=tv_scroll.set)
        tv_scroll.pack(side=RIGHT, fill=Y)
        self.tier_tree.pack(fill=X, padx=1, pady=1)

        # Quality Panel
        self.quality_frame = ttk.Labelframe(self.tab_stats, padding=12)
        self.quality_frame.pack(fill=X, pady=(0, 14))

        self.quality_tree = ttk.Treeview(self.quality_frame, show="tree headings", bootstyle=INFO)
        qv_scroll = ttk.Scrollbar(self.quality_frame, orient=VERTICAL, command=self.quality_tree.yview)
        self.quality_tree.configure(yscrollcommand=qv_scroll.set)
        qv_scroll.pack(side=RIGHT, fill=Y)
        self.quality_tree.pack(fill=X, padx=1, pady=1)

        # Filter bar
        self.filter_frame = ttk.Labelframe(self.tab_stats, padding=12)
        self.filter_frame.pack(fill=X, pady=(0, 14))

        self.lbl_min_occur = ttk.Label(self.filter_frame)
        self.lbl_min_occur.grid(row=0, column=0, padx=4)
        self.recur_var = tk.IntVar(value=2)
        ttk.Spinbox(self.filter_frame, from_=1, to=10, width=5, textvariable=self.recur_var).grid(row=0, column=1,
                                                                                                  padx=8)
        self.lbl_times = ttk.Label(self.filter_frame)
        self.lbl_times.grid(row=0, column=2, padx=4)

        self.lbl_min_tier = ttk.Label(self.filter_frame)
        self.lbl_min_tier.grid(row=0, column=3, padx=12)
        self.t_filter_var = tk.StringVar(value="All")
        self.t_combo = ttk.Combobox(self.filter_frame, textvariable=self.t_filter_var, width=10, state="readonly")
        self.t_combo.grid(row=0, column=4, padx=8)

        self.btn_filter = ttk.Button(self.filter_frame, bootstyle=PRIMARY, command=self.filter_recurrence)
        self.btn_filter.grid(row=0, column=5, padx=16)
        self.btn_sort = ttk.Button(self.filter_frame, bootstyle=WARNING, command=self.sort_by_weight)
        self.btn_sort.grid(row=0, column=6, padx=4)
        self.btn_export_res = ttk.Button(self.filter_frame, bootstyle=SECONDARY, command=self.export_filtered)
        self.btn_export_res.grid(row=0, column=7, padx=4)

        # Filter result list
        result_card = ttk.Frame(self.tab_stats, bootstyle=LIGHT, padding=2)
        result_card.pack(fill=BOTH, expand=True)
        self.filter_tree = ttk.Treeview(result_card, show="tree headings", bootstyle=INFO)
        fv_scroll = ttk.Scrollbar(result_card, orient=VERTICAL, command=self.filter_tree.yview)
        self.filter_tree.configure(yscrollcommand=fv_scroll.set)
        fv_scroll.pack(side=RIGHT, fill=Y)
        self.filter_tree.pack(fill=BOTH, expand=True, padx=1, pady=1)

        # ===== Tab 3: Heatmap =====
        self.tab_heatmap = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_heatmap, text="")

        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_heatmap)
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        self.btn_save_heatmap = ttk.Button(self.tab_heatmap, bootstyle=SECONDARY, command=self.save_heatmap)
        self.btn_save_heatmap.pack(pady=10)

    def _switch_language(self, event=None):
        selected = self.lang_var.get()
        self.lang = "zh" if selected == "中文" else "en"
        self._update_ui_text()
        if self.result_mat is not None:
            self._setup_matrix_headers()
            self._setup_tier_headers()
            self._setup_quality_headers()
            self._setup_filter_headers()
            self._update_tier_distribution()
            self._refresh_quality_tree()
            self._update_stats_visuals()
            if self.filtered_df is not None:
                self._refresh_filter_tree()

    def _update_ui_text(self):
        self.root.title(self.tr("app_title"))

        self.btn_import.config(text=self.tr("import_papers"))
        self.btn_scan.config(text=self.tr("scan_folder"))
        self.btn_rank.config(text=self.tr("import_rank"))
        self.btn_clear.config(text=self.tr("clear"))
        self.btn_calc.config(text=self.tr("calculate"))
        self.chk_rank.config(text=self.tr("enable_tier"))

        self.notebook.tab(0, text=self.tr("tab_matrix"))
        self.notebook.tab(1, text=self.tr("tab_stats"))
        self.notebook.tab(2, text=self.tr("tab_heatmap"))

        self.lbl_selected.config(text=self.tr("selected_files"))
        self.lbl_matrix_tip.config(text=self.tr("overlap_matrix_tip"))
        self.btn_export_matrix.config(text=self.tr("export_matrix"))
        self.status_label.config(text=self.tr("ready"))

        for key, lbl in self.stat_val_labels.items():
            lbl.config(text=self.tr(key))

        self.tier_frame.configure(text=self.tr("tier_dist_title"))
        self.quality_frame.configure(text=self.tr("quality_title"))
        self.filter_frame.configure(text=self.tr("filter_title"))

        if self.result_mat is not None:
            self._setup_matrix_headers()
        self._setup_tier_headers()
        self._setup_quality_headers()
        self._setup_filter_headers()

        self.lbl_min_occur.config(text=self.tr("min_occur"))
        self.lbl_times.config(text=self.tr("times"))
        self.lbl_min_tier.config(text=self.tr("min_tier"))
        self.btn_filter.config(text=self.tr("filter_btn"))
        self.btn_sort.config(text=self.tr("sort_score"))
        self.btn_export_res.config(text=self.tr("export_results"))

        self.t_combo["values"] = [
            self.tr("all_option"), self.tr("t0_option"),
            self.tr("t1_option"), self.tr("t2_option"),
            self.tr("t3_option"), self.tr("t4_option"),
        ]

        self.btn_save_heatmap.config(text=self.tr("save_heatmap"))

    # ========== Table header setups ==========
    def _setup_matrix_headers(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []

        col_names = list(self.result_mat.columns)
        data_cols = [f"col_{i}" for i in range(len(col_names))]
        self.tree["columns"] = data_cols

        self.tree.heading("#0", text=self.tr("literature_set"))
        self.tree.column("#0", width=220, anchor=W, stretch=False)

        for i, name in enumerate(col_names):
            self.tree.heading(data_cols[i], text=name)
            self.tree.column(data_cols[i], width=130, anchor=CENTER, stretch=False)

    def _setup_tier_headers(self):
        self.tier_tree.delete(*self.tier_tree.get_children())
        self.tier_tree["columns"] = []

        data_cols = ["total", "t0_cnt", "t0_pct", "t1_cnt", "t1_pct"]
        self.tier_tree["columns"] = data_cols

        self.tier_tree.heading("#0", text=self.tr("set"))
        self.tier_tree.column("#0", width=260, anchor=W)

        for key in data_cols:
            self.tier_tree.heading(key, text=self.tr(key))
            self.tier_tree.column(key, width=110, anchor=CENTER)

    def _setup_quality_headers(self):
        self.quality_tree.delete(*self.quality_tree.get_children())
        self.quality_tree["columns"] = []

        data_cols = ["total", "unique_rate", "top_tier_rate", "avg_overlap_col", "quality_score"]
        self.quality_tree["columns"] = data_cols

        self.quality_tree.heading("#0", text=self.tr("set"))
        self.quality_tree.column("#0", width=260, anchor=W)

        for key in data_cols:
            self.quality_tree.heading(key, text=self.tr(key))
            self.quality_tree.column(key, width=120, anchor=CENTER)

    def _setup_filter_headers(self):
        self.filter_tree.delete(*self.filter_tree.get_children())
        self.filter_tree["columns"] = []

        # 新增pubmed列
        data_cols = ["journal", "pubmed", "tier", "quartile", "count", "score"]
        self.filter_tree["columns"] = data_cols

        self.filter_tree.heading("#0", text=self.tr("paper_title"))
        self.filter_tree.column("#0", width=420, anchor=W)

        headers = ["journal_col", "pubmed_id", "tier_col", "quartile_col", "occurrences", "importance_score"]
        widths = [220, 100, 60, 70, 90, 110]
        for key, hdr, w in zip(data_cols, headers, widths):
            self.filter_tree.heading(key, text=self.tr(hdr))
            self.filter_tree.column(key, width=w, anchor=CENTER)

    def _update_stats_visuals(self):
        self.tier_tree.tag_configure("t0_high", background="#ffd7d7", foreground="#b00020",
                                     font=("Segoe UI", 9, "bold"))
        self.tier_tree.tag_configure("t0_mid", background="#ffe3e3", foreground="#c92a2a")
        self.tier_tree.tag_configure("small_sample", foreground="#868e96")
        self.quality_tree.tag_configure("q_high", background="#e3f7e7", foreground="#0b7a38",
                                        font=("Segoe UI", 9, "bold"))
        self.quality_tree.tag_configure("q_mid", background="#fff3bf", foreground="#e67700")

    def _refresh_quality_tree(self):
        """独立的质量表渲染函数：把self.quality_df的数据插入到表格中"""
        if self.quality_df is None:
            return
        self.quality_tree.delete(*self.quality_tree.get_children())
        for _, row in self.quality_df.iterrows():
            tag = ""
            if row["quality_score"] >= 70:
                tag = "q_high"
            elif row["quality_score"] >= 50:
                tag = "q_mid"
            self.quality_tree.insert("", tk.END, text=row["set"], values=(
                row["total"], row["unique_rate"], row["top_tier_rate"],
                row["avg_overlap_col"], row["quality_score"]
            ), tags=(tag,))

    def _refresh_filter_tree(self):
        self.filter_tree.delete(*self.filter_tree.get_children())
        for _, row in self.filtered_df.iterrows():
            tag = ""
            if row["tier"] == "T0":
                tag = "t0"
            elif row["tier"] == "T1":
                tag = "t1"
            elif row["tier"] == "T2":
                tag = "t2"
            # 新增pmid显示
            self.filter_tree.insert("", tk.END, text=row["title"], values=(
                row["journal"], row["pubmed"], row["tier"], row["quartile"], row["count"], row["score"]
            ), tags=(tag,))
        self.filter_tree.tag_configure("t0", background="#ffd7d7", foreground="#b00020", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t1", background="#ffe3e3", foreground="#c92a2a", font=("Segoe UI", 9, "bold"))
        self.filter_tree.tag_configure("t2", background="#fff3bf", foreground="#e67700")

    # ========== File Operations ==========
    def select_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[(self.tr("data_files"), "*.csv *.xlsx *.xls"),
                       (self.tr("csv_files"), "*.csv"),
                       (self.tr("excel_files"), "*.xlsx *.xls")]
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
        path = filedialog.askopenfilename(
            filetypes=[(self.tr("data_files"), "*.csv *.xlsx"),
                       (self.tr("csv_files"), "*.csv"),
                       (self.tr("excel_files"), "*.xlsx")]
        )
        if not path:
            return
        df = read_file_safe(path)
        if df is None or df.shape[1] < 2:
            messagebox.showerror(self.tr("error"), self.tr("rank_table_error"))
            return

        journal_col = df.columns[0]
        zone_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if re.search(r'journal|source|publication', col_lower):
                journal_col = col
            if re.search(r'quartile|zone|tier|分区', col_lower):
                zone_col = col

        if zone_col is None:
            messagebox.showerror(self.tr("error"), self.tr("quartile_not_found"))
            return

        count = 0
        for _, row in df.iterrows():
            j_norm = normalize_journal(row[journal_col])
            zone = str(row[zone_col]).strip()
            if "1" in zone:
                tier = "T1"
            elif "2" in zone:
                tier = "T2"
            elif "3" in zone:
                tier = "T3"
            elif "4" in zone:
                tier = "T4"
            else:
                tier = "Unmatched"
            self.journal_rank[j_norm] = (zone, tier)
            count += 1

        messagebox.showinfo(self.tr("success"), self.tr("imported_entries").format(count=count))
        self.status_label.config(text=self.tr("rank_updated").format(num=len(self.journal_rank)))

    def clear_list(self):
        self.file_list = []
        self.listbox.delete(0, tk.END)
        self.status_label.config(text=self.tr("cleared"))

    def _refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for f in self.file_list:
            self.listbox.insert(tk.END, "  " + os.path.basename(f))
        self.status_label.config(text=self.tr("files_selected").format(num=len(self.file_list)))

    # ========== Core Calculation ==========
    def calculate(self):
        if len(self.file_list) < 2:
            messagebox.showwarning(self.tr("warning"), self.tr("select_2_files"))
            return

        self.raw_sets = {}
        self.norm_sets = {}
        self.title_journal_map = {}
        self.title_pmid_map = {}  # 重置PMID映射
        self.rank_enabled = self.rank_var.get()

        for f in self.file_list:
            name = os.path.splitext(os.path.basename(f))[0]
            df = read_file_safe(f)
            if df is None or df.shape[1] == 0:
                continue

            title_col = df.columns[0]
            journal_col = None
            pmid_col = None

            for col in df.columns:
                col_str = str(col).strip()
                if col_str in ["标题", "title", "Title"]:
                    title_col = col
                if col_str in ["期刊", "journal", "Journal", "source", "Source"]:
                    journal_col = col
                # 精确匹配PMID列名
                if col_str.lower() in ["pmid", "pubmed id", "pubmed_id", "pubmedid", "pmid号"]:
                    pmid_col = col

            if journal_col is None:
                for col in df.columns:
                    col_lower = str(col).lower()
                    if re.search(r'journal|source|publication|期刊|出版物|来源', col_lower):
                        journal_col = col
                        break

            if pmid_col is None:
                for col in df.columns:
                    col_lower = str(col).lower()
                    if re.search(r'pmid|pubmed.*id|pubmed', col_lower):
                        pmid_col = col
                        break

            if journal_col is None and df.shape[1] >= 3:
                journal_col = df.columns[2]

            title_journal_pairs = []
            seen_titles = set()
            for _, row in df.iterrows():
                t = str(row[title_col]).strip()
                j = str(row[journal_col]).strip() if journal_col is not None else "Unknown"
                pmid = str(row[pmid_col]).strip() if pmid_col is not None else "-"

                if not t or t in ["标题", "Title", "title"]:
                    continue
                if t in seen_titles:
                    continue
                seen_titles.add(t)
                title_journal_pairs.append((t, j, pmid))

            if len(title_journal_pairs) == 0:
                continue
            self.raw_sets[name] = title_journal_pairs
            norm_titles = [normalize_title(t) for t, j, pmid in title_journal_pairs]
            self.norm_sets[name] = norm_titles
            for t, j, pmid in title_journal_pairs:
                nt = normalize_title(t)
                if nt not in self.title_journal_map:
                    self.title_journal_map[nt] = j
                if nt not in self.title_pmid_map:
                    self.title_pmid_map[nt] = pmid

        names = sorted(list(self.norm_sets.keys()))
        n = len(names)
        if n < 2:
            messagebox.showerror(self.tr("error"), self.tr("fewer_2_sets"))
            return

        mat = pd.DataFrame(0.0, index=names, columns=names)
        for i in range(n):
            mat.iloc[i, i] = 100.0
            for j in range(i + 1, n):
                val = calc_overlap(self.norm_sets[names[i]], self.norm_sets[names[j]])
                mat.iloc[i, j] = val
                mat.iloc[j, i] = val

        self.result_mat = mat
        # 矩阵：先建表头，再填数据
        self._setup_matrix_headers()
        self._populate_matrix()

        # 所有表头先初始化，再计算填充数据
        self._setup_tier_headers()
        self._setup_quality_headers()
        self._setup_filter_headers()

        # 再计算并填充所有数据
        self._update_stats()
        self._update_tier_distribution()

        self._draw_heatmap(mat)
        self.status_label.config(text=self.tr("calc_completed"))

    def _populate_matrix(self):
        self.tree.delete(*self.tree.get_children())
        for idx, (index, row) in enumerate(self.result_mat.iterrows()):
            tag = "odd" if idx % 2 == 0 else "even"
            self.tree.insert("", tk.END, text=index, values=row.tolist(), tags=(tag,))

        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("even", background="#f8f9fa")

    # ========== Detail Window ==========
    def on_cell_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        if col == "#0":
            return
        col_idx = int(col.replace("#", "")) - 1
        col_names = list(self.result_mat.columns)
        if col_idx < 0 or col_idx >= len(col_names):
            return
        col_name = col_names[col_idx]

        row_name = self.tree.item(item, "text")
        if row_name == col_name:
            return
        self._show_detail_window(row_name, col_name)

    def _show_detail_window(self, name_a, name_b):
        set_a_raw = {t for t, j, pmid in self.raw_sets[name_a]}
        set_b_raw = {t for t, j, pmid in self.raw_sets[name_b]}
        set_a_norm = set(self.norm_sets[name_a])
        set_b_norm = set(self.norm_sets[name_b])

        common_norm = set_a_norm & set_b_norm
        common_raw = [t for t in set_a_raw if normalize_title(t) in common_norm]
        a_only = [t for t in set_a_raw if normalize_title(t) not in set_b_norm]
        b_only = [t for t in set_b_raw if normalize_title(t) not in set_a_norm]

        win = ttk.Toplevel(self.root)
        win.title(self.tr("detail_title").format(a=name_a, b=name_b))
        win.geometry("900x600")

        detail_notebook = ttk.Notebook(win)
        detail_notebook.pack(fill=BOTH, expand=True, padx=12, pady=12)

        tabs = [
            (self.tr("common_papers").format(num=len(common_raw)), common_raw, True),
            (self.tr("only_in").format(name=name_a, num=len(a_only)), a_only, False),
            (self.tr("only_in").format(name=name_b, num=len(b_only)), b_only, False),
        ]

        for tab_name, data, is_common in tabs:
            tab = ttk.Frame(detail_notebook, padding=8)
            detail_notebook.add(tab, text=tab_name)

            tree = ttk.Treeview(tab, show="tree headings", bootstyle=INFO)
            tree["columns"] = ("journal", "pubmed", "tier")
            tree.heading("#0", text=self.tr("title_col"))
            tree.heading("journal", text=self.tr("journal_col"))
            tree.heading("pubmed", text=self.tr("pubmed_id"))
            tree.heading("tier", text=self.tr("tier_col"))
            tree.column("#0", anchor=W, width=380)
            tree.column("journal", anchor=W, width=220)
            tree.column("pubmed", anchor=CENTER, width=100)
            tree.column("tier", width=60, anchor=CENTER)
            tree.pack(fill=BOTH, expand=True)

            for item in data:
                nt = normalize_title(item)
                j = self.title_journal_map.get(nt, "Unknown")
                pmid = self.title_pmid_map.get(nt, "-")
                j_norm = normalize_journal(j)
                quartile, tier = self.journal_rank.get(j_norm, ("-", "Unmatched"))
                tag = ""
                if is_common and self.rank_enabled:
                    if tier == "T0":
                        tag = "t0"
                    elif tier == "T1":
                        tag = "t1"
                    elif tier == "T2":
                        tag = "t2"
                tree.insert("", tk.END, text=item, values=(j, pmid, tier), tags=(tag,))

            tree.tag_configure("t0", background="#ffd7d7", foreground="#b00020", font=("Segoe UI", 9, "bold"))
            tree.tag_configure("t1", background="#ffe3e3", foreground="#c92a2a", font=("Segoe UI", 9, "bold"))
            tree.tag_configure("t2", background="#fff3bf", foreground="#e67700")

        btn_frame = ttk.Frame(win, padding=(12, 0, 12, 12))
        btn_frame.pack(fill=X)

        def export_detail():
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[(self.tr("excel_files"), "*.xlsx"), (self.tr("csv_files"), "*.csv")],
                initialfile=self.tr("default_detail_file").format(a=name_a, b=name_b)
            )
            if path:
                with pd.ExcelWriter(path) as writer:
                    rows = []
                    for t in common_raw:
                        nt = normalize_title(t)
                        j = self.title_journal_map.get(nt, "Unknown")
                        pmid = self.title_pmid_map.get(nt, "-")
                        j_norm = normalize_journal(j)
                        quartile, tier = self.journal_rank.get(j_norm, ("-", "Unmatched"))
                        rows.append({"Title": t, "Journal": j, "PubMed ID": pmid, "Quartile": quartile, "Tier": tier})
                    pd.DataFrame(rows).to_excel(writer, sheet_name="Common", index=False)
                    pd.DataFrame({"Title": a_only}).to_excel(writer, sheet_name=f"Only_{name_a}", index=False)
                    pd.DataFrame({"Title": b_only}).to_excel(writer, sheet_name=f"Only_{name_b}", index=False)
                messagebox.showinfo(self.tr("success"), self.tr("detail_exported"))

        ttk.Button(btn_frame, text=self.tr("export_all_excel"), bootstyle=PRIMARY, command=export_detail).pack(
            side=RIGHT)

    # ========== Statistics (检索质量计算逻辑) ==========
    def _update_stats(self):
        names = sorted(list(self.norm_sets.keys()))
        n = len(names)
        total_raw = sum(len(v) for v in self.raw_sets.values())

        all_norm = []
        for v in self.norm_sets.values():
            all_norm.extend(v)
        self.all_titles_freq = pd.Series(all_norm).value_counts()
        unique_count = len(self.all_titles_freq)
        dup_rate = round((1 - unique_count / total_raw) * 100, 1)

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

        # ========== 检索质量四维评分模型 ==========
        # 1. 定义全域核心文献（出现次数>=2）
        core_papers = set(self.all_titles_freq[self.all_titles_freq >= 2].index)
        total_core = len(core_papers) if core_papers else 1

        # 2. 计算全域期刊质量先验均值（贝叶斯收缩基准）
        all_tiers = []
        for nt in all_norm:
            j = self.title_journal_map.get(nt, "")
            j_norm = normalize_journal(j)
            _, tier = self.journal_rank.get(j_norm, ("", "Unmatched"))
            all_tiers.append(tier)
        global_avg_journal_score = sum(JOURNAL_QUALITY_WEIGHT.get(t, 5) for t in all_tiers) / len(all_tiers)

        # 3. 计算所有集合的独有高质量文献最大值（归一化基准）
        max_unique_high_quality = 0
        set_unique_high = {}
        for name in names:
            set_norm = set(self.norm_sets[name])
            other_norm = set()
            for other in names:
                if other != name:
                    other_norm.update(self.norm_sets[other])
            unique_in_set = set_norm - other_norm
            # 统计独有且T2及以上的文献
            high_quality_count = 0
            for nt in unique_in_set:
                j = self.title_journal_map.get(nt, "")
                j_norm = normalize_journal(j)
                _, tier = self.journal_rank.get(j_norm, ("", "Unmatched"))
                if tier in ["T0", "T1", "T2"]:
                    high_quality_count += 1
            set_unique_high[name] = high_quality_count
            if high_quality_count > max_unique_high_quality:
                max_unique_high_quality = high_quality_count

        quality_rows = []
        for name in names:
            set_norm = set(self.norm_sets[name])
            total = len(set_norm)

            # 维度1：核心捕获分 - 衡量查全能力
            core_captured = len(set_norm & core_papers)
            core_capture_score = round(core_captured / total_core * 100, 1)

            # 维度2：期刊质量分 - 贝叶斯平滑解决小样本失真
            if self.rank_enabled:
                raw_scores = []
                for nt in set_norm:
                    j = self.title_journal_map.get(nt, "")
                    j_norm = normalize_journal(j)
                    _, tier = self.journal_rank.get(j_norm, ("", "Unmatched"))
                    raw_scores.append(JOURNAL_QUALITY_WEIGHT.get(tier, 5))
                raw_avg = sum(raw_scores) / len(raw_scores)
                # 贝叶斯收缩：样本量越小，越向全局均值收缩
                shrunk_score = (total * raw_avg + BAYESIAN_STRENGTH * global_avg_journal_score) / (
                            total + BAYESIAN_STRENGTH)
                journal_score = round(shrunk_score, 1)
            else:
                journal_score = 0

            # 维度3：独有贡献分 - 衡量检索的增量价值
            if max_unique_high_quality > 0:
                unique_contrib_score = round(set_unique_high[name] / max_unique_high_quality * 100, 1)
            else:
                unique_contrib_score = 0

            # 维度4：过度冗余惩罚 - 仅重叠率>70%时扣分
            overlap_vals = []
            for other_name in names:
                if other_name != name:
                    overlap_vals.append(self.result_mat.loc[name, other_name])
            avg_overlap_set = sum(overlap_vals) / len(overlap_vals) if overlap_vals else 0
            redundancy_penalty = max(0, (avg_overlap_set - 70) / 30 * 100)

            # 综合质量得分（加权百分制）
            quality_score = 0.35 * core_capture_score + 0.35 * journal_score + 0.2 * unique_contrib_score - 0.1 * redundancy_penalty
            quality_score = round(max(0, min(100, quality_score)), 1)

            quality_rows.append({
                "set": name,
                "total": total,
                "unique_rate": core_capture_score,
                "top_tier_rate": journal_score,
                "avg_overlap_col": round(avg_overlap_set, 1),
                "quality_score": quality_score
            })

        self.quality_df = pd.DataFrame(quality_rows).sort_values("quality_score", ascending=False).reset_index(
            drop=True)
        # 调用独立渲染函数插入数据
        self._refresh_quality_tree()
        self._update_stats_visuals()

    # ========== Tier分布（贝叶斯平滑+小样本标记） ==========
    def _update_tier_distribution(self):
        names = sorted(list(self.norm_sets.keys()))

        # 计算全域T级占比先验
        global_t0 = global_t1 = 0
        global_total = 0
        for name in names:
            for t, j, pmid in self.raw_sets[name]:
                j_norm = normalize_journal(j)
                _, tier = self.journal_rank.get(j_norm, ("", "Unmatched"))
                if tier == "T0":
                    global_t0 += 1
                elif tier == "T1":
                    global_t1 += 1
                global_total += 1
        global_t0_pct = global_t0 / global_total if global_total > 0 else 0
        global_t1_pct = global_t1 / global_total if global_total > 0 else 0

        tier_rows = []
        for name in names:
            total = len(self.raw_sets[name])
            t0_cnt = t1_cnt = 0
            if self.rank_enabled:
                for t, j, pmid in self.raw_sets[name]:
                    j_norm = normalize_journal(j)
                    _, tier = self.journal_rank.get(j_norm, ("", "Unmatched"))
                    if tier == "T0":
                        t0_cnt += 1
                    elif tier == "T1":
                        t1_cnt += 1

            # 贝叶斯平滑占比，解决小样本波动大
            t0_pct_smooth = (t0_cnt + BAYESIAN_STRENGTH * global_t0_pct) / (total + BAYESIAN_STRENGTH) * 100
            t1_pct_smooth = (t1_cnt + BAYESIAN_STRENGTH * global_t1_pct) / (total + BAYESIAN_STRENGTH) * 100
            t0_pct_smooth = round(t0_pct_smooth, 1)
            t1_pct_smooth = round(t1_pct_smooth, 1)

            # 小样本标记
            small_sample = total < 20

            tier_rows.append({
                "set": name,
                "total": total,
                "t0_cnt": t0_cnt,
                "t0_pct": t0_pct_smooth,
                "t1_cnt": t1_cnt,
                "t1_pct": t1_pct_smooth,
                "small_sample": small_sample
            })

        self.tier_dist_df = pd.DataFrame(tier_rows).sort_values("t0_pct", ascending=False).reset_index(drop=True)
        self.tier_tree.delete(*self.tier_tree.get_children())
        for _, row in self.tier_dist_df.iterrows():
            tag = ""
            if row["t0_pct"] >= 10:
                tag = "t0_high"
            elif row["t0_pct"] >= 5:
                tag = "t0_mid"
            if row["small_sample"]:
                tag += " small_sample" if tag else "small_sample"

            # 小样本加星号提示
            t0_display = f"{row['t0_pct']}*" if row["small_sample"] else f"{row['t0_pct']}"
            t1_display = f"{row['t1_pct']}*" if row["small_sample"] else f"{row['t1_pct']}"

            self.tier_tree.insert("", tk.END, text=row["set"], values=(
                row["total"], row["t0_cnt"], t0_display,
                row["t1_cnt"], t1_display
            ), tags=(tag,))
        self._update_stats_visuals()

    # ========== Filter ==========
    def filter_recurrence(self):
        if self.all_titles_freq is None:
            messagebox.showwarning(self.tr("warning"), self.tr("calc_first"))
            return

        min_count = self.recur_var.get()
        t_filter = self.t_filter_var.get()
        filtered = self.all_titles_freq[self.all_titles_freq >= min_count].reset_index()
        filtered.columns = ["norm_title", "count"]

        title_raw_map = {}
        for raw_list in self.raw_sets.values():
            for t, j, pmid in raw_list:
                norm = normalize_title(t)
                if norm not in title_raw_map:
                    title_raw_map[norm] = (t, j, pmid)

        filtered["title"] = filtered["norm_title"].map(lambda x: title_raw_map.get(x, (x, "Unknown", "-"))[0])
        filtered["journal"] = filtered["norm_title"].map(lambda x: title_raw_map.get(x, ("", "Unknown", "-"))[1])
        filtered["pubmed"] = filtered["norm_title"].map(lambda x: title_raw_map.get(x, ("", "", "-"))[2])

        if self.rank_enabled:
            def get_rank(row):
                j = row["journal"]
                j_norm = normalize_journal(j)
                quartile, tier = self.journal_rank.get(j_norm, ("-", "Unmatched"))
                return pd.Series([quartile, tier], index=["quartile", "tier"])

            filtered[["quartile", "tier"]] = filtered.apply(get_rank, axis=1)

            tier_map = {
                self.tr("t0_option"): ["T0"],
                self.tr("t1_option"): ["T0", "T1"],
                self.tr("t2_option"): ["T0", "T1", "T2"],
                self.tr("t3_option"): ["T0", "T1", "T2", "T3"],
                self.tr("t4_option"): ["T0", "T1", "T2", "T3", "T4"],
            }
            if t_filter in tier_map:
                filtered = filtered[filtered["tier"].isin(tier_map[t_filter])]

            filtered["score"] = filtered["tier"].map(T_WEIGHT) * filtered["count"]
        else:
            filtered["quartile"] = "-"
            filtered["tier"] = "-"
            filtered["score"] = filtered["count"]

        self.filtered_df = filtered[["title", "journal", "pubmed", "tier", "quartile", "count", "score"]].sort_values(
            "score",
            ascending=False).reset_index(
            drop=True)
        self._refresh_filter_tree()
        self.status_label.config(text=self.tr("filtered_papers").format(num=len(self.filtered_df)))

    def sort_by_weight(self):
        if self.filtered_df is None:
            messagebox.showwarning(self.tr("warning"), self.tr("filter_first"))
            return
        self.filtered_df = self.filtered_df.sort_values("score", ascending=False).reset_index(drop=True)
        self._refresh_filter_tree()

    def export_filtered(self):
        if self.filtered_df is None:
            messagebox.showwarning(self.tr("warning"), self.tr("filter_first"))
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[(self.tr("excel_files"), "*.xlsx"), (self.tr("csv_files"), "*.csv")],
            initialfile=self.tr("default_filter_file")
        )
        if path:
            if path.endswith(".csv"):
                self.filtered_df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                self.filtered_df.to_excel(path, index=False)
            messagebox.showinfo(self.tr("success"), self.tr("export_success"))

    # ========== Heatmap ==========
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
        self.fig.colorbar(im, ax=self.ax, label=self.tr("avg_overlap"))
        self.ax.set_title(self.tr("tab_matrix"), fontsize=12, pad=12)
        self.fig.tight_layout()
        self.canvas.draw()

    def save_heatmap(self):
        if self.result_mat is None:
            messagebox.showwarning(self.tr("warning"), self.tr("calc_first"))
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[(self.tr("png_image"), "*.png"), ("PDF", "*.pdf")],
            initialfile=self.tr("default_heatmap_file")
        )
        if path:
            self.fig.savefig(path, dpi=300, bbox_inches="tight")
            messagebox.showinfo(self.tr("success"), self.tr("heatmap_saved").format(path=path))

    # ========== Export Matrix ==========
    def export_matrix(self):
        if self.result_mat is None:
            messagebox.showwarning(self.tr("warning"), self.tr("calc_first"))
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[(self.tr("excel_files"), "*.xlsx"), (self.tr("csv_files"), "*.csv")],
            initialfile=self.tr("default_matrix_file")
        )
        if path:
            out = self.result_mat.copy()
            out.index.name = self.tr("literature_set")
            if path.endswith(".csv"):
                out.to_csv(path, encoding="utf-8-sig")
            else:
                out.to_excel(path)
            messagebox.showinfo(self.tr("success"), self.tr("matrix_exported"))


if __name__ == "__main__":
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "Segoe UI"]
    plt.rcParams["axes.unicode_minus"] = False

    root = ttk.Window(themename="cosmo")
    app = OverlapApp(root)
    root.mainloop()
