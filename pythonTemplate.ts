export const PYTHON_CODE = `# -*- coding: utf-8 -*-
"""
Jieba 分词统计工具 (轻量级离线版)
"""

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import threading
import sys
import os
import csv
import tempfile
from collections import Counter
import re
import datetime

# --- 依赖检查 ---
def check_dependencies():
    missing = []
    try:
        import jieba
        import jieba.analyse
        import jieba.posseg as pseg
    except ImportError:
        missing.append("jieba")
        
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")
        
    if missing:
        root = tk.Tk()
        root.withdraw()
        msg = "检测到缺少必要库：\\n" + "\\n".join(missing) + "\\n\\n请运行: pip install " + " ".join(missing)
        messagebox.showerror("缺少依赖", msg)
        sys.exit(1)

check_dependencies()
import jieba
import jieba.analyse
import jieba.posseg as pseg
import openpyxl

# 默认停用词
DEFAULT_STOPWORDS = "的,了,是,我,你,在,和,就,都,而,及,与,这,那,有,个,之,为,但,我们,公司,有限公司,\\n,，,。,！,？,、,“, ”,：,；"

# 词性映射表 (用于定义哪些tag属于哪类)
POS_MAP = {
    'n': ['n', 'nr', 'ns', 'nt', 'nz', 'f', 's', 'ng'],  # 名词
    'v': ['v', 'vd', 'vn', 'vf', 'vx', 'vi', 'vl', 'vg'], # 动词
    'a': ['a', 'an', 'ad', 'ag', 'al'],                   # 形容词
    'd': ['d', 'dg'],                                     # 副词
    'p': ['p', 'pba', 'pbei'],                            # 介词
    'u': ['u', 'ul', 'uj', 'uz', 'ug', 'uv', 'ud']        # 助词
}

class WordFreqApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 分词统计工具 (轻量版)")
        self.root.geometry("1000x900")
        
        # Win10 DPI适配
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        # 样式
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=25, font=('Microsoft YaHei UI', 9))
        style.configure("Treeview.Heading", font=('Microsoft YaHei UI', 9, 'bold'))
        
        # 变量
        self.filepath = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.sheet_headers = []
        self.workbook_path = ""
        self.wb = None
        self.current_preview_mode = None  # 'stats', 'seg', 'tfidf', 'row_tfidf'
        self.last_result_data = None      # 暂存用于导出的数据
        
        self.init_ui()

    def init_ui(self):
        main = tk.Frame(self.root, padx=15, pady=15, bg='#f0f2f5')
        main.pack(fill='both', expand=True)

        # 顶部：文件与列选择
        top_frame = tk.LabelFrame(main, text=" 第一步：数据源设置 ", font=('Microsoft YaHei UI', 10, 'bold'), bg='white', padx=10, pady=10)
        top_frame.pack(fill='x', pady=(0, 10))
        
        # 文件
        f_row = tk.Frame(top_frame, bg='white')
        f_row.pack(fill='x', pady=2)
        tk.Button(f_row, text="📂 选择 Excel/CSV", command=self.load_file, bg='#e3f2fd', relief='groove').pack(side='left')
        tk.Label(f_row, textvariable=self.filepath, bg='#f5f5f5', relief='sunken', anchor='w').pack(side='left', padx=10, fill='x', expand=True)
        
        # Sheet
        s_row = tk.Frame(top_frame, bg='white')
        s_row.pack(fill='x', pady=5)
        tk.Label(s_row, text="选择Sheet：", bg='white').pack(side='left')
        self.sheet_combo = ttk.Combobox(s_row, textvariable=self.selected_sheet, state="readonly", width=30)
        self.sheet_combo.pack(side='left', padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # 列
        c_row = tk.Frame(top_frame, bg='white')
        c_row.pack(fill='x', pady=5)
        tk.Label(c_row, text="分析目标列：", bg='white').pack(side='left')
        self.col_combo = ttk.Combobox(c_row, textvariable=self.selected_column, state="readonly", width=30)
        self.col_combo.pack(side='left', padx=5)

        # 中部：停用词与词性设置 (合并)
        mid_frame = tk.LabelFrame(main, text=" 第二步：停用词与过滤设置 ", font=('Microsoft YaHei UI', 10, 'bold'), bg='white', padx=10, pady=10)
        mid_frame.pack(fill='x', pady=(0, 10))
        
        # 1. 词性停用 (排除)
        pos_frame = tk.Frame(mid_frame, bg='white')
        pos_frame.pack(fill='x', pady=(0, 5))
        tk.Label(pos_frame, text="【词性与类型过滤】(勾选即排除，不参与统计)：", bg='white', font=('Microsoft YaHei UI', 9, 'bold'), fg='#555').pack(anchor='w')
        
        check_row = tk.Frame(pos_frame, bg='white')
        check_row.pack(fill='x', pady=2)
        
        self.var_noun = tk.BooleanVar(value=False)
        self.var_verb = tk.BooleanVar(value=False)
        self.var_adj = tk.BooleanVar(value=False)
        self.var_adv = tk.BooleanVar(value=False)
        self.var_prep = tk.BooleanVar(value=False)
        self.var_part = tk.BooleanVar(value=False)
        self.var_filter_en = tk.BooleanVar(value=False)
        self.var_filter_num = tk.BooleanVar(value=False)
        
        # Checkboxes
        tk.Checkbutton(check_row, text="名词(n*)", variable=self.var_noun, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="动词(v*)", variable=self.var_verb, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="形容词(a*)", variable=self.var_adj, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="副词(d*)", variable=self.var_adv, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="介词(p*)", variable=self.var_prep, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="助词(u*)", variable=self.var_part, bg='white').pack(side='left', padx=2)
        tk.Checkbutton(check_row, text="过滤英文", variable=self.var_filter_en, bg='white').pack(side='left', padx=10)
        tk.Checkbutton(check_row, text="过滤数字", variable=self.var_filter_num, bg='white').pack(side='left', padx=2)

        # 2. 具体词停用
        txt_frame = tk.Frame(mid_frame, bg='white')
        txt_frame.pack(fill='x', pady=(5, 0))
        tk.Label(txt_frame, text="【具体词停用】(每行一个，或逗号分隔)：", bg='white', font=('Microsoft YaHei UI', 9, 'bold'), fg='#555').pack(anchor='w')
        
        self.txt_stopwords = scrolledtext.ScrolledText(txt_frame, height=3, font=('Microsoft YaHei UI', 9))
        self.txt_stopwords.pack(fill='x', pady=2)
        self.txt_stopwords.insert('1.0', DEFAULT_STOPWORDS)

        # 3. 固定搭配
        tk.Label(txt_frame, text="【固定搭配】(不拆分，每行一个或逗号分隔)：", bg='white', font=('Microsoft YaHei UI', 9, 'bold'), fg='#555').pack(anchor='w', pady=(5,0))
        self.txt_fixed = scrolledtext.ScrolledText(txt_frame, height=2, font=('Microsoft YaHei UI', 9))
        self.txt_fixed.pack(fill='x', pady=2)

        # 底部：结果与操作
        btm_frame = tk.Frame(main, bg='#f0f2f5')
        btm_frame.pack(fill='both', expand=True)

        # 左侧：结果预览
        left_res = tk.LabelFrame(btm_frame, text=" 结果预览 (Top 30) ", font=('Microsoft YaHei UI', 10, 'bold'), bg='white', padx=5, pady=5)
        left_res.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        # 预览操作栏
        prev_bar = tk.Frame(left_res, bg='white')
        prev_bar.pack(fill='x', pady=(0,5))
        self.lbl_status = tk.Label(prev_bar, text="请在右侧选择分析功能", bg='white', fg='#666')
        self.lbl_status.pack(side='left')
        self.btn_save_file = tk.Button(prev_bar, text="💾 导出当前结果", command=self.save_current_result, state='disabled', bg='#2196f3', fg='white', font=('Microsoft YaHei UI', 9))
        self.btn_save_file.pack(side='right')

        # 表格
        self.tree = ttk.Treeview(left_res, show='headings')
        vsb = ttk.Scrollbar(left_res, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left_res, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        # 右侧：操作区
        right_ops = tk.Frame(btm_frame, bg='#f0f2f5', width=220)
        right_ops.pack(side='right', fill='y', padx=(5, 0))
        
        tk.Label(right_ops, text="分析功能区", bg='#f0f2f5', font=('Microsoft YaHei UI', 9, 'bold')).pack(anchor='w', pady=(0, 5))
        
        def make_btn(txt, cmd, color, icon):
            return tk.Button(right_ops, text=icon+" "+txt, command=cmd, bg=color, fg='white', font=('Microsoft YaHei UI', 10, 'bold'), pady=8)
        
        self.btn_stats = make_btn("全局词频统计", lambda: self.run_analysis('stats'), '#4caf50', '🚀')
        self.btn_stats.pack(fill='x', pady=5)

        self.btn_seg = make_btn("逐行分词预览", lambda: self.run_analysis('seg'), '#ff9800', '📝')
        self.btn_seg.pack(fill='x', pady=5)
        
        self.btn_row_tfidf = make_btn("逐行关键词提取", lambda: self.run_analysis('row_tfidf'), '#009688', '🔑')
        self.btn_row_tfidf.pack(fill='x', pady=5)

        self.btn_tfidf = make_btn("全局 TF-IDF 关键词", lambda: self.run_analysis('tfidf'), '#9c27b0', '📊')
        self.btn_tfidf.pack(fill='x', pady=5)

        self.btn_entropy = make_btn("逐行信息熵与长度", lambda: self.run_analysis('entropy'), '#e91e63', '📈')
        self.btn_entropy.pack(fill='x', pady=5)

        self.btn_ner_person = make_btn("提取人名 (nr)", lambda: self.run_analysis('ner_person'), '#3f51b5', '👤')
        self.btn_ner_person.pack(fill='x', pady=5)

        self.btn_ner_location = make_btn("提取地名 (ns)", lambda: self.run_analysis('ner_location'), '#00bcd4', '📍')
        self.btn_ner_location.pack(fill='x', pady=5)

        # 异常长度设置
        len_frame = tk.Frame(right_ops, bg='#f0f2f5')
        len_frame.pack(fill='x', pady=5)
        tk.Label(len_frame, text="异常长度阈值:", bg='#f0f2f5', font=('Microsoft YaHei UI', 9)).pack(side='left')
        self.var_min_len = tk.IntVar(value=2)
        self.var_max_len = tk.IntVar(value=50)
        tk.Entry(len_frame, textvariable=self.var_min_len, width=3).pack(side='left', padx=2)
        tk.Label(len_frame, text="-", bg='#f0f2f5').pack(side='left')
        tk.Entry(len_frame, textvariable=self.var_max_len, width=3).pack(side='left', padx=2)

    def load_file(self):
        fn = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")])
        if not fn: return
        self.filepath.set(fn)
        self.workbook_path = fn
        
        try:
            if fn.endswith('.csv'):
                self.sheet_combo['values'] = ['CSV']
                self.sheet_combo.current(0)
                self.on_sheet_change(None)
            else:
                if self.wb: self.wb.close()
                self.wb = openpyxl.load_workbook(fn, read_only=True)
                self.sheet_combo['values'] = self.wb.sheetnames
                if self.wb.sheetnames:
                    self.sheet_combo.current(0)
                    self.on_sheet_change(None)
        except Exception as e:
            messagebox.showerror("读取错误", f"无法读取文件: {e}")

    def on_sheet_change(self, event):
        sheet_name = self.selected_sheet.get()
        if not sheet_name: return
        headers = []
        try:
            if self.filepath.get().endswith('.csv'):
                with open(self.filepath.get(), 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    headers = next(reader)
            else:
                ws = self.wb[sheet_name]
                for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    headers = list(row)
                    break
            self.sheet_headers = [str(h) if h else f"Col_{i}" for i, h in enumerate(headers)]
            self.col_combo['values'] = self.sheet_headers
            if self.sheet_headers: self.col_combo.current(0)
        except Exception as e:
            messagebox.showerror("读取错误", f"无法读取表头: {e}")

    def get_stop_words(self):
        raw = self.txt_stopwords.get('1.0', tk.END)
        words = re.split(r'[,，\\s\\n]+', raw)
        return set([w.strip() for w in words if w.strip()])
        
    def get_fixed_words(self):
        raw = self.txt_fixed.get('1.0', tk.END)
        words = re.split(r'[,，\\s\\n]+', raw)
        return set([w.strip() for w in words if w.strip()])

    def get_pos_exclusion_config(self):
        # Determine excluded roots
        excluded_roots = []
        if self.var_noun.get(): excluded_roots.append('n')
        if self.var_verb.get(): excluded_roots.append('v')
        if self.var_adj.get(): excluded_roots.append('a')
        if self.var_adv.get(): excluded_roots.append('d')
        if self.var_prep.get(): excluded_roots.append('p')
        if self.var_part.get(): excluded_roots.append('u')
        
        # 1. Logic for pseg loop (Stats & Segmentation) - Return False if excluded
        def is_allowed(flag):
            # If flag starts with any excluded root, return False
            for r in excluded_roots:
                if flag.startswith(r):
                    return False
            return True

        # 2. Logic for extract_tags allowPOS (Keywords & TF-IDF) - Construct whitelist
        allow_pos_list = []
        for root_key, tags in POS_MAP.items():
            if root_key not in excluded_roots:
                allow_pos_list.extend(tags)
        
        if not allow_pos_list:
            allow_pos_list = ['__dummy__']
            
        return tuple(allow_pos_list), is_allowed

    def run_analysis(self, mode):
        if not self.filepath.get() or not self.selected_column.get():
            messagebox.showwarning("提示", "请先选择文件和列")
            return

        self.btn_stats.config(state='disabled')
        self.btn_seg.config(state='disabled')
        self.btn_tfidf.config(state='disabled')
        self.btn_row_tfidf.config(state='disabled')
        self.btn_entropy.config(state='disabled')
        self.btn_ner_person.config(state='disabled')
        self.btn_ner_location.config(state='disabled')
        self.btn_save_file.config(state='disabled')
        self.lbl_status.config(text="正在分析中，请稍候...")
        
        target_col = self.selected_column.get()
        stop_words = self.get_stop_words()
        fixed_words = self.get_fixed_words()
        allow_pos_tuple, is_pos_allowed_func = self.get_pos_exclusion_config()
        
        filter_en = self.var_filter_en.get()
        filter_num = self.var_filter_num.get()
        min_len = self.var_min_len.get()
        max_len = self.var_max_len.get()
        
        args = (target_col, stop_words, fixed_words, mode, allow_pos_tuple, is_pos_allowed_func, filter_en, filter_num, min_len, max_len)
        threading.Thread(target=self._process_thread, args=args, daemon=True).start()

    def _process_thread(self, col_name, stop_words, fixed_words, mode, allow_pos_tuple, is_pos_allowed_func, filter_en, filter_num, min_len, max_len):
        temp_sw_path = None
        try:
            # Add fixed words to jieba
            for w in fixed_words:
                if w: jieba.add_word(w)

            def is_english_number_allowed(w):
                if filter_en and re.match(r'^[a-zA-Z]+$', w): return False
                if filter_num and re.match(r'^[0-9]+(\\.[0-9]+)?$', w): return False
                return True

            # TF-IDF Setup
            if mode in ['tfidf', 'row_tfidf']:
                fd, temp_sw_path = tempfile.mkstemp(suffix='.txt', text=True)
                try:
                    with os.fdopen(fd, 'w', encoding='utf-8') as f:
                        f.write("\\n".join(stop_words))
                except Exception:
                    os.close(fd)
                    raise
                jieba.analyse.set_stop_words(temp_sw_path)

            # Read Data
            col_idx = -1
            try: col_idx = self.sheet_headers.index(col_name)
            except: raise Exception("列名未找到")

            data_rows = []
            is_csv = self.workbook_path.endswith('.csv')
            
            if is_csv:
                with open(self.workbook_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    next(reader) 
                    data_rows = list(reader)
            else:
                ws = self.wb[self.selected_sheet.get()]
                all_rows = list(ws.iter_rows(values_only=False))
                data_rows = all_rows[1:] # Skip header

            # Processing
            result_preview = [] 
            full_result_payload = None 
            
            texts = []
            for row in data_rows:
                cell_val = ""
                if is_csv:
                    if col_idx < len(row): cell_val = row[col_idx]
                else:
                    cell_obj = row[col_idx]
                    cell_val = str(cell_obj.value) if cell_obj.value else ""
                texts.append(str(cell_val).strip())

            total_rows = len(texts)

            if mode == 'stats':
                all_words = []
                row_counts = Counter()
                for t in texts:
                    words = pseg.lcut(t)
                    row_words = set()
                    for w, f in words:
                        clean_w = w.strip()
                        is_chinese = re.search(r'[\u4e00-\u9fa5]', clean_w)
                        if (len(clean_w) > 1 or clean_w in fixed_words or is_chinese) and clean_w not in stop_words and is_pos_allowed_func(f) and is_english_number_allowed(clean_w):
                            all_words.append(clean_w)
                            row_words.add(clean_w)
                    for w in row_words:
                        row_counts[w] += 1
                            
                counts = Counter(all_words).most_common(1000)
                full_result_payload = [(w, c, row_counts[w], f"{(row_counts[w]/total_rows)*100:.2f}%") for w, c in counts]
                result_preview = [(i+1, w, c, rp) for i, (w, c, rc, rp) in enumerate(full_result_payload[:30])]

            elif mode == 'seg':
                segmented_col = []
                for t in texts:
                    words = pseg.lcut(t)
                    valid = []
                    for w, f in words:
                        clean_w = w.strip()
                        is_chinese = re.search(r'[\u4e00-\u9fa5]', clean_w)
                        if (len(clean_w) > 1 or clean_w in fixed_words or is_chinese) and clean_w not in stop_words and is_pos_allowed_func(f) and is_english_number_allowed(clean_w):
                            valid.append(clean_w)
                    segmented_col.append((len(valid), " ".join(valid)))
                
                for i in range(min(30, len(texts))):
                    orig = texts[i]
                    if len(orig) > 20: orig = orig[:20] + "..."
                    wc, res = segmented_col[i]
                    if len(res) > 20: res = res[:20] + "..."
                    result_preview.append((i+1, orig, wc, res))
                
                full_result_payload = (data_rows, segmented_col, is_csv)

            elif mode == 'row_tfidf':
                keywords_col = []
                for t in texts:
                    kws = jieba.analyse.extract_tags(t, topK=5, allowPOS=allow_pos_tuple)
                    filtered_kws = [w for w in kws if is_english_number_allowed(w)]
                    keywords_col.append(",".join(filtered_kws))
                
                for i in range(min(30, len(texts))):
                    orig = texts[i]
                    if len(orig) > 20: orig = orig[:20] + "..."
                    res = keywords_col[i]
                    if len(res) > 20: res = res[:20] + "..."
                    result_preview.append((i+1, orig, res))
                
                full_result_payload = (data_rows, keywords_col, is_csv)

            elif mode == 'tfidf':
                full_text = "\\n".join(texts)
                tags = jieba.analyse.extract_tags(full_text, topK=200, withWeight=True, allowPOS=allow_pos_tuple)
                
                row_counts = Counter()
                for t in texts:
                    for w, _ in tags:
                        if w in t:
                            row_counts[w] += 1

                filtered_tags = []
                for w, we in tags:
                    if is_english_number_allowed(w):
                        filtered_tags.append((w, we, row_counts[w], f"{(row_counts[w]/total_rows)*100:.2f}%"))

                result_preview = [(i+1, w, round(we, 4), rp) for i, (w, we, rc, rp) in enumerate(filtered_tags[:30])]
                full_result_payload = filtered_tags

            elif mode == 'entropy':
                import math
                def calc_entropy(words):
                    if not words: return 0
                    counts = Counter(words)
                    total = len(words)
                    ent = 0
                    for c in counts.values():
                        p = c / total
                        ent -= p * math.log2(p)
                    return ent

                entropy_data = []
                for i, t in enumerate(texts):
                    words = pseg.lcut(t)
                    valid = []
                    for w, f in words:
                        clean_w = w.strip()
                        is_chinese = re.search(r'[\u4e00-\u9fa5]', clean_w)
                        if (len(clean_w) > 1 or clean_w in fixed_words or is_chinese) and clean_w not in stop_words and is_pos_allowed_func(f) and is_english_number_allowed(clean_w):
                            valid.append(clean_w)
                    
                    word_count = len(valid)
                    ent = calc_entropy(valid)
                    is_anomaly = word_count < min_len or word_count > max_len
                    
                    if is_anomaly:
                        entropy_data.append((i, word_count, round(ent, 4), "异常"))

                for item in entropy_data[:30]:
                    idx, wc, ent, st = item
                    orig = texts[idx]
                    if len(orig) > 20: orig = orig[:20] + "..."
                    result_preview.append((idx+1, orig, wc, ent, st))
                
                full_result_payload = (data_rows, entropy_data, is_csv)

            elif mode in ['ner_person', 'ner_location']:
                ner_data = []
                target_tag = 'nr' if mode == 'ner_person' else 'ns'
                for i, t in enumerate(texts):
                    if not t.strip(): continue
                    words = pseg.lcut(t)
                    entities = []
                    for w, f in words:
                        if f == target_tag and w not in stop_words:
                            entities.append(f"{w}({f})")
                    
                    if entities:
                        ner_data.append((i, ", ".join(entities)))

                for item in ner_data[:30]:
                    idx, ents = item
                    orig = texts[idx]
                    if len(orig) > 20: orig = orig[:20] + "..."
                    result_preview.append((idx+1, orig, ents))
                
                full_result_payload = (data_rows, ner_data, is_csv)

            self.root.after(0, lambda: self._update_ui_success(mode, result_preview, full_result_payload))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            if temp_sw_path and os.path.exists(temp_sw_path):
                try: os.remove(temp_sw_path)
                except: pass
            self.root.after(0, lambda: self._enable_btns())

    def _update_ui_success(self, mode, preview_data, full_payload):
        self.current_preview_mode = mode
        self.last_result_data = full_payload
        
        for i in self.tree.get_children(): self.tree.delete(i)
        
        if mode == 'stats':
            self.tree['columns'] = ("rank", "word", "count", "row_percent")
            self.tree.heading("rank", text="#")
            self.tree.heading("word", text="词语")
            self.tree.heading("count", text="频次")
            self.tree.heading("row_percent", text="行数占比")
            self.tree.column("rank", width=50, anchor='center')
            self.tree.column("word", width=150)
            self.tree.column("count", width=80, anchor='center')
            self.tree.column("row_percent", width=80, anchor='center')
            self.lbl_status.config(text="预览：全局词频 Top 30")

        elif mode == 'seg':
            self.tree['columns'] = ("idx", "orig", "wc", "seg")
            self.tree.heading("idx", text="行号")
            self.tree.heading("orig", text="原文(部分)")
            self.tree.heading("wc", text="分词数量")
            self.tree.heading("seg", text="分词结果(部分)")
            self.tree.column("idx", width=50, anchor='center')
            self.tree.column("orig", width=150)
            self.tree.column("wc", width=80, anchor='center')
            self.tree.column("seg", width=200)
            self.lbl_status.config(text="预览：分词结果前 30 行")
        
        elif mode == 'row_tfidf':
            self.tree['columns'] = ("idx", "orig", "kw")
            self.tree.heading("idx", text="行号")
            self.tree.heading("orig", text="原文(部分)")
            self.tree.heading("kw", text="提取关键词(Top5)")
            self.tree.column("idx", width=50, anchor='center')
            self.tree.column("orig", width=200)
            self.tree.column("kw", width=200)
            self.lbl_status.config(text="预览：逐行关键词前 30 行")
            
        elif mode == 'tfidf':
            self.tree['columns'] = ("rank", "kw", "weight", "row_percent")
            self.tree.heading("rank", text="#")
            self.tree.heading("kw", text="关键词")
            self.tree.heading("weight", text="权重")
            self.tree.heading("row_percent", text="行数占比")
            self.tree.column("rank", width=50, anchor='center')
            self.tree.column("kw", width=150)
            self.tree.column("weight", width=80, anchor='center')
            self.tree.column("row_percent", width=80, anchor='center')
            self.lbl_status.config(text="预览：全局 TF-IDF 关键词 Top 30")

        elif mode == 'entropy':
            self.tree['columns'] = ("idx", "orig", "wc", "ent", "status")
            self.tree.heading("idx", text="行号")
            self.tree.heading("orig", text="原文(部分)")
            self.tree.heading("wc", text="词数")
            self.tree.heading("ent", text="信息熵")
            self.tree.heading("status", text="状态")
            self.tree.column("idx", width=50, anchor='center')
            self.tree.column("orig", width=150)
            self.tree.column("wc", width=60, anchor='center')
            self.tree.column("ent", width=80, anchor='center')
            self.tree.column("status", width=60, anchor='center')
            self.lbl_status.config(text="预览：逐行信息熵与长度 Top 30")

        elif mode in ['ner_person', 'ner_location']:
            self.tree['columns'] = ("idx", "orig", "ents")
            self.tree.heading("idx", text="行号")
            self.tree.heading("orig", text="原文(部分)")
            self.tree.heading("ents", text="人名 (nr)" if mode == 'ner_person' else "地名 (ns)")
            self.tree.column("idx", width=50, anchor='center')
            self.tree.column("orig", width=200)
            self.tree.column("ents", width=200)
            self.lbl_status.config(text=f"预览：{'人名' if mode == 'ner_person' else '地名'}提取 Top 30")

        for row in preview_data:
            self.tree.insert('', 'end', values=row)

        self.btn_save_file.config(state='normal')

    def save_current_result(self):
        if not self.current_preview_mode or not self.last_result_data: return
        
        mode = self.current_preview_mode
        data = self.last_result_data
        base_name = os.path.splitext(self.workbook_path)[0]
        col_name = self.selected_column.get()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_base = f"{base_name}_{col_name}_{timestamp}"
        
        try:
            if mode == 'stats':
                out_path = out_base + "_stats.xlsx"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["Rank", "Word", "Count", "Row Count", "Row Percent"])
                for i, (w, c, rc, rp) in enumerate(data):
                    ws.append([i+1, w, c, rc, rp])
                wb.save(out_path)
                messagebox.showinfo("成功", f"词频统计已导出：\\n{out_path}")

            elif mode in ['seg', 'row_tfidf']:
                suffix = "_segmented.xlsx" if mode == 'seg' else "_row_keywords.xlsx"
                
                out_path = out_base + suffix
                orig_rows, new_col, is_csv = data
                
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                
                if mode == 'seg':
                    ws_out.append(self.sheet_headers + ["Word_Count", "Segmented_Text"])
                else:
                    ws_out.append(self.sheet_headers + ["Keywords_Top5"])
                
                for i, row_obj in enumerate(orig_rows):
                    row_vals = []
                    if is_csv: row_vals = row_obj
                    else: row_vals = [c.value for c in row_obj]
                    
                    if mode == 'seg':
                        wc, seg_str = new_col[i]
                        ws_out.append(list(row_vals) + [wc, seg_str])
                    else:
                        ws_out.append(list(row_vals) + [new_col[i]])
                    
                wb_out.save(out_path)
                messagebox.showinfo("成功", f"文件已导出：\\n{out_path}")

            elif mode == 'entropy':
                out_path = out_base + "_entropy.xlsx"
                orig_rows, entropy_data, is_csv = data
                
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.append(self.sheet_headers + ["Word_Count", "Entropy", "Status"])
                
                for item in entropy_data:
                    idx, wc, ent, st = item
                    row_obj = orig_rows[idx]
                    row_vals = []
                    if is_csv: row_vals = row_obj
                    else: row_vals = [c.value for c in row_obj]
                    ws_out.append(list(row_vals) + [wc, ent, st])
                    
                wb_out.save(out_path)
                messagebox.showinfo("成功", f"信息熵分析已导出：\\n{out_path}")

            elif mode in ['ner_person', 'ner_location']:
                out_path = out_base + f"_{mode}.xlsx"
                orig_rows, ner_data, is_csv = data
                
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                col_name = "Entities_Person_nr" if mode == 'ner_person' else "Entities_Location_ns"
                ws_out.append(self.sheet_headers + [col_name])
                
                for item in ner_data:
                    idx, ents = item
                    row_obj = orig_rows[idx]
                    row_vals = []
                    if is_csv: row_vals = row_obj
                    else: row_vals = [c.value for c in row_obj]
                    ws_out.append(list(row_vals) + [ents])
                    
                wb_out.save(out_path)
                messagebox.showinfo("成功", f"{'人名' if mode == 'ner_person' else '地名'}提取结果已导出：\\n{out_path}")

            elif mode == 'tfidf':
                out_path = out_base + "_tfidf.xlsx"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["Rank", "Keyword", "Weight", "Row Count", "Row Percent"])
                for i, (w, we, rc, rp) in enumerate(data):
                    ws.append([i+1, w, we, rc, rp])
                wb.save(out_path)
                messagebox.showinfo("成功", f"关键词表已导出：\\n{out_path}")
                
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _enable_btns(self):
        self.btn_stats.config(state='normal')
        self.btn_seg.config(state='normal')
        self.btn_row_tfidf.config(state='normal')
        self.btn_tfidf.config(state='normal')
        self.btn_entropy.config(state='normal')
        self.btn_ner_person.config(state='normal')
        self.btn_ner_location.config(state='normal')

if __name__ == '__main__':
    root = tk.Tk()
    app = WordFreqApp(root)
    root.mainloop()
`;
