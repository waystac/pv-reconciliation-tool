"""
光伏结算对账核销工具 v3.5

作者: waystac
许可证: MIT License
仓库: https://github.com/waystac/pv-reconciliation-tool

Copyright (c) 2026 waystac

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import os
from decimal import Decimal, ROUND_HALF_UP, getcontext
from itertools import combinations

getcontext().prec = 28

class CheckApp:
    def __init__(self, root):
        self.root = root
        self.root.title("光伏结算对账核销工具 v3.5")
        self.root.geometry("1000x920")
        
        self.file1 = tk.StringVar()
        self.file2 = tk.StringVar()
        self.history_file = tk.StringVar()
        self.use_history = tk.BooleanVar(value=False)
        
        self.name_col = tk.StringVar(value="用户名")
        self.total_col = tk.StringVar(value="总数")
        self.detail_col = tk.StringVar(value="明细")
        self.detail2_col = tk.StringVar(value="备用")
        self.use_detail2 = tk.BooleanVar(value=False)
        self.date_col1 = tk.StringVar(value="日期")
        self.date_col2 = tk.StringVar(value="日期")
        self.tax_col = tk.StringVar(value="")
        
        self.tax_list = tk.StringVar(value="0")
        self.tolerance = tk.StringVar(value="0")
        
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        
        self.create_widgets()
    
    def create_widgets(self):
        title = tk.Label(self.root, text="光伏结算对账核销工具 v3.5", font=("微软雅黑", 16, "bold"))
        title.pack(pady=(10, 0))
        copyright_label = tk.Label(self.root, text="Copyright (C) 2026 waystac", font=("微软雅黑", 9), fg="gray")
        copyright_label.pack()
        desc_label = tk.Label(self.root,
                              text="功能：将汇总表中的每笔总数，与明细表中同一用户的多条明细自动配对勾销。",
                              font=("微软雅黑", 9), fg="gray")
        desc_label.pack(pady=(0, 10))
        
        # 文件选择
        file_frame = tk.LabelFrame(self.root, text="选择 Excel 文件", padx=10, pady=10)
        file_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(file_frame, text="汇总表：").grid(row=0, column=0, sticky="w")
        tk.Entry(file_frame, textvariable=self.file1, width=50).grid(row=0, column=1, padx=5)
        tk.Button(file_frame, text="浏览...", command=self.browse_file1).grid(row=0, column=2)
        tk.Label(file_frame, text="明细表：").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(file_frame, textvariable=self.file2, width=50).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(file_frame, text="浏览...", command=self.browse_file2).grid(row=1, column=2)
        
        # 历史匹配
        hist_frame = tk.LabelFrame(self.root, text="历史匹配（可选）", padx=10, pady=10)
        hist_frame.pack(fill="x", padx=20, pady=5)
        tk.Checkbutton(hist_frame, text="加载历史匹配", variable=self.use_history, command=self.toggle_history).grid(row=0, column=0, sticky="w")
        tk.Label(hist_frame, text="历史结果文件：").grid(row=0, column=1, sticky="w", padx=(20,0))
        self.hist_entry = tk.Entry(hist_frame, textvariable=self.history_file, width=50, state="disabled")
        self.hist_entry.grid(row=0, column=2, padx=5)
        self.hist_btn = tk.Button(hist_frame, text="浏览...", command=self.browse_history, state="disabled")
        self.hist_btn.grid(row=0, column=3)
        tk.Label(hist_frame, text="（选择上次导出的结果文件，需包含“汇总表-核对结果”和“明细表-核对结果”工作表）",
                 font=("微软雅黑", 8), fg="gray").grid(row=1, column=1, columnspan=3, sticky="w", pady=(2,0))
        
        # 列名设置
        col_frame = tk.LabelFrame(self.root, text="列名设置（与表格表头完全一致）", padx=10, pady=10)
        col_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(col_frame, text="用户名列：").grid(row=0, column=0, sticky="w")
        tk.Entry(col_frame, textvariable=self.name_col, width=12).grid(row=0, column=1, padx=3)
        tk.Label(col_frame, text="汇总表 总数列：").grid(row=0, column=2, sticky="w")
        tk.Entry(col_frame, textvariable=self.total_col, width=12).grid(row=0, column=3, padx=3)
        tk.Label(col_frame, text="汇总表 日期列：").grid(row=0, column=4, sticky="w")
        tk.Entry(col_frame, textvariable=self.date_col1, width=12).grid(row=0, column=5, padx=3)
        tk.Label(col_frame, text="明细表 明细列：").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(col_frame, textvariable=self.detail_col, width=12).grid(row=1, column=1, padx=3, pady=5)
        tk.Label(col_frame, text="明细表 备用列：").grid(row=1, column=2, sticky="w", pady=5)
        tk.Entry(col_frame, textvariable=self.detail2_col, width=12).grid(row=1, column=3, padx=3, pady=5)
        tk.Checkbutton(col_frame, text="启用", variable=self.use_detail2).grid(row=1, column=4, sticky="w")
        tk.Label(col_frame, text="明细表 日期列：").grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(col_frame, textvariable=self.date_col2, width=12).grid(row=2, column=1, padx=3, pady=5)
        tk.Label(col_frame, text="明细表 税率列：").grid(row=2, column=2, sticky="w", pady=5)
        tk.Entry(col_frame, textvariable=self.tax_col, width=12).grid(row=2, column=3, padx=3, pady=5)
        tk.Label(col_frame, text="（留空则使用下方统一税率列表）", font=("微软雅黑", 8), fg="gray").grid(row=2, column=4, sticky="w")
        
        # 税率与容差
        tax_frame = tk.LabelFrame(self.root, text="统一税率列表（无明细税率列时使用）", padx=10, pady=10)
        tax_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(tax_frame, text="尝试的税率列表(%)：").pack(side="left")
        tk.Entry(tax_frame, textvariable=self.tax_list, width=30).pack(side="left", padx=5)
        
        tol_frame = tk.Frame(self.root)
        tol_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(tol_frame, text="匹配容差(元)：").pack(side="left")
        tk.Entry(tol_frame, textvariable=self.tolerance, width=8).pack(side="left", padx=3)
        
        # 按钮
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始核对", width=10, command=self.start_check)
        self.run_btn.pack(side="left", padx=3)
        self.pause_btn = tk.Button(btn_frame, text="暂停", width=10, command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", padx=3)
        self.stop_btn = tk.Button(btn_frame, text="结束", width=10, command=self.stop_check, state="disabled")
        self.stop_btn.pack(side="left", padx=3)
        self.export_btn = tk.Button(btn_frame, text="导出结果 Excel", width=14, command=self.export_result, state="disabled")
        self.export_btn.pack(side="left", padx=3)
        self.save_log_btn = tk.Button(btn_frame, text="保存日志", width=10, command=self.save_log, state="disabled")
        self.save_log_btn.pack(side="left", padx=3)
        
        # 结果展示
        result_frame = tk.LabelFrame(self.root, text="核对结果", padx=10, pady=10)
        result_frame.pack(fill="both", expand=True, padx=20, pady=5)
        self.result_text = tk.Text(result_frame, wrap="word", font=("微软雅黑", 10))
        self.result_text.pack(fill="both", expand=True)
        
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill="x", padx=20, pady=5)
        self.status = tk.Label(self.root, text="就绪", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="bottom", fill="x")
    
    def toggle_history(self):
        if self.use_history.get():
            self.hist_entry.config(state="normal")
            self.hist_btn.config(state="normal")
        else:
            self.hist_entry.config(state="disabled")
            self.hist_btn.config(state="disabled")
            self.history_file.set("")
    
    def browse_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.file1.set(path)
    
    def browse_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.file2.set(path)
    
    def browse_history(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.history_file.set(path)
    
    def log(self, msg):
        self.result_text.insert("end", msg + "\n")
        self.result_text.see("end")
    
    def start_check(self):
        if not self.file1.get() or not self.file2.get():
            messagebox.showwarning("缺少文件", "请先选择汇总表和明细表文件！")
            return
        if self.use_history.get() and not self.history_file.get():
            messagebox.showwarning("缺少历史文件", "请选择历史匹配结果文件！")
            return
        self.result_text.delete(1.0, "end")
        self.export_btn.config(state="disabled")
        self.save_log_btn.config(state="disabled")
        self.run_btn.config(state="disabled")
        self.pause_btn.config(state="normal", text="暂停")
        self.stop_btn.config(state="normal")
        self.progress.start()
        self.status.config(text="正在核对，请稍候...")
        self.stop_event.clear()
        self.pause_event.set()
        thread = threading.Thread(target=self.run_check)
        thread.daemon = True
        thread.start()
    
    def toggle_pause(self):
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.pause_btn.config(text="继续")
            self.status.config(text="已暂停")
            self.progress.stop()
        else:
            self.pause_event.set()
            self.pause_btn.config(text="暂停")
            self.status.config(text="正在核对，请稍候...")
            self.progress.start()
    
    def stop_check(self):
        self.stop_event.set()
        self.pause_event.set()
        self.status.config(text="正在停止...")
    
    @staticmethod
    def _to_decimal(value):
        try:
            if pd.isna(value) or value == '':
                return Decimal('0.00')
            s = f"{float(value):.2f}"
            return Decimal(s)
        except (ValueError, TypeError):
            return Decimal('0.00')
    
    @staticmethod
    def _apply_tax(amount, tax_rate):
        tax_amount = (amount * Decimal(str(tax_rate / 100))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        return amount + tax_amount
    
    @staticmethod
    def _to_fen(dec):
        return int(dec * 100)
    
    @staticmethod
    def _normalize_name(name):
        return str(name).strip().replace('（', '(').replace('）', ')')
    
    # ==================== 历史加载（支持重复项按顺序匹配） ====================
    def _load_history(self, df1, df2, name, total, detail, date1, date2,
                      df1_name_norm, df2_name_norm, used_rows):
        """加载历史匹配记录，重复项按顺序逐一匹配"""
        self.log("正在加载历史匹配记录...")
        try:
            hist_df1 = pd.read_excel(self.history_file.get(), sheet_name='汇总表-核对结果')
            hist_df2 = pd.read_excel(self.history_file.get(), sheet_name='明细表-核对结果')
            
            # --- 汇总表恢复 ---
            hist_matched_1 = hist_df1[hist_df1['匹配状态'] == '已匹配']
            match_count_1 = 0
            for _, hrow in hist_matched_1.iterrows():
                h_user = self._normalize_name(hrow[name])
                h_total = round(float(hrow[total]), 2)
                h_date = pd.to_datetime(hrow[date1])
                # 筛选出所有符合条件的行，然后取第一条"未匹配"的
                mask = (
                    (df1_name_norm == h_user) &
                    (df1[total].round(2) == h_total) &
                    (df1[date1] == h_date) &
                    (df1['匹配状态'] == '未匹配')
                )
                candidates = df1[mask]
                if len(candidates) >= 1:
                    idx = candidates.index[0]
                    df1.at[idx, '匹配状态'] = '已匹配'
                    df1.at[idx, '匹配组号'] = hrow.get('匹配组号', '')
                    df1.at[idx, '所用税率%'] = str(hrow.get('所用税率%', ''))
                    df1.at[idx, '实际含税合计'] = float(hrow.get('实际含税合计', 0))
                    if '匹配方式' in hrow:
                        if '匹配方式' not in df1.columns:
                            df1['匹配方式'] = ''
                        df1.at[idx, '匹配方式'] = hrow['匹配方式']
                    match_count_1 += 1
                elif len(candidates) == 0 and hrow.get('匹配状态') == '已匹配':
                    self.log(f"⚠️ 历史汇总未找到匹配：{h_user} {h_total} {h_date.date()}")
            
            self.log(f"从历史文件恢复了 {match_count_1} 条汇总匹配记录。")
            
            # --- 明细表恢复 ---
            hist_matched_2 = hist_df2[hist_df2['匹配组号'].notna()]
            match_count_2 = 0
            for _, hrow in hist_matched_2.iterrows():
                h_user = self._normalize_name(hrow[name])
                h_amount = round(float(hrow[detail]), 2)
                h_date = pd.to_datetime(hrow[date2])
                mask = (
                    (df2_name_norm == h_user) &
                    (df2[detail].round(2) == h_amount) &
                    (df2[date2] == h_date) &
                    (~df2.index.isin(used_rows))
                )
                candidates = df2[mask]
                if len(candidates) >= 1:
                    idx = candidates.index[0]
                    used_rows.add(idx)
                    df2.at[idx, '匹配总数'] = hrow.get('匹配总数')
                    df2.at[idx, '匹配组号'] = hrow.get('匹配组号')
                    df2.at[idx, '换算税率%'] = str(hrow.get('换算税率%', ''))
                    df2.at[idx, '匹配来源'] = '历史'
                    match_count_2 += 1
                elif len(candidates) == 0:
                    self.log(f"⚠️ 历史明细未找到匹配：{h_user} {h_amount} {h_date.date()}")
            
            self.log(f"从历史文件恢复了 {match_count_2} 条明细匹配记录。")
        except Exception as e:
            self.log(f"❌ 加载历史匹配失败：{str(e)}")
            self.log("将忽略历史匹配，继续全新匹配。")
    
    # ==================== 主流程 ====================
    def run_check(self):
        try:
            df1 = pd.read_excel(self.file1.get())
            df2 = pd.read_excel(self.file2.get())
            
            name = self.name_col.get()
            total = self.total_col.get()
            detail = self.detail_col.get()
            detail2 = self.detail2_col.get().strip()
            date1 = self.date_col1.get()
            date2 = self.date_col2.get()
            tax_col_name = self.tax_col.get().strip()
            
            use_detail2 = self.use_detail2.get() and bool(detail2)
            if self.use_detail2.get() and not detail2:
                self.log("⚠️ 已启用备用列但未填写列名，将忽略备用列。")
            
            required1 = [name, total, date1]
            required2 = [name, detail, date2]
            if use_detail2:
                required2.append(detail2)
            missing1 = [col for col in required1 if col not in df1.columns]
            missing2 = [col for col in required2 if col not in df2.columns]
            if missing1:
                self.log(f"❌ 汇总表缺少列: {missing1}")
                return
            if missing2:
                self.log(f"❌ 明细表缺少列: {missing2}")
                return
            
            self.log(f"汇总表行数: {len(df1)}，明细表行数: {len(df2)}")
            
            # 数据清洗
            df1[name] = df1[name].astype(str).str.strip()
            df2[name] = df2[name].astype(str).str.strip()
            df1_name_norm = df1[name].apply(self._normalize_name)
            df2_name_norm = df2[name].apply(self._normalize_name)
            
            df1[date1] = pd.to_datetime(df1[date1], errors='coerce')
            
            # 修复明细表日期
            date_str = df2[date2].astype(str)
            extracted = date_str.str.extract(r'(\d{6})')[0]
            parsed_date = pd.to_datetime(extracted, format='%Y%m', errors='coerce')
            mask_na = parsed_date.isna()
            if mask_na.any():
                fallback = pd.to_datetime(date_str[mask_na], errors='coerce')
                fallback = fallback.dt.to_period('M').dt.to_timestamp()
                parsed_date[mask_na] = fallback
            df2[date2] = parsed_date
            
            na_after = df2[date2].isna().sum()
            if na_after > 0:
                self.log(f"⚠️ 明细表日期解析失败: {na_after} 条")
                fail_samples = date_str[df2[date2].isna()].head().tolist()
                self.log(f"   失败样例: {fail_samples}")
            
            # 数值列清洗
            df1[total] = df1[total].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').fillna(0)
            df2[detail] = df2[detail].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').fillna(0)
            if use_detail2:
                df2[detail2] = df2[detail2].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').fillna(0)
            
            # 初始化状态列
            df1['匹配状态'] = '未匹配'
            df1['匹配组号'] = ''
            df1['所用税率%'] = ''
            df1['实际含税合计'] = 0.0
            df1['失败原因'] = ''
            df2['匹配总数'] = None
            df2['匹配组号'] = None
            df2['换算税率%'] = None
            df2['匹配来源'] = ''
            
            used_rows = set()
            
            # ---- 加载历史匹配 ----
            if self.use_history.get() and self.history_file.get():
                self._load_history(df1, df2, name, total, detail, date1, date2,
                                   df1_name_norm, df2_name_norm, used_rows)
            
            # Decimal 转换
            df1['__target_dec__'] = df1[total].apply(self._to_decimal)
            df2['__detail_dec__'] = df2[detail].apply(self._to_decimal)
            detail_dec_series = df2['__detail_dec__']
            detail2_dec_series = None
            if use_detail2:
                df2['__detail2_dec__'] = df2[detail2].apply(self._to_decimal)
                detail2_dec_series = df2['__detail2_dec__']
            
            use_detail_tax = bool(tax_col_name) and (tax_col_name in df2.columns)
            if use_detail_tax:
                df2[tax_col_name] = pd.to_numeric(df2[tax_col_name], errors='coerce').fillna(0)
                tax_list = []
            else:
                tax_str = self.tax_list.get().strip()
                tax_list = [float(t.strip()) for t in tax_str.split(',') if t.strip()] if tax_str else [0.0]
            
            tol_yuan = float(self.tolerance.get().strip() or 0)
            if tol_yuan < 0: tol_yuan = 0
            tol_fen = round(tol_yuan * 100)
            
            self.log(f"{'使用明细税率列: '+tax_col_name if use_detail_tax else '统一税率列表: '+str(tax_list)}，容差：{tol_yuan} 元")
            if use_detail2:
                self.log(f"备用列已启用：{detail2}")
            
            found, miss = 0, 0
            
            for idx1, row1 in df1.iterrows():
                if self.stop_event.is_set(): break
                self.pause_event.wait()
                
                if row1['匹配状态'] == '已匹配':
                    found += 1
                    continue
                
                user = row1[name]
                user_norm = df1_name_norm.loc[idx1]
                target_dec = row1['__target_dec__']
                target_date = row1[date1]
                
                mask = (df2_name_norm == user_norm) & (~df2.index.isin(used_rows)) & (df2[date2] <= target_date)
                cand_idx = df2[mask].index.tolist()
                
                if not cand_idx:
                    user_all = df2[df2_name_norm == user_norm]
                    if len(user_all) == 0:
                        reason = "用户在明细表中无记录"
                    else:
                        user_unused = user_all[~user_all.index.isin(used_rows)]
                        if len(user_unused) == 0:
                            reason = "所有明细已被占用"
                        else:
                            reason = f"无可用明细"
                    df1.at[idx1, '失败原因'] = reason
                    self.log(f"【未匹配】{user} {float(target_dec):.2f} ({target_date.date()})：{reason}")
                    miss += 1
                    continue
                
                # 主明细列匹配
                success, fail_reasons = self._try_match_with_detail(
                    df1, df2, idx1, target_dec, target_date,
                    cand_idx, detail_dec_series, detail, '明细',
                    tax_list, use_detail_tax, tax_col_name, tol_fen, used_rows
                )
                if success:
                    found += 1
                    continue
                
                # 备用列匹配
                if detail2_dec_series is not None:
                    success2, fail_reasons2 = self._try_match_with_detail(
                        df1, df2, idx1, target_dec, target_date,
                        cand_idx, detail2_dec_series, detail2, '备用列',
                        tax_list, use_detail_tax, tax_col_name, tol_fen, used_rows
                    )
                    if success2:
                        found += 1
                        continue
                    fail_reasons = fail_reasons + fail_reasons2
                else:
                    fail_reasons.append("备用列：未启用")
                
                df1.at[idx1, '失败原因'] = "；".join(fail_reasons)
                self.log(f"【未匹配】{user} {float(target_dec):.2f} ({target_date.date()})：{df1.at[idx1, '失败原因']}")
                miss += 1
            
            # 计算差值列
            df1['差值'] = round(df1[total] - df1['实际含税合计'], 2)
            df1.loc[df1['匹配状态'] == '未匹配', '差值'] = pd.NA
            df1.drop(columns=['__target_dec__', '实际含税合计'], inplace=True, errors='ignore')
            cols = df1.columns.tolist()
            if '差值' in cols:
                cols.insert(cols.index('匹配组号') + 1, cols.pop(cols.index('差值')))
            df1 = df1[cols]
            
            # 删除明细表中的临时列
            df2.drop(columns=['__detail_dec__'], inplace=True, errors='ignore')
            if use_detail2:
                df2.drop(columns=['__detail2_dec__'], inplace=True, errors='ignore')
            
            self.df1_result = df1
            self.df2_result = df2
            self.log(f"\n核对完成！成功 {found} 条，未匹配 {miss} 条。")
            self.root.after(0, lambda: self.status.config(text=f"完成，成功{found}，未匹配{miss}"))
            self.root.after(0, lambda: self.export_btn.config(state="normal"))
            self.root.after(0, lambda: self.save_log_btn.config(state="normal"))
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("运行出错", err_msg))
        finally:
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            self.root.after(0, lambda: self.pause_btn.config(state="disabled"))
            self.root.after(0, lambda: self.stop_btn.config(state="disabled"))
    
    def _try_match_with_detail(self, df1, df2, idx1, target_dec, target_date,
                               cand_idx, detail_dec_series, detail_col_name, source_label,
                               tax_list, use_detail_tax, tax_col_name, tol_fen, used_rows):
        cand_dec = detail_dec_series.loc[cand_idx].tolist()
        cand_dates = df2.loc[cand_idx, self.date_col2.get()]
        cand_tax_rates = df2.loc[cand_idx, tax_col_name].tolist() if use_detail_tax else None
        
        trial_tax_list = [None] if use_detail_tax else tax_list
        all_fail_reasons = []
        
        for tax in trial_tax_list:
            if use_detail_tax:
                taxed_dec = [self._apply_tax(d, r) for d, r in zip(cand_dec, cand_tax_rates)]
                taxed_fen = [self._to_fen(d) for d in taxed_dec]
                used_tax_str = "明细自带"
            else:
                taxed_dec = [self._apply_tax(d, tax) for d in cand_dec]
                taxed_fen = [self._to_fen(d) for d in taxed_dec]
                used_tax_str = str(tax)
            
            target_fen = self._to_fen(target_dec)
            fail_reasons = []
            
            # 月度包
            month_packages = {}
            for i, idx in enumerate(cand_idx):
                m = cand_dates.iloc[i].to_period('M')
                if m not in month_packages:
                    month_packages[m] = {'indices': [], 'total_fen': 0}
                month_packages[m]['indices'].append(idx)
                month_packages[m]['total_fen'] += taxed_fen[i]
            
            # ---- 1. 整月匹配 ----
            for m, pkg in month_packages.items():
                if abs(pkg['total_fen'] - target_fen) <= tol_fen:
                    self._apply_match(df1, df2, pkg['indices'], idx1, f"组{idx1+1}",
                                      float(target_dec), used_tax_str, self.date_col2.get(),
                                      used_rows, f"{source_label}-整月({m})", pkg['total_fen'], source_label)
                    return True, []
            fail_reasons.append(f"{source_label}-整月(税率{used_tax_str}%)匹配失败")
            
            # ---- 2. 连续月份组合 ----
            month_list = sorted(month_packages.keys())
            month_totals = [month_packages[m]['total_fen'] for m in month_list]
            combo = self._find_contiguous_sum(target_fen, month_totals, tol_fen)
            if combo is not None:
                start, end = combo
                combined_indices = []
                for i in range(start, end + 1):
                    combined_indices.extend(month_packages[month_list[i]]['indices'])
                months_str = f"{month_list[start]}-{month_list[end]}"
                combined_fen = sum(month_totals[start:end+1])
                self._apply_match(df1, df2, combined_indices, idx1, f"组{idx1+1}",
                                  float(target_dec), used_tax_str, self.date_col2.get(),
                                  used_rows, f"{source_label}-连续月({months_str})", combined_fen, source_label)
                return True, []
            fail_reasons.append(f"{source_label}-连续月(税率{used_tax_str}%)匹配失败")
            
            # ---- 3. 单条匹配 ----
            for i, fen in enumerate(taxed_fen):
                if abs(fen - target_fen) <= tol_fen:
                    real_idx = [cand_idx[i]]
                    self._apply_match(df1, df2, real_idx, idx1, f"组{idx1+1}",
                                      float(target_dec), used_tax_str, self.date_col2.get(),
                                      used_rows, f"{source_label}-单条", fen, source_label)
                    return True, []
            fail_reasons.append(f"{source_label}-单条(税率{used_tax_str}%)匹配失败")
            
            # ---- 4. 单条DP ----
            if target_fen > 5000000:
                fail_reasons.append(f"{source_label}-单条DP(税率{used_tax_str}%)超过5万元，已跳过")
            else:
                match_idx = self._subset_sum(target_fen, taxed_fen)
                if match_idx is not None:
                    real_idx = [cand_idx[i] for i in match_idx]
                    matched_fen = sum(taxed_fen[i] for i in match_idx)
                    self._apply_match(df1, df2, real_idx, idx1, f"组{idx1+1}",
                                      float(target_dec), used_tax_str, self.date_col2.get(),
                                      used_rows, f"{source_label}-单条DP", matched_fen, source_label)
                    return True, []
                fail_reasons.append(f"{source_label}-单条DP(税率{used_tax_str}%)匹配失败")
            
            # ---- 5. 回溯（不显示具体条数） ----
            search_items = taxed_fen
            search_indices = list(range(len(cand_idx)))
            
            if len(cand_idx) > 22:
                dates = [cand_dates.iloc[i] for i in range(len(cand_idx))]
                sorted_pairs = sorted(enumerate(zip(taxed_fen, dates)),
                                      key=lambda x: x[1][1], reverse=True)
                search_indices = [i for i, _ in sorted_pairs[:22]]
                search_items = [taxed_fen[i] for i in search_indices]
            
            match_idx_sub = self._search_by_count(target_fen, search_items, tol_fen, start_r=2)
            if match_idx_sub is not None:
                real_idx = [cand_idx[search_indices[i]] for i in match_idx_sub]
                matched_fen = sum(taxed_fen[search_indices[i]] for i in match_idx_sub)
                self._apply_match(df1, df2, real_idx, idx1, f"组{idx1+1}",
                                  float(target_dec), used_tax_str, self.date_col2.get(),
                                  used_rows, f"{source_label}-回溯", matched_fen, source_label)
                return True, []
            fail_reasons.append(f"{source_label}-回溯(税率{used_tax_str}%)匹配失败")
            
            all_fail_reasons.extend(fail_reasons)
        
        return False, all_fail_reasons
    
    def _apply_match(self, df1, df2, real_idx, row1_idx, group_name, target, tax, date2_col,
                     used_rows, match_type="整月", actual_fen=0, source='明细'):
        df2.loc[real_idx, '匹配总数'] = target
        df2.loc[real_idx, '匹配组号'] = group_name
        df2.loc[real_idx, '换算税率%'] = str(tax)
        df2.loc[real_idx, '匹配来源'] = source
        df1.at[row1_idx, '匹配状态'] = '已匹配'
        df1.at[row1_idx, '匹配组号'] = group_name
        df1.at[row1_idx, '所用税率%'] = str(tax)
        if '匹配方式' not in df1.columns:
            df1['匹配方式'] = ''
        df1.at[row1_idx, '匹配方式'] = match_type
        df1.at[row1_idx, '实际含税合计'] = round(actual_fen / 100.0, 2)
        df1.at[row1_idx, '失败原因'] = ''
        used_rows.update(real_idx)
        self.log(f"【已匹配-{match_type}】{df1.at[row1_idx, self.name_col.get()]} {target} 税率{tax}% ← {len(real_idx)}条")
    
    def _find_contiguous_sum(self, target, arr, tol):
        n = len(arr)
        for start in range(n):
            s = 0
            for end in range(start, n):
                s += arr[end]
                if abs(s - target) <= tol:
                    return (start, end)
                if s > target + tol:
                    break
        return None
    
    def _subset_sum(self, target_int, items_int):
        if target_int < 0 or any(i < 0 for i in items_int): return None
        if target_int > 5000000: return None   # 5万元保护
        n = len(items_int)
        dp = [False]*(target_int+1)
        dp[0]=True
        last = [[-1]*(target_int+1) for _ in range(n+1)]
        for i in range(1,n+1):
            val=items_int[i-1]
            for s in range(target_int, val-1, -1):
                if not dp[s] and dp[s-val]:
                    dp[s]=True
                    last[i][s]=1
        if not dp[target_int]: return None
        selected=[]
        s=target_int
        for i in range(n,0,-1):
            if last[i][s]==1:
                selected.append(i-1)
                s-=items_int[i-1]
        selected.reverse()
        return selected
    
    def _search_by_count(self, target_int, items_int, tol, start_r=2):
        n = len(items_int)
        for r in range(start_r, n + 1):
            for combo in combinations(range(n), r):
                s = sum(items_int[i] for i in combo)
                if abs(s - target_int) <= tol:
                    return list(combo)
        return None
    
    def save_log(self):
        log_content = self.result_text.get(1.0, "end")
        if not log_content.strip():
            messagebox.showwarning("无日志", "当前没有可保存的日志内容。")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if save_path:
            try:
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                messagebox.showinfo("保存成功", f"日志已保存至：{save_path}")
            except Exception as e:
                messagebox.showerror("保存失败", str(e))
    
    def export_result(self):
        if not hasattr(self, 'df1_result'):
            messagebox.showwarning("没有结果", "请先执行核对！")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    self.df1_result.to_excel(writer, sheet_name='汇总表-核对结果', index=False)
                    self.df2_result.to_excel(writer, sheet_name='明细表-核对结果', index=False)
                messagebox.showinfo("导出成功", f"结果已保存至：{save_path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = CheckApp(root)
    root.mainloop()
