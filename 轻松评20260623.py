import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import shutil
import re
from pathlib import Path
from datetime import datetime
import numpy as np
import threading
import warnings

# 忽略 openpyxl 的默认样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ===================== 公共工具函数 =====================
def sanitize_filename(filename):
    """清理Windows非法字符"""
    return re.sub(r'[\\/*?:"<>|]', '_', str(filename))


def validate_date(date_str):
    """验证日期格式"""
    try:
        return datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        raise ValueError("日期格式错误，请使用YYYY-MM-DD格式")


def safe_join(items):
    """安全拼接字符串列表，处理数值类型"""
    return '-'.join([str(item) for item in items])


def safe_convert_float(value):
    """安全转换为浮点数，处理各种异常情况"""
    if pd.isna(value) or value in ['', ' ', None]:
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        try:
            cleaned = str(value).replace(',', '').strip()
            return float(cleaned) if cleaned else 0.0
        except:
            return 0.0


def extract_data_source(filename):
    """从文件名中提取数据来源（A股/B股/港股通/信用账户）和业务类型（证券持有变更/证券持有余额/证券冻结情况）"""
    pattern = r'^(证券持有变更|证券持有余额|证券冻结情况).*?（(A股|B股|港股通|信用账户)）'
    match = re.search(pattern, filename)
    if match:
        biz_type = match.group(1)
        source = match.group(2)
        return biz_type, source
    return None, None


# ===================== 主界面 =====================
class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("轻松评")
        self.geometry("400x400")
        self.cleaning_system = None
        self.organizer_system = None
        self.calculator = None
        self.risk_identifier = None
        self.price_matching = None
        self.create_widgets()
        self.add_watermark()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def create_widgets(self):
        ttk.Label(self, text="请选择操作模块", font=('微软雅黑', 14)).pack(pady=20)
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="数据清分", command=self.open_cleaning).pack(pady=10, fill=tk.X)
        ttk.Button(btn_frame, text="数据筛选", command=self.open_organizer).pack(pady=10, fill=tk.X)
        ttk.Button(btn_frame, text="价格匹配", command=self.open_price_matching).pack(pady=10, fill=tk.X)
        ttk.Button(btn_frame, text="收益计算", command=self.open_calculator).pack(pady=10, fill=tk.X)
        ttk.Button(btn_frame, text="风险识别", command=self.open_risk_identifier).pack(pady=10, fill=tk.X)

    def open_cleaning(self):
        if not self.cleaning_system or not self.cleaning_system.window.winfo_exists():
            self.cleaning_system = DataCleaningSystem(self)
        self.cleaning_system.window.lift()

    def open_organizer(self):
        if not self.organizer_system or not self.organizer_system.window.winfo_exists():
            self.organizer_system = DataOrganizerSystem(self)
        self.organizer_system.window.lift()

    def open_calculator(self):
        if not self.calculator or not self.calculator.winfo_exists():
            self.calculator = TaxCalculator(self)
        self.calculator.deiconify()

    def open_risk_identifier(self):
        if not self.risk_identifier or not self.risk_identifier.winfo_exists():
            self.risk_identifier = RiskIdentificationSystem(self)
        self.risk_identifier.deiconify()

    def open_price_matching(self):
        if not self.price_matching or not self.price_matching.winfo_exists():
            self.price_matching = PriceMatchingSystem(self)
        self.price_matching.deiconify()


# ===================== 数据清分模块（修复版 - 解决沪市0行问题 + 数值格式修复）=====================
class DataCleaningSystem:
    def __init__(self, master):
        self.master = master
        self.window = tk.Toplevel(master)
        self.window.title("数据清分系统")
        self.window.geometry("900x700")

        # 组件变量
        self.input_files = []
        self.mapping_file = tk.StringVar()
        self.output_path = tk.StringVar()
        self.group_method = tk.StringVar(value="tax_office")
        self.progress_value = tk.IntVar()

        self.create_widgets()
        self.add_watermark()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self.window,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.window.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def create_widgets(self):
        # 文件列表区域
        file_frame = ttk.LabelFrame(self.window, text="待清分文件列表")
        file_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        self.file_listbox = tk.Listbox(file_frame, height=6)
        scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        btn_file_frame = ttk.Frame(file_frame)
        btn_file_frame.pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_file_frame, text="添加文件", command=self.add_files).pack(pady=2)
        ttk.Button(btn_file_frame, text="移除选中", command=self.remove_selected).pack(pady=2)
        ttk.Button(btn_file_frame, text="清空列表", command=self.clear_files).pack(pady=2)

        # 映射文件和输出路径
        ttk.Label(self.window, text="主管税务机关映射文件:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.window, textvariable=self.mapping_file, width=50).grid(row=1, column=1)
        ttk.Button(self.window, text="浏览", command=lambda: self.select_file(self.mapping_file)).grid(row=1, column=2)

        ttk.Label(self.window, text="输出路径:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.window, textvariable=self.output_path, width=50).grid(row=2, column=1)
        ttk.Button(self.window, text="浏览", command=self.select_output).grid(row=2, column=2)

        # 归集方式选择
        group_frame = ttk.LabelFrame(self.window, text="归集方式")
        group_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        ttk.Radiobutton(group_frame, text="按主管税务机关归集", variable=self.group_method,
                        value="tax_office").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="按数据类型归集", variable=self.group_method,
                        value="data_type").pack(side=tk.LEFT)

        # 进度条
        ttk.Label(self.window, text="处理进度:").grid(row=4, column=0, pady=10, sticky="e")
        self.progress = ttk.Progressbar(self.window, orient="horizontal",
                                        length=400, mode="determinate",
                                        variable=self.progress_value)
        self.progress.grid(row=4, column=1, columnspan=2, pady=5, sticky="w")

        # 日志区域
        ttk.Label(self.window, text="处理明细:").grid(row=5, column=0, pady=5, sticky="ne")
        self.log_area = tk.Text(self.window, height=12, width=80)
        self.log_area.grid(row=5, column=1, columnspan=2, padx=5, pady=5)
        scrollbar2 = ttk.Scrollbar(self.window, command=self.log_area.yview)
        scrollbar2.grid(row=5, column=3, sticky="ns")
        self.log_area.config(yscrollcommand=scrollbar2.set)

        ttk.Button(self.window, text="开始处理", command=self.start_process_thread).grid(row=6, column=1, pady=20)

    def add_files(self):
        files = filedialog.askopenfilenames(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
                size = os.path.getsize(f) / 1024 / 1024
                self.file_listbox.insert(tk.END, f"{os.path.basename(f)} ({size:.2f} MB)")

    def remove_selected(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.input_files.pop(index)
            self.file_listbox.delete(index)

    def clear_files(self):
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)

    def select_file(self, var):
        file = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file:
            var.set(file)

    def select_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path.set(path)

    def start_process_thread(self):
        if not self.input_files:
            messagebox.showwarning("警告", "请先添加待清分文件")
            return
        if not self.mapping_file.get():
            messagebox.showwarning("警告", "请选择主管税务机关映射文件")
            return
        if not self.output_path.get():
            messagebox.showwarning("警告", "请选择输出路径")
            return
        self.progress_value.set(0)
        self.log_area.delete(1.0, tk.END)
        threading.Thread(target=self.process, daemon=True).start()

    def update_progress(self, value, log):
        self.window.after(0, lambda: self.progress_value.set(value))
        self.window.after(0, lambda: self.log_area.insert(tk.END, log + "\n"))
        self.window.after(0, lambda: self.log_area.see(tk.END))

    def process(self):
        """执行完整处理流程 - 修复沪市0行问题 + 数值格式修复"""
        try:
            # 加载映射
            self.update_progress(0, "正在加载主管税务机关映射文件...")
            mapping_df = pd.read_excel(self.mapping_file.get(), engine='openpyxl')
            company_to_bureau = dict(zip(mapping_df["公司名称"], mapping_df["主管税务机关"]))
            self.update_progress(5, f"映射文件加载成功，共{len(company_to_bureau)}条映射关系")

            today_str = datetime.now().strftime("%Y%m%d")
            base_output_dir = Path(self.output_path.get()) / f"数据清分结果{today_str}"
            base_output_dir.mkdir(parents=True, exist_ok=True)
            self.update_progress(5, f"输出根目录：{base_output_dir}")

            summary_data = []
            total_steps = len(self.input_files)
            self.update_progress(5, f"开始处理，共{total_steps}个文件...")

            for idx, file_path in enumerate(self.input_files):
                file_progress_start = 5 + int(idx / total_steps * 90)
                file_progress_end = 5 + int((idx + 1) / total_steps * 90)

                # 识别数据来源
                biz_type, source = self.extract_data_source(os.path.basename(file_path))
                if not biz_type or not source:
                    self.update_progress(file_progress_start, f"文件名格式不正确：{os.path.basename(file_path)}")
                    continue

                self.update_progress(file_progress_start, f"开始处理文件：{os.path.basename(file_path)} (数据来源：{source})")

                # 检查Excel文件和工作表
                try:
                    xl = pd.ExcelFile(file_path, engine='openpyxl')
                    self.update_progress(file_progress_start, f"  Excel文件工作表列表：{xl.sheet_names}")
                except Exception as e:
                    self.update_progress(file_progress_start, f"  无法读取文件：{str(e)}")
                    continue

                # 支持的市场类型（更宽松匹配）
                market_patterns = {
                    "沪市": r"^沪市.*",
                    "深市": r"^深市.*",
                    "北京市场": r"^北京市场.*",
                    "沪市港股通": r"^沪市港股通.*",
                    "深市港股通": r"^深市港股通.*",
                    "H股全流通": r"^H股.*全流通.*",
                }

                market_sheets = []
                for actual_sheet_name in xl.sheet_names:
                    for market, pattern in market_patterns.items():
                        if re.match(pattern, actual_sheet_name, re.IGNORECASE):
                            market_sheets.append((actual_sheet_name, market))
                            break

                if not market_sheets:
                    self.update_progress(file_progress_start, f"  警告：未找到可识别的工作表")
                    continue

                self.update_progress(file_progress_start, f"  发现{len(market_sheets)}个有效工作表：{[s[0] for s in market_sheets]}")

                all_data = pd.DataFrame()
                for sheet_name, market_type in market_sheets:
                    try:
                        self.update_progress(file_progress_start, f"    正在读取工作表 '{sheet_name}'...")

                        # 尝试多种方式读取工作表
                        df = None
                        read_method = ""

                        # 方式1：默认读取
                        try:
                            df = xl.parse(sheet_name=sheet_name)
                            read_method = "默认"
                            if not df.empty and len(df) > 0:
                                self.update_progress(file_progress_start, f"      方式1成功：{len(df)}行 x {len(df.columns)}列")
                        except Exception as e1:
                            self.update_progress(file_progress_start, f"      方式1失败：{str(e1)[:50]}")

                        # 方式2：如果为空，尝试跳过标题行
                        if df is None or df.empty or len(df) == 0:
                            for skip in [1, 2, 3, 4, 5]:
                                try:
                                    df_test = xl.parse(sheet_name=sheet_name, skiprows=skip)
                                    if not df_test.empty and len(df_test) > 0:
                                        df = df_test
                                        read_method = f"skiprows={skip}"
                                        self.update_progress(file_progress_start, f"      方式2成功（跳过{skip}行）：{len(df)}行")
                                        break
                                except:
                                    continue

                        # 方式3：尝试指定header行号
                        if df is None or df.empty or len(df) == 0:
                            for header_row in [0, 1, 2, 3, 4]:
                                try:
                                    df_test = xl.parse(sheet_name=sheet_name, header=header_row)
                                    if not df_test.empty and len(df_test) > 0:
                                        df = df_test
                                        read_method = f"header={header_row}"
                                        self.update_progress(file_progress_start, f"      方式3成功（header={header_row}）：{len(df)}行")
                                        break
                                except:
                                    continue

                        # 方式4：无header模式（最后一招）
                        if df is None or df.empty or len(df) == 0:
                            try:
                                df = xl.parse(sheet_name=sheet_name, header=None)
                                if not df.empty:
                                    read_method = "header=None"
                                    self.update_progress(file_progress_start, f"      方式4成功（无header）：{len(df)}行，列名需手动设置")
                                    if len(df.columns) >= 10:
                                        generic_cols = [f"Col_{i}" for i in range(len(df.columns))]
                                        df.columns = generic_cols
                                        self.update_progress(file_progress_start, f"      已设置通用列名：{generic_cols[:5]}...")
                            except Exception as e4:
                                self.update_progress(file_progress_start, f"      方式4失败：{str(e4)[:50]}")

                        # 最终检查
                        if df is None or df.empty or len(df) == 0:
                            self.update_progress(file_progress_start, f"    工作表'{sheet_name}'所有读取方式均失败，已跳过")
                            continue

                        self.update_progress(file_progress_start, f"    最终使用[{read_method}]读取'{sheet_name}'：{len(df)}行 x {len(df.columns)}列")
                        self.update_progress(file_progress_start, f"    列名：{list(df.columns)[:10]}{'...' if len(df.columns) > 10 else ''}")

                        # 数据类型转换（读取后处理）
                        if '证券代码' in df.columns:
                            df['证券代码'] = df['证券代码'].astype(str).apply(lambda x: self.clean_stock_code(x, market_type))
                        if '合约编码' in df.columns:
                            df['合约编码'] = df['合约编码'].astype(str).apply(lambda x: self.clean_contract_code(x))

                        # 宽松的数据检查
                        has_company_col = any(col in df.columns for col in ["实际权益拥有人名称", "持有人名称"])
                        if not has_company_col:
                            self.update_progress(file_progress_start, f"    工作表'{sheet_name}'缺少公司列（无实际权益拥有人名称/持有人名称），已跳过")
                            continue

                        df['市场类型'] = market_type
                        all_data = pd.concat([all_data, df], ignore_index=True)
                        self.update_progress(file_progress_start, f"    ✓ 成功读取工作表：{sheet_name} → {market_type}，累计{len(all_data)}行")

                    except Exception as e:
                        self.update_progress(file_progress_start, f"    ✗ 读取工作表{sheet_name}失败：{str(e)}")
                        continue

                if all_data.empty:
                    self.update_progress(file_progress_start, f"  ✗ 警告：文件未能解析到任何有效数据")
                    continue

                # 查找公司列
                company_col = None
                for col in ["实际权益拥有人名称", "持有人名称"]:
                    if col in all_data.columns:
                        company_col = col
                        break

                if company_col is None:
                    self.update_progress(file_progress_start, f"  ✗ 错误：未找到公司名称列")
                    continue

                self.update_progress(file_progress_start, f"  使用公司列：{company_col}，总行数：{len(all_data)}")

                original_name = Path(file_path).stem
                company_groups = all_data.groupby(company_col, dropna=False)
                total_companies = len(company_groups)
                self.update_progress(file_progress_start, f"  共发现{total_companies}家公司")

                for company_idx, (company, company_group) in enumerate(company_groups):
                    company = str(company).strip()
                    bureau = company_to_bureau.get(company, "未知主管税务机关")

                    progress = file_progress_start + int((company_idx / total_companies) * (file_progress_end - file_progress_start))
                    if company_idx < 5 or company_idx % 10 == 0:
                        self.update_progress(progress, f"    处理公司[{company_idx+1}/{total_companies}]：{company[:30]}... (主管税务机关：{bureau[:20]}...)")

                    for market, market_group in company_group.groupby('市场类型'):
                        code_column = '合约编码' if '合约编码' in market_group.columns else '证券代码'

                        for code, code_group in market_group.groupby(code_column, dropna=False):
                            if code_column == '合约编码':
                                code_str = self.format_contract_code(code)
                            else:
                                code_str = self.format_stock_code(code, market)

                            name_column = '合约简称' if '合约简称' in code_group.columns else '证券简称'
                            names = []
                            if name_column in code_group.columns:
                                for n in code_group[name_column].unique():
                                    if pd.isna(n):
                                        names.append(f'{name_column}缺失')
                                    elif isinstance(n, (float, int)):
                                        names.append(str(int(n)) if n == int(n) else str(n))
                                    else:
                                        names.append(str(n))
                            else:
                                names.append(f'{name_column}缺失')

                            name_part = safe_join(sorted(names)) or f'{name_column}缺失'

                            filename = f"{bureau}-{company}-{source}-{market}-{code_str}-{name_part}.xlsx"
                            safe_name = sanitize_filename(filename)

                            if self.group_method.get() == "tax_office":
                                target_dir = base_output_dir / bureau / company / original_name
                            else:
                                target_dir = base_output_dir / original_name / bureau / company

                            target_dir.mkdir(parents=True, exist_ok=True)
                            output_file = target_dir / safe_name

                            # 数值格式转换（解决绿色三角问题）
                            numeric_columns = ['过户数量', '期末余额', '成交价格']
                            for col in numeric_columns:
                                if col in code_group.columns:
                                    cleaned = code_group[col].astype(str).str.replace(',', '').str.replace(' ', '').str.strip()
                                    converted = pd.to_numeric(cleaned, errors='coerce')
                                    mask = converted.notna()
                                    code_group.loc[mask, col] = converted[mask]

                            code_group.to_excel(output_file, index=False, engine='openpyxl')

                            summary_data.append({
                                '主管税务机关': bureau,
                                '公司名称': company,
                                '数据来源': source,
                                '市场类型': market,
                                '证券代码': code_str,
                                '证券简称': name_part,
                                '文件路径': str(output_file)
                            })

                self.update_progress(file_progress_end, f"✓ 文件处理完成：{os.path.basename(file_path)}")

            # 生成汇总表
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_df = summary_df[['主管税务机关', '公司名称', '数据来源', '市场类型', '证券代码', '证券简称', '文件路径']]
                summary_path = base_output_dir / f"数据清分结果{today_str}.xlsx"
                summary_df.to_excel(summary_path, index=False, engine='openpyxl')
                self.update_progress(100, f"✓ 处理完成！共生成{len(summary_data)}个文件，汇总表：{summary_path}")
                messagebox.showinfo("成功", f"数据清分完成！\n共处理{len(self.input_files)}个文件\n生成{len(summary_data)}个结果文件")
            else:
                self.update_progress(100, "⚠ 未生成任何文件")
                messagebox.showinfo("完成", "处理完成，但未生成任何文件，请检查文件内容和映射关系。")

        except Exception as e:
            self.update_progress(0, f"✗ 发生错误: {str(e)}")
            import traceback
            error_detail = traceback.format_exc()
            self.update_progress(0, f"错误详情：{error_detail[:500]}...")
            messagebox.showerror("错误", str(e))

    def extract_data_source(self, filename):
        """从文件名中提取数据来源"""
        pattern = r'^(证券持有变更|证券持有余额|证券冻结情况).*?[（(](A股|B股|港股通|信用账户)[）)]'
        match = re.search(pattern, filename)
        if match:
            return match.group(1), match.group(2)
        if 'A股' in filename:
            return '证券持有变更', 'A股'
        elif 'B股' in filename:
            return '证券持有变更', 'B股'
        elif '港股通' in filename:
            return '证券持有变更', '港股通'
        elif '信用账户' in filename:
            return '证券持有变更', '信用账户'
        return None, None

    def clean_stock_code(self, code, market_type):
        """清洗证券代码"""
        if pd.isna(code):
            return '未知代码'
        code_str = str(code).strip()
        if code_str.endswith('.0'):
            code_str = code_str[:-2]
        if market_type in ('深市', '深市港股通'):
            if len(code_str) < 6:
                code_str = code_str.zfill(6)
        elif market_type in ('沪市', '沪市港股通'):
            if len(code_str) < 6 and code_str.isdigit():
                code_str = code_str.zfill(6)
        return code_str

    def clean_contract_code(self, code):
        """清洗合约编码"""
        if pd.isna(code):
            return '未知合约编码'
        code_str = str(code).strip()
        if code_str.endswith('.0'):
            code_str = code_str[:-2]
        return code_str

    def format_stock_code(self, code, market_type):
        """格式化证券代码用于文件名"""
        code_str = str(code).strip()
        if pd.isna(code) or code_str == 'nan':
            return '证券代码缺失'
        if '.' in code_str:
            code_str = code_str.split('.')[0]
        if market_type in ('深市', '深市港股通'):
            return code_str.zfill(6)
        elif market_type in ('沪市', '沪市港股通'):
            return code_str.zfill(6) if code_str.isdigit() else code_str
        else:
            return code_str

    def format_contract_code(self, code):
        """格式化合约编码用于文件名"""
        code_str = str(code).strip()
        if pd.isna(code) or code_str == 'nan':
            return '合约编码缺失'
        if '.' in code_str:
            code_str = code_str.split('.')[0]
        return code_str.strip()


# ===================== 数据筛选模块（优化版）=====================
class DataOrganizerSystem:
    def __init__(self, master):
        self.master = master
        self.window = tk.Toplevel(master)
        self.window.title("数据筛选系统")
        self.window.geometry("900x700")
        self.input_folder = tk.StringVar()
        self.filter_type = tk.StringVar(value="time")
        self.cutoff_date = tk.StringVar()
        self.output_path = tk.StringVar()
        self.rules_file = tk.StringVar()
        self.progress_value = tk.IntVar()

        self.window.columnconfigure(0, weight=0, minsize=120)
        self.window.columnconfigure(1, weight=1, minsize=400)
        self.window.columnconfigure(2, weight=0, minsize=80)

        self.create_widgets()
        self.add_watermark()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self.window,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.window.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def create_widgets(self):
        filter_frame = ttk.LabelFrame(self.window, text="筛选方式")
        filter_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        ttk.Radiobutton(filter_frame, text="按减持时间筛选", variable=self.filter_type,
                        value="time", command=self.toggle_inputs).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(filter_frame, text="按股票类型筛选", variable=self.filter_type,
                        value="stock_type", command=self.toggle_inputs).pack(side=tk.LEFT)

        ttk.Label(self.window, text="输入文件夹:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.window, textvariable=self.input_folder, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.window, text="浏览", command=self.select_input).grid(row=1, column=2, padx=5, pady=5, sticky="ew")

        self.time_frame = ttk.LabelFrame(self.window, text="减持时间筛选参数")
        self.time_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        self.time_frame.columnconfigure(0, weight=0, minsize=120)
        self.time_frame.columnconfigure(1, weight=1, minsize=150)
        self.time_frame.columnconfigure(2, weight=0, minsize=80)
        self.time_frame.columnconfigure(3, weight=0, minsize=80)
        self.time_frame.columnconfigure(4, weight=1, minsize=200)
        self.time_frame.columnconfigure(5, weight=0, minsize=80)

        ttk.Label(self.time_frame, text="截止日期 (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.time_frame, textvariable=self.cutoff_date, width=20).grid(row=0, column=1, padx=5, sticky="w")
        ttk.Label(self.time_frame, text="输出路径:").grid(row=0, column=3, padx=5, pady=5, sticky="e")
        ttk.Entry(self.time_frame, textvariable=self.output_path, width=30).grid(row=0, column=4, padx=5, sticky="w")
        ttk.Button(self.time_frame, text="浏览", command=self.select_output).grid(row=0, column=5, padx=5)

        self.stock_frame = ttk.LabelFrame(self.window, text="股票类型筛选参数")
        self.stock_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        self.stock_frame.columnconfigure(0, weight=0, minsize=120)
        self.stock_frame.columnconfigure(1, weight=1, minsize=150)
        self.stock_frame.columnconfigure(2, weight=0, minsize=80)
        self.stock_frame.columnconfigure(3, weight=0, minsize=80)
        self.stock_frame.columnconfigure(4, weight=1, minsize=200)
        self.stock_frame.columnconfigure(5, weight=0, minsize=80)

        ttk.Label(self.stock_frame, text="限售股筛选规则文件:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.stock_frame, textvariable=self.rules_file, width=40).grid(row=0, column=1, padx=5, sticky="w")
        ttk.Button(self.stock_frame, text="浏览", command=self.select_rules).grid(row=0, column=2, padx=5)
        ttk.Label(self.stock_frame, text="输出路径:").grid(row=0, column=3, padx=5, pady=5, sticky="e")
        ttk.Entry(self.stock_frame, textvariable=self.output_path, width=30).grid(row=0, column=4, padx=5, sticky="w")
        ttk.Button(self.stock_frame, text="浏览", command=self.select_output).grid(row=0, column=5, padx=5)

        ttk.Label(self.window, text="处理进度:").grid(row=3, column=0, pady=10, sticky="e")
        self.progress = ttk.Progressbar(self.window, orient="horizontal",
                                        length=400, mode="determinate",
                                        variable=self.progress_value)
        self.progress.grid(row=3, column=1, columnspan=2, pady=5, sticky="w")

        ttk.Label(self.window, text="处理明细:").grid(row=4, column=0, pady=5, sticky="ne")
        self.log_area = tk.Text(self.window, height=10, width=80)
        self.log_area.grid(row=4, column=1, columnspan=2, padx=5, pady=5)
        scrollbar = ttk.Scrollbar(self.window, command=self.log_area.yview)
        scrollbar.grid(row=4, column=3, sticky="ns")
        self.log_area.config(yscrollcommand=scrollbar.set)

        ttk.Button(self.window, text="开始筛选", command=self.start_process_thread).grid(row=5, column=1, pady=20)

        self.toggle_inputs()

    def toggle_inputs(self):
        if self.filter_type.get() == "time":
            self.time_frame.grid()
            self.stock_frame.grid_remove()
        else:
            self.time_frame.grid_remove()
            self.stock_frame.grid()

    def select_input(self):
        path = filedialog.askdirectory()
        if path:
            self.input_folder.set(path)

    def select_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path.set(path)

    def select_rules(self):
        file = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file:
            self.rules_file.set(file)

    def start_process_thread(self):
        if not self.input_folder.get():
            messagebox.showwarning("警告", "请选择输入文件夹")
            return
        if not self.output_path.get():
            messagebox.showwarning("警告", "请选择输出路径")
            return
        if self.filter_type.get() == "time" and not self.cutoff_date.get():
            messagebox.showwarning("警告", "请输入截止日期")
            return
        if self.filter_type.get() == "stock_type" and not self.rules_file.get():
            messagebox.showwarning("警告", "请选择限售股筛选规则文件")
            return

        self.progress_value.set(0)
        self.log_area.delete(1.0, tk.END)
        threading.Thread(target=self.process, daemon=True).start()

    def update_progress(self, value, log):
        self.window.after(0, lambda: self.progress_value.set(value))
        self.window.after(0, lambda: self._append_log(log))

    def _append_log(self, log):
        self.log_area.insert(tk.END, log + "\n")
        self.log_area.see(tk.END)

    def process(self):
        try:
            input_path = Path(self.input_folder.get())
            output_root = Path(self.output_path.get())
            output_root.mkdir(parents=True, exist_ok=True)

            if not input_path.exists():
                raise ValueError("输入文件夹不存在")

            if self.filter_type.get() == "time":
                self.process_by_time(input_path, output_root)
            else:
                self.process_by_stock_type(input_path, output_root)
        except Exception as e:
            self.update_progress(0, f"发生错误: {str(e)}")
            messagebox.showerror("错误", str(e))

    def process_by_time(self, input_path, output_root):
        """按减持时间筛选"""
        cutoff = validate_date(self.cutoff_date.get())
        date_str = self.cutoff_date.get().replace("-", "")

        result_folder_name = f"数据筛选结果（{date_str}后存在减持行为）"
        result_folder = output_root / result_folder_name
        result_folder.mkdir(parents=True, exist_ok=True)

        self.update_progress(5, f"开始按减持时间筛选，截止日期：{self.cutoff_date.get()}")
        self.update_progress(5, f"输出文件夹：{result_folder}")

        all_files = list(input_path.rglob('*.xlsx')) + list(input_path.rglob('*.xls'))
        self.total_files = len(all_files)
        if self.total_files == 0:
            raise ValueError("输入文件夹中没有Excel文件")

        self.update_progress(10, f"共发现 {self.total_files} 个待处理文件")

        results = []
        processed_count = 0
        matched_count = 0

        for idx, file in enumerate(all_files):
            progress = 10 + int((idx + 1) / self.total_files * 80)
            self.update_progress(progress, f"[{idx + 1}/{self.total_files}] 正在处理: {file.name}")

            result = self.process_file_time(file, cutoff, result_folder, input_path)
            processed_count += 1

            if result:
                results.append(result)
                matched_count += 1
                self.update_progress(progress, f"  ✓ 符合条件，已复制: {result['证券简称']}")
            else:
                self.update_progress(progress, f"  ✗ 不符合条件，已跳过")

        if results:
            summary_df = pd.DataFrame(results, columns=[
                "主管税务机关", "公司名称", "数据来源", "市场类型", "证券代码", "证券简称", "文件路径"
            ])
            summary_df.insert(0, "序号", range(1, len(summary_df) + 1))

            summary_filename = f"数据筛选结果（{date_str}后存在减持行为）.xlsx"
            summary_path = result_folder / summary_filename
            summary_df.to_excel(summary_path, index=False, engine='openpyxl')

            self.update_progress(100, f"✓ 处理完成！共处理 {processed_count} 个文件，筛选出 {matched_count} 个符合条件")
            self.update_progress(100, f"✓ 汇总表：{summary_path}")
            messagebox.showinfo("完成",
                                f"按减持时间筛选完成！\n共处理 {processed_count} 个文件\n筛选出 {matched_count} 个符合条件\n汇总表：{summary_path}")
        else:
            self.update_progress(100, f"⚠ 处理完成，但未找到符合条件的文件（共处理 {processed_count} 个文件）")
            messagebox.showinfo("完成", f"处理完成，但未找到符合条件的文件\n共处理 {processed_count} 个文件")

    def safe_float_convert(self, x):
        """安全转换为浮点数，处理各种异常情况"""
        if pd.isna(x) or x in ['', ' ', None, 'None', 'nan', 'NaN', 'NULL', '-']:
            return 0.0
        try:
            cleaned = str(x).replace(',', '').strip()
            return float(cleaned) if cleaned else 0.0
        except (ValueError, TypeError):
            return 0.0

    def process_file_time(self, file_path, cutoff_date, output_root, input_path):
        """处理单个文件（按减持时间筛选）"""
        try:
            df = pd.read_excel(
                file_path,
                engine='openpyxl',
                converters={
                    '过户数量': self.safe_float_convert,
                    '过户日期': str
                }
            )

            if '过户日期' not in df.columns or '过户数量' not in df.columns:
                return None

            df['过户日期'] = pd.to_datetime(df['过户日期'], errors='coerce')
            df = df.dropna(subset=['过户日期'])

            negative_df = df[df['过户数量'] < 0]
            if negative_df.empty:
                return None

            last_negative = negative_df.iloc[-1]
            last_date = last_negative['过户日期']

            if last_date <= cutoff_date:
                return None

            parts = file_path.stem.split('-')
            if len(parts) < 6:
                return None

            bureau, company, source, market, code, name = parts[0], parts[1], parts[2], parts[3], parts[4], '-'.join(parts[5:])

            try:
                relative_path = file_path.relative_to(input_path).parent
                target_folder = output_root / relative_path
                target_folder.mkdir(parents=True, exist_ok=True)
            except ValueError:
                target_folder = output_root
                target_folder.mkdir(parents=True, exist_ok=True)

            dest_path = target_folder / file_path.name
            shutil.copy2(file_path, dest_path)

            return {
                "主管税务机关": bureau,
                "公司名称": company,
                "数据来源": source,
                "市场类型": market,
                "证券代码": code,
                "证券简称": name,
                "文件路径": str(dest_path)
            }
        except Exception as e:
            self.update_progress(0, f"处理文件 {file_path.name} 时出错: {str(e)}")
            return None

    def process_by_stock_type(self, input_path, output_root):
        """按股票类型筛选"""
        rules_df = pd.read_excel(self.rules_file.get(), engine='openpyxl')
        sh_keywords = rules_df['证券类别'].dropna().str.strip().tolist()
        other_keywords = rules_df['股份性质'].dropna().str.strip().tolist()

        result_folder_name = "数据筛选结果（限售股）"
        result_folder = output_root / result_folder_name
        result_folder.mkdir(parents=True, exist_ok=True)

        self.update_progress(5, "开始按股票类型筛选（限售股）")
        self.update_progress(5, f"输出文件夹：{result_folder}")
        self.update_progress(5, f"沪市关键词：{sh_keywords}")
        self.update_progress(5, f"深市/北京市场关键词：{other_keywords}")

        all_files = list(input_path.rglob('*.xlsx')) + list(input_path.rglob('*.xls'))
        self.total_files = len(all_files)
        if self.total_files == 0:
            raise ValueError("输入文件夹中没有Excel文件")

        self.update_progress(10, f"共发现 {self.total_files} 个待处理文件")

        results = []
        processed_count = 0
        matched_count = 0

        for idx, file in enumerate(all_files):
            progress = 10 + int((idx + 1) / self.total_files * 80)
            self.update_progress(progress, f"[{idx + 1}/{self.total_files}] 正在处理: {file.name}")

            result = self.process_file_stock_type(file, sh_keywords, other_keywords, result_folder, input_path)
            processed_count += 1

            if result:
                results.append(result)
                matched_count += 1
                self.update_progress(progress, f"  ✓ 匹配限售股，已复制: {result['证券简称']}")
            else:
                self.update_progress(progress, f"  ✗ 非限售股，已跳过")

        if results:
            summary_df = pd.DataFrame(results, columns=[
                "主管税务机关", "公司名称", "数据来源", "市场类型", "证券代码", "证券简称", "文件路径"
            ])
            summary_df.insert(0, "序号", range(1, len(summary_df) + 1))

            summary_filename = "数据筛选结果（限售股）.xlsx"
            summary_path = result_folder / summary_filename
            summary_df.to_excel(summary_path, index=False, engine='openpyxl')

            self.update_progress(100, f"✓ 处理完成！共处理 {processed_count} 个文件，筛选出 {matched_count} 个限售股文件")
            self.update_progress(100, f"✓ 汇总表：{summary_path}")
            messagebox.showinfo("完成",
                                f"按股票类型筛选完成！\n共处理 {processed_count} 个文件\n筛选出 {matched_count} 个限售股文件\n汇总表：{summary_path}")
        else:
            self.update_progress(100, f"⚠ 处理完成，但未找到限售股文件（共处理 {processed_count} 个文件）")
            messagebox.showinfo("完成", f"处理完成，但未找到限售股文件\n共处理 {processed_count} 个文件")

    def process_file_stock_type(self, file_path, sh_keywords, other_keywords, output_root, input_path):
        """处理单个文件（按股票类型筛选）"""
        try:
            parts = file_path.stem.split('-')
            if len(parts) < 6:
                return None

            bureau, company, source, market, code, name = parts[0], parts[1], parts[2], parts[3], parts[4], '-'.join(parts[5:])

            if market == '沪市':
                key_column = '证券类别'
                keywords = sh_keywords
            elif market in ('深市', '北京市场'):
                key_column = '股份性质'
                keywords = other_keywords
            else:
                return None

            try:
                temp_df = pd.read_excel(file_path, nrows=1, engine='openpyxl')
                if key_column not in temp_df.columns:
                    return None
            except Exception:
                return None

            df = pd.read_excel(file_path, usecols=[key_column], engine='openpyxl')
            df[key_column] = df[key_column].astype(str).str.strip()

            matched = any(val in keywords for val in df[key_column])
            if not matched:
                return None

            try:
                relative_path = file_path.relative_to(input_path).parent
                target_folder = output_root / relative_path
                target_folder.mkdir(parents=True, exist_ok=True)
            except ValueError:
                target_folder = output_root
                target_folder.mkdir(parents=True, exist_ok=True)

            dest_path = target_folder / file_path.name
            shutil.copy2(file_path, dest_path)

            return {
                "主管税务机关": bureau,
                "公司名称": company,
                "数据来源": source,
                "市场类型": market,
                "证券代码": code,
                "证券简称": name,
                "文件路径": str(dest_path)
            }
        except Exception as e:
            self.update_progress(0, f"处理文件 {file_path.name} 时出错: {str(e)}")
            return None


# ===================== 价格匹配模块（修复版 - 按公司/来源/市场/代码四维分组，与清分维度对齐）=====================
class PriceMatchingSystem(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("价格匹配系统")
        self.geometry("1000x700")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)
        self.match_type = tk.StringVar(value="market")
        self.input_mode = tk.StringVar(value="file")
        self.input_files = []
        self.input_folder = tk.StringVar()
        self.price_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.progress_value = tk.IntVar()
        self.price_cache = {}
        self.price_cache_by_name = {}
        self.create_widgets()
        self.add_watermark()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        type_frame = ttk.LabelFrame(main_frame, text="匹配方式")
        type_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(type_frame, text="市场日均价匹配", variable=self.match_type,
                        value="market", command=self.toggle_inputs).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="券商交易价格匹配", variable=self.match_type,
                        value="broker", command=self.toggle_inputs).pack(side=tk.LEFT)
        tk.Label(type_frame, text="（券商交易价格匹配暂未开放）", fg="gray").pack(side=tk.LEFT, padx=10)

        mode_frame = ttk.LabelFrame(main_frame, text="输入模式")
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(mode_frame, text="批量导入文件", variable=self.input_mode,
                        value="file", command=self.toggle_inputs).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="导入文件夹", variable=self.input_mode,
                        value="folder", command=self.toggle_inputs).pack(side=tk.LEFT)

        self.file_frame = ttk.LabelFrame(main_frame, text="待匹配文件列表")
        self.file_listbox = tk.Listbox(self.file_frame, height=6)
        scrollbar = ttk.Scrollbar(self.file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        btn_file_frame = ttk.Frame(self.file_frame)
        btn_file_frame.pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_file_frame, text="添加文件", command=self.add_files).pack(pady=2)
        ttk.Button(btn_file_frame, text="移除选中", command=self.remove_selected).pack(pady=2)
        ttk.Button(btn_file_frame, text="清空列表", command=self.clear_files).pack(pady=2)

        self.folder_frame = ttk.LabelFrame(main_frame, text="文件夹输入")
        ttk.Label(self.folder_frame, text="输入文件夹:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.folder_frame, textvariable=self.input_folder, width=50).grid(row=0, column=1)
        ttk.Button(self.folder_frame, text="浏览", command=self.select_input_folder).grid(row=0, column=2)

        self.param_frame = ttk.LabelFrame(main_frame, text="参数设置")
        self.param_frame.pack(fill=tk.X, pady=5)
        ttk.Label(self.param_frame, text="价格文件夹（每日均价）:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.param_frame, textvariable=self.price_folder, width=50).grid(row=0, column=1)
        ttk.Button(self.param_frame, text="浏览", command=self.select_price_folder).grid(row=0, column=2)
        ttk.Label(self.param_frame, text="输出文件夹:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(self.param_frame, textvariable=self.output_folder, width=50).grid(row=1, column=1)
        ttk.Button(self.param_frame, text="浏览", command=self.select_output_folder).grid(row=1, column=2)

        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        ttk.Label(progress_frame, text="处理进度:").pack(side=tk.LEFT, padx=5)
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal",
                                        length=600, mode="determinate",
                                        variable=self.progress_value)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        log_frame = ttk.LabelFrame(main_frame, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        import tkinter.scrolledtext as st
        self.log_area = st.ScrolledText(log_frame, wrap=tk.WORD, height=10, width=100)
        self.log_area.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.log_area.tag_config("timestamp", foreground="gray")
        self.log_area.tag_config("info", foreground="black")
        self.log_area.tag_config("success", foreground="green")
        self.log_area.tag_config("warning", foreground="orange")
        self.log_area.tag_config("error", foreground="red")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="开始匹配", command=self.start_process_thread).pack(side=tk.LEFT, padx=5)

        self.toggle_inputs()

    def toggle_inputs(self):
        if self.match_type.get() == "broker":
            messagebox.showinfo("提示", "券商交易价格匹配功能暂未开放")
            self.match_type.set("market")
        if self.input_mode.get() == "file":
            self.file_frame.pack(fill=tk.X, pady=5, before=self.param_frame)
            self.folder_frame.pack_forget()
        else:
            self.file_frame.pack_forget()
            self.folder_frame.pack(fill=tk.X, pady=5, before=self.param_frame)

    def add_files(self):
        files = filedialog.askopenfilenames(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
                size = os.path.getsize(f) / 1024
                self.file_listbox.insert(tk.END, f"{os.path.basename(f)} ({size:.1f} KB)")

    def remove_selected(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.input_files.pop(index)
            self.file_listbox.delete(index)

    def clear_files(self):
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)

    def select_input_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.input_folder.set(path)

    def select_price_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.price_folder.set(path)

    def select_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder.set(path)

    def start_process_thread(self):
        if not self.price_folder.get():
            messagebox.showwarning("警告", "请选择价格文件夹")
            return
        if not self.output_folder.get():
            messagebox.showwarning("警告", "请选择输出文件夹")
            return
        if self.input_mode.get() == "file" and not self.input_files:
            messagebox.showwarning("警告", "请添加待匹配文件")
            return
        if self.input_mode.get() == "folder" and not self.input_folder.get():
            messagebox.showwarning("警告", "请选择输入文件夹")
            return

        self.progress_value.set(0)
        self.log_area.delete(1.0, tk.END)
        threading.Thread(target=self.process, daemon=True).start()

    def log(self, message, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_area.insert(tk.END, f"{message}\n", level)
        self.log_area.yview(tk.END)
        self.update()

    def update_progress(self, value):
        self.after(0, lambda: self.progress_value.set(value))

    def load_all_price_data(self, price_folder):
        """加载所有价格文件到内存缓存（代码+简称双映射），统一日期格式为 YYYY-MM-DD"""
        price_path = Path(price_folder)
        if not price_path.exists():
            raise FileNotFoundError(f"价格文件夹不存在: {price_folder}")

        self.price_cache.clear()
        self.price_cache_by_name.clear()

        price_files = list(price_path.glob("*.xlsx")) + list(price_path.glob("*.xls"))
        total_files = len(price_files)
        self.log(f"开始加载价格数据，共发现 {total_files} 个价格文件...", "info")

        for idx, pf in enumerate(price_files):
            try:
                price_df = pd.read_excel(pf, dtype=str)
                if price_df.empty:
                    continue

                date_col = None
                price_col = None
                name_col = None
                code_col = None
                for col in price_df.columns:
                    if "日期" in col:
                        date_col = col
                    if "均价" in col:
                        price_col = col
                    if "简称" in col:
                        name_col = col
                    if "代码" in col and "证券" not in col:
                        code_col = col

                if not date_col or not price_col:
                    self.log(f"  跳过 {pf.name}: 缺少必要列（日期/均价）", "warning")
                    continue

                file_name = None
                if name_col:
                    file_name = str(price_df[name_col].iloc[0]).strip()

                file_code = None
                if code_col:
                    first_code = str(price_df[code_col].iloc[0]).strip()
                    match = re.search(r"(\d{6})", first_code)
                    if match:
                        file_code = match.group(1)
                if not file_code:
                    match = re.search(r"\b(\d{6})\b", pf.stem.upper())
                    if match:
                        file_code = match.group(1)

                record_count = 0
                for _, row in price_df.iterrows():
                    raw_date = str(row[date_col]).strip()
                    try:
                        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y.%m.%d"):
                            try:
                                parsed = datetime.strptime(raw_date, fmt)
                                date_val = parsed.strftime("%Y-%m-%d")
                                break
                            except:
                                continue
                        else:
                            date_val = raw_date.replace("/", "-").split()[0]
                    except:
                        date_val = raw_date.replace("/", "-").split()[0]

                    price_val = str(row[price_col]).strip()
                    if file_code:
                        if file_code not in self.price_cache:
                            self.price_cache[file_code] = {}
                        self.price_cache[file_code][date_val] = price_val
                        record_count += 1
                    if file_name:
                        if file_name not in self.price_cache_by_name:
                            self.price_cache_by_name[file_name] = {}
                        self.price_cache_by_name[file_name][date_val] = price_val

                progress = int((idx + 1) / total_files * 30)
                self.update_progress(progress)
                self.log(
                    f"  ✓ 已加载 {pf.name} (代码:{file_code or '未知'}, 简称:{file_name or '未知'}, {record_count}条记录)",
                    "success")

            except Exception as e:
                self.log(f"  ✗ 加载 {pf.name} 失败: {str(e)[:50]}", "error")

        self.log(
            f"价格数据加载完成！代码映射: {len(self.price_cache)} 支股票，简称映射: {len(self.price_cache_by_name)} 支股票",
            "success")

    def parse_filename_info(self, file_path):
        """从文件名解析主管税务机关和公司名称，以及市场类型"""
        stem = Path(file_path).stem
        # 移除已匹配标记，避免干扰
        stem = stem.replace('（已匹配市场日均价）', '').replace('（匹配状态：无需匹配）', '').replace('（匹配状态：部分匹配）',
                                                                                                 '').replace(
            '（匹配状态：全部匹配）', '').replace('（匹配状态：未匹配）', '').replace(' - 计算结果', '')

        parts = stem.split('-')
        result = {
            'bureau': '未知',
            'company': '未知',
            'source': '未知',
            'market': '未知',
            'code': '未知',
            'name': '未知'
        }

        if len(parts) >= 6:
            result['bureau'] = parts[0]
            result['company'] = parts[1]
            result['source'] = parts[2]
            result['market'] = parts[3]  # 沪市/深市/北京市场等
            result['code'] = parts[4]
            result['name'] = '-'.join(parts[5:])
        elif len(parts) >= 2:
            result['bureau'] = parts[0]
            result['company'] = parts[1]

        # 从文件名直接提取市场类型（作为备选）
        if '沪市' in stem or '600' in stem[:3] or '601' in stem[:3] or '603' in stem[:3] or '688' in stem[:3]:
            result['market'] = '沪市'
        elif '深市' in stem or '深圳' in stem or '000' in stem[:3] or '001' in stem[:3] or '002' in stem[
                                                                                                    :3] or '300' in stem[
                                                                                                                    :3]:
            result['market'] = '深市'
        elif '北京市场' in stem or '北交所' in stem or '北证' in stem or '835' in stem[:3] or '920' in stem[:3]:
            result['market'] = '北京市场'

        return result

    def get_market_type_from_sheetname(self, sheet_name):
        """根据工作表名称返回标准市场类型"""
        sn = str(sheet_name).strip()
        if "沪市" in sn:
            return "沪市"
        if "深市" in sn:
            return "深市"
        if "北京市场" in sn:
            return "北京市场"
        if "沪市港股通" in sn:
            return "沪市港股通"
        if "深市港股通" in sn:
            return "深市港股通"
        if "H股全流通" in sn:
            return "H股全流通"
        return None

    def detect_data_source(self, df):
        """根据列名判断数据来源：A股 或 信用账户"""
        cols = [str(c) for c in df.columns]
        if "持有人名称" in cols:
            return "A股"
        if "实际权益拥有人名称" in cols:
            return "信用账户"
        return "未知"

    def get_company_name(self, df, idx):
        """获取指定行的公司名称（持有人/实际权益拥有人）"""
        if "持有人名称" in df.columns:
            val = df.loc[idx, "持有人名称"]
            if pd.notna(val):
                return str(val).strip()
        if "实际权益拥有人名称" in df.columns:
            val = df.loc[idx, "实际权益拥有人名称"]
            if pd.notna(val):
                return str(val).strip()
        return "未知"

    def match_market_price(self, transaction_file, output_root, input_root=None):
        """
        返回结果列表，每个元素对应一个 (证券代码, 市场类型, 数据来源, 公司名称) 的匹配明细。
        与清分模块维度对齐：清分按 (bureau, company, source, market, code) 拆分文件，
        匹配汇总按 (code, market, source, company) 分组，确保行数和条数口径完全一致。
        """
        try:
            xl = pd.ExcelFile(transaction_file)
            all_sheets_data = {}
            all_missing_records = []
            all_stock_info = []
            total_transactions = 0

            # 首先尝试从文件名获取市场类型（作为备选）
            file_market_type = self.parse_filename_info(transaction_file).get('market')

            for sheet_name in xl.sheet_names:
                raw_df = xl.parse(sheet_name, header=None, dtype=str)
                if raw_df.empty:
                    continue
                header_row = None
                for i in range(min(5, len(raw_df))):
                    row_vals = raw_df.iloc[i].astype(str).str.lower()
                    if any("证券代码" in v for v in row_vals) or any("过户日期" in v for v in row_vals):
                        header_row = i
                        break
                if header_row is None:
                    header_row = 0
                df = xl.parse(sheet_name, header=header_row, dtype=str)
                if df.empty:
                    continue

                all_sheets_data[sheet_name] = df
                total_transactions += len(df)

                # 优先从工作表名称识别，失败则使用文件名中的市场类型
                market_type = self.get_market_type_from_sheetname(sheet_name)
                if market_type is None:
                    market_type = file_market_type

                if market_type is None:
                    self.log(f"  Sheet '{sheet_name}' 无法识别市场类型且文件名不包含市场信息，跳过", "warning")
                    continue

                code_col = None
                for col in df.columns:
                    if "证券代码" in str(col):
                        code_col = col
                        break
                date_col = None
                for col in df.columns:
                    if "过户日期" in str(col):
                        date_col = col
                        break
                price_col = None
                for col in df.columns:
                    if "成交价格" in str(col):
                        price_col = col
                        break
                short_name_col = None
                for col in df.columns:
                    if "证券简称" in str(col) or ("简称" in str(col) and "证券" not in str(col)):
                        short_name_col = col
                        break

                if not code_col or not date_col or not price_col:
                    self.log(f"  Sheet '{sheet_name}' 缺少必要列（证券代码/过户日期/成交价格），跳过", "warning")
                    continue

                df[code_col] = df[code_col].astype(str).str.strip().str.zfill(6)

                data_source = self.detect_data_source(df)

                for idx in df.index:
                    stock_code = str(df.loc[idx, code_col]).strip().zfill(6)
                    short_name = ""
                    if short_name_col:
                        short_name = str(df.loc[idx, short_name_col]).strip()
                    company_name = self.get_company_name(df, idx)

                    all_stock_info.append((
                        sheet_name, idx, stock_code, short_name, company_name, data_source, market_type
                    ))

                    if market_type in ("深市", "深市港股通", "北京市场"):
                        price_val = df.loc[idx, price_col]
                        is_blank = False
                        if pd.isna(price_val) or str(price_val).strip() == "":
                            is_blank = True
                        elif str(price_val).strip().lower() in ["nan", "null", "none"]:
                            is_blank = True

                        if is_blank:
                            trans_date = str(df.loc[idx, date_col]).strip()
                            if trans_date and trans_date != "nan":
                                try:
                                    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y.%m.%d"):
                                        try:
                                            parsed = datetime.strptime(trans_date, fmt)
                                            date_clean = parsed.strftime("%Y-%m-%d")
                                            break
                                        except:
                                            continue
                                    else:
                                        date_clean = trans_date.replace("/", "-").split()[0]
                                except:
                                    date_clean = trans_date.replace("/", "-").split()[0]

                                # 关键修复：缺失记录带上 company_name，确保分组维度完整
                                all_missing_records.append((
                                    sheet_name, idx, stock_code, short_name, date_clean, market_type, data_source,
                                    company_name
                                ))

            if not all_stock_info:
                return []

            # ==================== 修复核心：按 (stock_code, market_type, data_source, company_name) 复合键分组 ====================
            # 该维度与清分模块输出的原子文件维度 (bureau, company, source, market, code) 对齐
            from collections import defaultdict
            group_records = defaultdict(list)  # key: (stock_code, market_type, data_source, company_name)
            group_info = {}

            for (sheet_name, idx, stock_code, short_name, company_name, data_source, market_type) in all_stock_info:
                key = (stock_code, market_type, data_source, company_name)
                group_records[key].append((sheet_name, idx))
                if key not in group_info:
                    group_info[key] = {"short_names": set()}
                if short_name:
                    group_info[key]["short_names"].add(short_name)

            # 价格缓存简称映射（保持按 stock_code 维度）
            code_to_names = {}
            for _, _, stock_code, short_name, _, _, _, _ in all_missing_records:
                if stock_code not in code_to_names:
                    code_to_names[stock_code] = set()
                if short_name:
                    code_to_names[stock_code].add(short_name)

            name_mapped_codes = set()
            for stock_code, names in code_to_names.items():
                if stock_code in self.price_cache:
                    continue
                for name in names:
                    if name in self.price_cache_by_name:
                        self.price_cache[stock_code] = self.price_cache_by_name[name].copy()
                        name_mapped_codes.add(stock_code)
                        self.log(f"    股票 {stock_code} 通过简称 '{name}' 建立价格文件映射", "info")
                        break

            # 填充价格
            filled_sheets = {sheet: df.copy() for sheet, df in all_sheets_data.items()}

            # 匹配统计也按复合键
            group_match_stats = defaultdict(
                lambda: {"need": 0, "matched": 0, "unmatched_file": 0, "unmatched_date": 0, "unmatched_name": 0})

            for sheet_name, idx, stock_code, short_name, date_str, mkt, src, comp in all_missing_records:
                df = filled_sheets[sheet_name]
                price_col = None
                for col in df.columns:
                    if "成交价格" in str(col):
                        price_col = col
                        break
                if price_col is None:
                    continue

                key = (stock_code, mkt, src, comp)
                group_match_stats[key]["need"] += 1
                price_found = None
                if stock_code in self.price_cache and date_str in self.price_cache[stock_code]:
                    price_found = self.price_cache[stock_code][date_str]

                if price_found is not None:
                    df.loc[idx, price_col] = price_found
                    group_match_stats[key]["matched"] += 1
                else:
                    if stock_code not in self.price_cache:
                        group_match_stats[key]["unmatched_file"] += 1
                    elif stock_code in self.price_cache and date_str not in self.price_cache[stock_code]:
                        group_match_stats[key]["unmatched_date"] += 1
                    else:
                        group_match_stats[key]["unmatched_name"] += 1

            # 计算整体匹配状态（基于复合键综合判断）
            overall_status = "无需匹配"
            if group_info:
                statuses = []
                for key in group_info.keys():
                    stock_code, market_type, data_source, company_name = key
                    if market_type in ("沪市", "沪市港股通"):
                        statuses.append("无需匹配")
                    else:
                        stats = group_match_stats.get(key, {"need": 0, "matched": 0})
                        need = stats["need"]
                        matched = stats["matched"]

                        if need == 0:
                            statuses.append("全部匹配")
                        elif matched == need:
                            statuses.append("全部匹配")
                        elif matched == 0:
                            statuses.append("未匹配")
                        else:
                            statuses.append("部分匹配")

                if all(s == "无需匹配" for s in statuses):
                    overall_status = "无需匹配"
                elif all(s == "全部匹配" for s in statuses):
                    overall_status = "全部匹配"
                elif all(s == "未匹配" for s in statuses):
                    overall_status = "未匹配"
                else:
                    overall_status = "部分匹配"

            # 生成输出文件路径
            today_str = datetime.now().strftime("%Y%m%d")
            if input_root:
                rel_path = Path(transaction_file).relative_to(input_root).parent
                out_file = output_root / rel_path / f"{Path(transaction_file).stem}（匹配状态：{overall_status}）{Path(transaction_file).suffix}"
            else:
                result_subfolder = output_root / f"市场日均价匹配结果{today_str}"
                result_subfolder.mkdir(parents=True, exist_ok=True)
                out_file = result_subfolder / f"{Path(transaction_file).stem}（匹配状态：{overall_status}）{Path(transaction_file).suffix}"

            out_file.parent.mkdir(parents=True, exist_ok=True)

            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                for sheet_name, df in filled_sheets.items():
                    numeric_cols = ['过户数量', '期末余额', '成交价格']
                    for col in numeric_cols:
                        if col in df.columns:
                            cleaned = df[col].astype(str).str.replace(',', '').str.strip()
                            converted = pd.to_numeric(cleaned, errors='coerce')
                            mask = converted.notna()
                            df.loc[mask, col] = converted[mask]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            summary_rows = []
            file_info = self.parse_filename_info(transaction_file)
            tax_bureau = file_info['bureau']

            # ==================== 按复合键遍历生成汇总行，全部交易条数精确对应该公司/来源/市场下的记录数 ====================
            for key, info in group_info.items():
                stock_code, market_type, data_source, company_name = key
                short_names = "/".join(sorted(info["short_names"])) if info["short_names"] else "未知"
                total_stock_trans = len(group_records[key])  # 仅该公司+来源+市场+代码下的条数

                stats = group_match_stats.get(key,
                                              {"need": 0, "matched": 0, "unmatched_file": 0, "unmatched_date": 0,
                                               "unmatched_name": 0})
                need = stats["need"]
                matched = stats["matched"]
                unmatched = need - matched

                if market_type in ("沪市", "沪市港股通"):
                    status = "无需匹配"
                    detail = "沪市股票，无需匹配价格"
                elif need == 0:
                    status = "全部匹配"
                    detail = "所有交易记录均已存在成交价格"
                elif matched == need:
                    status = "全部匹配"
                    detail = f"需匹配{need}条，全部成功"
                    if stock_code in name_mapped_codes:
                        detail += "（通过简称匹配建立映射）"
                elif matched > 0:
                    status = "部分匹配"
                    reasons = []
                    if stats["unmatched_file"] > 0:
                        reasons.append(f"价格文件不存在:{stats['unmatched_file']}")
                    if stats["unmatched_date"] > 0:
                        reasons.append(f"日期无数据:{stats['unmatched_date']}")
                    if stats["unmatched_name"] > 0:
                        reasons.append(f"简称匹配失败:{stats['unmatched_name']}")
                    detail = f"需匹配{need}条，成功{matched}条，失败{unmatched}条（{','.join(reasons)}）"
                    if stock_code in name_mapped_codes:
                        detail += "（通过简称匹配建立映射）"
                else:
                    status = "未匹配"
                    reasons = []
                    if stats["unmatched_file"] > 0:
                        reasons.append(f"价格文件不存在:{stats['unmatched_file']}")
                    if stats["unmatched_date"] > 0:
                        reasons.append(f"日期无数据:{stats['unmatched_date']}")
                    if stats["unmatched_name"] > 0:
                        reasons.append(f"简称匹配失败:{stats['unmatched_name']}")
                    detail = f"需匹配{need}条，全部失败（{','.join(reasons)}）"

                summary_rows.append({
                    "主管税务机关": tax_bureau,
                    "公司名称": company_name,
                    "数据来源": data_source,
                    "市场类型": market_type,
                    "证券代码": stock_code,
                    "证券简称": short_names,
                    "匹配状态": status,
                    "全部交易条数": total_stock_trans,
                    "需匹配交易条数": need,
                    "已匹配交易条数": matched,
                    "未匹配交易条数": unmatched,
                    "匹配详情": detail,
                    "文件路径": str(out_file)
                })

            return summary_rows

        except Exception as e:
            self.log(f"处理文件 {transaction_file} 时出错: {str(e)}", "error")
            import traceback
            self.log(traceback.format_exc()[:300], "error")
            return []

    def process(self):
        """主处理流程：加载价格数据，处理每个文件，生成汇总表"""
        try:
            today_str = datetime.now().strftime("%Y%m%d")

            if self.input_mode.get() == "file":
                files_to_process = self.input_files
                output_root = Path(self.output_folder.get())
                output_root.mkdir(parents=True, exist_ok=True)
                input_root = None
            else:
                input_root = Path(self.input_folder.get())
                if not input_root.exists():
                    raise ValueError("输入文件夹不存在")
                files_to_process = list(input_root.rglob("*.xlsx")) + list(input_root.rglob("*.xls"))
                output_root = Path(self.output_folder.get())
                output_root.mkdir(parents=True, exist_ok=True)

            total_files = len(files_to_process)
            if total_files == 0:
                raise ValueError("没有找到Excel文件")

            self.log(f"共发现 {total_files} 个待处理交易文件", "info")

            self.load_all_price_data(self.price_folder.get())

            all_summary_rows = []
            for idx, input_file in enumerate(files_to_process):
                progress = 30 + int((idx + 1) / total_files * 70)
                self.update_progress(progress)
                self.log(f"[{idx + 1}/{total_files}] 正在处理: {os.path.basename(input_file)}", "info")

                rows = self.match_market_price(input_file, output_root, input_root)
                if rows:
                    all_summary_rows.extend(rows)
                    for row in rows:
                        status = row["匹配状态"]
                        comp_display = row['公司名称'][:15] if len(row['公司名称']) > 15 else row['公司名称']
                        if status == "无需匹配":
                            self.log(
                                f"  ℹ {row['证券代码']}({row['证券简称']}) [{row['市场类型']}-{row['数据来源']}-{comp_display}] - {status}",
                                "info")
                        elif status == "全部匹配":
                            self.log(
                                f"  ✓ {row['证券代码']}({row['证券简称']}) [{row['市场类型']}-{row['数据来源']}-{comp_display}] - {status} ({row['已匹配交易条数']}/{row['需匹配交易条数']})",
                                "success")
                        elif status == "部分匹配":
                            self.log(
                                f"  ⚠ {row['证券代码']}({row['证券简称']}) [{row['市场类型']}-{row['数据来源']}-{comp_display}] - {status} ({row['已匹配交易条数']}/{row['需匹配交易条数']})",
                                "warning")
                        else:
                            self.log(
                                f"  ✗ {row['证券代码']}({row['证券简称']}) [{row['市场类型']}-{row['数据来源']}-{comp_display}] - {status} (0/{row['需匹配交易条数']})",
                                "error")
                else:
                    self.log(f"  ⚠ 文件未生成任何汇总记录（可能无有效数据）", "warning")

            if all_summary_rows:
                for i, rec in enumerate(all_summary_rows, 1):
                    rec["序号"] = i
                summary_df = pd.DataFrame(all_summary_rows)
                column_order = [
                    "序号", "主管税务机关", "公司名称", "数据来源", "市场类型",
                    "证券代码", "证券简称", "匹配状态",
                    "全部交易条数", "需匹配交易条数", "已匹配交易条数", "未匹配交易条数",
                    "匹配详情", "文件路径"
                ]
                summary_df = summary_df[column_order]
                summary_filename = f"市场日均价匹配结果{today_str}.xlsx"
                summary_path = output_root / summary_filename

                numeric_cols = ['全部交易条数', '需匹配交易条数', '已匹配交易条数', '未匹配交易条数']
                for col in numeric_cols:
                    if col in summary_df.columns:
                        cleaned = summary_df[col].astype(str).str.replace(',', '').str.strip()
                        converted = pd.to_numeric(cleaned, errors='coerce')
                        mask = converted.notna()
                        summary_df.loc[mask, col] = converted[mask]

                summary_df.to_excel(summary_path, index=False, engine="openpyxl")

                self.log(f"\n{'=' * 50}", "info")
                self.log(f"处理完成！汇总表：{summary_path}", "success")
                self.log(f"共处理 {total_files} 个文件，生成 {len(all_summary_rows)} 条明细记录", "success")

                stats = summary_df["匹配状态"].value_counts().to_dict()
                self.log(
                    f"匹配统计：无需匹配 {stats.get('无需匹配', 0)} 条，全部匹配 {stats.get('全部匹配', 0)} 条，部分匹配 {stats.get('部分匹配', 0)} 条，未匹配 {stats.get('未匹配', 0)} 条",
                    "info")
                messagebox.showinfo("完成",
                                    f"处理完成！\n共处理{total_files}个文件\n生成{len(all_summary_rows)}条明细记录\n汇总表：{summary_path}")
            else:
                self.log("处理完成，但无匹配结果明细", "warning")
                messagebox.showinfo("完成", "处理完成，但未生成匹配明细（可能所有文件无有效数据）")

        except Exception as e:
            self.log(f"发生错误: {str(e)}", "error")
            messagebox.showerror("错误", str(e))


# ===================== 收益计算模块（修改版 - 汇总所有减持记录） =====================
class TaxCalculator(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("收益计算系统")
        self.geometry("1400x900")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)
        self.df = None
        self.result_df = None
        self.current_file_path = None
        self.batch_files = []
        self.batch_results = {}
        self.batch_output_dir = tk.StringVar()

        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.single_frame = ttk.LabelFrame(main_frame, text="单个文件处理")
        self.single_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.single_frame.columnconfigure(0, weight=1)

        row = 0

        file_frame = ttk.LabelFrame(self.single_frame, text="文件操作")
        file_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=5)
        self.btn_load = ttk.Button(file_frame, text="导入Excel文件", command=self.load_file)
        self.btn_load.pack(side=tk.LEFT, padx=5)
        row += 1

        type_frame = ttk.LabelFrame(self.single_frame, text="计算类型选择")
        type_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=5)
        self.calc_type = tk.StringVar(value="tradable")
        ttk.Radiobutton(type_frame, text="流通股减持", variable=self.calc_type,
                        value="tradable", command=self.toggle_inputs).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="限售股减持", variable=self.calc_type,
                        value="restricted", command=self.toggle_inputs).pack(side=tk.LEFT)
        row += 1

        self.input_frame = ttk.LabelFrame(self.single_frame, text="参数输入")
        ttk.Label(self.input_frame, text="根据以下政策确定的增值税买入价:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.ipo_entry = ttk.Entry(self.input_frame, width=15)
        self.ipo_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(self.input_frame, text="解禁日期 (YYYY-MM-DD):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.unlock_entry = ttk.Entry(self.input_frame, width=15)
        self.unlock_entry.grid(row=0, column=3, padx=5, pady=5)
        self.input_frame_row = row
        row += 1

        self.policy_frame = ttk.LabelFrame(self.single_frame, text="参考政策")
        policy_container = ttk.Frame(self.policy_frame)
        policy_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.policy_text = tk.Text(policy_container, height=8, wrap=tk.WORD)
        policy_scrollbar = ttk.Scrollbar(policy_container, orient="vertical", command=self.policy_text.yview)
        self.policy_text.configure(yscrollcommand=policy_scrollbar.set)
        self.policy_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        policy_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        policy_content = """国家税务总局公告2016年第53号
五、单位将其持有的限售股在解禁流通后对外转让的，按照以下规定确定买入价：
(一)上市公司实施股权分置改革时，在股票复牌之前形成的原非流通股股份，以及股票复牌首日至解禁日期间由上述股份孳生的送、转股，以该上市公司完成股权分置改革后股票复牌首日的开盘价为买入价。
(二)公司首次公开发行股票并上市形成的限售股，以及上市首日至解禁日期间由上述股份孳生的送、转股，以该上市公司股票首次公开发行(IPO)的发行价为买入价。
(三)因上市公司实施重大资产重组形成的限售股，以及股票复牌首日至解禁日期间由上述股份孳生的送、转股，以该上市公司因重大资产重组股票停牌前一交易日的收盘价为买入价。

国家税务总局公告2018年第42号
四、上市公司因实施重大资产重组形成的限售股，以及股票复牌首日至解禁日期间由上述股份孳生的送、转股，因重大资产重组停牌的，按照《国家税务总局关于营改增试点若干征管问题的公告》(国家税务总局公告2016年第53号)第五条第(三)项的规定确定买入价；在重大资产重组前已经暂停上市的，以上市公司完成资产重组后股票恢复上市首日的开盘价为买入价。

国家税务总局公告2019年第31号
十、关于限售股买入价的确定
(一)纳税人转让因同时实施股权分置改革和重大资产重组而首次公开发行股票并上市形成的限售股，以及上市首日至解禁日期间由上述股份孳生的送、转股，以该上市公司股票上市首日开盘价为买入价，按照"金融商品转让"缴纳增值税。

国家税务总局公告2020年第9号
四、单位将其持有的限售股在解禁流通后对外转让，按照《国家税务总局关于营改增试点若干征管问题的公告》(2016年第53号)第五条规定确定的买入价，低于该单位取得限售股的实际成本价的，以实际成本价为买入价计算缴纳增值税。"""
        self.policy_text.insert(tk.END, policy_content)
        self.policy_text.config(state=tk.DISABLED)
        self.policy_frame_row = row
        row += 1

        self.result_frame = ttk.LabelFrame(self.single_frame, text="单个文件计算结果")
        tree_container = ttk.Frame(self.result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(tree_container, orient="vertical")
        hsb = ttk.Scrollbar(tree_container, orient="horizontal")
        self.tree = ttk.Treeview(
            tree_container,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode="extended",
            columns=[],
            show="headings"
        )
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill=tk.BOTH, expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.result_frame_row = row
        self.single_frame.rowconfigure(row, weight=1)
        row += 1

        self.btn_frame = ttk.Frame(self.single_frame)
        self.btn_calculate = ttk.Button(self.btn_frame, text="开始计算", command=self.calculate)
        self.btn_calculate.pack(side=tk.LEFT, padx=5)
        self.btn_export = ttk.Button(self.btn_frame, text="导出结果", command=self.export_result)
        self.btn_export.pack(side=tk.LEFT, padx=5)
        self.btn_frame_row = row
        row += 1

        batch_frame = ttk.LabelFrame(main_frame, text="批量文件处理（适用于流通股）")
        batch_frame.pack(fill=tk.X, pady=5)

        list_frame = ttk.Frame(batch_frame)
        list_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(list_frame, text="批量文件列表:").pack(anchor=tk.W)
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill=tk.X, pady=5)
        self.file_listbox = tk.Listbox(list_container, height=6)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        path_frame = ttk.Frame(batch_frame)
        path_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(path_frame, text="导出文件夹:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(path_frame, textvariable=self.batch_output_dir, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="浏览", command=self.select_batch_output_dir).pack(side=tk.LEFT, padx=5)

        progress_frame = ttk.Frame(batch_frame)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(progress_frame, text="计算进度:").pack(side=tk.LEFT, padx=5)
        self.batch_progress_var = tk.DoubleVar()
        self.batch_progress = ttk.Progressbar(progress_frame, variable=self.batch_progress_var, maximum=100, length=400)
        self.batch_progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.batch_status_label = ttk.Label(progress_frame, text="就绪")
        self.batch_status_label.pack(side=tk.LEFT, padx=5)

        batch_btn_frame = ttk.Frame(batch_frame)
        batch_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(batch_btn_frame, text="添加文件", command=self.batch_add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(batch_btn_frame, text="移除选中", command=self.batch_remove_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(batch_btn_frame, text="清空列表", command=self.batch_clear_files).pack(side=tk.LEFT, padx=5)
        self.btn_batch_calc_export = ttk.Button(batch_btn_frame, text="批量计算并导出",
                                                command=self.batch_calculate_and_export)
        self.btn_batch_calc_export.pack(side=tk.LEFT, padx=5)

        self.add_watermark()
        self.toggle_inputs()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def toggle_inputs(self):
        if self.calc_type.get() == "restricted":
            self.input_frame.grid(row=self.input_frame_row, column=0, sticky="ew", padx=5, pady=5)
            self.policy_frame.grid(row=self.policy_frame_row, column=0, sticky="ew", padx=5, pady=5)
            self.result_frame.grid(row=self.result_frame_row, column=0, sticky="nsew", padx=5, pady=5)
            self.btn_frame.grid(row=self.btn_frame_row, column=0, pady=10)
        else:
            self.input_frame.grid_remove()
            self.policy_frame.grid_remove()
            self.result_frame.grid(row=self.result_frame_row, column=0, sticky="nsew", padx=5, pady=5)
            self.btn_frame.grid(row=self.btn_frame_row, column=0, pady=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.current_file_path = file_path
                self.df = pd.read_excel(file_path, parse_dates=['过户日期'], engine='openpyxl')
                required_cols = {'过户日期', '过户数量', '成交价格'}
                if not required_cols.issubset(self.df.columns):
                    missing = required_cols - set(self.df.columns)
                    raise ValueError("缺少必要列：{}".format(', '.join(missing)))
                messagebox.showinfo("成功", "文件加载成功！")
            except Exception as e:
                messagebox.showerror("错误", "文件加载失败：{}".format(str(e)))
                self.df = None
                self.current_file_path = None

    def calculate_tradable(self, df):
        result_cols = [
            '年度', '月份', '持仓数量',
            '单股成交价', '单笔卖出金额',
            '增值税-单股买入价', '增值税-单笔总买入价', '增值税-单笔减持收益',
            '所得税-单股成本', '所得税-单笔总成本', '所得税-单笔减持收益'
        ]
        df = df.copy()
        original_cols = df.columns.tolist()
        for col in result_cols:
            df[col] = np.nan

        vat_q, vat_avg = 0, 0.0
        tax_q, tax_avg = 0, 0.0

        for idx in df.index:
            row = df.loc[idx]
            qty = row['过户数量']
            price = row['成交价格']
            trans_date = row['过户日期']

            df.at[idx, '年度'] = trans_date.year
            df.at[idx, '月份'] = trans_date.month
            df.at[idx, '单股成交价'] = price
            if qty < 0:
                df.at[idx, '单笔卖出金额'] = -qty * price
            else:
                df.at[idx, '单笔卖出金额'] = 0.0

            if qty > 0:
                new_q = vat_q + qty
                vat_avg = (vat_q * vat_avg + qty * price) / new_q if new_q != 0 else 0
                vat_q = new_q
                df.at[idx, '增值税-单股买入价'] = vat_avg
                df.at[idx, '增值税-单笔总买入价'] = 0.0
                df.at[idx, '增值税-单笔减持收益'] = 0.0
            else:
                sell_qty = -qty
                df.at[idx, '增值税-单股买入价'] = vat_avg
                df.at[idx, '增值税-单笔总买入价'] = sell_qty * vat_avg
                df.at[idx, '增值税-单笔减持收益'] = (price - vat_avg) * sell_qty
                vat_q -= sell_qty

            if qty > 0:
                new_q = tax_q + qty
                tax_avg = (tax_q * tax_avg + qty * price) / new_q if new_q != 0 else 0
                tax_q = new_q
                df.at[idx, '所得税-单股成本'] = tax_avg
                df.at[idx, '所得税-单笔总成本'] = 0.0
                df.at[idx, '所得税-单笔减持收益'] = 0.0
            else:
                sell_qty = -qty
                df.at[idx, '所得税-单股成本'] = tax_avg
                df.at[idx, '所得税-单笔总成本'] = sell_qty * tax_avg
                df.at[idx, '所得税-单笔减持收益'] = (price - tax_avg) * sell_qty
                tax_q -= sell_qty

            df.at[idx, '持仓数量'] = max(vat_q, 0)

        final_cols = original_cols + result_cols
        return df[final_cols]

    def calculate_restricted(self, df, ipo_price, unlock_date):
        result_cols = [
            '年度', '月份', '持仓数量',
            '单股成交价', '单笔卖出金额',
            '增值税-单股买入价', '增值税-单笔总买入价', '增值税-单笔减持收益',
            '所得税-单股成本', '所得税-单笔总成本', '所得税-单笔减持收益'
        ]
        df = df.copy()
        original_cols = df.columns.tolist()
        for col in result_cols:
            df[col] = np.nan

        df['过户日期'] = pd.to_datetime(df['过户日期'])
        unlock_date = pd.to_datetime(unlock_date)

        vat_q, vat_avg = 0, 0.0
        tax_q, tax_avg = 0, 0.0

        for idx in df.index:
            row = df.loc[idx]
            qty = row['过户数量']
            price = row['成交价格']
            trans_date = row['过户日期']

            df.at[idx, '年度'] = trans_date.year
            df.at[idx, '月份'] = trans_date.month
            df.at[idx, '单股成交价'] = price
            if qty < 0:
                df.at[idx, '单笔卖出金额'] = -qty * price
            else:
                df.at[idx, '单笔卖出金额'] = 0.0

            if qty > 0:
                cost_price = ipo_price if trans_date <= unlock_date else price
                new_q = vat_q + qty
                vat_avg = (vat_q * vat_avg + qty * cost_price) / new_q if new_q != 0 else 0
                vat_q = new_q
                df.at[idx, '增值税-单股买入价'] = vat_avg
                df.at[idx, '增值税-单笔总买入价'] = 0.0
                df.at[idx, '增值税-单笔减持收益'] = 0.0
            else:
                sell_qty = -qty
                df.at[idx, '增值税-单股买入价'] = vat_avg
                df.at[idx, '增值税-单笔总买入价'] = sell_qty * vat_avg
                df.at[idx, '增值税-单笔减持收益'] = (price - vat_avg) * sell_qty
                vat_q -= sell_qty

            if qty > 0:
                new_q = tax_q + qty
                tax_avg = (tax_q * tax_avg + qty * price) / new_q if new_q != 0 else 0
                tax_q = new_q
                df.at[idx, '所得税-单股成本'] = tax_avg
                df.at[idx, '所得税-单笔总成本'] = 0.0
                df.at[idx, '所得税-单笔减持收益'] = 0.0
            else:
                sell_qty = -qty
                df.at[idx, '所得税-单股成本'] = tax_avg
                df.at[idx, '所得税-单笔总成本'] = sell_qty * tax_avg
                df.at[idx, '所得税-单笔减持收益'] = (price - tax_avg) * sell_qty
                tax_q -= sell_qty

            df.at[idx, '持仓数量'] = max(vat_q, 0)

        final_cols = original_cols + result_cols
        return df[final_cols]

    def calculate(self):
        if self.df is None:
            messagebox.showwarning("警告", "请先导入Excel文件")
            return

        try:
            if self.calc_type.get() == "restricted":
                ipo_price = float(self.ipo_entry.get())
                unlock_date = datetime.strptime(self.unlock_entry.get(), "%Y-%m-%d")
                self.result_df = self.calculate_restricted(self.df, ipo_price, unlock_date)
            else:
                self.result_df = self.calculate_tradable(self.df)

            self.show_results()
            messagebox.showinfo("成功", "计算完成！")
        except ValueError as e:
            messagebox.showerror("输入错误", "参数格式错误：{}".format(str(e)))
        except Exception as e:
            messagebox.showerror("计算错误", "发生错误：{}".format(str(e)))

    def show_results(self):
        self.tree.delete(*self.tree.get_children())
        if self.result_df is None:
            return
        columns = list(self.result_df.columns)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor=tk.CENTER)

        for idx in self.result_df.index:
            row = self.result_df.loc[idx]
            values = []
            for col in columns:
                val = row[col]
                if pd.isnull(val):
                    values.append("")
                elif isinstance(val, (float, np.float64)):
                    if col in ['增值税-单股买入价', '所得税-单股成本']:
                        values.append("{0:.4f}".format(val))
                    else:
                        values.append("{0:.2f}".format(val) if not val.is_integer() else str(int(val)))
                elif isinstance(val, datetime):
                    values.append(val.strftime("%Y-%m-%d"))
                else:
                    values.append(str(val))
            self.tree.insert("", "end", values=values)

    def parse_file_info(self, file_path):
        """从文件名解析信息：主管税务机关、公司名称、数据来源、市场类型、证券代码、证券简称"""
        if not file_path:
            return {
                'bureau': '未知',
                'company': '未知',
                'source': '未知',
                'market': '未知',
                'code': '未知',
                'name': '未知'
            }

        stem = Path(file_path).stem
        stem = stem.replace('（已匹配市场日均价）', '').replace(' - 计算结果', '')

        parts = stem.split('-')
        if len(parts) >= 6:
            return {
                'bureau': parts[0],
                'company': parts[1],
                'source': parts[2],
                'market': parts[3],
                'code': parts[4],
                'name': '-'.join(parts[5:])
            }
        elif len(parts) >= 2:
            return {
                'bureau': parts[0],
                'company': parts[1],
                'source': '未知',
                'market': '未知',
                'code': '未知',
                'name': '未知'
            }
        else:
            return {
                'bureau': '未知',
                'company': '未知',
                'source': '未知',
                'market': '未知',
                'code': '未知',
                'name': '未知'
            }

    def get_stock_info_from_df(self, df):
        """从DataFrame中提取证券代码和证券简称（所有唯一值）"""
        code = '未知'
        name = '未知'

        code_col = None
        for col in df.columns:
            if '证券代码' in str(col):
                code_col = col
                break

        name_col = None
        for col in df.columns:
            if '证券简称' in str(col):
                name_col = col
                break

        if code_col:
            codes = df[code_col].dropna().astype(str).unique()
            if len(codes) > 0:
                code = codes[0] if len(codes) == 1 else '-'.join(sorted(set(codes)))

        if name_col:
            names = df[name_col].dropna().astype(str).unique()
            if len(names) > 0:
                name = names[0] if len(names) == 1 else '-'.join(sorted(set(names)))

        return code, name

    # ========== 修改：汇总所有减持记录（过户数量 < 0） ==========
    def calculate_vat_summary(self, df, bureau, company, code='未知', name='未知'):
        """增值税汇总：统计所有减持记录，按月汇总，并计算应申报额（跨月结转负数，跨年清零）"""
        # [修改] 筛选过户数量 < 0（所有减持），而非单笔卖出金额 > 0
        sales_df = df[df['过户数量'] < 0].copy()
        if sales_df.empty:
            return pd.DataFrame(columns=['序号', '主管税务机关', '公司名称', '证券代码', '证券简称', '年度', '月份',
                                         '增值税-股票减持总收入', '增值税-股票减持总成本',
                                         '增值税-股票减持总收益', '增值税-当月应申报金融商品转让销售额（价税合计）'])

        monthly = sales_df.groupby(['年度', '月份']).agg({
            '单笔卖出金额': 'sum',
            '增值税-单笔总买入价': 'sum',
            '增值税-单笔减持收益': 'sum'
        }).reset_index()
        monthly.columns = ['年度', '月份', '增值税-股票减持总收入', '增值税-股票减持总成本', '增值税-股票减持总收益']

        monthly = monthly.sort_values(['年度', '月份']).reset_index(drop=True)

        result = []
        carry_over = 0.0
        current_year = None
        for _, row in monthly.iterrows():
            year = row['年度']
            month = row['月份']
            if current_year is None:
                current_year = year
            if year != current_year:
                carry_over = 0.0
                current_year = year
            monthly_gain = row['增值税-股票减持总收益']
            total = carry_over + monthly_gain
            if total > 0:
                taxable = total
                carry_over = 0.0
            else:
                taxable = 0.0
                carry_over = total
            result.append({
                '序号': len(result) + 1,
                '主管税务机关': bureau,
                '公司名称': company,
                '证券代码': code,
                '证券简称': name,
                '年度': year,
                '月份': month,
                '增值税-股票减持总收入': round(row['增值税-股票减持总收入'], 2),
                '增值税-股票减持总成本': round(row['增值税-股票减持总成本'], 2),
                '增值税-股票减持总收益': round(monthly_gain, 2),
                '增值税-当月应申报金融商品转让销售额（价税合计）': round(taxable, 2)
            })
        return pd.DataFrame(result)

    # [修改] 所得税汇总同样筛选所有减持记录
    def calculate_income_summary(self, df, bureau, company, code='未知', name='未知'):
        """所得税汇总：统计所有减持记录，按年汇总（不跨年结转）"""
        sales_df = df[df['过户数量'] < 0].copy()
        if sales_df.empty:
            return pd.DataFrame(columns=['序号', '主管税务机关', '公司名称', '证券代码', '证券简称', '年度',
                                         '所得税-股票减持总收入', '所得税-股票减持总成本', '所得税-股票减持总收益'])

        annual = sales_df.groupby('年度').agg({
            '单笔卖出金额': 'sum',
            '所得税-单笔总成本': 'sum',
            '所得税-单笔减持收益': 'sum'
        }).reset_index()
        annual.columns = ['年度', '所得税-股票减持总收入', '所得税-股票减持总成本', '所得税-股票减持总收益']

        annual = annual.sort_values('年度').reset_index(drop=True)
        annual['序号'] = range(1, len(annual) + 1)
        annual['主管税务机关'] = bureau
        annual['公司名称'] = company
        annual['证券代码'] = code
        annual['证券简称'] = name

        cols = ['序号', '主管税务机关', '公司名称', '证券代码', '证券简称', '年度',
                '所得税-股票减持总收入', '所得税-股票减持总成本', '所得税-股票减持总收益']
        annual = annual[cols]
        return annual

    def export_result(self):
        if self.result_df is None:
            messagebox.showwarning("警告", "请先进行计算")
            return

        if self.current_file_path:
            path_obj = Path(self.current_file_path)
            default_name = f"{path_obj.stem} - 计算结果.xlsx"
        else:
            default_name = "计算结果.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=default_name,
            title="保存计算结果"
        )
        if not save_path:
            return

        try:
            file_info = self.parse_file_info(self.current_file_path)
            bureau = file_info['bureau']
            company = file_info['company']

            code, name = self.get_stock_info_from_df(self.result_df)
            if code == '未知':
                code = file_info['code']
            if name == '未知':
                name = file_info['name']

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                self.result_df.to_excel(writer, sheet_name="交易明细", index=False)

                vat_summary = self.calculate_vat_summary(self.result_df, bureau, company, code, name)
                vat_summary.to_excel(writer, sheet_name="增值税汇总", index=False)

                income_summary = self.calculate_income_summary(self.result_df, bureau, company, code, name)
                income_summary.to_excel(writer, sheet_name="所得税汇总", index=False)

            messagebox.showinfo("成功", f"结果已保存至：\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")

    def batch_add_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        for f in files:
            if f not in self.batch_files:
                self.batch_files.append(f)
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def batch_remove_file(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.batch_files.pop(index)
            self.file_listbox.delete(index)

    def batch_clear_files(self):
        self.batch_files.clear()
        self.file_listbox.delete(0, tk.END)

    def select_batch_output_dir(self):
        path = filedialog.askdirectory(title="选择批量导出文件夹")
        if path:
            self.batch_output_dir.set(path)

    # [修改] 批量处理中收集汇总数据时同样筛选所有减持记录
    def batch_calculate_and_export(self):
        if not self.batch_files:
            messagebox.showwarning("警告", "请先添加文件")
            return

        if self.calc_type.get() == "restricted":
            messagebox.showwarning("警告", "限售股批量计算请先设置买入价和解禁日期")
            return

        output_dir = self.batch_output_dir.get()
        if not output_dir:
            messagebox.showwarning("警告", "请先选择导出文件夹")
            return

        self.batch_progress_var.set(0)
        self.batch_status_label.config(text="准备计算...")
        self.btn_batch_calc_export.config(state=tk.DISABLED)

        def calculate_and_export_in_thread():
            total = len(self.batch_files)
            success_count = 0

            all_detail_records = []
            all_vat_monthly = []
            all_income_annual = []

            for i, file_path in enumerate(self.batch_files):
                progress = (i / total) * 100
                self.after(0, lambda p=progress, idx=i, t=total: (
                    self.batch_progress_var.set(p),
                    self.batch_status_label.config(
                        text=f"正在计算 ({idx + 1}/{t}): {os.path.basename(file_path)[:30]}...")
                ))

                try:
                    df = pd.read_excel(file_path, parse_dates=['过户日期'], engine='openpyxl')
                    result_df = self.calculate_tradable(df)

                    file_info = self.parse_file_info(file_path)
                    code, name = self.get_stock_info_from_df(result_df)
                    if code == '未知':
                        code = file_info['code']
                    if name == '未知':
                        name = file_info['name']
                    file_info['code'] = code
                    file_info['name'] = name

                    bureau = file_info['bureau']
                    company = file_info['company']
                    source = file_info['source']
                    market = file_info['market']

                    path_obj = Path(file_path)
                    new_name = f"{path_obj.stem} - 计算结果.xlsx"
                    output_path = os.path.join(output_dir, new_name)

                    vat_summary = self.calculate_vat_summary(result_df, bureau, company, code, name)
                    income_summary = self.calculate_income_summary(result_df, bureau, company, code, name)

                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name="交易明细", index=False)
                        vat_summary.to_excel(writer, sheet_name="增值税汇总", index=False)
                        income_summary.to_excel(writer, sheet_name="所得税汇总", index=False)

                    # [修改] 收集明细数据时使用过户数量 < 0
                    sales_df = result_df[result_df['过户数量'] < 0].copy()
                    if not sales_df.empty:
                        monthly_data = sales_df.groupby(['年度', '月份']).agg({
                            '单笔卖出金额': 'sum',
                            '增值税-单笔总买入价': 'sum',
                            '增值税-单笔减持收益': 'sum',
                            '所得税-单笔总成本': 'sum',
                            '所得税-单笔减持收益': 'sum'
                        }).reset_index()

                        for _, row in monthly_data.iterrows():
                            all_detail_records.append({
                                '主管税务机关': bureau,
                                '公司名称': company,
                                '数据来源': source,
                                '市场类型': market,
                                '证券代码': code,
                                '证券简称': name,
                                '年度': row['年度'],
                                '月份': row['月份'],
                                '增值税-股票减持总收入': row['单笔卖出金额'],
                                '增值税-股票减持总成本': row['增值税-单笔总买入价'],
                                '增值税-股票减持总收益': row['增值税-单笔减持收益'],
                                '增值税-当月应申报金融商品转让销售额（价税合计）': 0,
                                '所得税-股票减持总收入': row['单笔卖出金额'],
                                '所得税-股票减持总成本': row['所得税-单笔总成本'],
                                '所得税-股票减持总收益': row['所得税-单笔减持收益']
                            })

                    if not vat_summary.empty:
                        all_vat_monthly.append(vat_summary)
                    if not income_summary.empty:
                        all_income_annual.append(income_summary)

                    success_count += 1

                except Exception as e:
                    print(f"处理文件 {file_path} 出错: {str(e)}")

            self.after(0, lambda: self.batch_status_label.config(text="正在生成汇总表..."))

            try:
                summary_path = os.path.join(output_dir, "汇总计算结果.xlsx")
                with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
                    if all_detail_records:
                        detail_df = pd.DataFrame(all_detail_records)
                        detail_df = detail_df.sort_values(['年度', '月份']).reset_index(drop=True)
                        detail_df = self._calculate_vat_taxable_for_detail(detail_df)
                        detail_df.insert(0, '序号', range(1, len(detail_df) + 1))

                        detail_cols = ['序号', '主管税务机关', '公司名称', '数据来源', '市场类型',
                                       '证券代码', '证券简称', '年度', '月份',
                                       '增值税-股票减持总收入', '增值税-股票减持总成本',
                                       '增值税-股票减持总收益', '增值税-当月应申报金融商品转让销售额（价税合计）',
                                       '所得税-股票减持总收入', '所得税-股票减持总成本', '所得税-股票减持总收益']
                        detail_df = detail_df[detail_cols]
                        detail_df.to_excel(writer, sheet_name="收益明细", index=False)

                    if all_vat_monthly:
                        combined_vat = pd.concat(all_vat_monthly, ignore_index=True)
                        vat_grouped = combined_vat.groupby(['主管税务机关', '公司名称', '年度', '月份']).agg({
                            '增值税-股票减持总收入': 'sum',
                            '增值税-股票减持总成本': 'sum',
                            '增值税-股票减持总收益': 'sum'
                        }).reset_index()

                        vat_grouped = vat_grouped.sort_values(['年度', '月份']).reset_index(drop=True)
                        vat_grouped = self._calculate_vat_taxable_for_group(vat_grouped)
                        vat_grouped.insert(0, '序号', range(1, len(vat_grouped) + 1))

                        vat_cols = ['序号', '主管税务机关', '公司名称', '年度', '月份',
                                    '增值税-股票减持总收入', '增值税-股票减持总成本',
                                    '增值税-股票减持总收益', '增值税-当月应申报金融商品转让销售额（价税合计）']
                        vat_grouped = vat_grouped[vat_cols]
                        vat_grouped.to_excel(writer, sheet_name="增值税汇总", index=False)

                    if all_income_annual:
                        combined_income = pd.concat(all_income_annual, ignore_index=True)
                        income_grouped = combined_income.groupby(['主管税务机关', '公司名称', '年度']).agg({
                            '所得税-股票减持总收入': 'sum',
                            '所得税-股票减持总成本': 'sum',
                            '所得税-股票减持总收益': 'sum'
                        }).reset_index()

                        income_grouped = income_grouped.sort_values('年度').reset_index(drop=True)
                        income_grouped.insert(0, '序号', range(1, len(income_grouped) + 1))

                        income_cols = ['序号', '主管税务机关', '公司名称', '年度',
                                       '所得税-股票减持总收入', '所得税-股票减持总成本', '所得税-股票减持总收益']
                        income_grouped = income_grouped[income_cols]
                        income_grouped.to_excel(writer, sheet_name="所得税汇总", index=False)

            except Exception as e:
                print(f"生成汇总表出错: {str(e)}")

            self.after(0, lambda: (
                self.batch_progress_var.set(100),
                self.batch_status_label.config(text=f"完成！成功 {success_count}/{total}"),
                self.btn_batch_calc_export.config(state=tk.NORMAL)
            ))

        threading.Thread(target=calculate_and_export_in_thread, daemon=True).start()

    # 以下两个辅助方法无需修改（它们处理已汇总的数据）
    def _calculate_vat_taxable_for_detail(self, df):
        """为收益明细计算增值税应申报额（按公司分别计算跨月结转）"""
        result_records = []

        for (bureau, company), group in df.groupby(['主管税务机关', '公司名称']):
            group = group.sort_values(['年度', '月份']).reset_index(drop=True)

            carry_over = 0.0
            current_year = None

            for _, row in group.iterrows():
                year = row['年度']
                month = row['月份']

                if current_year is None:
                    current_year = year
                if year != current_year:
                    carry_over = 0.0
                    current_year = year

                monthly_gain = row['增值税-股票减持总收益']
                total = carry_over + monthly_gain

                if total > 0:
                    taxable = total
                    carry_over = 0.0
                else:
                    taxable = 0.0
                    carry_over = total

                record = row.to_dict()
                record['增值税-当月应申报金融商品转让销售额（价税合计）'] = round(taxable, 2)
                result_records.append(record)

        return pd.DataFrame(result_records)

    def _calculate_vat_taxable_for_group(self, df):
        """为增值税汇总计算应申报额（按公司分别计算跨月结转）"""
        result_records = []

        for (bureau, company), group in df.groupby(['主管税务机关', '公司名称']):
            group = group.sort_values(['年度', '月份']).reset_index(drop=True)

            carry_over = 0.0
            current_year = None

            for _, row in group.iterrows():
                year = row['年度']
                month = row['月份']

                if current_year is None:
                    current_year = year
                if year != current_year:
                    carry_over = 0.0
                    current_year = year

                monthly_gain = row['增值税-股票减持总收益']
                total = carry_over + monthly_gain

                if total > 0:
                    taxable = total
                    carry_over = 0.0
                else:
                    taxable = 0.0
                    carry_over = total

                record = {
                    '主管税务机关': bureau,
                    '公司名称': company,
                    '年度': year,
                    '月份': month,
                    '增值税-股票减持总收入': round(row['增值税-股票减持总收入'], 2),
                    '增值税-股票减持总成本': round(row['增值税-股票减持总成本'], 2),
                    '增值税-股票减持总收益': round(monthly_gain, 2),
                    '增值税-当月应申报金融商品转让销售额（价税合计）': round(taxable, 2)
                }
                result_records.append(record)

        return pd.DataFrame(result_records)


# ===================== 风险识别模块（修正版） =====================
class RiskIdentificationSystem(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("风险识别系统")
        self.geometry("1400x800")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)
        # 初始化变量
        self.reduction_df = None  # 减持汇总数据
        self.declaration_df = None  # 申报数据
        self.risk_result_df = None  # 风险比对结果
        self.company_col = None  # 汇总数据中的公司列（持有人/实际权益拥有人）
        self.create_widgets()
        self.add_watermark()

    def add_watermark(self):
        self.watermark_label = tk.Label(
            self,
            text="国家税务总局湖北省税务局",
            font=("宋体", 24),
            fg="#cccccc",
            bg=self.cget("bg")
        )
        self.watermark_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 文件操作区域
        file_frame = ttk.LabelFrame(main_frame, text="文件导入")
        file_frame.pack(fill=tk.X, pady=5)
        self.btn_import_reduction = ttk.Button(file_frame, text="批量导入减持数据", command=self.import_reduction_data)
        self.btn_import_reduction.pack(side=tk.LEFT, padx=5, pady=5)
        self.btn_import_declaration = ttk.Button(file_frame, text="导入申报数据", command=self.import_declaration_data)
        self.btn_import_declaration.pack(side=tk.LEFT, padx=5, pady=5)

        # 进度条区域
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        ttk.Label(progress_frame, text="处理进度:").pack(side=tk.LEFT, padx=5)
        self.progress_value = tk.IntVar()
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal",
                                        length=800, mode="determinate",
                                        variable=self.progress_value)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志")
        log_frame.pack(fill=tk.X, pady=5)
        self.log_area = tk.Text(log_frame, height=5, width=120)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_area.yview)
        self.log_area.config(yscrollcommand=scrollbar.set)
        self.log_area.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="风险比对结果")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(tree_container, orient="vertical")
        hsb = ttk.Scrollbar(tree_container, orient="horizontal")
        self.tree = ttk.Treeview(
            tree_container,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode="extended",
            columns=[],
            show="headings"
        )
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill=tk.BOTH, expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # 操作按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        self.btn_compare = ttk.Button(btn_frame, text="开始比对", command=self.start_compare_thread)
        self.btn_compare.pack(side=tk.LEFT, padx=5)
        self.btn_export = ttk.Button(btn_frame, text="导出结果", command=self.export_result)
        self.btn_export.pack(side=tk.LEFT, padx=5)

        # 初始化日志
        self.log_area.insert(tk.END, "请先导入减持数据和申报数据，然后点击开始比对\n")

    def log(self, message):
        """添加日志"""
        self.log_area.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_area.see(tk.END)

    def update_progress(self, value):
        """更新进度条"""
        self.after(0, lambda: self.progress_value.set(value))

    def get_company_column(self, df):
        """自动识别汇总数据的公司列（持有人名称/实际权益拥有人名称）"""
        for col in ["实际权益拥有人名称", "持有人名称"]:
            if col in df.columns:
                return col
        raise ValueError("减持数据缺少公司列：需包含「持有人名称」或「实际权益拥有人名称」")

    def import_reduction_data(self):
        """批量导入减持数据（收益计算系统导出结果）"""
        file_paths = filedialog.askopenfilenames(
            title="选择减持数据Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if not file_paths:
            return

        try:
            self.update_progress(0)
            self.log(f"开始导入{len(file_paths)}个减持数据文件...")
            all_data = []
            for idx, file_path in enumerate(file_paths):
                df = pd.read_excel(file_path, engine='openpyxl')
                # 验证必要列（含公司列、过户日期、增值税销售额）
                required_cols = {'过户日期', '增值税_销售额', '所得税_投资收益'}
                temp_company_col = self.get_company_column(df)  # 临时确认公司列
                required_cols.add(temp_company_col)
                if not required_cols.issubset(df.columns):
                    missing = required_cols - set(df.columns)
                    raise ValueError(f"文件{os.path.basename(file_path)}缺少必要列：{', '.join(missing)}")
                all_data.append(df)
                self.update_progress(int((idx + 1) / len(file_paths) * 50))

            # 汇总数据并生成年度/月份列
            self.reduction_df = pd.concat(all_data, ignore_index=True)
            self.company_col = self.get_company_column(self.reduction_df)  # 确认最终公司列
            self.reduction_df['过户日期'] = pd.to_datetime(self.reduction_df['过户日期'], errors='coerce')
            self.reduction_df = self.reduction_df.dropna(subset=['过户日期'])
            self.reduction_df['年度'] = self.reduction_df['过户日期'].dt.year
            self.reduction_df['月份'] = self.reduction_df['过户日期'].dt.month

            self.update_progress(100)
            self.log(f"减持数据导入完成！共汇总{len(self.reduction_df)}条记录，公司列：{self.company_col}")
            messagebox.showinfo("成功", f"减持数据批量导入完成！\n公司匹配列：{self.company_col}")
        except Exception as e:
            self.log(f"减持数据导入失败：{str(e)}")
            messagebox.showerror("错误", f"减持数据导入失败：{str(e)}")

    def import_declaration_data(self):
        """导入申报数据（含纳税人名称字段）"""
        file_path = filedialog.askopenfilename(
            title="选择申报数据Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            self.log("开始导入申报数据...")
            self.declaration_df = pd.read_excel(file_path, engine='openpyxl')
            # 验证必要字段（含纳税人名称）
            required_fields = [
                '纳税人名称',  # 用于匹配减持数据的公司列
                '是否为一般纳税人',
                '《增值税纳税申报表》（一般纳税人适用）附列资料（三）_6%税率的金融商品转让项目_本期服务、不动产和无形资产价税合计额（免税销售额）',
                '《增值税纳税申报表》（小规模纳税人适用）附列资料_全部含税收入（适用3%征收率）',
                '《增值税纳税申报表》（一般纳税人适用）附列资料（三）_6%税率的金融商品转让项目_服务、不动产和无形资产扣除项目_本期发生额',
                '《增值税纳税申报表》（小规模纳税人适用）附列资料_应税行为（3%征收率）扣除额计算_本期发生额',
                '年度',
                '月份',
                '季度'
            ]
            # 检查所有必要字段是否存在
            missing_fields = [f for f in required_fields if f not in self.declaration_df.columns]
            if missing_fields:
                raise ValueError(f"申报数据缺少必要字段：{', '.join(missing_fields)}")

            # 数据类型转换（数值字段处理为float，空值填0）
            numeric_fields = required_fields[2:6]  # 4个申报数值字段
            for field in numeric_fields:
                self.declaration_df[field] = pd.to_numeric(self.declaration_df[field], errors='coerce').fillna(0.0)

            self.log("申报数据导入完成！匹配字段：「纳税人名称」")
            messagebox.showinfo("成功",
                                f"申报数据导入完成！\n匹配关系：申报数据「纳税人名称」 ↔ 减持数据「{self.company_col}」")
        except Exception as e:
            self.log(f"申报数据导入失败：{str(e)}")
            messagebox.showerror("错误", f"申报数据导入失败：{str(e)}")

    def calculate_taxable_sales_general(self, group_df):
        """一般纳税人应申报额计算：每年重新计算，不跨年结转"""
        group_df = group_df.sort_values(['年度', '月份']).reset_index(drop=True)
        result = []

        # 按年度分组，每年独立计算
        for year, year_group in group_df.groupby('年度'):
            carry_over = 0.0  # 每年重新开始，结转余额清零

            for _, row in year_group.iterrows():
                taxpayer = row[self.company_col]
                month = row['月份']
                current_sales = row['增值税_销售额']
                total = carry_over + current_sales  # 当月汇总值（结转+当月）

                # 12月特殊处理：负余额不结转下年
                is_december = (month == 12)
                if total > 0:
                    taxable_sales = total
                    new_carry_over = 0.0  # 汇总为正，次月重置
                else:
                    taxable_sales = 0.0
                    new_carry_over = 0.0 if is_december else total  # 12月不结转

                result.append({
                    '纳税人': taxpayer,
                    '纳税人类型': '一般纳税人',
                    '年度': year,
                    '月份': month,
                    '季度': '',
                    '当月增值税销售额': round(current_sales, 2),
                    '结转余额': round(carry_over, 2),
                    '当月汇总值': round(total, 2),
                    '应申报增值税销售额': round(taxable_sales, 2)
                })

                carry_over = new_carry_over  # 更新结转余额

        return pd.DataFrame(result)

    def calculate_taxable_sales_small(self, group_df):
        """小规模纳税人应申报额计算：每年重新计算，不跨年结转"""
        # 步骤1：按年度和季度汇总销售额
        group_df['季度'] = group_df['过户日期'].dt.quarter
        result = []

        # 按年度分组，每年独立计算
        for year, year_group in group_df.groupby('年度'):
            quarter_carry_over = 0.0  # 每年重新开始，季度结转余额清零

            # 按季度汇总当季销售额
            quarterly_summary = year_group.groupby([self.company_col, '季度']).agg({
                '增值税_销售额': 'sum',  # 当季总销售额
                '所得税_投资收益': 'sum'
            }).reset_index()

            quarterly_summary = quarterly_summary.sort_values('季度').reset_index(drop=True)

            # 按季度计算应申报额
            for _, row in quarterly_summary.iterrows():
                taxpayer = row[self.company_col]
                quarter = row['季度']
                quarterly_sales = row['增值税_销售额']
                total = quarter_carry_over + quarterly_sales  # 当季汇总值（结转+当季）

                # 第四季度特殊处理：负余额不结转下年
                is_q4 = (quarter == 4)
                if total > 0:
                    taxable_sales = total
                    new_carry_over = 0.0  # 汇总为正，下季度重置
                else:
                    taxable_sales = 0.0
                    new_carry_over = 0.0 if is_q4 else total  # Q4不结转

                result.append({
                    '纳税人': taxpayer,
                    '纳税人类型': '小规模纳税人',
                    '年度': year,
                    '月份': '',
                    '季度': quarter,
                    '当季增值税销售额': round(quarterly_sales, 2),
                    '季度结转余额': round(quarter_carry_over, 2),
                    '当季汇总值': round(total, 2),
                    '应申报增值税销售额': round(taxable_sales, 2)
                })

                quarter_carry_over = new_carry_over  # 更新季度结转余额

        return pd.DataFrame(result)

    def compare_data(self):
        """执行数据比对（按纳税人名称匹配）"""
        try:
            self.update_progress(20)
            self.log("开始数据透视和应申报额计算...")

            # 步骤1：减持数据按「公司列-年度-月份」分组
            pivot_df = self.reduction_df.groupby([self.company_col, '年度', '月份']).agg({
                '增值税_销售额': 'sum',
                '所得税_投资收益': 'sum',
                '过户日期': 'first'  # 用于小规模季度计算
            }).reset_index()

            self.update_progress(40)
            self.log("开始按纳税人类型计算应申报额...")

            # 步骤2：按纳税人类型拆分计算
            all_taxable_results = []
            all_taxpayers = pivot_df[self.company_col].unique()  # 所有减持数据中的纳税人

            for taxpayer in all_taxpayers:
                # 用申报数据的「纳税人名称」匹配减持数据的「公司列」
                taxpayer_decl = self.declaration_df[self.declaration_df['纳税人名称'] == taxpayer]
                if taxpayer_decl.empty:
                    self.log(f"警告：纳税人「{taxpayer}」未找到对应申报数据，跳过比对")
                    continue

                # 判断纳税人类型
                tax_type_flag = str(taxpayer_decl.iloc[0]['是否为一般纳税人']).strip().upper()
                is_general = tax_type_flag in ['是', 'Y']

                # 筛选该纳税人的减持数据
                taxpayer_reduction = pivot_df[pivot_df[self.company_col] == taxpayer].copy()

                # 按类型计算应申报额
                if is_general:
                    taxable_df = self.calculate_taxable_sales_general(taxpayer_reduction)
                else:
                    taxable_df = self.calculate_taxable_sales_small(taxpayer_reduction)

                all_taxable_results.append(taxable_df)

            if not all_taxable_results:
                raise ValueError("无有效比对数据（所有纳税人未找到申报数据）")

            # 合并所有纳税人的应申报额结果
            taxable_combined = pd.concat(all_taxable_results, ignore_index=True)
            self.update_progress(60)
            self.log(f"应申报额计算完成，共{len(taxable_combined)}条记录，开始与申报数据比对...")

            # 步骤3：与申报数据比对，生成风险结果
            risk_results = []
            for _, taxable_row in taxable_combined.iterrows():
                taxpayer = taxable_row['纳税人']
                tax_type = taxable_row['纳税人类型']
                year = taxable_row['年度']
                month = taxable_row['月份']
                quarter = taxable_row['季度']
                taxable_sales = taxable_row['应申报增值税销售额']

                # 按类型筛选对应申报数据
                if tax_type == '一般纳税人':
                    # 一般纳税人：逐月比对
                    decl_filter = (
                            (self.declaration_df['纳税人名称'] == taxpayer) &
                            (self.declaration_df['年度'] == year) &
                            (self.declaration_df['月份'] == month)
                    )
                    decl_cols = [
                        '《增值税纳税申报表》（一般纳税人适用）附列资料（三）_6%税率的金融商品转让项目_本期服务、不动产和无形资产价税合计额（免税销售额）',
                        '《增值税纳税申报表》（一般纳税人适用）附列资料（三）_6%税率的金融商品转让项目_服务、不动产和无形资产扣除项目_本期发生额'
                    ]
                else:
                    # 小规模纳税人：逐季比对
                    decl_filter = (
                            (self.declaration_df['纳税人名称'] == taxpayer) &
                            (self.declaration_df['年度'] == year) &
                            (self.declaration_df['季度'] == quarter)
                    )
                    decl_cols = [
                        '《增值税纳税申报表》（小规模纳税人适用）附列资料_全部含税收入（适用3%征收率）',
                        '《增值税纳税申报表》（小规模纳税人适用）附列资料_应税行为（3%征收率）扣除额计算_本期发生额'
                    ]

                # 获取申报数据（无申报数据则视为0）
                decl_data = self.declaration_df[decl_filter]
                if decl_data.empty:
                    declare_sales = 0.0
                    declare_deduction = 0.0
                else:
                    declare_sales = decl_data.iloc[0][decl_cols[0]]
                    declare_deduction = decl_data.iloc[0][decl_cols[1]]

                # 计算申报收益、少申报额、少交税额
                declare_income = declare_sales - declare_deduction  # 申报收益
                under_declare = max(taxable_sales - declare_income, 0.0)  # 少申报额（非负）
                # 计算少交税额（按不同税率）
                if tax_type == '一般纳税人':
                    under_tax = under_declare / 1.06 * 0.06 if under_declare > 0 else 0.0
                else:
                    under_tax = under_declare / 1.03 * 0.03 if under_declare > 0 else 0.0
                # 风险标记
                risk_flag = '是' if under_declare > 0 else '否'

                # 记录风险结果
                risk_results.append({
                    '纳税人': taxpayer,
                    '纳税人类型': tax_type,
                    '年度': year,
                    '月份': month,
                    '季度': quarter,
                    '应申报增值税销售额': round(taxable_sales, 2),
                    '申报金融商品转让增值税收益': round(declare_income, 2),
                    '少申报增值税销售额': round(under_declare, 2),
                    '测算少交增值税额': round(under_tax, 2),
                    '风险标记': risk_flag
                })

            # 生成最终风险结果表
            self.risk_result_df = pd.DataFrame(risk_results)
            self.update_progress(80)
            self.show_results()  # 显示结果到表格
            self.update_progress(100)

            # 统计风险数量并提示
            risk_count = len(self.risk_result_df[self.risk_result_df['风险标记'] == '是'])
            self.log(f"风险比对完成！共{len(self.risk_result_df)}条记录，其中风险记录{risk_count}条")
            messagebox.showinfo("成功", f"风险比对完成！\n总记录数：{len(self.risk_result_df)}\n风险记录数：{risk_count}")

        except Exception as e:
            self.log(f"比对失败：{str(e)}")
            messagebox.showerror("错误", f"比对失败：{str(e)}")

    def show_results(self):
        """显示比对结果到表格"""
        self.tree.delete(*self.tree.get_children())
        # 定义结果表格字段
        columns = [
            '纳税人', '纳税人类型', '年度', '月份', '季度',
            '应申报增值税销售额', '申报金融商品转让增值税收益',
            '少申报增值税销售额', '测算少交增值税额', '风险标记'
        ]
        self.tree.config(columns=columns)

        # 设置表头与列宽
        for col in columns:
            self.tree.heading(col, text=col)
            width = 180 if col == '纳税人' else 120
            self.tree.column(col, width=width, anchor=tk.CENTER)

        # 插入结果数据
        for _, row in self.risk_result_df.iterrows():
            values = [
                row['纳税人'], row['纳税人类型'], str(row['年度']),
                str(row['月份']) if row['月份'] != '' else '',
                str(row['季度']) if row['季度'] != '' else '',
                f"{row['应申报增值税销售额']:.2f}",
                f"{row['申报金融商品转让增值税收益']:.2f}",
                f"{row['少申报增值税销售额']:.2f}",
                f"{row['测算少交增值税额']:.2f}",
                row['风险标记']
            ]
            self.tree.insert("", "end", values=values)

    def export_result(self):
        """导出风险结果"""
        if self.risk_result_df is None:
            messagebox.showwarning("警告", "请先进行比对")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            title="保存风险比对结果"
        )
        if not save_path:
            return

        try:
            self.log(f"开始导出结果到：{save_path}")
            self.risk_result_df.to_excel(save_path, index=False, engine='openpyxl')
            self.log("结果导出完成！")
            messagebox.showinfo("成功", f"结果已保存至：\n{save_path}")
        except Exception as e:
            self.log(f"导出失败：{str(e)}")
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def start_compare_thread(self):
        """启动比对线程"""
        if self.reduction_df is None:
            messagebox.showwarning("警告", "请先导入减持数据")
            return
        if self.declaration_df is None:
            messagebox.showwarning("警告", "请先导入申报数据")
            return

        self.progress_value.set(0)
        self.log("开始风险比对...")
        threading.Thread(target=self.compare_data, daemon=True).start()


# ===================== 程序入口 =====================
if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()