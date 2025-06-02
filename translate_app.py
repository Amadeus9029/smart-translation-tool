import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
from pathlib import Path
import pandas as pd
import threading
import logging
import queue
import deepl_selenium_translate
from datetime import datetime
import time
from openai import OpenAI
from PyQt5.QtCore import pyqtSignal, QObject
from constants import SUPPORTED_LANGUAGES  # 从constants导入
from text_translate import TextTranslateFrame  # 导入TextTranslateFrame
from subtitle_translate import SubtitleTranslateFrame  # 添加这行导入
from subtitle_result import SubtitleResultFrame  # 添加导入

class LightTheme:
    """明亮主题样式"""
    BACKGROUND = "#ffffff"      # 主背景色
    FOREGROUND = "#333333"      # 主文字色
    SIDEBAR_BG = "#f0f0f0"      # 侧边栏背景色
    SIDEBAR_FG = "#333333"      # 侧边栏文字色
    SIDEBAR_HOVER = "#e0e0e0"   # 侧边栏悬停色
    SIDEBAR_SELECT = "#e0e0e0"  # 侧边栏选中色
    CONTENT_BG = "#ffffff"      # 内容区背景色
    BUTTON_BG = "#0d6efd"       # 按钮背景色
    BUTTON_ACTIVE = "#0b5ed7"   # 按钮激活色
    ENTRY_BG = "#ffffff"        # 输入框背景色
    PROGRESS_BG = "#e9ecef"     # 进度条背景色
    PROGRESS_BAR = "#0d6efd"    # 进度条颜色
    FRAME_BG = "#f8f9fa"        # 框架背景色

class DarkTheme:
    """深色主题样式"""
    BACKGROUND = "#1e1e1e"      # 主背景色
    FOREGROUND = "#ffffff"      # 主文字色
    SIDEBAR_BG = "#252526"      # 侧边栏背景色
    SIDEBAR_FG = "#ffffff"      # 侧边栏文字色
    SIDEBAR_HOVER = "#37373d"   # 侧边栏悬停色
    SIDEBAR_SELECT = "#37373d"  # 侧边栏选中色
    CONTENT_BG = "#1e1e1e"      # 内容区背景色
    BUTTON_BG = "#0d6efd"       # 按钮背景色
    BUTTON_ACTIVE = "#0b5ed7"   # 按钮激活色
    ENTRY_BG = "#2d2d2d"        # 输入框背景色
    PROGRESS_BG = "#3d3d3d"     # 进度条背景色
    PROGRESS_BAR = "#0d6efd"    # 进度条颜色
    FRAME_BG = "#252526"        # 框架背景色

class TranslateApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("智能翻译工具")
        self.root.geometry("1200x800")

        # 添加 DeepSeek API 配置
        self.DEEPSEEK_BASE_URL = "https://api.deepseek.com"
        
        # 加载配置
        self.config = self.load_config()
        
        # 设置当前主题
        self.current_theme = self.config.get("theme", "light")
        self.theme = LightTheme if self.current_theme == "light" else DarkTheme
        
        # 配置窗口背景色
        self.root.configure(bg=self.theme.BACKGROUND)
        
        # 初始化变量
        self.file_path = tk.StringVar()
        self.save_path = tk.StringVar(value=self.config.get("save_path", ""))
        self.source_lang = tk.StringVar(value=self.config.get("source_lang", "英语"))
        self.api_key = tk.StringVar(value=self.config.get("api_key", ""))
        self.progress_var = tk.DoubleVar()
        self.progress_percent = tk.StringVar(value="0%")
        self.remaining_time = tk.StringVar(value="预计剩余时间: --:--")
        self.progress_detail = tk.StringVar(value="等待开始翻译...")
        self.theme_var = tk.StringVar(value=self.current_theme)
        
        # 参考相关变量
        self.ref_mode = tk.StringVar(value="none")  # 默认不使用参考源
        self.ref_file_path = tk.StringVar()
        self.ref_lang = tk.StringVar()  # 外部参考语言
        self.internal_ref_lang = tk.StringVar()  # 内置参考语言
        
        self.translation_start_time = None
        self.cancel_flag = False
        self.translation_thread = None
        self.is_translating = False
        self.last_update_time = 0
        
        # 创建消息队列
        self.message_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        
        # 配置现代化样式
        self.setup_styles()
        
        # 创建日志处理器
        self.setup_logger()
        
        # 创建主布局
        self.create_main_layout()
        
        # 启动消息处理
        self.process_messages()
        
        # 定期刷新文件列表
        self.refresh_file_list()
        
        # 加载历史日志
        self.load_history_log()
        
    def setup_styles(self):
        """配置自定义样式"""
        style = ttk.Style()
        
        # 侧边栏按钮样式
        style.configure("Sidebar.TButton",
                       background=self.theme.SIDEBAR_BG,
                       foreground=self.theme.SIDEBAR_FG,
                       font=("微软雅黑", 10),
                       padding=(20, 10),
                       relief="flat",
                       anchor="w")
                       
        style.map("Sidebar.TButton",
                 background=[("active", self.theme.SIDEBAR_HOVER),
                           ("selected", self.theme.SIDEBAR_SELECT)])
        
        # 主框架样式
        style.configure("Modern.TFrame", 
                       background=self.theme.BACKGROUND)
        
        # 标签框架样式
        style.configure("Modern.TLabelframe", 
                       background=self.theme.FRAME_BG,
                       foreground=self.theme.FOREGROUND)
        style.configure("Modern.TLabelframe.Label", 
                       background=self.theme.BACKGROUND,
                       foreground=self.theme.FOREGROUND,
                       font=("微软雅黑", 10))
        
        # 按钮样式
        style.configure("Modern.TButton",
                       background=self.theme.BUTTON_BG,
                       foreground="#ffffff",
                       padding=(10, 5),
                       font=("微软雅黑", 9))
        
        # 标签样式
        style.configure("Modern.TLabel",
                       background=self.theme.FRAME_BG,
                       foreground=self.theme.FOREGROUND,
                       font=("微软雅黑", 9))
        
        # 进度条样式
        style.configure("Modern.Horizontal.TProgressbar",
                       background=self.theme.PROGRESS_BAR,
                       troughcolor=self.theme.PROGRESS_BG)
                       
        # 输入框样式
        style.configure("Modern.TEntry",
                       fieldbackground=self.theme.ENTRY_BG,
                       foreground=self.theme.FOREGROUND)

    def create_main_layout(self):
        """创建主布局"""
        # 创建左侧边栏
        self.sidebar = ttk.Frame(self.root, style="Modern.TFrame")
        self.sidebar.pack(side="left", fill="y", padx=0, pady=0)
        self.sidebar.configure(width=200)
        
        # 设置侧边栏背景色
        sidebar_bg = ttk.Label(self.sidebar, background=self.theme.SIDEBAR_BG)
        sidebar_bg.place(relwidth=1, relheight=1)
        
        # 创建侧边栏按钮
        self.current_page = tk.StringVar(value="doc_translate")  # 修改默认页面
        
        # 文本翻译按钮
        self.text_translate_btn = ttk.Button(self.sidebar, text="文本翻译",
                                 style="Sidebar.TButton",
                                 command=lambda: self.show_page("text_translate"))
        self.text_translate_btn.pack(fill="x", pady=(0, 1))
        
        # 文档翻译按钮（原首页）
        self.doc_translate_btn = ttk.Button(self.sidebar, text="文档翻译",
                                style="Sidebar.TButton",
                                command=lambda: self.show_page("doc_translate"))
        self.doc_translate_btn.pack(fill="x", pady=(0, 1))
        
        # 视频字幕翻译按钮
        self.subtitle_translate_btn = ttk.Button(self.sidebar, text="视频字幕翻译",
                                  style="Sidebar.TButton",
                                  command=lambda: self.show_page("subtitle_translate"))
        self.subtitle_translate_btn.pack(fill="x", pady=(0, 1))
        
        # 字幕结果按钮
        self.subtitle_result_btn = ttk.Button(self.sidebar, text="字幕翻译结果",
                                  style="Sidebar.TButton",
                                  command=lambda: self.show_page("subtitle_result"))
        self.subtitle_result_btn.pack(fill="x", pady=(0, 1))
        
        # 结果列表按钮
        self.results_btn = ttk.Button(self.sidebar, text="翻译结果",
                                   style="Sidebar.TButton",
                                   command=lambda: self.show_page("results"))
        self.results_btn.pack(fill="x", pady=(0, 1))
        
        # 日志按钮
        self.log_btn = ttk.Button(self.sidebar, text="日志",
                               style="Sidebar.TButton",
                               command=lambda: self.show_page("log"))
        self.log_btn.pack(fill="x", pady=(0, 1))
        
        # 设置按钮
        self.settings_btn = ttk.Button(self.sidebar, text="设置",
                                    style="Sidebar.TButton",
                                    command=lambda: self.show_page("settings"))
        self.settings_btn.pack(fill="x", pady=(0, 1))
        
        # 创建右侧内容区
        self.content = ttk.Frame(self.root, style="Modern.TFrame")
        self.content.pack(side="right", fill="both", expand=True)
        
        # 创建各个页面
        self.pages = {}
        self.create_text_translate_page()  # 添加文本翻译页面
        self.create_home_page()  # 文档翻译页面
        self.create_results_page()
        self.create_log_page()
        self.create_settings_page()
        self.create_subtitle_translate_page()  # 添加字幕翻译页面
        self.create_subtitle_result_page()  # 添加字幕结果页面
        
        # 修改页面字典的键名
        self.pages["doc_translate"] = self.pages.pop("home")
        
        # 默认显示文档翻译页面
        # self.show_page("doc_translate")
        # 默认显示文本翻译页面
        self.show_page("text_translate")
        
    def show_page(self, page_name):
        """显示指定页面"""
        # 隐藏所有页面
        for page in self.pages.values():
            page.pack_forget()
            
        # 重置所有按钮状态
        self.text_translate_btn.configure(style="Sidebar.TButton")
        self.doc_translate_btn.configure(style="Sidebar.TButton")
        self.subtitle_translate_btn.configure(style="Sidebar.TButton")
        self.subtitle_result_btn.configure(style="Sidebar.TButton")
        self.results_btn.configure(style="Sidebar.TButton")
        self.log_btn.configure(style="Sidebar.TButton")
        self.settings_btn.configure(style="Sidebar.TButton")
        
        # 显示选中页面并高亮对应按钮
        if page_name == "subtitle_result":
            self.subtitle_result_btn.configure(style="Sidebar.TButton")
            # 切换到结果页面时自动刷新列表
            self.pages[page_name].refresh_list()
        elif page_name == "text_translate":
            self.text_translate_btn.configure(style="Sidebar.TButton")
        elif page_name == "doc_translate":
            self.doc_translate_btn.configure(style="Sidebar.TButton")
        elif page_name == "subtitle_translate":
            self.subtitle_translate_btn.configure(style="Sidebar.TButton")
        elif page_name == "results":
            self.results_btn.configure(style="Sidebar.TButton")
        elif page_name == "log":
            self.log_btn.configure(style="Sidebar.TButton")
        elif page_name == "settings":
            self.settings_btn.configure(style="Sidebar.TButton")
        
        # 显示选中的页面
        self.pages[page_name].pack(fill="both", expand=True)
        
    def create_text_translate_page(self):
        """创建文本翻译页面"""
        text_translate_frame = TextTranslateFrame(self.content, self.theme, self.api_key.get())
        self.pages["text_translate"] = text_translate_frame
        
    def create_home_page(self):
        """创建首页"""
        home_page = ttk.Frame(self.content, style="Modern.TFrame")
        self.pages["home"] = home_page
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(home_page, text="文件设置", 
                                  style="Modern.TLabelframe",
                                  padding=15)
        file_frame.pack(fill="x", pady=(0, 10))
        
        # Excel文件选择
        ttk.Label(file_frame, text="Excel文件:", 
                 style="Modern.TLabel").grid(row=0, column=0, sticky="w")
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        file_entry.grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="选择文件", 
                  style="Modern.TButton",
                  command=self.select_file).grid(row=0, column=2, padx=5)
                  
        # 翻译参考设置区域
        ref_frame = ttk.LabelFrame(home_page, text="翻译参考设置", 
                                 style="Modern.TLabelframe",
                                 padding=15)
        ref_frame.pack(fill="x", pady=(0, 10))
        
        # 参考模式选择
        mode_frame = ttk.Frame(ref_frame)
        mode_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Radiobutton(mode_frame, text="不使用参考源",
                       variable=self.ref_mode,
                       value="none",
                       command=self.toggle_ref_mode).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="使用内置参考源",
                       variable=self.ref_mode,
                       value="internal",
                       command=self.toggle_ref_mode).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="使用外置参考源",
                       variable=self.ref_mode,
                       value="external",
                       command=self.toggle_ref_mode).pack(side="left", padx=5)
        
        # 内置参考源设置
        self.internal_ref_frame = ttk.Frame(ref_frame)
        # 初始不显示内置参考设置
        
        ttk.Label(self.internal_ref_frame, text="参考语言:",
                 style="Modern.TLabel").pack(side="left")
        internal_ref_combo = ttk.Combobox(self.internal_ref_frame,
                                      textvariable=self.internal_ref_lang,
                                      values=list(SUPPORTED_LANGUAGES.keys()),
                                      state="readonly",
                                      width=20)
        internal_ref_combo.pack(side="left", padx=5)
                       
        # 外部参考文件设置
        self.ref_file_frame = ttk.Frame(ref_frame)
        # 初始不显示外部参考设置
        
        # 参考文件选择
        ref_file_frame = ttk.Frame(self.ref_file_frame)
        ref_file_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(ref_file_frame, text="参考文件:",
                 style="Modern.TLabel").pack(side="left")
        ttk.Entry(ref_file_frame,
                 textvariable=self.ref_file_path,
                 width=50).pack(side="left", padx=5)
        ttk.Button(ref_file_frame, text="选择文件",
                  style="Modern.TButton",
                  command=self.select_ref_file).pack(side="left")
                  
        # 参考语言选择
        ref_lang_frame = ttk.Frame(self.ref_file_frame)
        ref_lang_frame.pack(fill="x")
        
        ttk.Label(ref_lang_frame, text="参考语言:",
                 style="Modern.TLabel").pack(side="left")
        ref_lang_combo = ttk.Combobox(ref_lang_frame,
                                    textvariable=self.ref_lang,
                                    values=list(SUPPORTED_LANGUAGES.keys()),
                                    state="readonly",
                                    width=20)
        ref_lang_combo.pack(side="left", padx=5)
        
        # 初始隐藏外部参考设置
        self.ref_file_frame.pack_forget()
        
        # 语言设置区域
        lang_frame = ttk.LabelFrame(home_page, text="语言设置", 
                                  style="Modern.TLabelframe",
                                  padding=15)
        lang_frame.pack(fill="x", pady=10)
        
        # 源语言选择
        ttk.Label(lang_frame, text="源语言:", 
                 style="Modern.TLabel").grid(row=0, column=0, sticky="w")
        source_combo = ttk.Combobox(lang_frame, 
                                  textvariable=self.source_lang,
                                  values=list(SUPPORTED_LANGUAGES.keys()),
                                  state="readonly",
                                  width=20)
        source_combo.grid(row=0, column=1, padx=5, sticky="w")
        
        # 目标语言选择框架
        target_frame = ttk.Frame(lang_frame)
        target_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        # 创建目标语言复选框
        self.target_langs_vars = {}
        target_checkbuttons = {}
        for i, lang in enumerate(SUPPORTED_LANGUAGES.keys()):
            var = tk.BooleanVar(value=False)
            self.target_langs_vars[lang] = var
            cb = ttk.Checkbutton(target_frame, text=lang, variable=var)
            cb.grid(row=i//12, column=i%12, padx=5, pady=2, sticky="we")
            target_checkbuttons[lang] = cb
            
        # 翻译进度框架（初始隐藏）
        self.progress_frame = ttk.LabelFrame(home_page, text="翻译进度", 
                                         style="Modern.TLabelframe",
                                         padding=15)
        
        # 进度信息框架
        info_frame = ttk.Frame(self.progress_frame)
        info_frame.pack(fill="x", pady=(0, 5))
        
        # 进度百分比标签
        percent_label = ttk.Label(info_frame, 
                                textvariable=self.progress_percent,
                                style="Modern.TLabel")
        percent_label.pack(side="left")
        
        # 剩余时间标签
        time_label = ttk.Label(info_frame, 
                             textvariable=self.remaining_time,
                             style="Modern.TLabel")
        time_label.pack(side="right")
        
        # 进度条
        self.progress_bar = ttk.Progressbar(self.progress_frame, 
                                          style="Modern.Horizontal.TProgressbar",
                                          variable=self.progress_var,
                                          maximum=100)
        self.progress_bar.pack(fill="x")
        
        # 详细进度标签
        detail_label = ttk.Label(self.progress_frame, 
                               textvariable=self.progress_detail,
                               style="Modern.TLabel")
        detail_label.pack(fill="x", pady=(5, 0))
        
        # 操作按钮
        btn_frame = ttk.Frame(home_page, style="Modern.TFrame")
        btn_frame.pack(fill="x", pady=10)
        
        # 创建翻译按钮
        self.translate_btn = ttk.Button(btn_frame, text="开始翻译",
                                      style="Modern.TButton",
                                      command=self.start_translation)
        self.translate_btn.pack(side="left", padx=5)
        
        # 创建取消按钮
        self.cancel_btn = ttk.Button(btn_frame, text="取消翻译",
                                   style="Modern.TButton",
                                   command=self.cancel_translation)
        
        # 运行日志区域
        self.home_log_frame = ttk.LabelFrame(home_page, text="运行日志", 
                                         style="Modern.TLabelframe",
                                         padding=15)
        self.home_log_frame.pack(fill="both", expand=True, pady=10)
        
        # 创建日志文本框
        self.home_log_text = tk.Text(self.home_log_frame, height=8,
                                 bg=self.theme.ENTRY_BG,
                                 fg=self.theme.FOREGROUND,
                                 insertbackground=self.theme.FOREGROUND,
                                 relief="flat",
                                 font=("微软雅黑", 9))
        self.home_log_text.pack(fill="both", expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.home_log_frame, orient="vertical", 
                                command=self.home_log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.home_log_text.configure(yscrollcommand=scrollbar.set)
        
        # 初始状态隐藏进度框
        self.show_progress_frame(False)
        
    def show_progress_frame(self, show=True):
        """显示或隐藏进度框"""
        if show:
            # 在按钮上方显示进度框
            self.progress_frame.pack(fill="x", pady=10, before=self.translate_btn.master)
            self.cancel_btn.pack(side="left", padx=5)
        else:
            self.progress_frame.pack_forget()
            self.cancel_btn.pack_forget()
            
    def cancel_translation(self):
        """取消翻译"""
        if self.is_translating and hasattr(self, 'translation_thread') and self.translation_thread.is_alive():
            self.cancel_flag = True
            self.is_translating = False
            self.message_queue.put(('log', "正在取消翻译..."))
            self.progress_detail.set("正在取消翻译...")
            self.translate_btn["state"] = "disabled"
            
            # 设置翻译取消状态
            deepl_selenium_translate.set_translation_cancelled(True)
            
            # 等待线程结束
            if self.translation_thread:
                self.translation_thread.join(timeout=2.0)  # 最多等待2秒
            
            # 重置UI状态
            self.translate_btn["state"] = "normal"
            self.show_progress_frame(False)  # 隐藏进度条和取消按钮
            self.progress_var.set(0)
            self.progress_percent.set("0%")
            self.progress_detail.set("等待开始翻译...")
            self.message_queue.put(('log', "翻译已取消"))

    def start_translation(self):
        """开始翻译流程"""
        if self.is_translating:
            return
            
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择Excel文件")
            return
            
        if not self.save_path.get():
            messagebox.showerror("错误", "请设置存储位置")
            return
            
        if not self.api_key.get().strip():
            messagebox.showerror("错误", "请在设置中配置DeepSeek API Key")
            return
            
        # 获取选中的目标语言，并去重，同时排除源语言
        source_lang = self.source_lang.get()
        target_langs = [lang for lang, var in self.target_langs_vars.items() 
                       if var.get() and lang != source_lang]
        target_langs = list(set(target_langs))  # 去重
        
        if not target_langs:
            messagebox.showerror("错误", "请选择至少一个目标语言（不包括源语言）")
            return
            
        try:
            # 检查文件是否存在
            if not os.path.exists(self.file_path.get()):
                messagebox.showerror("错误", "选择的Excel文件不存在")
                return
                
            # 检查是否可以读取文件
            try:
                df = pd.read_excel(self.file_path.get())
                # 检查是否已经存在目标语言列
                existing_langs = []
                selected_lang_codes = [SUPPORTED_LANGUAGES[lang] for lang in target_langs]
                
                for col in df.columns:
                    if col in selected_lang_codes:
                        # 找到对应的中文名称
                        for zh_name, en_name in SUPPORTED_LANGUAGES.items():
                            if en_name == col:
                                existing_langs.append(zh_name)
                                break
                                
                if existing_langs:
                    if not messagebox.askyesno("警告", 
                        f"文件中已存在以下语言列：\n{', '.join(existing_langs)}\n" + 
                        "这些列将被删除并重新翻译，是否继续？"):
                        return
            except Exception as e:
                messagebox.showerror("错误", f"无法读取Excel文件: {str(e)}")
                return
                
            # 检查保存路径是否存在
            save_dir = Path(self.save_path.get())
            if not save_dir.exists():
                try:
                    save_dir.mkdir(parents=True)
                except Exception as e:
                    messagebox.showerror("错误", f"无法创建保存目录: {str(e)}")
                    return
            
            # 生成新的输出文件名
            input_file = Path(self.file_path.get())
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = save_dir / f"{input_file.stem}_translated_{timestamp}{input_file.suffix}"
            
            # 清空日志显示
            if hasattr(self, 'home_log_text'):
                self.home_log_text.delete("1.0", "end")
            
            # 显示进度框并重置进度显示
            self.show_progress_frame(True)
            self.progress_var.set(0)
            self.progress_percent.set("0%")
            self.remaining_time.set("预计剩余时间: --:--")
            self.progress_detail.set("准备开始翻译...")
            
            # 记录开始信息
            self.message_queue.put(('log', f"Excel文件共有 {len(df)} 行"))
            self.message_queue.put(('log', f"添加目标语言列: {', '.join(target_langs)}"))
            
            # 重置状态
            self.cancel_flag = False
            self.is_translating = True
            
            # 记录开始时间
            self.translation_start_time = time.time()
            
            # 禁用开始翻译按钮
            self.translate_btn["state"] = "disabled"
            
            # 在新线程中运行翻译任务
            self.translation_thread = threading.Thread(
                target=self.run_translation,
                args=(self.file_path.get(), 
                      str(output_file),
                      self.source_lang.get(),
                      target_langs),
                daemon=True
            )
            self.translation_thread.start()
            
            # 启动进度处理
            self.root.after(50, self.process_progress)
            
            # 启动进度更新检查
            self.root.after(100, self.check_translation_progress)
                
        except Exception as e:
            self.handle_error("启动翻译时出错", e)
            
    def handle_error(self, context, error):
        """统一处理错误"""
        error_msg = str(error)
        
        # 检查是否是余额不足错误
        if "insufficient balance" in error_msg.lower() or "余额不足" in error_msg:
            error_msg = "DeepSeek API 余额不足，请充值后再试。"
        elif "unauthorized" in error_msg.lower() or "invalid api key" in error_msg.lower():
            error_msg = "API Key 无效或未授权，请检查API Key是否正确。"
        else:
            error_msg = f"{context}: {error_msg}"
            
        self.message_queue.put(('log', error_msg))
        self.progress_detail.set("翻译过程中出现错误")
        self.translate_btn["state"] = "normal"
        self.is_translating = False
        self.show_progress_frame(False)  # 隐藏进度条和取消按钮
        messagebox.showerror("错误", error_msg)
            
    def check_translation_progress(self):
        """检查翻译进度"""
        if self.is_translating and hasattr(self, 'translation_thread') and self.translation_thread.is_alive():
            # 如果翻译线程还在运行，继续检查
            self.root.after(100, self.check_translation_progress)
        else:
            # 如果翻译线程已结束
            if not self.cancel_flag:
                # 如果不是用户取消的，则更新UI
                self.translate_btn["state"] = "normal"
                self.is_translating = False
                if self.progress_var.get() >= 100:
                    self.show_progress_frame(False)
                
    def run_translation(self, input_file, output_file, source_lang, target_langs):
        try:
            # 检查API Key
            api_key = self.api_key.get().strip()
            if not api_key:
                self.handle_error("API设置错误", "请在设置中配置DeepSeek API Key")
                return
            
            # 记录开始信息
            start_time = datetime.now()
            total_rows = len(pd.read_excel(input_file))
            
            # 使用logger记录开始信息
            self.logger.info(f"\n{'='*50}")
            self.logger.info(f"开始翻译文件: {os.path.basename(input_file)}")
            self.logger.info(f"源语言: {source_lang}")
            self.logger.info(f"目标语言: {', '.join(target_langs)}")
            self.logger.info(f"需要翻译的行数: {total_rows}")
            self.logger.info(f"开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
            self.logger.info(f"{'='*50}")
            
            # 设置进度回调
            deepl_selenium_translate.progress_callback = self.update_progress
            
            # 准备参考源参数
            reference_params = {}
            if self.ref_mode.get() == "internal":
                # 检查内置参考源语言是否选择
                if not self.internal_ref_lang.get():
                    self.handle_error("参考源设置错误", "请选择内置参考源的语言")
                    return
                    
                # 读取Excel文件获取参考列
                df = pd.read_excel(input_file)
                ref_lang_display = self.internal_ref_lang.get()
                ref_lang_code = SUPPORTED_LANGUAGES[ref_lang_display]
                
                # 检查文件中的列名，支持中文名和英文代码
                ref_col_name = None
                for col in df.columns:
                    # 检查是否是支持的语言（中文名或英文代码）
                    if col == ref_lang_code:  # 英文代码匹配
                        ref_col_name = col
                        break
                    # 检查是否是中文名匹配
                    for zh_name, en_code in SUPPORTED_LANGUAGES.items():
                        if col == zh_name:
                            ref_col_name = col
                            ref_lang_code = en_code  # 更新为实际的语言代码
                            break
                    if ref_col_name:
                        break
                
                if not ref_col_name:
                    self.handle_error("参考源设置错误", f"在Excel文件中未找到{ref_lang_display}列作为参考源")
                    return
                    
                reference_params = {
                    "reference_lang": ref_lang_code,
                    "reference_column": ref_col_name
                }
                self.message_queue.put(('log', f"使用内置参考源，参考语言: {ref_lang_display} (列名: {ref_col_name})"))
            
            elif self.ref_mode.get() == "external":
                # 检查外部参考源设置
                if not self.ref_file_path.get() or not self.ref_lang.get():
                    self.handle_error("参考源设置错误", "请选择外部参考文件和参考语言")
                    return
                if not os.path.exists(self.ref_file_path.get()):
                    self.handle_error("参考源设置错误", "选择的外部参考文件不存在")
                    return
                
                # 读取参考文件的第一行（标题行）
                try:
                    ref_df = pd.read_excel(self.ref_file_path.get(), nrows=1)  # 只读取第一行
                    ref_lang_code = SUPPORTED_LANGUAGES[self.ref_lang.get()]
                    
                    # 只检查参考语言列
                    if ref_lang_code not in ref_df.columns and self.ref_lang.get() not in ref_df.columns:
                        self.handle_error("参考源设置错误", f"在参考文件中未找到{self.ref_lang.get()}列")
                        return
                    
                    reference_params = {
                        "reference_file": self.ref_file_path.get(),
                        "reference_lang": ref_lang_code
                    }
                    self.message_queue.put(('log', f"使用外部参考文件: {os.path.basename(self.ref_file_path.get())}"))
                    self.message_queue.put(('log', f"参考语言: {self.ref_lang.get()}"))
                    
                except Exception as e:
                    self.handle_error("读取参考文件失败", str(e))
                    return
            else:
                # 不使用参考源
                self.message_queue.put(('log', "不使用参考源进行翻译"))
            
            # 设置全局 API 配置
            deepl_selenium_translate.api_key = api_key  # 设置全局 API Key
            deepl_selenium_translate.DEEPSEEK_BASE_URL = self.DEEPSEEK_BASE_URL  # 设置全局 base_url
            deepl_selenium_translate.set_translation_cancelled(False)  # 重置取消状态
            
            # 设置翻译参数配置
            translate_config = {
                'max_workers': int(self.max_workers_var.get()),
                'batch_size': int(self.batch_size_var.get()),
                'max_retries': int(self.max_retries_var.get()),
                'save_interval': int(self.save_interval_var.get()),
                'progress_interval': int(self.progress_interval_var.get())
            }
            
            # 设置全局配置
            deepl_selenium_translate.set_config(translate_config)
            
            success = deepl_selenium_translate.process_excel_with_threading(
                excel_file=input_file,
                output_file=output_file,
                source_lang=SUPPORTED_LANGUAGES[source_lang],
                target_languages=[SUPPORTED_LANGUAGES[lang] for lang in target_langs],
                api_key_param=api_key,
                **reference_params
            )
            
            # 使用logger记录完成信息
            if not self.cancel_flag:
                if success:
                    # 计算用时
                    end_time = datetime.now()
                    duration = end_time - start_time
                    duration_str = str(duration).split('.')[0]
                    
                    self.logger.info(f"\n{'='*50}")
                    self.logger.info(f"翻译完成！")
                    self.logger.info(f"总用时：{duration_str}")
                    self.logger.info(f"翻译结果已保存到：{os.path.basename(output_file)}")
                    self.logger.info(f"{'='*50}")
                    
                    # 添加到历史记录
                    history_msg = (
                        f"翻译完成\n"
                        f"时间：{start_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"文件：{os.path.basename(input_file)}\n"
                        f"源语言：{source_lang}\n"
                        f"目标语言：{', '.join(target_langs)}\n"
                        f"翻译行数：{total_rows}\n"
                        f"用时：{duration_str}\n"
                        f"保存位置：{os.path.basename(output_file)}\n"
                        f"{'-' * 50}\n"
                    )
                    self.message_queue.put(('history', history_msg))
                else:
                    error_msg = "翻译过程返回失败状态，请检查浏览器是否正常运行"
                    self.message_queue.put(('log', error_msg))
                    self.message_queue.put(('complete', False))
            else:
                # 取消翻译的记录
                end_time = datetime.now()
                duration = end_time - start_time
                duration_str = str(duration).split('.')[0]
                
                # 使用logger记录取消信息
                self.logger.info(f"\n{'='*50}")
                self.logger.info("翻译已取消")
                self.logger.info(f"已用时间：{duration_str}")
                self.logger.info(f"{'='*50}")
                
                # 添加到历史记录
                history_msg = (
                    f"翻译已取消\n"
                    f"时间：{start_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"文件：{os.path.basename(input_file)}\n"
                    f"源语言：{source_lang}\n"
                    f"目标语言：{', '.join(target_langs)}\n"
                    f"已完成进度：{self.progress_var.get():.0f}%\n"
                    f"已用时间：{duration_str}\n"
                    f"{'-' * 50}\n"
                )
                self.message_queue.put(('history', history_msg))
                    
        except ValueError as e:
            error_msg = str(e).lower()
            if "api key无效" in error_msg:
                self.handle_error("API认证失败", "API Key 无效或未授权，请检查API Key是否正确")
            else:
                self.handle_error("翻译过程出现异常", e)
        except Exception as e:
            error_msg = str(e).lower()
            if "insufficient balance" in error_msg or "余额不足" in error_msg:
                self.handle_error("API余额不足", "DeepSeek API 余额不足，请充值后再试")
            else:
                self.handle_error("翻译过程出现异常", e)
                
    def process_messages(self):
        """处理消息队列中的消息"""
        try:
            while True:
                msg_type, msg_content = self.message_queue.get_nowait()
                if msg_type == 'log':
                    # 更新首页的运行日志
                    if hasattr(self, 'home_log_text'):
                        # 移除INFO前缀
                        msg_content = msg_content.replace(" - INFO - ", " - ")
                        self.home_log_text.insert("end", f"{msg_content}\n")
                        self.home_log_text.see("end")
                        self.home_log_text.update_idletasks()
                elif msg_type == 'history':
                    # 更新历史记录页面
                    if hasattr(self, 'log_text'):
                        self.log_text.insert("end", f"{msg_content}\n")
                        self.log_text.see("end")
                        self.log_text.update_idletasks()
                elif msg_type == 'progress':
                    current, total = msg_content
                    self.update_progress(current, total)
                elif msg_type == 'complete':
                    success = msg_content
                    if success:
                        self.progress_var.set(100)  # 确保进度条显示100%
                        self.progress_percent.set("100%")
                        self.progress_detail.set("翻译已完成!")
                        self.remaining_time.set("完成时间: " + datetime.now().strftime("%H:%M:%S"))
                        
                        # 添加翻译完成记录
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        history_msg = (
                            f"[{timestamp}] 完成翻译：\n"
                            f"源文件：{os.path.basename(self.file_path.get())}\n"
                            f"源语言：{self.source_lang.get()}\n"
                            f"目标语言：{', '.join(lang for lang, var in self.target_langs_vars.items() if var.get())}\n"
                            f"保存位置：{os.path.basename(str(self.save_path.get()))}\n"
                            f"{'-' * 50}\n"
                        )
                        self.message_queue.put(('history', history_msg))
                        
                        # 在运行日志中添加完成信息
                        self.message_queue.put(('log', "翻译完成！"))
                        self.message_queue.put(('log', f"翻译结果已保存到：{os.path.basename(str(self.save_path.get()))}"))
                        
                        # 显示完成提示
                        messagebox.showinfo("提示", "翻译已完成！\n请在翻译结果页面查看生成的文件。")
                    else:
                        self.progress_detail.set("翻译过程中出现错误")
                        if hasattr(self, 'home_log_text'):
                            last_error = self.home_log_text.get("end-2l", "end-1l")
                            messagebox.showerror("错误", f"翻译过程中出现错误:\n{last_error}")
                        else:
                            messagebox.showerror("错误", "翻译过程中出现错误")
                    
                    # 恢复界面状态
                    self.translate_btn["state"] = "normal"
                    self.is_translating = False
                    self.cancel_btn.pack_forget()
                    self.translation_start_time = None
                    
                    # 如果成功完成，延迟3秒后隐藏进度框
                    if success:
                        self.root.after(3000, lambda: self.show_progress_frame(False))
        except queue.Empty:
            pass
        finally:
            # 每50ms检查一次消息队列
            self.root.after(50, self.process_messages)
    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_path.set(file_path)
            
    def select_save_path(self):
        """选择保存目录"""
        save_path = filedialog.askdirectory()
        if save_path:
            self.save_path.set(save_path)
            
    def load_config(self):
        """加载配置文件"""
        config_path = Path.home() / ".translate_config.json"
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}
        
    def save_config(self):
        """保存配置"""
        # 获取选中的目标语言
        selected_langs = [lang for lang, var in self.target_langs_vars.items() 
                         if var.get()]
        
        # 验证并保存翻译参数
        try:
            max_workers = int(self.max_workers_var.get())
            batch_size = int(self.batch_size_var.get())
            max_retries = int(self.max_retries_var.get())
            save_interval = int(self.save_interval_var.get())
            progress_interval = int(self.progress_interval_var.get())
            
            # 参数验证
            if not (1 <= max_workers <= 20):
                raise ValueError("并发线程数必须在1-20之间")
            if not (1 <= batch_size <= 50):
                raise ValueError("批处理大小必须在1-50之间")
            if not (1 <= max_retries <= 10):
                raise ValueError("最大重试次数必须在1-10之间")
            if not (10 <= save_interval <= 1000):
                raise ValueError("保存间隔必须在10-1000之间")
            if not (1 <= progress_interval <= 100):
                raise ValueError("进度显示间隔必须在1-100之间")
            
            config = {
                "save_path": self.save_path.get(),
                "source_lang": self.source_lang.get(),
                "target_langs": selected_langs,
                "theme": self.current_theme,
                "api_key": self.api_key.get(),
                "max_workers": max_workers,
                "batch_size": batch_size,
                "max_retries": max_retries,
                "save_interval": save_interval,
                "progress_interval": progress_interval
            }
            
            config_path = Path.home() / ".translate_config.json"
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("提示", "配置已保存")
            
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            return

    def update_progress(self, current, total, finished=False):
        """更新进度显示"""
        try:
            if total > 0:
                percentage = (current / total) * 100
                self.progress_var.set(percentage)
                self.progress_percent.set(f"{percentage:.1f}%")
                self.progress_detail.set(f"已翻译: {current}/{total}")
                
                # 计算剩余时间
                if self.translation_start_time and current > 0:
                    elapsed_time = time.time() - self.translation_start_time
                    items_per_second = current / elapsed_time
                    remaining_items = total - current
                    
                    if items_per_second > 0:
                        remaining_seconds = remaining_items / items_per_second
                        
                        # 格式化剩余时间
                        if remaining_seconds < 60:
                            time_str = f"约{int(remaining_seconds)}秒"
                        elif remaining_seconds < 3600:
                            minutes = int(remaining_seconds / 60)
                            time_str = f"约{minutes}分钟"
                        else:
                            hours = int(remaining_seconds / 3600)
                            minutes = int((remaining_seconds % 3600) / 60)
                            time_str = f"约{hours}小时{minutes}分钟"
                        
                        self.remaining_time.set(f"预计剩余时间: {time_str}")
                
                if finished:
                    self.cancel_btn["state"] = "disabled"
                    self.cancel_btn.pack_forget()
                    self.progress_detail.set("翻译完成!")
                    self.progress_var.set(100)
                    self.root.after(3000, lambda: self.show_progress_frame(False))
            
            # 强制更新UI
            self.root.update_idletasks()
            
        except Exception as e:
            logger.error(f"更新进度显示时出错: {e}")

    def process_progress(self):
        """处理进度更新队列"""
        try:
            while True:
                current, total = self.progress_queue.get_nowait()
                
                # 限制更新频率，避免UI过度刷新
                current_time = time.time()
                if current_time - self.last_update_time < 0.1:  # 限制为每0.1秒更新一次
                    continue
                self.last_update_time = current_time
                
                if total > 0:
                    # 计算进度百分比
                    progress = (current / total) * 100
                    self.progress_var.set(progress)
                    self.progress_percent.set(f"{progress:.1f}%")
                    
                    # 计算剩余时间
                    if self.translation_start_time and current > 0:
                        elapsed_time = time.time() - self.translation_start_time
                        items_per_second = current / elapsed_time
                        remaining_items = total - current
                        
                        if items_per_second > 0:
                            remaining_seconds = remaining_items / items_per_second
                            
                            # 格式化剩余时间
                            if remaining_seconds < 60:
                                time_str = f"约{int(remaining_seconds)}秒"
                            elif remaining_seconds < 3600:
                                minutes = int(remaining_seconds / 60)
                                time_str = f"约{minutes}分钟"
                            else:
                                hours = int(remaining_seconds / 3600)
                                minutes = int((remaining_seconds % 3600) / 60)
                                time_str = f"约{hours}小时{minutes}分钟"
                            
                            self.remaining_time.set(f"预计剩余时间: {time_str}")
                    
                    # 更新详细进度
                    self.progress_detail.set(f"已翻译: {current}/{total}")
                
                # 强制更新UI
                self.root.update_idletasks()
                
        except queue.Empty:
            pass
        finally:
            # 每50ms检查一次进度队列
            self.root.after(50, self.process_progress)
            
    def change_theme(self):
        """切换主题"""
        self.current_theme = self.theme_var.get()
        self.theme = LightTheme if self.current_theme == "light" else DarkTheme
        
        # 更新窗口背景色
        self.root.configure(bg=self.theme.BACKGROUND)
        
        # 重新配置样式
        self.setup_styles()
        
        # 更新所有文本框的颜色
        self.update_text_widgets()
        
        # 保存主题设置
        self.save_config()

    def update_text_widgets(self):
        """更新所有文本框的颜色"""
        # 更新首页日志文本框
        if hasattr(self, 'home_log_text'):
            self.home_log_text.configure(
                bg=self.theme.ENTRY_BG,
                fg=self.theme.FOREGROUND,
                insertbackground=self.theme.FOREGROUND
            )
        
        # 更新历史记录文本框
        if hasattr(self, 'log_text'):
            self.log_text.configure(
                bg=self.theme.ENTRY_BG,
                fg=self.theme.FOREGROUND,
                insertbackground=self.theme.FOREGROUND
            )
            
        # 更新文件列表框
        if hasattr(self, 'file_listbox'):
            self.file_listbox.configure(
                bg=self.theme.ENTRY_BG,
                fg=self.theme.FOREGROUND
            )

    def refresh_file_list(self):
        """刷新翻译结果文件列表"""
        if hasattr(self, 'file_listbox'):
            # 清空列表
            self.file_listbox.delete(0, tk.END)
            
            # 获取保存路径
            save_path = self.config.get("save_path")
            if save_path and os.path.exists(save_path):
                # 获取所有翻译结果文件
                files = []
                for file in os.listdir(save_path):
                    if file.endswith((".xlsx", ".xls")) and "_translated_" in file:
                        # 获取文件修改时间
                        file_path = os.path.join(save_path, file)
                        mod_time = os.path.getmtime(file_path)
                        files.append((file, mod_time))
                
                # 按修改时间排序，最新的在前
                files.sort(key=lambda x: x[1], reverse=True)
                
                # 添加到列表中
                for file, _ in files:
                    self.file_listbox.insert(tk.END, file)
        
        # 每60秒自动刷新一次
        self.root.after(60000, self.refresh_file_list)
        
    def open_selected_file(self, event=None):
        """打开选中的文件"""
        selection = self.file_listbox.curselection()
        if selection:
            file_name = self.file_listbox.get(selection[0])
            save_path = self.config.get("save_path")
            if save_path:
                file_path = os.path.join(save_path, file_name)
                if os.path.exists(file_path):
                    try:
                        os.startfile(file_path)  # Windows
                    except AttributeError:
                        try:
                            import subprocess
                            subprocess.run(['xdg-open', file_path])  # Linux
                        except:
                            try:
                                subprocess.run(['open', file_path])  # macOS
                            except:
                                messagebox.showerror("错误", "无法打开文件")
                else:
                    messagebox.showerror("错误", "文件不存在")

    def setup_logger(self):
        """设置日志处理器"""
        # 在项目目录下创建logs文件夹
        log_dir = Path("logs")
        log_dir.mkdir(parents=True, exist_ok=True)
        
        # 创建日志文件路径
        self.log_file = log_dir / "translation_history.log"
        
        # 获取logger实例
        logger = logging.getLogger("translation_history")
        logger.setLevel(logging.INFO)
        
        # 创建文件处理器，用于写入历史记录
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        logger.addHandler(file_handler)
        
        # 创建队列处理器，用于更新界面
        class QueueHandler(logging.Handler):
            def __init__(self, queue):
                super().__init__()
                self.queue = queue

            def emit(self, record):
                try:
                    msg = self.format(record)
                    self.queue.put(('log', msg))
                except Exception:
                    self.handleError(record)
        
        # 添加队列处理器
        queue_handler = QueueHandler(self.message_queue)
        queue_handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(queue_handler)
        
        # 设置不传播到父logger
        logger.propagate = False
        
        # 保存logger实例以供后续使用
        self.logger = logger

    def load_history_log(self):
        """加载历史日志"""
        if hasattr(self, 'log_text') and self.log_file.exists():
            try:
                with open(self.log_file, 'r', encoding='utf-8') as f:
                    log_content = f.read()
                    self.log_text.delete("1.0", "end")
                    self.log_text.insert("1.0", log_content)
                    self.log_text.see("end")
            except Exception as e:
                print(f"加载历史日志失败: {str(e)}")

    def create_settings_page(self):
        """创建设置页面"""
        settings_page = ttk.Frame(self.content, style="Modern.TFrame")
        self.pages["settings"] = settings_page
        
        # API设置
        api_frame = ttk.LabelFrame(settings_page, text="DeepSeek API设置", 
                                style="Modern.TLabelframe",
                                padding=15)
        api_frame.pack(fill="x", pady=(0, 10))
        
        # API Key输入框
        api_key_frame = ttk.Frame(api_frame)
        api_key_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Label(api_key_frame, text="API Key:", 
                style="Modern.TLabel").pack(side="left")
        
        # 创建带密码遮罩的输入框
        api_key_entry = ttk.Entry(api_key_frame, 
                               textvariable=self.api_key,
                               width=50,
                               show="*")  # 使用*遮罩显示
        api_key_entry.pack(side="left", padx=5)
        
        # 显示/隐藏密码按钮
        self.show_password = tk.BooleanVar(value=False)
        def toggle_password():
            api_key_entry.configure(show="" if self.show_password.get() else "*")
        
        ttk.Checkbutton(api_key_frame, text="显示",
                     variable=self.show_password,
                     command=toggle_password).pack(side="left")
        
        # 添加API Key说明
        ttk.Label(api_frame, 
                text="注意：请确保API Key有足够的余额，余额不足时翻译将会失败。",
                style="Modern.TLabel",
                wraplength=500).pack(fill="x", pady=5)
        
        # 主题设置
        theme_frame = ttk.LabelFrame(settings_page, text="主题设置", 
                                   style="Modern.TLabelframe",
                                   padding=15)
        theme_frame.pack(fill="x", pady=10)
        
        ttk.Label(theme_frame, text="界面主题:", 
                 style="Modern.TLabel").grid(row=0, column=0, sticky="w")
        
        self.theme_var = tk.StringVar(value=self.current_theme)
        theme_light = ttk.Radiobutton(theme_frame, text="明亮模式",
                                    variable=self.theme_var,
                                    value="light",
                                    command=self.change_theme)
        theme_light.grid(row=0, column=1, padx=5)
        
        theme_dark = ttk.Radiobutton(theme_frame, text="深色模式",
                                   variable=self.theme_var,
                                   value="dark",
                                   command=self.change_theme)
        theme_dark.grid(row=0, column=2, padx=5)
        
        # 存储位置设置
        save_frame = ttk.LabelFrame(settings_page, text="存储设置", 
                                  style="Modern.TLabelframe",
                                  padding=15)
        save_frame.pack(fill="x", pady=10)
        
        ttk.Label(save_frame, text="存储位置:", 
                 style="Modern.TLabel").grid(row=0, column=0, sticky="w")
        save_entry = ttk.Entry(save_frame, textvariable=self.save_path, 
                             width=50, style="Modern.TEntry")
        save_entry.grid(row=0, column=1, padx=5)
        ttk.Button(save_frame, text="选择目录", 
                  style="Modern.TButton",
                  command=self.select_save_path).grid(row=0, column=2, padx=5)
        
        # 添加翻译参数设置
        translate_params_frame = ttk.LabelFrame(settings_page, text="翻译参数设置", 
                                              style="Modern.TLabelframe",
                                              padding=15)
        translate_params_frame.pack(fill="x", pady=10)

        # 并发线程数
        ttk.Label(translate_params_frame, text="并发线程数:", 
                 style="Modern.TLabel").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.max_workers_var = tk.StringVar(value=str(self.config.get("max_workers", 5)))
        ttk.Entry(translate_params_frame, textvariable=self.max_workers_var, 
                 width=10).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(translate_params_frame, text="（建议：1-10，默认：5）", 
                 style="Modern.TLabel").grid(row=0, column=2, sticky="w", padx=5)

        # 批处理大小
        ttk.Label(translate_params_frame, text="批处理大小:", 
                 style="Modern.TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.batch_size_var = tk.StringVar(value=str(self.config.get("batch_size", 10)))
        ttk.Entry(translate_params_frame, textvariable=self.batch_size_var, 
                 width=10).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(translate_params_frame, text="（建议：5-20，默认：10）", 
                 style="Modern.TLabel").grid(row=1, column=2, sticky="w", padx=5)

        # 最大重试次数
        ttk.Label(translate_params_frame, text="最大重试次数:", 
                 style="Modern.TLabel").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.max_retries_var = tk.StringVar(value=str(self.config.get("max_retries", 3)))
        ttk.Entry(translate_params_frame, textvariable=self.max_retries_var, 
                 width=10).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(translate_params_frame, text="（建议：1-5，默认：3）", 
                 style="Modern.TLabel").grid(row=2, column=2, sticky="w", padx=5)

        # 保存间隔
        ttk.Label(translate_params_frame, text="保存间隔:", 
                 style="Modern.TLabel").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.save_interval_var = tk.StringVar(value=str(self.config.get("save_interval", 100)))
        ttk.Entry(translate_params_frame, textvariable=self.save_interval_var, 
                 width=10).grid(row=3, column=1, sticky="w", padx=5)
        ttk.Label(translate_params_frame, text="（每处理多少个单元格保存一次，默认：100）", 
                 style="Modern.TLabel").grid(row=3, column=2, sticky="w", padx=5)

        # 进度显示间隔
        ttk.Label(translate_params_frame, text="进度显示间隔:", 
                 style="Modern.TLabel").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.progress_interval_var = tk.StringVar(value=str(self.config.get("progress_interval", 10)))
        ttk.Entry(translate_params_frame, textvariable=self.progress_interval_var, 
                 width=10).grid(row=4, column=1, sticky="w", padx=5)
        ttk.Label(translate_params_frame, text="（每处理多少个单元格更新一次进度，默认：10）", 
                 style="Modern.TLabel").grid(row=4, column=2, sticky="w", padx=5)
        
        # 保存配置按钮
        ttk.Button(settings_page, text="保存配置",
                  style="Modern.TButton",
                  command=self.save_config).pack(pady=10)
                  
    def create_results_page(self):
        """创建翻译结果页面"""
        results_page = ttk.Frame(self.content, style="Modern.TFrame")
        self.pages["results"] = results_page
        
        # 创建工具栏
        toolbar = ttk.Frame(results_page, style="Modern.TFrame")
        toolbar.pack(fill="x", padx=10, pady=5)
        
        # 刷新按钮
        refresh_btn = ttk.Button(toolbar, text="刷新列表",
                               style="Modern.TButton",
                               command=self.refresh_file_list)
        refresh_btn.pack(side="left", padx=5)
        
        # 创建文件列表框架
        list_frame = ttk.LabelFrame(results_page, text="翻译结果文件",
                                  style="Modern.TLabelframe",
                                  padding=15)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建文件列表
        self.file_listbox = tk.Listbox(list_frame,
                                     bg=self.theme.ENTRY_BG,
                                     fg=self.theme.FOREGROUND,
                                     selectmode="single",
                                     font=("微软雅黑", 9))
        self.file_listbox.pack(fill="both", expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                command=self.file_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 绑定双击事件
        self.file_listbox.bind("<Double-Button-1>", self.open_selected_file)
        
    def create_log_page(self):
        """创建日志页面"""
        log_page = ttk.Frame(self.content, style="Modern.TFrame")
        self.pages["log"] = log_page
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(log_page, text="翻译历史记录", 
                                 style="Modern.TLabelframe",
                                 padding=15)
        log_frame.pack(fill="both", expand=True)
        
        # 创建日志文本框
        self.log_text = tk.Text(log_frame, height=20,
                              bg=self.theme.ENTRY_BG,
                              fg=self.theme.FOREGROUND,
                              insertbackground=self.theme.FOREGROUND,
                              relief="flat",
                              font=("微软雅黑", 9))
        self.log_text.pack(fill="both", expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", 
                                command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 添加说明文本
        self.log_text.insert("1.0", "翻译历史记录：\n\n")

    def toggle_ref_mode(self):
        """切换参考模式"""
        if self.ref_mode.get() == "external":
            self.internal_ref_frame.pack_forget()  # 隐藏内置参考设置
            self.ref_file_frame.pack(fill="x", pady=5)  # 显示外部参考文件设置
            self.internal_ref_lang.set("")  # 清空内置参考语言
        elif self.ref_mode.get() == "internal":
            self.ref_file_frame.pack_forget()  # 隐藏外部参考文件设置
            self.internal_ref_frame.pack(fill="x", pady=5)  # 显示内置参考设置
            self.ref_file_path.set("")  # 清空参考文件路径
            self.ref_lang.set("")  # 清空外部参考语言
        else:  # none
            self.internal_ref_frame.pack_forget()  # 隐藏内置参考设置
            self.ref_file_frame.pack_forget()  # 隐藏外部参考文件设置
            self.internal_ref_lang.set("")  # 清空内置参考语言
            self.ref_file_path.set("")  # 清空参考文件路径
            self.ref_lang.set("")  # 清空外部参考语言
            
    def select_ref_file(self):
        """选择参考文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.ref_file_path.set(file_path)
            
    def check_reference_source(self):
        """检查参考源"""
        if self.ref_mode.get() == "internal":
            # 检查内置参考源
            try:
                if not self.internal_ref_lang.get():
                    messagebox.showerror("错误", "请选择内置参考源的参考语言")
                    return False
                    
                reference_file = os.path.join(os.path.dirname(self.file_path.get()), "reference.xlsx")
                has_reference = os.path.exists(reference_file)
                if not has_reference:
                    return messagebox.askyesno("提示",
                        "未在Excel文件所在目录找到reference.xlsx参考源文件，是否继续翻译？\n" +
                        "注意：没有参考源可能会影响翻译质量。")
                return True
            except Exception as e:
                messagebox.showerror("错误", f"检查内置参考源时出错: {str(e)}")
                return False
        else:
            # 检查外部参考文件
            if not self.ref_file_path.get():
                messagebox.showerror("错误", "请选择参考文件")
                return False
            if not self.ref_lang.get():
                messagebox.showerror("错误", "请选择参考语言")
                return False
            if not os.path.exists(self.ref_file_path.get()):
                messagebox.showerror("错误", "选择的参考文件不存在")
                return False
            return True

    def create_subtitle_translate_page(self):
        """创建字幕翻译页面"""
        subtitle_translate_frame = SubtitleTranslateFrame(
            self.content,
            self.theme,
            self.api_key.get()
        )
        self.pages["subtitle_translate"] = subtitle_translate_frame
        
    def create_subtitle_result_page(self):
        """创建字幕结果页面"""
        subtitle_result_frame = SubtitleResultFrame(
            self.content,
            self.theme
        )
        self.pages["subtitle_result"] = subtitle_result_frame

def main():
    root = tk.Tk()
    app = TranslateApplication(root)
    root.mainloop()

if __name__ == "__main__":
    main()