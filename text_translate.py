import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
from openai import OpenAI
import threading
import os
from constants import SUPPORTED_LANGUAGES  # 从constants导入

class TextTranslateFrame(ttk.Frame):
    def __init__(self, master, theme, api_key):
        super().__init__(master, style="Modern.TFrame")
        self.theme = theme
        self.api_key = api_key
        self.DEEPSEEK_BASE_URL = "https://api.deepseek.com"
        self.is_translating = False
        self.animation_chars = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"  # 加载动画字符
        self.animation_index = 0
        
        # 初始化术语表目录
        self.glossary_dir = Path("glossary")
        self.glossary_dir.mkdir(exist_ok=True)
        
        # 创建界面
        self.create_layout()
        
    def get_glossary_file(self, source_lang, target_lang):
        """获取术语表文件路径"""
        return self.glossary_dir / f"{source_lang}_{target_lang}.json"
        
    def load_terminology(self, source_lang, target_lang):
        """加载特定语言对的术语表"""
        glossary_file = self.get_glossary_file(source_lang, target_lang)
        if glossary_file.exists():
            try:
                with open(glossary_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                messagebox.showerror("错误", f"加载术语表失败: {str(e)}")
                return {}
        return {}
        
    def save_terminology(self, source_lang, target_lang, terms):
        """保存术语表"""
        glossary_file = self.get_glossary_file(source_lang, target_lang)
        try:
            with open(glossary_file, 'w', encoding='utf-8') as f:
                json.dump(terms, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存术语表失败: {str(e)}")

    def create_layout(self):
        """创建主界面布局"""
        # 创建主框架
        main_frame = ttk.Frame(self, style="Modern.TFrame")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 顶部语言选择区域
        lang_frame = ttk.Frame(main_frame, style="Modern.TFrame")
        lang_frame.pack(fill="x", pady=(0, 10))
        
        # 源语言选择
        source_frame = ttk.Frame(lang_frame)
        source_frame.pack(side="left")
        self.source_lang = tk.StringVar(value="英语")
        source_combo = ttk.Combobox(source_frame, 
                                  textvariable=self.source_lang,
                                  values=list(SUPPORTED_LANGUAGES.keys()),
                                  state="readonly",
                                  width=15)
        source_combo.pack(side="left", padx=5)
        # 绑定语言变更事件
        source_combo.bind('<<ComboboxSelected>>', self.on_language_change)
        
        # 交换语言按钮
        ttk.Button(lang_frame, text="⇄",
                  style="Modern.TButton",
                  command=self.swap_languages).pack(side="left", padx=10)
        
        # 目标语言选择
        target_frame = ttk.Frame(lang_frame)
        target_frame.pack(side="left")
        self.target_lang = tk.StringVar(value="中文")
        target_combo = ttk.Combobox(target_frame,
                                  textvariable=self.target_lang,
                                  values=list(SUPPORTED_LANGUAGES.keys()),
                                  state="readonly",
                                  width=15)
        target_combo.pack(side="left", padx=5)
        # 绑定语言变更事件
        target_combo.bind('<<ComboboxSelected>>', self.on_language_change)
        
        # 术语表按钮
        ttk.Button(lang_frame, text="术语表管理",
                  style="Modern.TButton",
                  command=self.show_terminology_manager).pack(side="right", padx=5)
        
        # 创建左右分栏
        paned = ttk.PanedWindow(main_frame, orient="horizontal")
        paned.pack(fill="both", expand=True)
        
        # 左侧源文本区域
        left_frame = ttk.Frame(paned, width=350)  # 设置初始宽度
        left_frame.pack_propagate(False)  # 防止子组件影响框架大小
        
        # 添加源文本标签
        ttk.Label(left_frame, text="源文本").pack(anchor="w", padx=5, pady=(0, 5))
        
        # 添加滚动条
        left_scroll = ttk.Scrollbar(left_frame)
        left_scroll.pack(side="right", fill="y")
        
        self.source_text = tk.Text(left_frame, wrap="word",
                                 bg=self.theme.ENTRY_BG,
                                 fg=self.theme.FOREGROUND,
                                 insertbackground=self.theme.FOREGROUND,
                                 yscrollcommand=left_scroll.set)
        self.source_text.pack(fill="both", expand=True)
        left_scroll.config(command=self.source_text.yview)
        paned.add(left_frame, weight=1)  # weight=1 表示均分空间
        
        # 右侧翻译结果区域
        right_frame = ttk.Frame(paned, width=350)  # 设置初始宽度
        right_frame.pack_propagate(False)  # 防止子组件影响框架大小
        
        # 添加翻译结果标签
        ttk.Label(right_frame, text="翻译结果").pack(anchor="w", padx=5, pady=(0, 5))
        
        # 添加滚动条
        right_scroll = ttk.Scrollbar(right_frame)
        right_scroll.pack(side="right", fill="y")
        
        self.target_text = tk.Text(right_frame, wrap="word",
                                 bg=self.theme.ENTRY_BG,
                                 fg=self.theme.FOREGROUND,
                                 insertbackground=self.theme.FOREGROUND,
                                 yscrollcommand=right_scroll.set)
        self.target_text.pack(fill="both", expand=True)
        right_scroll.config(command=self.target_text.yview)
        paned.add(right_frame, weight=1)  # weight=1 表示均分空间
        
        # 底部按钮和状态区域
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=10)
        
        # 左侧按钮区域
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(side="left")
        
        # 翻译按钮
        self.translate_btn = ttk.Button(button_frame, text="翻译",
                                      style="Modern.TButton",
                                      command=self.translate_text)
        self.translate_btn.pack(side="left", padx=5)
        
        # 清空按钮
        ttk.Button(button_frame, text="清空",
                  style="Modern.TButton",
                  command=self.clear_text).pack(side="left", padx=5)
        
        # 复制按钮
        ttk.Button(button_frame, text="复制结果",
                  style="Modern.TButton",
                  command=self.copy_result).pack(side="left", padx=5)
        
        # 右侧状态标签
        self.status_label = ttk.Label(bottom_frame, text="就绪",
                                    style="Modern.TLabel")
        self.status_label.pack(side="right", padx=5)
        
        # 绑定快捷键
        self.bind_shortcuts()
        
    def bind_shortcuts(self):
        """绑定快捷键"""
        self.source_text.bind("<Control-Return>", lambda e: self.translate_text())
        self.source_text.bind("<Control-Key-c>", lambda e: self.copy_result())
        
    def translate_text(self):
        """执行翻译"""
        if self.is_translating:
            return
            
        source_text = self.source_text.get("1.0", "end-1c").strip()
        if not source_text:
            messagebox.showwarning("提示", "请输入要翻译的文本")
            return
            
        # 获取相关术语
        terms = self.get_relevant_terms(source_text)
        
        # 禁用翻译按钮并开始动画
        self.translate_btn.configure(state="disabled")
        self.is_translating = True
        self.update_animation()
        
        # 在新线程中执行翻译
        threading.Thread(target=self._do_translate,
                       args=(source_text, terms),
                       daemon=True).start()
        
    def update_animation(self):
        """更新加载动画"""
        if self.is_translating:
            self.animation_index = (self.animation_index + 1) % len(self.animation_chars)
            char = self.animation_chars[self.animation_index]
            self.status_label.configure(text=f"{char} 正在翻译...")
            self.after(100, self.update_animation)
        else:
            self.status_label.configure(text="就绪")

    def _do_translate(self, source_text, terms):
        """执行翻译的具体实现"""
        try:
            client = OpenAI(api_key=self.api_key, base_url=self.DEEPSEEK_BASE_URL)
            
            # 将文本分段，每段最多1000个字符
            segments = self._split_text(source_text, 1000)
            translated_segments = []
            
            for i, segment in enumerate(segments):
                # 构建提示词
                prompt = f"请将以下{self.source_lang.get()}文本翻译成{self.target_lang.get()}。\n\n"
                
                # 添加术语表提示
                if terms:
                    prompt += "请注意以下专业术语的翻译：\n"
                    for source, target in terms:
                        prompt += f"- {source} → {target}\n"
                    prompt += "\n原文：\n"
                
                prompt += segment
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "你是一个专业的翻译专家，请准确翻译用户的文本。"},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                    max_tokens=2000
                )
                
                # 获取翻译结果
                translation = response.choices[0].message.content.strip()
                translated_segments.append(translation)
                
                # 更新进度
                self.after(0, self.status_label.configure, 
                          {"text": f"已完成 {i+1}/{len(segments)} 段"})
            
            # 合并所有翻译结果
            final_translation = "\n".join(translated_segments)
            
            # 在主线程中更新UI
            self.after(0, self._update_translation, final_translation)
            
        except Exception as e:
            self.after(0, self._show_error, str(e))
        finally:
            self.is_translating = False
            self.after(0, self.translate_btn.configure, {"state": "normal"})

    def _split_text(self, text, max_length):
        """将文本分段"""
        segments = []
        lines = text.split('\n')
        current_segment = []
        current_length = 0
        
        for line in lines:
            line_length = len(line)
            if current_length + line_length > max_length and current_segment:
                segments.append('\n'.join(current_segment))
                current_segment = [line]
                current_length = line_length
            else:
                current_segment.append(line)
                current_length += line_length
        
        if current_segment:
            segments.append('\n'.join(current_segment))
        
        return segments
        
    def _update_translation(self, translation):
        """更新翻译结果"""
        self.target_text.delete("1.0", "end")
        self.target_text.insert("1.0", translation)
        
    def _show_error(self, error):
        """显示错误信息"""
        messagebox.showerror("翻译错误", error)
        
    def get_relevant_terms(self, text):
        """获取与文本相关的术语"""
        source_lang = self.source_lang.get()
        target_lang = self.target_lang.get()
        terms = self.load_terminology(source_lang, target_lang)
        
        relevant_terms = []
        for source_term, target_term in terms.items():
            if source_term.lower() in text.lower():
                relevant_terms.append((source_term, target_term))
        return relevant_terms
        
    def show_terminology_manager(self):
        """显示术语表管理窗口"""
        TerminologyManager(self, self.theme)
        
    def clear_text(self):
        """清空文本"""
        self.source_text.delete("1.0", "end")
        self.target_text.delete("1.0", "end")
        
    def copy_result(self):
        """复制翻译结果"""
        result = self.target_text.get("1.0", "end-1c")
        if result:
            self.clipboard_clear()
            self.clipboard_append(result)
            messagebox.showinfo("提示", "翻译结果已复制到剪贴板")
            
    def swap_languages(self):
        """交换源语言和目标语言"""
        source = self.source_lang.get()
        target = self.target_lang.get()
        self.source_lang.set(target)
        self.target_lang.set(source)
        
        # 交换文本内容
        source_text = self.source_text.get("1.0", "end-1c")
        target_text = self.target_text.get("1.0", "end-1c")
        self.source_text.delete("1.0", "end")
        self.target_text.delete("1.0", "end")
        self.source_text.insert("1.0", target_text)
        self.target_text.insert("1.0", source_text)

    def on_language_change(self, event=None):
        """当语言选择改变时的处理"""
        # 清空翻译结果
        self.target_text.delete("1.0", "end")
        # 可以在这里添加其他语言变更时的处理逻辑

class TerminologyManager(tk.Toplevel):
    def __init__(self, parent, theme):
        super().__init__(parent)
        self.parent = parent
        self.theme = theme
        self.title("术语表管理")
        self.geometry("800x600")
        
        # 分页相关变量
        self.page_size = 10  # 每页显示数量
        self.current_page = 1  # 当前页码
        self.total_terms = 0  # 术语总数
        self.filtered_terms = []  # 搜索过滤后的术语列表
        
        # 加载当前语言对的术语表
        self.current_terms = self.parent.load_terminology(
            self.parent.source_lang.get(),
            self.parent.target_lang.get()
        )
        
        self.create_layout()
        self.refresh_terminology_list()

    def create_layout(self):
        """创建术语表管理界面"""
        # 创建主框架
        main_frame = ttk.Frame(self, style="Modern.TFrame")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 添加术语区域
        add_frame = ttk.LabelFrame(main_frame, text="添加术语", padding=10)
        add_frame.pack(fill="x", pady=(0, 10))
        
        # 源语言术语
        source_frame = ttk.Frame(add_frame)
        source_frame.pack(fill="x", pady=5)
        ttk.Label(source_frame, text="源语言术语:").pack(side="left")
        self.source_term = ttk.Entry(source_frame, width=30)
        self.source_term.pack(side="left", padx=5)
        
        # 目标语言术语
        target_frame = ttk.Frame(add_frame)
        target_frame.pack(fill="x", pady=5)
        ttk.Label(target_frame, text="目标语言术语:").pack(side="left")
        self.target_term = ttk.Entry(target_frame, width=30)
        self.target_term.pack(side="left", padx=5)
        
        # 语言选择
        lang_frame = ttk.Frame(add_frame)
        lang_frame.pack(fill="x", pady=5)
        
        # 源语言选择
        ttk.Label(lang_frame, text="源语言:").pack(side="left")
        self.source_lang = ttk.Combobox(lang_frame,
                                      values=list(SUPPORTED_LANGUAGES.keys()),
                                      state="readonly",
                                      width=15)
        self.source_lang.set(self.parent.source_lang.get())
        self.source_lang.pack(side="left", padx=5)
        self.source_lang.bind('<<ComboboxSelected>>', self.on_language_change)
        
        # 目标语言选择
        ttk.Label(lang_frame, text="目标语言:").pack(side="left", padx=(10, 0))
        self.target_lang = ttk.Combobox(lang_frame,
                                      values=list(SUPPORTED_LANGUAGES.keys()),
                                      state="readonly",
                                      width=15)
        self.target_lang.set(self.parent.target_lang.get())
        self.target_lang.pack(side="left", padx=5)
        self.target_lang.bind('<<ComboboxSelected>>', self.on_language_change)
        
        # 添加按钮
        ttk.Button(add_frame, text="添加术语",
                  command=self.add_term).pack(pady=5)
        
        # 在术语列表框架上方添加搜索和统计区域
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=(0, 5))
        
        # 左侧搜索框
        search_left = ttk.Frame(search_frame)
        search_left.pack(side="left")
        ttk.Label(search_left, text="搜索:").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_left, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left")
        self.search_var.trace("w", self.on_search_change)
        
        # 右侧统计信息
        self.stats_label = ttk.Label(search_frame, text="总计: 0 个术语")
        self.stats_label.pack(side="right")
        
        # 术语列表
        list_frame = ttk.LabelFrame(main_frame, text="术语列表", padding=10)
        list_frame.pack(fill="both", expand=True)
        
        # 创建表格
        columns = ("源术语", "目标术语")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 添加分页控制区域
        page_frame = ttk.Frame(main_frame)
        page_frame.pack(fill="x", pady=(5, 10))
        
        # 分页按钮和页码显示
        ttk.Button(page_frame, text="<<", command=self.first_page).pack(side="left", padx=2)
        ttk.Button(page_frame, text="<", command=self.prev_page).pack(side="left", padx=2)
        self.page_label = ttk.Label(page_frame, text="第 1 页")
        self.page_label.pack(side="left", padx=10)
        ttk.Button(page_frame, text=">", command=self.next_page).pack(side="left", padx=2)
        ttk.Button(page_frame, text=">>", command=self.last_page).pack(side="left", padx=2)

    def on_language_change(self, event=None):
        """当语言选择改变时重新加载术语表"""
        self.refresh_terminology_list()

    def delete_term(self):
        """删除选中的术语"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择要删除的术语")
            return
            
        for item in selected:
            values = self.tree.item(item)["values"]
            source_term = values[0]
            if source_term in self.current_terms:
                del self.current_terms[source_term]
            
        # 保存更改
        self.save_terminology()
        
        # 直接刷新列表显示
        self.refresh_terminology_list()

    def add_term(self):
        """添加新术语"""
        source = self.source_term.get().strip()
        target = self.target_term.get().strip()
        
        if not source or not target:
            messagebox.showwarning("提示", "请填写完整的术语信息")
            return
            
        # 检查是否存在冲突的术语
        if source in self.current_terms and self.current_terms[source] != target:
            if not messagebox.askyesno("确认", 
                f"源术语 '{source}' 已存在，当前翻译为 '{self.current_terms[source]}'。\n是否要更新为新的翻译 '{target}'？"):
                return
        
        # 添加到当前术语表
        self.current_terms[source] = target
        
        # 先保存术语
        self.save_terminology()
        
        # 刷新列表显示
        self.refresh_terminology_list()
        
        # 清空输入框
        self.source_term.delete(0, "end")
        self.target_term.delete(0, "end")
        
    def on_search_change(self, *args):
        """搜索框内容变化时触发"""
        self.current_page = 1  # 重置到第一页
        self.refresh_terminology_list()

    def filter_terms(self):
        """根据搜索条件过滤术语"""
        search_text = self.search_var.get().lower()
        if not search_text:
            self.filtered_terms = list(sorted(self.current_terms.items()))
        else:
            self.filtered_terms = [
                (source, target) for source, target in sorted(self.current_terms.items())
                if search_text in source.lower() or search_text in target.lower()
            ]
        self.total_terms = len(self.filtered_terms)

    def refresh_terminology_list(self):
        """刷新术语列表显示"""
        # 过滤术语
        self.filter_terms()
        
        # 计算分页
        start_idx = (self.current_page - 1) * self.page_size
        end_idx = start_idx + self.page_size
        current_terms = self.filtered_terms[start_idx:end_idx]
        
        # 清空现有项目
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 添加当前页的术语到表格中
        for source_term, target_term in current_terms:
            self.tree.insert("", "end", values=(source_term, target_term))
        
        # 更新统计信息
        self.stats_label.config(text=f"总计: {self.total_terms} 个术语")
        
        # 更新页码信息
        total_pages = max(1, (self.total_terms + self.page_size - 1) // self.page_size)
        self.page_label.config(text=f"第 {self.current_page}/{total_pages} 页")

    def first_page(self):
        """跳转到第一页"""
        if self.current_page != 1:
            self.current_page = 1
            self.refresh_terminology_list()

    def last_page(self):
        """跳转到最后一页"""
        total_pages = (self.total_terms + self.page_size - 1) // self.page_size
        if self.current_page != total_pages:
            self.current_page = total_pages
            self.refresh_terminology_list()

    def prev_page(self):
        """上一页"""
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_terminology_list()

    def next_page(self):
        """下一页"""
        total_pages = (self.total_terms + self.page_size - 1) // self.page_size
        if self.current_page < total_pages:
            self.current_page += 1
            self.refresh_terminology_list()

    def save_terminology(self):
        """保存术语表（自动保存）"""
        try:
            source_lang = self.source_lang.get()
            target_lang = self.target_lang.get()
            
            # 保存当前方向的术语表
            self.parent.save_terminology(
                source_lang,
                target_lang,
                self.current_terms
            )
            
            # 创建并保存反向的术语表
            reverse_terms = {v: k for k, v in self.current_terms.items()}
            self.parent.save_terminology(
                target_lang,
                source_lang,
                reverse_terms
            )
        except Exception as e:
            messagebox.showerror("错误", f"保存术语表失败: {str(e)}")
        