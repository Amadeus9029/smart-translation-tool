import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
from pathlib import Path
import chardet
from openai import OpenAI
import threading
from constants import SUPPORTED_LANGUAGES
import datetime

class SubtitleTranslateFrame(ttk.Frame):
    def __init__(self, master, theme, api_key):
        super().__init__(master, style="Modern.TFrame")
        self.theme = theme
        self.api_key = api_key
        self.DEEPSEEK_BASE_URL = "https://api.deepseek.com"
        self.is_translating = False
        
        # 字幕文件相关变量
        self.source_file = None
        self.subtitle_content = []
        self.translated_content = []
        
        self.create_layout()
        
    def create_layout(self):
        """创建字幕翻译界面"""
        # 主框架
        main_frame = ttk.Frame(self, style="Modern.TFrame")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 顶部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(control_frame, text="字幕文件", padding=10)
        file_frame.pack(fill="x", pady=(0, 10))
        
        # 文件路径显示
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, state="readonly", width=50).pack(side="left", padx=5)
        
        # 选择文件按钮
        ttk.Button(file_frame, text="选择文件", command=self.select_file).pack(side="left", padx=5)
        
        # 语言选择区域
        lang_frame = ttk.Frame(control_frame)
        lang_frame.pack(fill="x", pady=5)
        
        # 源语言选择
        ttk.Label(lang_frame, text="源语言:").pack(side="left")
        self.source_lang = ttk.Combobox(lang_frame,
                                      values=list(SUPPORTED_LANGUAGES.keys()),
                                      state="readonly",
                                      width=15)
        self.source_lang.set("英语")
        self.source_lang.pack(side="left", padx=5)
        
        # 目标语言选择
        ttk.Label(lang_frame, text="目标语言:").pack(side="left", padx=(10, 0))
        self.target_lang = ttk.Combobox(lang_frame,
                                      values=list(SUPPORTED_LANGUAGES.keys()),
                                      state="readonly",
                                      width=15)
        self.target_lang.set("中文")
        self.target_lang.pack(side="left", padx=5)
        
        # 翻译按钮
        self.translate_btn = ttk.Button(lang_frame, text="开始翻译",
                                      command=self.start_translation)
        self.translate_btn.pack(side="right", padx=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="字幕预览", padding=10)
        preview_frame.pack(fill="both", expand=True)
        
        # 创建左右分栏
        paned = ttk.PanedWindow(preview_frame, orient="horizontal")
        paned.pack(fill="both", expand=True)
        
        # 左侧原文
        left_frame = ttk.Frame(paned, width=350)  # 设置初始宽度
        left_frame.pack_propagate(False)  # 防止子组件影响框架大小
        ttk.Label(left_frame, text="原文字幕").pack()
        
        # 添加滚动条
        left_scroll = ttk.Scrollbar(left_frame)
        left_scroll.pack(side="right", fill="y")
        
        self.source_text = tk.Text(left_frame, wrap="word",
                                 bg=self.theme.ENTRY_BG,
                                 fg=self.theme.FOREGROUND,
                                 yscrollcommand=left_scroll.set)
        self.source_text.pack(fill="both", expand=True)
        left_scroll.config(command=self.source_text.yview)
        paned.add(left_frame, weight=1)  # weight=1 表示均分空间
        
        # 右侧译文
        right_frame = ttk.Frame(paned, width=350)  # 设置初始宽度
        right_frame.pack_propagate(False)  # 防止子组件影响框架大小
        ttk.Label(right_frame, text="翻译字幕").pack()
        
        # 添加滚动条
        right_scroll = ttk.Scrollbar(right_frame)
        right_scroll.pack(side="right", fill="y")
        
        self.target_text = tk.Text(right_frame, wrap="word",
                                 bg=self.theme.ENTRY_BG,
                                 fg=self.theme.FOREGROUND,
                                 yscrollcommand=right_scroll.set)
        self.target_text.pack(fill="both", expand=True)
        right_scroll.config(command=self.target_text.yview)
        paned.add(right_frame, weight=1)  # weight=1 表示均分空间
        
        # 底部状态栏
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.pack(fill="x", pady=(5, 0))
        
    def select_file(self):
        """选择字幕文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("字幕文件", "*.srt;*.ass"),
                ("SRT字幕", "*.srt"),
                ("ASS字幕", "*.ass"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.source_file = Path(file_path)
            self.file_path.set(file_path)
            self.load_subtitle()
            
    def load_subtitle(self):
        """加载字幕文件"""
        try:
            # 检测文件编码
            with open(self.source_file, 'rb') as f:
                raw_data = f.read()
                result = chardet.detect(raw_data)
                encoding = result['encoding']
            
            # 读取文件内容
            with open(self.source_file, 'r', encoding=encoding) as f:
                content = f.read()
            
            # 根据文件类型解析字幕
            if self.source_file.suffix.lower() == '.srt':
                self.subtitle_content = self.parse_srt(content)
            else:  # .ass
                self.subtitle_content = self.parse_ass(content)
            
            # 显示原文字幕
            self.source_text.delete('1.0', tk.END)
            preview_text = '\n'.join([item['text'] for item in self.subtitle_content])
            self.source_text.insert('1.0', preview_text)
            
            self.status_label.config(text=f"已加载字幕文件: {len(self.subtitle_content)} 条字幕")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载字幕文件失败: {str(e)}")
            
    def parse_srt(self, content):
        """解析SRT格式字幕"""
        subtitle_list = []
        pattern = re.compile(r'(\d+)\n(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})\n((?:.*\n)*?)\n')
        
        matches = pattern.finditer(content)
        for match in matches:
            subtitle_list.append({
                'index': match.group(1),
                'start': match.group(2),
                'end': match.group(3),
                'text': match.group(4).strip(),
                'type': 'srt'
            })
        
        return subtitle_list
        
    def parse_ass(self, content):
        """解析ASS格式字幕"""
        subtitle_list = []
        dialogue_pattern = re.compile(r'Dialogue: [^,]*,([^,]*),([^,]*),([^,]*),([^,]*),([^,]*),([^,]*),([^,]*),([^,]*),(.*)')
        
        for line in content.split('\n'):
            if line.startswith('Dialogue:'):
                match = dialogue_pattern.match(line)
                if match:
                    subtitle_list.append({
                        'start': match.group(1),
                        'end': match.group(2),
                        'style': match.group(3),
                        'text': match.group(9),
                        'type': 'ass'
                    })
        
        return subtitle_list
        
    def start_translation(self):
        """开始翻译"""
        if not self.subtitle_content:
            messagebox.showwarning("提示", "请先选择字幕文件")
            return
            
        if self.is_translating:
            return
            
        self.is_translating = True
        self.translate_btn.config(state="disabled")
        self.status_label.config(text="正在翻译...")
        
        # 在新线程中执行翻译
        threading.Thread(target=self._do_translate, daemon=True).start()
        
    def _do_translate(self):
        """执行翻译"""
        try:
            client = OpenAI(api_key=self.api_key, base_url=self.DEEPSEEK_BASE_URL)
            self.translated_content = []
            
            # 批量翻译，每次处理10条字幕
            batch_size = 10
            for i in range(0, len(self.subtitle_content), batch_size):
                batch = self.subtitle_content[i:i + batch_size]
                texts = [item['text'] for item in batch]
                
                # 构建提示词
                prompt = f"请将以下{self.source_lang.get()}字幕翻译成{self.target_lang.get()}，保持原文的格式和语气：\n\n"
                prompt += "\n---\n".join(texts)
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "你是一个专业的字幕翻译专家。"},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                    max_tokens=2000
                )
                
                # 解析翻译结果
                translations = response.choices[0].message.content.strip().split("\n---\n")
                
                # 更新翻译结果
                for j, translation in enumerate(translations):
                    item = batch[j].copy()
                    item['translation'] = translation.strip()
                    self.translated_content.append(item)
                
                # 更新进度
                progress = min(100, int((i + len(batch)) / len(self.subtitle_content) * 100))
                self.status_label.config(text=f"翻译进度: {progress}%")
            
            # 显示翻译结果
            self.show_translation()
            
            # 保存翻译结果
            self.save_translation()
            
        except Exception as e:
            messagebox.showerror("错误", f"翻译失败: {str(e)}")
        finally:
            self.is_translating = False
            self.translate_btn.config(state="normal")
            self.status_label.config(text="翻译完成")
            
    def show_translation(self):
        """显示翻译结果"""
        self.target_text.delete('1.0', tk.END)
        preview_text = '\n'.join([item['translation'] for item in self.translated_content])
        self.target_text.insert('1.0', preview_text)
        
    def save_translation(self):
        """保存翻译后的字幕文件"""
        try:
            # 创建结果目录
            result_dir = Path("subtitle_result")
            result_dir.mkdir(exist_ok=True)
            
            # 构建输出文件路径（使用原文件名+目标语言作为新文件名）
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = result_dir / f"{self.source_file.stem}_{self.target_lang.get()}_{timestamp}{self.source_file.suffix}"
            
            # 根据字幕类型生成内容
            if self.subtitle_content[0]['type'] == 'srt':
                self.save_srt(output_path)
            else:
                self.save_ass(output_path)
                
            # 更新状态
            self.status_label.config(text=f"翻译完成，已保存至: {output_path.name}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存翻译文件失败: {str(e)}")
            
    def save_srt(self, output_path):
        """保存SRT格式字幕"""
        with open(output_path, 'w', encoding='utf-8') as f:
            for i, item in enumerate(self.translated_content, 1):
                f.write(f"{i}\n")
                f.write(f"{item['start']} --> {item['end']}\n")
                f.write(f"{item['translation']}\n\n")
                
    def save_ass(self, output_path):
        """保存ASS格式字幕"""
        # 首先复制原文件的样式部分
        with open(self.source_file, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 找到[Events]部分
        parts = content.split('[Events]')
        if len(parts) != 2:
            raise ValueError("无效的ASS文件格式")
            
        header = parts[0] + '[Events]\n'
        
        # 写入新文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(header)
            # 写入翻译后的对话行
            for item in self.translated_content:
                dialogue = f"Dialogue: 0,{item['start']},{item['end']},{item['style']},,0,0,0,,{item['translation']}\n"
                f.write(dialogue) 