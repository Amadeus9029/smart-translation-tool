import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil
from pathlib import Path
import os
import datetime

class SubtitleResultFrame(ttk.Frame):
    def __init__(self, master, theme):
        super().__init__(master, style="Modern.TFrame")
        self.theme = theme
        self.create_layout()
        
    def create_layout(self):
        """创建字幕结果管理界面"""
        main_frame = ttk.Frame(self, style="Modern.TFrame")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 顶部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        
        # 刷新按钮
        ttk.Button(control_frame, text="刷新列表",
                  command=self.refresh_list).pack(side="left", padx=5)
        
        # 文件列表区域
        list_frame = ttk.LabelFrame(main_frame, text="翻译结果文件", padding=10)
        list_frame.pack(fill="both", expand=True)
        
        # 创建表格
        columns = ("文件名", "类型", "创建时间")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # 设置列标题和宽度
        self.tree.heading("文件名", text="文件名")
        self.tree.heading("类型", text="类型")
        self.tree.heading("创建时间", text="创建时间")
        
        self.tree.column("文件名", width=300)
        self.tree.column("类型", width=100)
        self.tree.column("创建时间", width=150)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 底部按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="导出选中文件",
                  command=self.export_file).pack(side="left", padx=5)
        ttk.Button(button_frame, text="删除选中文件",
                  command=self.delete_file).pack(side="left", padx=5)
        
        # 初始加载文件列表
        self.refresh_list()
        
    def refresh_list(self):
        """刷新文件列表"""
        # 清空现有项目
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 获取结果目录中的文件
        result_dir = Path("subtitle_result")
        if not result_dir.exists():
            result_dir.mkdir(exist_ok=True)
            return
            
        # 获取所有字幕文件并按修改时间排序
        files = []
        for file_path in result_dir.glob("*.*"):
            if file_path.suffix.lower() in ['.srt', '.ass']:
                files.append(file_path)
        
        # 按修改时间降序排序
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        
        # 添加到列表中
        for file_path in files:
            # 获取文件信息
            create_time = os.path.getmtime(file_path)
            create_time_str = datetime.datetime.fromtimestamp(create_time).strftime('%Y-%m-%d %H:%M:%S')
            
            self.tree.insert("", "end", values=(
                file_path.name,
                file_path.suffix[1:].upper(),
                create_time_str
            ))
            
    def export_file(self):
        """导出选中的文件"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择要导出的文件")
            return
            
        from tkinter import filedialog
        
        # 获取选中的第一个文件
        file_name = self.tree.item(selected[0])["values"][0]
        source_path = Path("subtitle_result") / file_name
        
        # 选择保存位置
        target_path = filedialog.asksaveasfilename(
            defaultextension=source_path.suffix,
            initialfile=file_name,
            filetypes=[
                ("字幕文件", "*.srt;*.ass"),
                ("SRT字幕", "*.srt"),
                ("ASS字幕", "*.ass"),
                ("所有文件", "*.*")
            ]
        )
        
        if target_path:
            try:
                shutil.copy2(source_path, target_path)
                messagebox.showinfo("成功", f"文件已导出至:\n{target_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出文件失败: {str(e)}")
                
    def delete_file(self):
        """删除选中的文件"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请选择要删除的文件")
            return
            
        if not messagebox.askyesno("确认", "确定要删除选中的文件吗？"):
            return
            
        try:
            for item in selected:
                file_name = self.tree.item(item)["values"][0]
                file_path = Path("subtitle_result") / file_name
                file_path.unlink()
                
            self.refresh_list()
            messagebox.showinfo("成功", "文件已删除")
        except Exception as e:
            messagebox.showerror("错误", f"删除文件失败: {str(e)}") 