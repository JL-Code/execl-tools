#!/usr/bin/env python3
"""
Excel拆分GUI工具
用于将大Excel文件拆分为多个小Excel文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from pathlib import Path
import threading
import subprocess
import platform


class ExcelSplitterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel文件拆分工具")
        self.root.geometry("800x650")
        self.root.resizable(True, True)
        self.root.minsize(750, 600)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 配置自定义样式
        self.style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#2c3e50')
        self.style.configure('Heading.TLabel', font=('Arial', 10, 'bold'), foreground='#34495e')
        self.style.configure('Accent.TButton', font=('Arial', 11, 'bold'))
        self.style.map('Accent.TButton', 
                      background=[('active', '#27ae60'), ('!active', '#2ecc71')],
                      foreground=[('active', 'white'), ('!active', 'white')])
        self.style.configure('Secondary.TButton', font=('Arial', 10))
        self.style.map('Secondary.TButton',
                      background=[('active', '#95a5a6'), ('!active', '#bdc3c7')],
                      foreground=[('active', '#2c3e50'), ('!active', '#2c3e50')])
        
        # 配置进度条样式
        self.style.configure('TProgressbar', 
                           background='#3498db',
                           troughcolor='#ecf0f1',
                           borderwidth=0,
                           lightcolor='#3498db',
                           darkcolor='#3498db')
        
        # 设置整体背景色
        self.root.configure(bg='#ecf0f1')
        
        # 变量
        self.input_file_path = tk.StringVar()
        self.rows_per_file = tk.StringVar(value="50")
        self.output_dir = tk.StringVar()
        
        # 存储文件信息
        self.current_file_info = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)  # 让文件信息区域可以扩展
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel文件拆分工具", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # 输入文件选择
        ttk.Label(main_frame, text="选择Excel文件:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=(0, 10))
        input_entry = ttk.Entry(main_frame, textvariable=self.input_file_path, 
                               width=60, font=('Arial', 10))
        input_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=(0, 10))
        ttk.Button(main_frame, text="浏览", command=self.browse_input_file,
                  style='Secondary.TButton').grid(row=1, column=2, pady=(0, 10))
        
        # 每个文件行数设置
        ttk.Label(main_frame, text="每个小文件行数:", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=(0, 10))
        rows_entry = ttk.Entry(main_frame, textvariable=self.rows_per_file, 
                              width=25, font=('Arial', 10))
        rows_entry.grid(row=2, column=1, sticky=tk.W, padx=(15, 0), pady=(0, 10))
        
        # 绑定行数输入框的变化事件
        self.rows_per_file.trace_add('write', self.on_rows_changed)
        
        # 输出目录选择
        ttk.Label(main_frame, text="输出目录:", style='Heading.TLabel').grid(
            row=3, column=0, sticky=tk.W, pady=(0, 15))
        output_entry = ttk.Entry(main_frame, textvariable=self.output_dir, 
                                width=60, font=('Arial', 10))
        output_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=(0, 15))
        ttk.Button(main_frame, text="浏览", command=self.browse_output_dir,
                  style='Secondary.TButton').grid(row=3, column=2, pady=(0, 15))
        
        # 文件信息显示区域
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="15")
        info_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)
        
        self.info_text = tk.Text(info_frame, height=8, width=80, wrap=tk.WORD,
                                font=('Consolas', 10), bg='#f8f9fa', fg='#2c3e50',
                                relief=tk.FLAT, borderwidth=1)
        scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          maximum=100, length=500)
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(0, 10))
        
        self.split_button = ttk.Button(button_frame, text="开始拆分", 
                                     command=self.start_split, style='Accent.TButton',
                                     width=15)
        self.split_button.pack(side=tk.LEFT, padx=(0, 15))
        
        clear_button = ttk.Button(button_frame, text="清空信息", command=self.clear_info,
                                 style='Secondary.TButton', width=12)
        clear_button.pack(side=tk.LEFT, padx=(0, 15))
        
        self.open_folder_button = ttk.Button(button_frame, text="打开输出目录", 
                                           command=self.open_output_folder,
                                           style='Secondary.TButton', width=15)
        self.open_folder_button.pack(side=tk.LEFT)
        # 初始状态下禁用按钮，直到有输出目录
        self.open_folder_button.config(state='disabled')
        
        # 初始化信息
        self.add_info("欢迎使用Excel文件拆分工具！")
        self.add_info("请选择要拆分的Excel文件，设置每个小文件的行数，然后点击开始拆分。")
        
    def browse_input_file(self):
        """浏览输入文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            # 自动设置输出目录为输入文件所在目录
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(file_path))
                self.open_folder_button.config(state='enabled')
            self.analyze_file()
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir.set(dir_path)
            # 当选择了输出目录后，启用打开目录按钮
            self.open_folder_button.config(state='normal')
    
    def analyze_file(self):
        """分析选中的Excel文件"""
        try:
            file_path = self.input_file_path.get()
            if not file_path:
                return
                
            # 读取Excel文件信息
            if file_path.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl')
            elif file_path.lower().endswith('.xls'):
                # .xls 文件使用 xlrd 引擎
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                df = pd.read_excel(file_path)
            
            total_rows = len(df)
            columns = len(df.columns)
            
            # 存储文件信息
            self.current_file_info = {
                'path': file_path,
                'total_rows': total_rows,
                'columns': columns
            }
            
            self.add_info(f"文件分析完成:")
            self.add_info(f"  文件路径: {file_path}")
            self.add_info(f"  总行数: {total_rows}")
            self.add_info(f"  列数: {columns}")
            
            # 计算拆分信息
            self.update_split_info()
                
        except Exception as e:
            self.add_info(f"文件分析失败: {str(e)}")
            if "xlrd" in str(e).lower():
                self.add_info("提示: 如果是.xls文件，请确保已安装xlrd包")
                self.add_info("可以运行: pip install xlrd>=2.0.1")
    
    def update_split_info(self):
        """更新拆分信息显示"""
        if not self.current_file_info:
            return
            
        try:
            rows_per_file = int(self.rows_per_file.get())
            if rows_per_file > 0:
                total_rows = self.current_file_info['total_rows']
                num_files = (total_rows + rows_per_file - 1) // rows_per_file
                self.add_info(f"  按每文件{rows_per_file}行拆分，将生成{num_files}个文件")
        except ValueError:
            pass
    
    def on_rows_changed(self, *args):
        """当行数输入发生变化时的回调函数"""
        if self.current_file_info:
            self.add_info("--- 重新计算拆分信息 ---")
            self.update_split_info()
    
    def add_info(self, message):
        """添加信息到信息显示区域"""
        self.info_text.insert(tk.END, message + "\n")
        self.info_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_info(self):
        """清空信息显示"""
        self.info_text.delete(1.0, tk.END)
        # 清空信息时也禁用打开目录按钮
        if not self.output_dir.get():
            self.open_folder_button.config(state='disabled')
    
    def start_split(self):
        """开始拆分文件"""
        # 验证输入
        if not self.input_file_path.get():
            messagebox.showerror("错误", "请选择要拆分的Excel文件")
            return
        
        if not self.output_dir.get():
            messagebox.showerror("错误", "请选择输出目录")
            return
        
        try:
            rows_per_file = int(self.rows_per_file.get())
            if rows_per_file <= 0:
                raise ValueError("行数必须大于0")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的行数（正整数）")
            return
        
        # 在新线程中执行拆分操作
        self.split_button.config(state='disabled')
        thread = threading.Thread(target=self.split_excel_file)
        thread.daemon = True
        thread.start()
    
    def split_excel_file(self):
        """拆分Excel文件的核心逻辑"""
        try:
            input_file = self.input_file_path.get()
            output_dir = self.output_dir.get()
            rows_per_file = int(self.rows_per_file.get())
            
            self.add_info("开始拆分文件...")
            self.progress_var.set(0)
            
            # 读取Excel文件
            if input_file.lower().endswith('.xlsx'):
                df = pd.read_excel(input_file, engine='openpyxl')
            elif input_file.lower().endswith('.xls'):
                # .xls 文件使用 xlrd 引擎
                df = pd.read_excel(input_file, engine='xlrd')
            else:
                df = pd.read_excel(input_file)
            
            total_rows = len(df)
            
            # 获取原文件名（不含扩展名）
            input_filename = Path(input_file).stem
            file_extension = Path(input_file).suffix
            
            # 计算需要生成的文件数量
            num_files = (total_rows + rows_per_file - 1) // rows_per_file
            
            self.add_info(f"总共{total_rows}行数据，将拆分为{num_files}个文件")
            
            # 创建输出目录（如果不存在）
            os.makedirs(output_dir, exist_ok=True)
            
            # 拆分文件
            for i in range(num_files):
                start_row = i * rows_per_file
                end_row = min((i + 1) * rows_per_file, total_rows)
                
                # 提取数据
                chunk_df = df.iloc[start_row:end_row]
                
                # 生成输出文件名 - 统一输出为xlsx格式
                output_filename = f"{input_filename}_{i+1:03d}.xlsx"
                output_path = os.path.join(output_dir, output_filename)
                
                # 保存文件 - 统一使用openpyxl引擎保存为xlsx格式
                chunk_df.to_excel(output_path, index=False, engine='openpyxl')
                
                # 更新进度
                progress = ((i + 1) / num_files) * 100
                self.progress_var.set(progress)
                
                self.add_info(f"已生成: {output_filename} ({end_row - start_row}行)")
            
            self.add_info("拆分完成！")
            self.add_info(f"输出目录: {output_dir}")
            messagebox.showinfo("完成", f"文件拆分完成！\n共生成{num_files}个文件")
            
        except Exception as e:
            error_msg = f"拆分失败: {str(e)}"
            self.add_info(error_msg)
            messagebox.showerror("错误", error_msg)
        
        finally:
            self.split_button.config(state='normal')
            self.progress_var.set(0)
    
    def open_output_folder(self):
        """打开输出文件夹"""
        output_path = self.output_dir.get()
        
        if not output_path:
            messagebox.showwarning("警告", "请先选择输出目录")
            return
        
        if not os.path.exists(output_path):
            messagebox.showerror("错误", "输出目录不存在")
            return
        
        try:
            # 根据操作系统选择合适的命令打开文件夹
            system = platform.system()
            if system == "Windows":
                os.startfile(output_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", output_path])
            else:  # Linux and other Unix-like systems
                subprocess.run(["xdg-open", output_path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")


def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelSplitterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()