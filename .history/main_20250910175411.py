#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 图片批量下载器
支持导入 Excel 表格，预览数据，并批量下载图片
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import requests
import os
from urllib.parse import urlparse
import re
from threading import Thread
import time
from PIL import Image, ImageTk

class ImageDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("通用 Excel 图片批量下载器")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f0f0f0")
        
        # 数据存储
        self.df = None
        self.download_folder = ""
        self.filename_pattern = "{column0}_{column1}.png"
        self.available_columns = []
        self.url_column = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主窗口的网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = tk.Label(main_frame, text="Excel 图片批量下载器", 
                              font=("Arial", 16, "bold"), bg="#f0f0f0")
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Excel 文件选择区域
        excel_frame = ttk.LabelFrame(main_frame, text="选择 Excel 文件", padding="10")
        excel_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        excel_frame.columnconfigure(1, weight=1)
        
        ttk.Label(excel_frame, text="Excel 文件:").grid(row=0, column=0, padx=(0, 10))
        self.excel_path_var = tk.StringVar()
        self.excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_path_var, width=50)
        self.excel_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(excel_frame, text="浏览", 
                  command=self.browse_excel_file).grid(row=0, column=2)
        
        ttk.Button(excel_frame, text="导入数据", 
                  command=self.import_excel_data).grid(row=0, column=3, padx=(10, 0))
        
        # 数据预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览", padding="10")
        preview_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # 表格
        self.tree = ttk.Treeview(preview_frame, columns=("name", "club", "photo_url"), show="headings")
        self.tree.heading("name", text="姓名")
        self.tree.heading("club", text="俱乐部名称")
        self.tree.heading("photo_url", text="照片链接")
        
        # 设置列宽
        self.tree.column("name", width=150)
        self.tree.column("club", width=200)
        self.tree.column("photo_url", width=400)
        
        # 滚动条
        v_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 下载设置区域
        download_frame = ttk.LabelFrame(main_frame, text="下载设置", padding="10")
        download_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        download_frame.columnconfigure(1, weight=1)
        
        # 下载文件夹选择
        ttk.Label(download_frame, text="下载文件夹:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.folder_path_var = tk.StringVar()
        ttk.Entry(download_frame, textvariable=self.folder_path_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(download_frame, text="选择文件夹", 
                  command=self.browse_download_folder).grid(row=0, column=2)
        
        # 文件命名格式设置
        ttk.Label(download_frame, text="文件命名格式:").grid(row=1, column=0, padx=(0, 10), sticky=tk.W, pady=(10, 0))
        
        format_frame = ttk.Frame(download_frame)
        format_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        format_frame.columnconfigure(0, weight=1)
        
        self.pattern_var = tk.StringVar(value=self.filename_pattern)
        pattern_entry = ttk.Entry(format_frame, textvariable=self.pattern_var, width=30)
        pattern_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(format_frame, text="可用变量: {name} {club}").grid(row=0, column=1)
        
        # 示例显示
        self.example_label = ttk.Label(format_frame, text="示例: ", foreground="blue")
        self.example_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # 绑定格式变化事件
        self.pattern_var.trace('w', self.update_filename_example)
        self.update_filename_example()
        
        # 下载控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=4, column=0, columnspan=3, pady=(0, 10))
        
        self.download_button = ttk.Button(control_frame, text="开始批量下载", 
                                        command=self.start_download, state="disabled")
        self.download_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 10))
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(control_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT)
        
        # 配置主框架的网格权重
        main_frame.rowconfigure(2, weight=1)
    
    def browse_excel_file(self):
        """浏览选择 Excel 文件"""
        filename = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
    
    def browse_download_folder(self):
        """浏览选择下载文件夹"""
        folder = filedialog.askdirectory(title="选择下载文件夹")
        if folder:
            self.folder_path_var.set(folder)
            self.download_folder = folder
    
    def import_excel_data(self):
        """导入 Excel 数据"""
        excel_path = self.excel_path_var.get()
        if not excel_path:
            messagebox.showerror("错误", "请先选择 Excel 文件")
            return
        
        try:
            # 读取 Excel 文件
            self.df = pd.read_excel(excel_path)
            
            # 检查必需的列是否存在
            required_columns = ['姓名', '俱乐部名称', '照片链接']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                messagebox.showerror("错误", f"Excel 文件缺少以下列: {', '.join(missing_columns)}")
                return
            
            # 清空现有数据
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 填充数据到表格
            for index, row in self.df.iterrows():
                self.tree.insert("", "end", values=(
                    row['姓名'], 
                    row['俱乐部名称'], 
                    row['照片链接']
                ))
            
            self.status_var.set(f"成功导入 {len(self.df)} 条记录")
            self.download_button.config(state="normal")
            
        except Exception as e:
            messagebox.showerror("错误", f"导入数据失败: {str(e)}")
    
    def update_filename_example(self, *args):
        """更新文件名示例"""
        pattern = self.pattern_var.get()
        try:
            example = pattern.format(name="王晓燕", club="天天俱乐部")
            self.example_label.config(text=f"示例: {example}")
        except:
            self.example_label.config(text="示例: 格式错误")
    
    def start_download(self):
        """开始下载"""
        if self.df is None or len(self.df) == 0:
            messagebox.showerror("错误", "没有数据可下载")
            return
        
        if not self.download_folder:
            messagebox.showerror("错误", "请选择下载文件夹")
            return
        
        # 在新线程中执行下载
        self.download_button.config(state="disabled")
        thread = Thread(target=self.download_images)
        thread.daemon = True
        thread.start()
    
    def download_images(self):
        """下载图片（在后台线程中执行）"""
        total_count = len(self.df)
        success_count = 0
        failed_count = 0
        
        try:
            for index, row in self.df.iterrows():
                # 更新进度
                progress = (index / total_count) * 100
                self.progress_var.set(progress)
                
                name = str(row['姓名']).strip()
                club = str(row['俱乐部名称']).strip()
                photo_url = str(row['照片链接']).strip()
                
                self.status_var.set(f"正在下载: {name}")
                
                try:
                    # 生成文件名
                    filename = self.pattern_var.get().format(name=name, club=club)
                    # 清理文件名中的非法字符
                    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                    
                    # 确保有扩展名
                    if not filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                        filename += '.png'
                    
                    filepath = os.path.join(self.download_folder, filename)
                    
                    # 下载图片
                    if self.download_single_image(photo_url, filepath):
                        success_count += 1
                    else:
                        failed_count += 1
                        
                except Exception as e:
                    print(f"下载 {name} 的图片时出错: {str(e)}")
                    failed_count += 1
            
            # 下载完成
            self.progress_var.set(100)
            self.status_var.set(f"下载完成: 成功 {success_count}, 失败 {failed_count}")
            
            # 显示完成消息
            self.root.after(0, lambda: messagebox.showinfo(
                "下载完成", 
                f"批量下载完成!\n成功: {success_count} 张\n失败: {failed_count} 张"
            ))
            
        except Exception as e:
            self.status_var.set(f"下载过程中出错: {str(e)}")
        
        finally:
            # 重新启用下载按钮
            self.root.after(0, lambda: self.download_button.config(state="normal"))
    
    def download_single_image(self, url, filepath):
        """下载单张图片"""
        try:
            # 发送请求
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            
            # 保存图片
            with open(filepath, 'wb') as f:
                f.write(response.content)
            
            return True
            
        except Exception as e:
            print(f"下载图片失败 {url}: {str(e)}")
            return False


def main():
    """主函数"""
    root = tk.Tk()
    app = ImageDownloader(root)
    root.mainloop()


if __name__ == "__main__":
    main()
