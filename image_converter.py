# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import threading

class ImageConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("图片格式转换器")
        self.root.geometry("550x480")
        self.root.resizable(False, False)
        
        self.input_files = []
        self.output_dir = tk.StringVar()
        self.output_format = tk.StringVar(value="png")
        
        self.setup_ui()
    
    def setup_ui(self):
        style = ttk.Style()
        style.configure("TLabel", font=("微软雅黑", 10))
        style.configure("TButton", font=("微软雅黑", 9))
        
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        row = 0
        
        ttk.Label(main_frame, text="选择图片文件:").grid(row=row, column=0, sticky=tk.W, pady=8)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=row, column=1, columnspan=2, sticky=tk.W, pady=5)
        ttk.Button(btn_frame, text="浏览文件", command=self.select_files, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="浏览文件夹", command=self.select_input, width=12).pack(side=tk.LEFT, padx=2)
        row += 1
        
        self.file_count_label = ttk.Label(main_frame, text="已选择: 0 个文件", foreground="blue")
        self.file_count_label.grid(row=row, column=1, columnspan=2, sticky=tk.W, pady=5)
        row += 1
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        ttk.Label(main_frame, text="输出文件夹:").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=45).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_output, width=10).grid(row=row, column=2, pady=5)
        row += 1
        
        ttk.Label(main_frame, text="目标格式:").grid(row=row, column=0, sticky=tk.W, pady=8)
        formats = ["png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp", "ico"]
        ttk.Combobox(main_frame, textvariable=self.output_format, values=formats, state="readonly", width=42).grid(row=row, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)
        row += 1
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        self.convert_btn = ttk.Button(main_frame, text="开始转换", command=self.start_convert, width=15)
        self.convert_btn.grid(row=row, column=1, pady=15)
        row += 1
        
        ttk.Label(main_frame, text="转换日志:").grid(row=row, column=0, sticky=tk.NW, pady=5)
        row += 1
        
        self.log_text = tk.Text(main_frame, width=58, height=13, font=("Consolas", 9))
        self.log_text.grid(row=row, column=0, columnspan=3, pady=5)
        
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=row, column=3, sticky=tk.NS, pady=5)
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择图片文件",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.webp *.ico"), ("所有文件", "*.*")]
        )
        if files:
            self.input_files = list(files)
            self.file_count_label.config(text=f"已选择: {len(files)} 个文件")
            self.log(f"已添加 {len(files)} 个文件")
    
    def select_input(self):
        path = filedialog.askdirectory(title="选择文件夹")
        if path:
            supported = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp', '.ico'}
            files = [os.path.join(path, f) for f in os.listdir(path) 
                     if os.path.splitext(f)[1].lower() in supported]
            if files:
                self.input_files = files
                self.file_count_label.config(text=f"已选择: {len(files)} 个文件 (来自文件夹)")
                self.log(f"已添加文件夹: {len(files)} 个文件")
            else:
                messagebox.showwarning("警告", "文件夹中没有找到支持的图片文件")
    
    def select_output(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir.set(path)
    
    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def convert_image(self, input_path, output_path):
        try:
            img = Image.open(input_path)
            img.save(output_path)
            return True
        except Exception as e:
            return False
    
    def start_convert(self):
        output_dir = self.output_dir.get()
        output_format = self.output_format.get()
        
        if not self.input_files:
            messagebox.showerror("错误", "请先选择图片文件或文件夹")
            return
        
        if not output_dir:
            messagebox.showerror("错误", "请选择输出文件夹")
            return
        
        self.convert_btn.config(state=tk.DISABLED)
        files_copy = self.input_files.copy()
        thread = threading.Thread(target=self.run_convert, args=(files_copy, output_dir, output_format))
        thread.start()
    
    def run_convert(self, files_to_convert, output_dir, output_format):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        total = len(files_to_convert)
        success_count = 0
        
        self.log(f"开始转换 {total} 个文件...")
        self.log("-" * 40)
        
        for i, input_path in enumerate(files_to_convert, 1):
            filename = os.path.basename(input_path)
            name = os.path.splitext(filename)[0]
            output_filename = f"{name}.{output_format}"
            output_path = os.path.join(output_dir, output_filename)
            
            self.log(f"[{i}/{total}] 转换中: {filename}")
            if self.convert_image(input_path, output_path):
                success_count += 1
                self.log(f"    ✓ 成功 -> {output_filename}")
            else:
                self.log(f"    ✗ 失败")
        
        self.log("-" * 40)
        self.log(f"转换完成: 成功 {success_count}/{total}")
        
        self.root.after(0, lambda: messagebox.showinfo("完成", f"转换完成!\n成功: {success_count}/{total}"))
        self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.input_files.clear())
        self.root.after(0, lambda: self.file_count_label.config(text="已选择: 0 个文件"))

def main():
    if len(sys.argv) >= 4:
        input_dir = sys.argv[1]
        output_dir = sys.argv[2]
        output_format = sys.argv[3].lower()
        
        if not os.path.exists(input_dir):
            print(f"输入目录不存在: {input_dir}")
            sys.exit(1)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        supported = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp', '.ico'}
        files = [f for f in os.listdir(input_dir) 
                 if os.path.splitext(f)[1].lower() in supported]
        
        for filename in files:
            input_path = os.path.join(input_dir, filename)
            name = os.path.splitext(filename)[0]
            output_filename = f"{name}.{output_format}"
            output_path = os.path.join(output_dir, output_filename)
            
            try:
                img = Image.open(input_path)
                img.save(output_path)
                print(f"成功: {filename} -> {output_filename}")
            except Exception as e:
                print(f"失败: {filename} - {e}")
        
        print(f"\n完成: {len(files)} 个文件")
    else:
        root = tk.Tk()
        app = ImageConverterGUI(root)
        root.mainloop()

if __name__ == "__main__":
    main()
