#!/usr/bin/env python3
# -*- coding -*-
"""
PDF转Word脚本
需要安装: pip install: utf-8 pdf2docx
"""

from pdf2docx import Converter
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading


def pdf_to_word(pdf_file, docx_file=None):
    """
    将PDF文件转换为Word文档

    Args:
        pdf_file: PDF文件路径
        docx_file: 输出的Word文件路径（可选，默认同目录同名）
    """
    if not os.path.exists(pdf_file):
        print(f"错误: 文件不存在 - {pdf_file}")
        return False

    if docx_file is None:
        docx_file = os.path.splitext(pdf_file)[0] + '.docx'

    try:
        print(f"正在转换: {pdf_file}")
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=0, end=None)
        cv.close()
        print(f"转换完成: {docx_file}")
        return True
    except Exception as e:
        print(f"转换失败: {e}")
        return False


def batch_convert(pdf_folder):
    """
    批量转换文件夹下所有PDF文件

    Args:
        pdf_folder: 包含PDF文件的文件夹路径
    """
    if not os.path.isdir(pdf_folder):
        print(f"错误: 目录不存在 - {pdf_folder}")
        return

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print("未找到PDF文件")
        return

    print(f"找到 {len(pdf_files)} 个PDF文件，开始转换...\n")

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        pdf_to_word(pdf_path)


class PDFToWordGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF转Word工具")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        self.pdf_file = tk.StringVar()

        tk.Label(root, text="PDF转Word转换器", font=("Arial", 16, "bold")).pack(pady=10)

        frame = tk.Frame(root)
        frame.pack(pady=10, padx=20, fill=tk.X)

        tk.Label(frame, text="选择PDF文件:").grid(row=0, column=0, sticky=tk.W)
        tk.Entry(frame, textvariable=self.pdf_file, width=40).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="浏览", command=self.select_file).grid(row=0, column=2)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="开始转换", command=self.start_convert, width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="批量转换", command=self.batch_convert, width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="退出", command=root.quit, width=10, height=2).pack(side=tk.LEFT, padx=10)

        tk.Label(root, text="转换日志:").pack(pady=5)
        self.log_text = scrolledtext.ScrolledText(root, width=70, height=18, font=("Consolas", 9))
        self.log_text.pack(pady=5, padx=20, fill=tk.BOTH, expand=True)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if filename:
            self.pdf_file.set(filename)
            self.log(f"已选择: {filename}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_convert(self):
        pdf_path = self.pdf_file.get().strip()
        if not pdf_path:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return

        threading.Thread(target=self._convert_file, args=(pdf_path,), daemon=True).start()

    def _convert_file(self, pdf_path):
        self.log(f"\n开始转换: {pdf_path}")
        docx_path = os.path.splitext(pdf_path)[0] + '.docx'
        success = pdf_to_word(pdf_path, docx_path)
        if success:
            self.log(f"转换成功!")
            self.root.after(0, lambda: messagebox.showinfo("完成", f"转换成功!\n保存至: {docx_path}"))
        else:
            self.log("转换失败!")
            self.root.after(0, lambda: messagebox.showerror("错误", "转换失败，请查看日志"))

    def batch_convert(self):
        folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if not folder:
            return

        threading.Thread(target=self._batch_convert_folder, args=(folder,), daemon=True).start()

    def _batch_convert_folder(self, folder):
        self.log(f"\n批量转换文件夹: {folder}")
        pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
        if not pdf_files:
            self.log("未找到PDF文件")
            self.root.after(0, lambda: messagebox.showwarning("警告", "未找到PDF文件"))
            return

        self.log(f"找到 {len(pdf_files)} 个PDF文件\n")
        count = 0
        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder, pdf_file)
            self.log(f"正在转换: {pdf_file}")
            if pdf_to_word(pdf_path):
                count += 1

        self.log(f"\n批量转换完成: {count}/{len(pdf_files)} 个文件转换成功")
        self.root.after(0, lambda: messagebox.showinfo("完成", f"批量转换完成!\n{count}/{len(pdf_files)} 个文件转换成功"))


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        root = tk.Tk()
        app = PDFToWordGUI(root)
        root.mainloop()
    else:
        pdf_to_word("JPT网口打标卡网络配置说明.pdf")
