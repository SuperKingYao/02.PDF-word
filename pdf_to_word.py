#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF转Word脚本
需要安装: pip install pdf2docx
"""

from pdf2docx import Converter
import os
import sys


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


if __name__ == "__main__":
    # 使用示例
    # 1. 单个文件转换
    pdf_to_word("JPT网口打标卡网络配置说明.pdf")

    # 2. 或者指定输出文件名
    # pdf_to_word("input.pdf", "output.docx")

    # 3. 批量转换（取消下面注释使用）
    # batch_convert("./pdfs")
