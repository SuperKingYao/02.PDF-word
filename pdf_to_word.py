#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF转Word脚本
使用 pdf2docx 库将PDF文件转换为Word文档
"""

import os
import sys
from pathlib import Path
from pdf2docx import convert


def pdf_to_word(pdf_path, output_path=None, pages=None):
    """
    将PDF文件转换为Word文档
    
    参数:
        pdf_path (str): PDF文件路径
        output_path (str): 输出Word文件路径，默认为同名.docx文件
        pages (list): 要转换的页面列表，如 [0, 1, 2]，默认转换所有页面
    
    返回:
        bool: 转换成功返回True，失败返回False
    """
    
    # 检查PDF文件是否存在
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在 - {pdf_path}")
        return False
    
    # 如果没有指定输出路径，则使用同名.docx文件
    if output_path is None:
        pdf_name = Path(pdf_path).stem
        output_path = os.path.join(os.path.dirname(pdf_path), f"{pdf_name}.docx")
    
    try:
        print(f"开始转换: {pdf_path}")
        print(f"输出文件: {output_path}")
        
        # 执行转换
        if pages:
            convert(pdf_path, output_path, pages=pages)
            print(f"已转换指定页面 {pages}")
        else:
            convert(pdf_path, output_path)
            print(f"已转换所有页面")
        
        # 检查输出文件
        if os.path.exists(output_path):
            size = os.path.getsize(output_path) / 1024  # 转换为KB
            print(f"✓ 转换成功! 文件大小: {size:.2f} KB")
            return True
        else:
            print("✗ 转换失败: 输出文件未生成")
            return False
            
    except Exception as e:
        print(f"✗ 转换过程中出错: {str(e)}")
        return False


def batch_convert(pdf_folder, output_folder=None):
    """
    批量转换PDF文件夹中的所有PDF文件
    
    参数:
        pdf_folder (str): 包含PDF文件的文件夹路径
        output_folder (str): 输出文件夹路径，默认为输入文件夹
    """
    
    if not os.path.isdir(pdf_folder):
        print(f"错误: 文件夹不存在 - {pdf_folder}")
        return
    
    if output_folder is None:
        output_folder = pdf_folder
    else:
        os.makedirs(output_folder, exist_ok=True)
    
    # 查找所有PDF文件
    pdf_files = list(Path(pdf_folder).glob("*.pdf"))
    
    if not pdf_files:
        print(f"未找到PDF文件 - {pdf_folder}")
        return
    
    print(f"找到 {len(pdf_files)} 个PDF文件，开始批量转换...\n")
    
    success_count = 0
    fail_count = 0
    
    for pdf_file in pdf_files:
        output_path = os.path.join(output_folder, f"{pdf_file.stem}.docx")
        if pdf_to_word(str(pdf_file), output_path):
            success_count += 1
        else:
            fail_count += 1
        print()
    
    print(f"批量转换完成! 成功: {success_count}, 失败: {fail_count}")


if __name__ == "__main__":
    
    # 使用示例
    if len(sys.argv) > 1:
        # 命令行模式
        pdf_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        pdf_to_word(pdf_path, output_path)
    else:
        # 交互模式
        print("=" * 50)
        print("PDF转Word转换工具")
        print("=" * 50)
        print("\n请选择转换模式:")
        print("1. 单个文件转换")
        print("2. 批量转换文件夹")
        
        choice = input("\n请输入选择 (1 或 2): ").strip()
        
        if choice == "1":
            pdf_path = input("请输入PDF文件路径: ").strip()
            output_path = input("请输入输出Word文件路径 (回车使用默认名称): ").strip()
            output_path = output_path if output_path else None
            pdf_to_word(pdf_path, output_path)
            
        elif choice == "2":
            folder_path = input("请输入包含PDF文件的文件夹路径: ").strip()
            output_folder = input("请输入输出文件夹路径 (回车使用输入文件夹): ").strip()
            output_folder = output_folder if output_folder else None
            batch_convert(folder_path, output_folder)
            
        else:
            print("无效的选择!")
