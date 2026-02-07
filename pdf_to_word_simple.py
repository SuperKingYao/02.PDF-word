#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF转Word脚本 - 纯Python版本
仅使用Python标准库，无外部依赖
"""

import os
import sys
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from io import BytesIO
from datetime import datetime


def create_docx_from_scratch(text_content, output_path):
    """
    使用标准库创建DOCX文件（DOCX是ZIP格式）
    
    参数:
        text_content (str): 要写入的文本内容
        output_path (str): 输出文件路径
    """
    
    # 转义文本中的特殊XML字符
    text_escaped = text_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')
    
    # DOCX的核心XML文件内容
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:sz w:val="32"/>
                    <w:szCs w:val="32"/>
                </w:rPr>
                <w:t>PDF转换文档</w:t>
            </w:r>
        </w:p>
        <w:p><w:t/></w:p>
'''
    
    # 添加文本内容
    for line in text_escaped.split('\n'):
        document_xml += f'''        <w:p>
            <w:r>
                <w:t>{line}</w:t>
            </w:r>
        </w:p>
'''
    
    document_xml += '''    </w:body>
</w:document>'''
    
    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    
    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    
    # 创建DOCX (实际是ZIP)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        # 添加文件
        docx.writestr('[Content_Types].xml', content_types_xml)
        docx.writestr('_rels/.rels', rels_xml)
        docx.writestr('word/document.xml', document_xml)
    
    print(f"✓ 已创建Word文档: {output_path}")


def pdf_to_word_using_libreoffice(pdf_path, output_path=None):
    """
    使用LibreOffice通过命令行转换PDF
    
    参数:
        pdf_path (str): PDF文件路径
        output_path (str): 输出Word文件路径
    
    返回:
        bool: 转换成功返回True，失败返回False
    """
    
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在 - {pdf_path}")
        return False
    
    if output_path is None:
        pdf_name = Path(pdf_path).stem
        output_path = os.path.join(os.path.dirname(pdf_path), f"{pdf_name}.docx")
    
    output_dir = os.path.dirname(output_path) or '.'
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"开始转换: {pdf_path}")
    print(f"输出文件: {output_path}")
    
    # 尝试使用LibreOffice
    try:
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'docx:MS Word 2007 XML', 
             '--outdir', output_dir, pdf_path],
            check=True,
            capture_output=True,
            timeout=120
        )
        if os.path.exists(output_path):
            size = os.path.getsize(output_path) / 1024
            print(f"✓ 转换成功 (使用LibreOffice)!")
            print(f"  文件大小: {size:.2f} KB")
            return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("LibreOffice不可用，使用备选方案...")
    
    return False


def pdf_to_word_manual_method(pdf_path, output_path=None):
    """
    手动创建Word文档的方法（基于纯Python）
    
    参数:
        pdf_path (str): PDF文件路径
        output_path (str): 输出Word文件路径
    
    返回:
        bool: 转换成功返回True，失败返回False
    """
    
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在 - {pdf_path}")
        return False
    
    if output_path is None:
        pdf_name = Path(pdf_path).stem
        output_path = os.path.join(os.path.dirname(pdf_path), f"{pdf_name}.docx")
    
    try:
        print(f"开始处理: {pdf_path}")
        print(f"输出文件: {output_path}")
        
        # 创建Word文档
        text_content = f"""PDF文件转换说明

源文件: {os.path.basename(pdf_path)}
文件大小: {os.path.getsize(pdf_path) / 1024:.2f} KB
转换时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

此Word文档由PDF文件转换而来。

完整转换方案:

1. 安装LibreOffice (推荐，免费)
   - 访问: https://www.libreoffice.org/download
   - 安装后运行本脚本即可自动转换

2. 安装PDF转换库 (需要网络)
   - 运行: pip install pdf2docx
   - 然后运行本脚本

3. 使用在线工具
   - CloudConvert, Zamzar 等在线PDF转Word工具

当前已生成基础Word框架，您可以手动编辑内容。"""
        
        create_docx_from_scratch(text_content, output_path)
        
        size = os.path.getsize(output_path) / 1024
        print(f"  文件大小: {size:.2f} KB")
        print(f"\n提示: 要完整转换PDF内容和格式，请安装LibreOffice")
        return True
        
    except Exception as e:
        print(f"✗ 处理过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    
    print("=" * 60)
    print("PDF转Word转换工具 (标准库版)")
    print("=" * 60)
    print()
    
    if len(sys.argv) > 1:
        # 命令行模式
        pdf_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        
        # 先尝试LibreOffice方法
        if not pdf_to_word_using_libreoffice(pdf_path, output_path):
            # 如果失败，使用手动方法
            pdf_to_word_manual_method(pdf_path, output_path)
    else:
        # 交互模式
        print("支持的转换方式:")
        print("1. LibreOffice (完整格式保留，需先安装LibreOffice)")
        print("2. 手动创建Word文档 (创建框架)")
        print()
        
        pdf_path = input("请输入PDF文件路径: ").strip()
        
        if pdf_path:
            output_path = input("输出Word文件路径 (回车使用默认): ").strip()
            output_path = output_path if output_path else None
            
            print()
            
            # 先尝试LibreOffice
            if not pdf_to_word_using_libreoffice(pdf_path, output_path):
                # 备选方案
                pdf_to_word_manual_method(pdf_path, output_path)
        else:
            print("未提供文件路径")
