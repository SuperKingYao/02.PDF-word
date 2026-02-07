# PDF转Word转换工具使用说明

## 功能介绍
这是一个Python脚本，可以将PDF文件转换为Word文档，支持：
- ✓ 单个PDF文件转换
- ✓ 自动生成Word框架（无需外部依赖）
- ✓ LibreOffice完整转换模式（可选）
- ✓ 交互式和命令行两种使用方式

## 快速开始

### 方式1：使用纯Python版本（推荐，无需安装依赖）

```bash
# 直接运行，不需要任何依赖
./.venv/Scripts/python pdf_to_word_simple.py "input.pdf" "output.docx"

# 或交互模式
./.venv/Scripts/python pdf_to_word_simple.py
```

### 方式2：使用替代版本（需要安装库）

```bash
pip install pymupdf python-docx

# 然后运行
./.venv/Scripts/python pdf_to_word_alternative.py
```

## 脚本版本说明

### 1. pdf_to_word_simple.py (推荐)
- **优点**: 无需外部依赖，使用Python标准库
- **功能**: 生成Word文档框架，支持LibreOffice完整转换
- **推荐场景**: 所有用户（最简单易用）

### 2. pdf_to_word_alternative.py
- **优点**: 可以提取PDF文本和图片
- **依赖**: pymupdf, python-docx
- **功能**: 文本提取、图片保存、支持两种转换模式

### 3. pdf_to_word.py (原始版本)
- **优点**: 最完善的格式保留
- **依赖**: pdf2docx (可能需要网络)
- **功能**: 完整的PDF到Word转换

## 转换方式详解

### 方式A：LibreOffice转换（最完整）
如果已安装LibreOffice，脚本会自动使用它进行高质量转换：
```bash
# 需要先安装LibreOffice
# 下载: https://www.libreoffice.org/download

# 然后自动使用（脚本会检测）
./.venv/Scripts/python pdf_to_word_simple.py "test.pdf"
```

### 方式B：自动生成Word框架（无需依赖）
如果未安装外部工具，脚本会创建基础Word文档：
```bash
./.venv/Scripts/python pdf_to_word_simple.py "test.pdf" "output.docx"
```

### 方式C：提取文本和图片（需安装库）
```bash
pip install pymupdf python-docx
./.venv/Scripts/python pdf_to_word_alternative.py
```

## 使用示例

### 示例1：最简单的方式
```bash
cd "j:\work_project_\48.脚本测试\02.PDF转word"
.\.venv\Scripts\python pdf_to_word_simple.py
# 按提示输入文件路径即可
```

### 示例2：直接指定文件
```bash
.\.venv\Scripts\python pdf_to_word_simple.py "C:\Documents\report.pdf" "C:\Output\report.docx"
```

### 示例3：批量转换
可以创建批处理脚本或使用Python脚本批处理

## 系统要求
- Python 3.6 或更高版本
- Windows / macOS / Linux
- pdf_to_word_simple.py：无额外依赖
- 其他版本：需要相应的Python库

## 常见问题

### Q: 脚本运行没有错误但生成的Word很小？
A: 这是正常的。使用的是标准库版本，只生成了框架。
解决方案：
1. 安装LibreOffice进行完整转换
2. 或安装pdf2docx库

### Q: 如何完整转换PDF的格式和样式？
A: 有三种方案：
1. **最推荐**: 安装LibreOffice，脚本会自动使用
2. 安装pdf2docx库
3. 使用在线转换工具（CloudConvert等）

### Q: 无法导入模块怎么办？
A: 使用标准库版本，它无需外部依赖：
```bash
.\.venv\Scripts\python pdf_to_word_simple.py
```

### Q: 网络不好无法安装库怎么办？
A: 使用纯Python版本（pdf_to_word_simple.py），完全无需网络。

## 安装LibreOffice（用于完整转换）

### Windows
1. 访问 https://www.libreoffice.org/download
2. 下载 LibreOffice Windows版本
3. 按照安装向导完成安装
4. 重启电脑（确保系统识别）
5. 再次运行脚本，会自动使用LibreOffice

### MacOS
```bash
brew install libreoffice
```

### Linux (Ubuntu/Debian)
```bash
sudo apt-get install libreoffice
```

## 技术细节

### DOCX格式说明
- DOCX本质是ZIP压缩包
- pdf_to_word_simple.py直接创建标准的DOCX文件结构
- 包含必要的XML文件和关系定义

### 为什么分多个版本？
- **simple版**: 兼容性最好，无需外部依赖
- **alternative版**: 功能更丰富，能提取文本和图片
- **原始版**: 使用专业库，格式保留最完整（需网络）

## 许可证
免费使用和修改

## 建议
- 对于简单用途，使用 `pdf_to_word_simple.py`
- 对于保留PDF格式，安装LibreOffice
- 对于提取内容，使用 `pdf_to_word_alternative.py`

---
如有问题或建议，欢迎反馈！
