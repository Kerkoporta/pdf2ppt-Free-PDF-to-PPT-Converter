# pdf2ppt-Free-PDF-to-PPT-Converter
中文: 智能PDF转PPT工具 - 保持原始布局，支持表格识别，文本可编辑的精准转换工具  English: Intelligent PDF to PPT Converter - Preserves original layout with editable text conversion

MARKDOWN
# 📄 PDF to PPT Converter

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![PyMuPDF](https://img.shields.io/badge/PyMuPDF-1.23%2B-orange)](https://github.com/pymupdf/PyMuPDF)

一个轻量级的PDF转PPT转换工具，能够保持原始布局并生成完全可编辑的PPTX文件。

## ✨ 特性

- 🎯 **精准布局保持** - 1:1保持PDF原始页面布局和尺寸
- 🖼️ **智能图片处理** - 自动过滤阴影和重复图片
- 📝 **文本可编辑** - 所有文本转换为可编辑内容，保持字体样式
- 🎨 **友好GUI界面** - 简洁易用的图形用户界面
- 🔢 **自动页码** - 为每页添加页码标注
- ⚡ **快速转换** - 高效的转换算法

## 🚀 快速开始

### 安装依赖

```bash
pip install PyMuPDF python-pptx
使用方法
图形界面方式 (推荐):
BASH
python pdf_to_ppt_tk.py
命令行方式:
PYTHON
from pdf_to_ppt_core import pdf_to_ppt
pdf_to_ppt("input.pdf", "output.pptx")
📁 项目结构
TEXT
PDF-to-PPT-Converter/
├── pdf_to_ppt_tk.py      # 图形界面主程序
├── pdf_to_ppt_core.py    # 核心转换逻辑
├── requirements.txt      # 依赖包列表
├── README.md            # 项目说明文档
└── LICENSE              # MIT许可证文件
🔧 核心功能
pdf_to_ppt_core.py
布局保持: 精确转换PDF页面尺寸到PPT
文本提取: 提取文本块并保持字体样式
图片处理: 智能识别并过滤阴影图片
坐标转换: 准确的pt到英寸单位转换
pdf_to_ppt_tk.py
文件选择: 支持PDF文件选择和PPTX保存路径设置
进度显示: 实时显示转换进度
错误处理: 友好的错误提示信息
自动命名: 根据PDF文件名自动生成PPTX文件名
🛠️ 技术栈
PyMuPDF - PDF解析和处理
python-pptx - PPTX文件生成
Tkinter - 图形用户界面
多线程处理 - 避免界面卡顿
📊 转换效果
功能	支持情况	说明
文本转换	✅ 完全支持	保持字体、大小和位置
图片转换	✅ 完全支持	自动过滤阴影图片
布局保持	✅ 高度保持	1:1比例还原
多页处理	✅ 支持	完整PDF文档转换
批量处理	⚡ 通过GUI支持	逐个文件处理
🎯 使用场景
将PDF报告转换为可编辑的PPT演示文稿
学术论文和文档的格式转换
商业文档的重新编辑和演示
任何需要将PDF内容转换为PPT的场景
⚙️ 安装
克隆项目：
BASH
git clone https://github.com/your-username/PDF-to-PPT-Converter.git
cd PDF-to-PPT-Converter
安装依赖：
BASH
pip install -r requirements.txt
🚀 使用方法
图形界面模式
运行 python pdf_to_ppt_tk.py
点击"浏览"选择PDF文件
选择或输入PPTX保存路径
点击"开始转换"
等待转换完成提示
编程方式
PYTHON
from pdf_to_ppt_core import pdf_to_ppt

# 基本转换
success = pdf_to_ppt("input.pdf", "output.pptx")

# 带错误处理
try:
    if pdf_to_ppt("document.pdf", "presentation.pptx"):
        print("转换成功!")
    else:
        print("转换失败!")
except Exception as e:
    print(f"错误: {e}")
⚠️ 注意事项
目前主要支持文本和图片内容的转换
复杂的PDF格式可能需要手动调整
建议在使用前备份原始文件
🤝 贡献
欢迎提交Issue和Pull Request来改进这个项目！

Fork 本项目
创建特性分支 (git checkout -b feature/AmazingFeature)
提交更改 (git commit -m 'Add some AmazingFeature')
推送到分支 (git push origin feature/AmazingFeature)
打开Pull Request
📄 许可证
本项目采用 MIT 许可证 - 查看 LICENSE 文件了解详情

🙏 致谢
PyMuPDF - 优秀的PDF处理库
python-pptx - 强大的PPTX生成库
⭐ 如果这个项目对您有帮助，请给它一个Star！
