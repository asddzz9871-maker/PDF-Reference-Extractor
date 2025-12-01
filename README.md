# PDF 文献引用提取工具

这是一个用于从 PDF 论文中批量提取元数据并生成 GB/T 7714 格式引用的 Python 工具。

## 功能
- 自动识别 PDF 中的 DOI。
- 使用 Crossref API 联网获取精准的标题、作者、期刊和年份。
- 支持 DOI 识别失败时的本地正则提取兜底。
- 导出 Excel 表格，包含直接可用的引用格式。

## 使用方法
1. 安装依赖：`pip install -r requirements.txt`
2. 运行脚本：`python pdf_extractor.py`
3. 输入你的 PDF 文件夹路径即可。