import os
import re
import json
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from datetime import datetime
from habanero import Crossref

class PDFReferenceExtractor:
    def __init__(self, pdf_dir):
        self.pdf_dir = Path(pdf_dir)
        self.papers = []
        self.cr = Crossref() # 【优化】只需实例化一次，复用连接池，大幅提升批量查询速度

    def process_all_pdfs(self):
        """核心主控流水线"""
        pdf_files = list(self.pdf_dir.glob('*.pdf'))
        print(f"🚀 开始处理 {len(pdf_files)} 个PDF文件...\n" + "="*50)
        
        for idx, pdf_path in enumerate(pdf_files, 1):
            print(f"[{idx}/{len(pdf_files)}] 解析中: {pdf_path.name[:30]}...", end=" ")
            try:
                # 1. 提取文本
                text = self._extract_text(pdf_path)
                # 2. 寻找 DOI
                doi = self._find_doi(text)
                # 3. 核心获取逻辑 (优先 API，API 失败则用本地正则兜底)
                meta = self._fetch_from_api(doi) if doi else self._fallback_parse(text, pdf_path.name)
                # 4. 生成国标引用并保存
                meta['filename'] = pdf_path.name
                meta['doi'] = doi if doi else "未找到"
                meta['citation_gb7714'] = self._format_gb7714(meta)
                
                self.papers.append(meta)
                print("✅ 成功" if meta.get('_source') == 'api' else "⚠️ 使用本地兜底")
            except Exception as e:
                print(f"❌ 失败 ({e})")

    # ================= 内部工具函数 (流水线车间) =================
    
    def _extract_text(self, pdf_path, max_pages=3):
        """提取 PDF 前几页文本"""
        with fitz.open(pdf_path) as doc:
            return "".join([page.get_text() for page in doc[:max_pages]]).replace('-\n', '').replace('\n', ' ')

    def _find_doi(self, text):
        """增强版 DOI 提取"""
        match = re.search(r'(10\.\d{4,}/[-._;()/:A-Za-z0-9]+)', text, re.IGNORECASE)
        return match.group(1).rstrip('.,;)]}') if match else None

    def _fetch_from_api(self, doi):
        """通过 Crossref API 获取精准数据"""
        try:
            item = self.cr.works(ids=doi)['message']
            authors = [f"{a.get('family', '')} {' '.join([p[0].upper() for p in a.get('given', '').split() if p])}".strip() 
                       for a in item.get('author', []) if a.get('family')]
            
            date_parts = item.get('published-print', {}).get('date-parts') or item.get('published-online', {}).get('date-parts')
            
            return {
                '_source': 'api',
                'title': item.get('title', [''])[0],
                'authors': authors,
                'journal': item.get('container-title', [''])[0],
                'year': str(date_parts[0][0]) if date_parts else "",
                'volume': item.get('volume', ''),
                'issue': item.get('issue', ''),
                'pages': item.get('page', '')
            }
        except:
            return self._fallback_parse("", "") # API报错时无缝切换兜底

    def _fallback_parse(self, text, filename):
        """【优化】将原来的多个冗长正则合并为一个紧凑的兜底函数"""
        year_match = re.search(r'(\d{4})', filename)
        vol_match = re.search(r'[Vv]ol(?:ume)?\.?\s*(\d+)', text)
        page_match = re.search(r'(?:pp?\.?|[Pp]ages?)[:\s]+(\d+(?:\s*[-–]\s*\d+)?)', text)
        
        return {
            '_source': 'fallback',
            'title': re.sub(r' - \d{4}.*|\.pdf', '', filename), # 简单拿文件名当标题
            'authors': [re.match(r'^([A-Za-z]+)', filename).group(1)] if re.match(r'^([A-Za-z]+)', filename) else [],
            'journal': "未知期刊",
            'year': year_match.group(1) if year_match else "",
            'volume': vol_match.group(1) if vol_match else "",
            'issue': "",
            'pages': page_match.group(1) if page_match else ""
        }

    def _format_gb7714(self, meta):
        """生成 GB/T 7714-2015 格式引用"""
        author_str = ', '.join(meta['authors'][:3]) + (', et al.' if len(meta['authors']) > 3 else '.') if meta['authors'] else '[Unknown].'
        pub_info = f"{meta['year']}, {meta['volume']}({meta['issue']}): {meta['pages']}".strip(' ,():')
        return f"{author_str} {meta['title']} [J]. {meta['journal']}, {pub_info}."

    def export_results(self):
        """导出为 Excel 和 JSON"""
        if not self.papers: return print("⚠️ 无数据可保存！")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        df = pd.DataFrame(self.papers).drop(columns=['_source']) # 移除内部标记
        
        excel_path = f"文献提取_{timestamp}.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            for idx, col in enumerate(df.columns):
                writer.sheets['Sheet1'].column_dimensions[chr(65+idx)].width = 20
                
        json_path = f"文献提取_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.papers, f, ensure_ascii=False, indent=2)
            
        print(f"\n🎉 处理完毕！数据已导出至:\n 📊 {excel_path}\n 📄 {json_path}")

if __name__ == "__main__":
    target_dir = input("📁 请输入PDF文件夹路径 (直接回车默认当前目录): ").strip() or os.getcwd()
    if os.path.exists(target_dir):
        extractor = PDFReferenceExtractor(target_dir)
        extractor.process_all_pdfs()
        extractor.export_results()
    else:
        print("❌ 路径不存在！")
