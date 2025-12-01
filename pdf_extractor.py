#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF文献信息提取与引用生成工具 (PDF Reference Extractor)
功能：
1. 自动提取PDF中的 DOI、作者、期刊、年份等元数据
2. 优先使用 Crossref API 联网获取精准数据
3. 自动生成国标 GB/T 7714-2015 引用格式
4. 导出 Excel 和 JSON 报告

依赖库：pymupdf, pandas, habanero, openpyxl
"""

import os
import re
import json
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from datetime import datetime
from habanero import Crossref
import time

class PDFReferenceExtractor:
    def __init__(self, pdf_dir):
        self.pdf_dir = Path(pdf_dir)
        self.papers = []
    
    def extract_text_from_pdf(self, pdf_path, max_pages=3):
        """从PDF前几页提取文本"""
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page_num in range(min(max_pages, len(doc))):
                page = doc[page_num]
                text += page.get_text()
            doc.close()
            return text
        except Exception as e:
            print(f"[错误] 读取PDF失败: {e}")
            return ""
    
    def parse_filename(self, filename):
        """从文件名解析基础信息（作为兜底）"""
        name = filename.replace('.pdf', '')
        
        # 提取年份
        year_match = re.search(r'(\d{4})', name)
        year = year_match.group(1) if year_match else ""
        
        # 提取第一作者姓氏
        author_match = re.match(r'^([A-Za-z]+)', name)
        first_author_last = author_match.group(1) if author_match else ""
        
        # 提取标题
        title_match = re.search(r' - \d{4} - (.+)$', name)
        if not title_match:
            title_match = re.search(r' - (.+)$', name)
        title = title_match.group(1).strip() if title_match else name
        
        return {
            'first_author_last': first_author_last,
            'year': year,
            'title_from_filename': title
        }
    
    def extract_doi(self, text):
        """增强版 DOI 提取"""
        # 1. 预处理：修复被换行符切断的 DOI
        clean_text = re.sub(r'-\n', '', text)
        clean_text = re.sub(r'\n', ' ', clean_text)
        
        doi_patterns = [
            r'https?://(?:dx\.)?doi\.org/(10\.\d{4,}/[-._;()/:A-Za-z0-9]+)',
            r'(?:DOI|doi)(?:\s*[:.])?\s*(10\.\d{4,}/[-._;()/:A-Za-z0-9]+)',
            r'\b(10\.\d{4,}/[-._;()/:A-Za-z0-9]+)\b'
        ]
        
        for pattern in doi_patterns:
            matches = re.findall(pattern, clean_text, re.IGNORECASE)
            for match in matches:
                clean_doi = match.strip().strip('.,;)]}')
                if '/' in clean_doi and len(clean_doi) > 5:
                    return clean_doi
        return ""
    
    def extract_authors_regex(self, text):
        """(本地正则) 提取作者 - 仅作为 API 失败时的备选"""
        authors = []
        lines = text.split('\n')
        
        for line in lines[:30]:
            line = line.strip()
            if not line or len(line) > 200: continue
            
            # 跳过明显不是作者的行
            if any(skip in line.lower() for skip in ['abstract', 'introduction', 'doi', 'published', 'received', 'elsevier']):
                continue
            
            pattern1 = r'([A-Z][a-z]+)\s+([A-Z])\.?\s*([A-Z])?\.?'
            matches1 = re.findall(pattern1, line)
            
            if matches1:
                for match in matches1:
                    last_name = match[0]
                    if last_name in ['Journal', 'Science', 'Technology', 'Research', 'International']: continue
                    
                    author = f"{last_name} {match[1]}"
                    if match[2]: author += f". {match[2]}."
                    
                    if author not in authors:
                        authors.append(author)
                if authors and len(authors) >= 3: break
        return authors[:10]
    
    def extract_journal_regex(self, text):
        """(本地正则) 提取期刊名"""
        lines = text.split('\n')
        journal_keywords = ['Journal of', 'International Journal', 'Applied', 'Chemical', 'Environmental', 'Science', 'Technology']
        
        for line in lines[:50]:
            line = line.strip()
            for keyword in journal_keywords:
                if keyword in line:
                    cleaned = re.sub(r'\d{4}', '', line)
                    cleaned = re.sub(r'[,;].*$', '', cleaned).strip()
                    if 10 < len(cleaned) < 100:
                        return cleaned
        return ""
    
    def extract_details_regex(self, text):
        """(本地正则) 提取卷、期、页码"""
        volume = issue = pages = ""
        
        vol_match = re.search(r'[Vv]ol(?:ume)?\.?\s*(\d+)', text)
        if vol_match: volume = vol_match.group(1)
        
        issue_match = re.search(r'[Nn]o\.?\s*(\d+)', text)
        if issue_match: issue = issue_match.group(1)
        
        page_match = re.search(r'(?:pp?\.?|[Pp]ages?)[:\s]+(\d+(?:\s*[-–]\s*\d+)?)', text)
        if page_match: 
            pages = page_match.group(1)
        elif re.search(r'\b(\d{6,})\b', text): # 6位文章号
             pages = re.search(r'\b(\d{6,})\b', text).group(1)

        return volume, issue, pages
    
    def format_citation_gb7714(self, paper_info):
        """生成 GB/T 7714-2015 格式引用"""
        parts = []
        
        # 1. 作者
        authors = paper_info.get('authors', [])
        if authors:
            author_str = ', '.join(authors[:3]) + (', et al' if len(authors) > 3 else '')
        else:
            author_str = paper_info.get('first_author_last', '[Unknown]')
        parts.append(author_str + '.')
        
        # 2. 标题
        title = paper_info.get('title', '').strip()
        if title: parts.append(title)
        
        # 3. 文献类型
        parts.append('[J].')
        
        # 4. 期刊
        journal = paper_info.get('journal', '').strip()
        if journal: parts.append(journal + ',')
        
        # 5. 年卷期页
        year = str(paper_info.get('year', '')).strip()
        volume = str(paper_info.get('volume', '')).strip()
        issue = str(paper_info.get('issue', '')).strip()
        pages = str(paper_info.get('pages', '')).strip()
        
        pub_info = []
        if year: pub_info.append(year)
        if volume: pub_info.append(f"{volume}({issue})" if issue else volume)
        
        if pub_info:
            parts.append(', '.join(pub_info))
            if pages: parts[-1] += f": {pages}"
            
        citation = ' '.join(parts)
        if not citation.endswith('.'): citation += '.'
        return citation
    
    def process_pdf(self, pdf_path):
        """处理单个PDF的主逻辑"""
        print(f"\n[处理] {pdf_path.name}")
        
        # 基础解析
        filename_info = self.parse_filename(pdf_path.name)
        text = self.extract_text_from_pdf(pdf_path)
        doi = self.extract_doi(text)
        
        # 初始化
        data = {
            'title': filename_info['title_from_filename'],
            'authors': [],
            'year': filename_info['year'],
            'journal': '', 'volume': '', 'issue': '', 'pages': ''
        }
        
        # 联网查询
        api_success = False
        if doi:
            print(f"  -> 识别到 DOI: {doi}，正在联网查询...", end="")
            meta = get_meta_from_doi(doi)
            if meta["状态"] == "成功":
                print(" [成功]")
                api_success = True
                data.update({
                    'title': meta['标题'],
                    'journal': meta['期刊'],
                    'year': meta['年份'],
                    'volume': meta['卷'],
                    'issue': meta['期'],
                    'pages': meta['页码']
                })
                if meta['作者']:
                    data['authors'] = [a.strip() for a in meta['作者'].split(',')]
            else:
                print(f" [API失败: {meta.get('错误信息')}]")
        else:
            print("  -> 未识别到有效 DOI")

        # 本地兜底
        if not api_success:
            print("  -> 使用本地正则解析")
            data['authors'] = self.extract_authors_regex(text)
            data['journal'] = self.extract_journal_regex(text)
            data['volume'], data['issue'], data['pages'] = self.extract_details_regex(text)
            
            # 尝试从正文首行提取标题
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            if lines and 10 < len(lines[0]) < 200:
                data['title'] = lines[0]

        # 整合结果
        paper_info = {
            'filename': pdf_path.name,
            'doi': doi,
            'file_path': str(pdf_path),
            'first_author_last': filename_info['first_author_last'],
            **data
        }
        
        paper_info['citation_gb7714'] = self.format_citation_gb7714(paper_info)
        
        print(f"  - 标题: {paper_info['title'][:40]}...")
        print(f"  - 来源: {'Crossref API' if api_success else '本地正则'}")
        return paper_info

    def process_all_pdfs(self):
        pdf_files = sorted(self.pdf_dir.glob('*.pdf'))
        total = len(pdf_files)
        print(f"{'='*60}\n开始处理 {total} 个PDF文件\n{'='*60}")
        
        for idx, pdf_path in enumerate(pdf_files, 1):
            print(f"\n[{idx}/{total}]", end=' ')
            try:
                self.papers.append(self.process_pdf(pdf_path))
            except Exception as e:
                print(f"[错误] {e}")
                
    def save_results(self, excel_path, json_path):
        if not self.papers:
            print("[警告] 没有数据可保存")
            return
            
        # 保存 Excel
        df = pd.DataFrame([{
            '序号': i+1,
            '文件名': p['filename'],
            '标题': p['title'],
            '作者': '; '.join(p['authors']),
            '年份': p['year'],
            '期刊': p['journal'],
            '卷': p['volume'],
            '期': p['issue'],
            '页码': p['pages'],
            'DOI': p['doi'],
            '国标引用': p['citation_gb7714']
        } for i, p in enumerate(self.papers)])
        
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='文献信息', index=False)
            # 自动调整列宽
            ws = writer.sheets['文献信息']
            for idx, col in enumerate(df.columns):
                ws.column_dimensions[chr(65+idx)].width = 20
        
        # 保存 JSON
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.papers, f, ensure_ascii=False, indent=2)
            
        print(f"\n[完成] 结果已保存至:\n - {excel_path}\n - {json_path}")

def get_meta_from_doi(doi_string):
    """通过 Crossref API 获取标准元数据"""
    cr = Crossref()
    clean_doi = doi_string.strip()
    try:
        item = cr.works(ids=clean_doi)['message']
        
        # 作者处理
        authors = []
        for a in item.get('author', []):
            family = a.get('family', '')
            given = a.get('given', '')
            if family:
                initials = ' '.join([p[0].upper() for p in given.split() if p]) if given else ''
                authors.append(f"{family} {initials}".strip())
        
        # 年份处理
        date_parts = item.get('published-print', {}).get('date-parts') or item.get('published-online', {}).get('date-parts')
        year = date_parts[0][0] if date_parts else ""

        return {
            "状态": "成功",
            "标题": item.get('title', [''])[0],
            "期刊": item.get('container-title', [''])[0],
            "作者": ", ".join(authors),
            "年份": year,
            "卷": item.get('volume', ''),
            "期": item.get('issue', ''),
            "页码": item.get('page', '')
        }
    except Exception as e:
        return {"状态": "失败", "错误信息": str(e)}

def main():
    print("--- PDF 文献引用提取工具 ---")
    
    # 修改：交互式输入路径，而不是硬编码
    default_dir = os.getcwd()
    input_path = input(f"请输入PDF文件夹路径 (直接回车默认: {default_dir}): ").strip()
    if not input_path:
        input_path = default_dir
        
    if not os.path.exists(input_path):
        print(f"[错误] 路径不存在: {input_path}")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = f"文献引用提取结果_{timestamp}.xlsx"
    output_json = f"文献引用提取结果_{timestamp}.json"
    
    extractor = PDFReferenceExtractor(input_path)
    extractor.process_all_pdfs()
    extractor.save_results(output_excel, output_json)

if __name__ == "__main__":
    main()