#!/usr/bin/env python3
"""
解析宁波银行客户洞察报告Word文档
提取前两个章节的表格结构和内容格式
"""

import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("Error: python-docx not installed")
    sys.exit(1)

def extract_tables(doc, start_para_idx, end_para_idx):
    """提取指定段落范围内的表格"""
    tables_info = []
    
    # 获取所有段落和表格的相对位置
    elements = []
    for i, para in enumerate(doc.paragraphs):
        elements.append(('para', i, para.text[:100] if para.text else ""))
    
    for i, table in enumerate(doc.tables):
        # 尝试找到表格的位置
        table_text = table.rows[0].cells[0].text[:50] if table.rows else ""
        elements.append(('table', i, table_text))
    
    # 打印前100个段落，帮助定位
    print("\n【文档段落结构】")
    for i, para in enumerate(doc.paragraphs[:100]):
        if para.text.strip():
            style = para.style.name if para.style else "Normal"
            print(f"{i}. [{style}] {para.text[:80]}")
    
    return tables_info

def analyze_structure(doc):
    """分析文档结构"""
    print("=" * 80)
    print("宁波银行客户洞察报告结构分析")
    print("=" * 80)
    
    # 查找第一章和第二章的位置
    chapter1_start = None
    chapter2_start = None
    chapter2_end = None
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if "洞察发现与建议总结" in text or "一、洞察发现" in text:
            chapter1_start = i
            print(f"\n第一章开始位置: 段落 {i}")
            print(f"  内容: {text[:80]}")
        elif ("基础信息" in text and para.style.name.startswith('Heading')) or text == "一、基础信息":
            chapter2_start = i
            print(f"\n第二章开始位置: 段落 {i}")
            print(f"  内容: {text[:80]}")
        elif ("业务与战略" in text and para.style.name.startswith('Heading')) or text == "二、业务与战略":
            chapter2_end = i
            print(f"\n第二章结束位置(第三章开始): 段落 {i}")
            print(f"  内容: {text[:80]}")
            break
    
    return chapter1_start, chapter2_start, chapter2_end

def extract_tables_in_range(doc, start_idx, end_idx):
    """提取指定段落范围内的所有表格"""
    print(f"\n\n提取第{start_idx}段到第{end_idx}段之间的表格")
    print("=" * 80)
    
    tables_info = []
    table_idx = 0
    
    # 由于docx库无法直接获取表格的段落位置，我们通过分析段落文本来推断
    # 先打印这个范围内的段落，看看表格标题
    print("\n【范围内段落】")
    for i in range(start_idx, min(end_idx, len(doc.paragraphs))):
        para = doc.paragraphs[i]
        if para.text.strip():
            print(f"{i}: {para.text[:100]}")
    
    # 提取所有表格并尝试匹配
    print(f"\n【文档中所有表格】共{len(doc.tables)}个")
    
    for i, table in enumerate(doc.tables[:20]):  # 先分析前20个表格
        print(f"\n--- 表格 {i+1} ---")
        print(f"尺寸: {len(table.rows)} 行 x {len(table.columns)} 列")
        
        if table.rows:
            # 打印表头
            header = [cell.text[:30] for cell in table.rows[0].cells]
            print(f"表头: {' | '.join(header)}")
            
            # 打印前3行数据（如果有）
            for row_idx in range(1, min(4, len(table.rows))):
                row_data = [cell.text[:30] for cell in table.rows[row_idx].cells]
                print(f"行{row_idx}: {' | '.join(row_data)}")
            
            if len(table.rows) > 4:
                print(f"... 还有 {len(table.rows) - 4} 行")
    
    return tables_info

if __name__ == "__main__":
    file_path = "/workspace/projects/assets/宁波银行客户洞察及相关商机分析报告（完备版） _Kimi生成_20260303_发布后优化版20260310.docx"
    
    if not Path(file_path).exists():
        print(f"Error: 文件不存在 - {file_path}")
        sys.exit(1)
    
    try:
        doc = Document(file_path)
        print(f"文档加载成功，共 {len(doc.paragraphs)} 个段落，{len(doc.tables)} 个表格")
        
        # 分析结构
        ch1_start, ch2_start, ch2_end = analyze_structure(doc)
        
        # 提取第二章的表格
        if ch2_start and ch2_end:
            tables = extract_tables_in_range(doc, ch2_start, ch2_end)
        
        print("\n\n解析完成！")
    except Exception as e:
        print(f"Error: 解析失败 - {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
