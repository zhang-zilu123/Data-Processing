# 使用PyMuPDF提取文本，自动合并纵坐标相近的行，只返回文本内容
from typing import List
import fitz
import re

def extract_text_lines_from_pdf(pdf_path: str, y_threshold: float = 20.0) -> List[str]:
    """
    使用PyMuPDF提取文本，自动合并纵坐标相近的行，只返回文本内容
    
    参数:
        pdf_path: PDF文件路径
        y_threshold: 纵坐标合并阈值，默认20.0
    
    返回:
        文本行列表
    """
    merged_lines = []
    
    try:
        with fitz.open(pdf_path) as doc:
            for page_num in range(len(doc)):
                page = doc[page_num]
                text_dict = page.get_text("dict")
                raw_lines = []
                
                # 提取原始文本行及其坐标
                for block in text_dict.get("blocks", []):
                    if "lines" not in block:
                        continue
                        
                    for line in block["lines"]:
                        line_text = ""
                        min_y0 = float('inf')
                        min_x0 = float('inf')
                        
                        for span in line["spans"]:
                            if not span["text"]:
                                continue
                                
                            line_text += span["text"]
                            # 更新最小坐标
                            if span["bbox"][1] < min_y0:
                                min_y0 = span["bbox"][1]
                            if span["bbox"][0] < min_x0:
                                min_x0 = span["bbox"][0]
                        
                        line_text = line_text
                        if line_text and not re.match(r'^[=\-\s]+$', line_text):
                            raw_lines.append({
                                'text': line_text,
                                'y0': min_y0,
                                'x0': min_x0
                            })
                
                # 如果没有提取到行，跳过后续处理
                if not raw_lines:
                    continue
                
                # 按纵坐标排序
                raw_lines.sort(key=lambda x: x['y0'])
                
                # 合并纵坐标相近的行
                current_group = [raw_lines[0]]
                for i in range(1, len(raw_lines)):
                    current_line = raw_lines[i]
                    last_line = current_group[-1]
                    
                    if abs(current_line['y0'] - last_line['y0']) <= y_threshold:
                        current_group.append(current_line)
                    else:
                        # 合并当前组
                        current_group.sort(key=lambda x: x['x0'])
                        merged_text = ' '.join([item['text'] for item in current_group])
                        merged_lines.append(merged_text)
                        current_group = [current_line]
                
                # 处理最后一组
                if current_group:
                    current_group.sort(key=lambda x: x['x0'])
                    merged_text = ' '.join([item['text'] for item in current_group])
                    merged_lines.append(merged_text)
                    
    except Exception as e:
        raise RuntimeError(f"使用PyMuPDF处理PDF文件时出错: {str(e)}")
    
    return merged_lines


if __name__ == "__main__":
    # 实际使用示例 - 输出所有行
    print(f"PyMuPDF版本: {fitz.__version__}")
    pdf_path = r"data\test\5555\上海纽恩特实业股份有限公司湖南享同实业有限公司1750645560.pdf"  # 替换为实际PDF路径
    try:
        # 使用推荐阈值20进行PyMuPDF提取
        print("使用推荐阈值20进行PyMuPDF提取:")
        pymupdf_lines = extract_text_lines_from_pdf(pdf_path, y_threshold=20.0)
        print(f"合并后提取到 {len(pymupdf_lines)} 行文本:")
        print("-" * 80)
        print(f"{'序号':<4} {'文本内容'}")
        print("-" * 80)
        
        # 输出所有行
        for i, line in enumerate(pymupdf_lines, 1):
            print(f"{i:<4} {line}")
        
        print("-" * 80)
        print(f"总共提取了 {len(pymupdf_lines)} 行文本")
            
    except Exception as e:
        print(f"提取文本时出错: {e}")