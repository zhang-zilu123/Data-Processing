# 从文件路径中提取标签信息
# 功能：从文件路径中提取标签信息
# 支持多页处理、文本智能识别、AI模型解析、微信二维码提取、数据标准化转换


import os
import re
import logging
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- 函数：从文件路径中提取标签信息 ---
def extract_tags(filepath: str, extra_num: int, extra_location: list) -> str:
    """
    从文件路径指定位置中智能提取分类标签信息
    
    参数：
        filepath (str): 完整的文件路径字符串
        extra_num (int): 提取的标签数量限制
        extra_location (list): 要检查的路径段位置索引列表（0-based）
        
    返回：
        str: 提取的标签字符串，多个标签用""连接，失败时返回"标签提取失败"
    """
    try:
        # 标准化路径并按路径分隔符分割
        normalized_path = os.path.normpath(filepath)
        path_segments = normalized_path.split(os.sep)
        
        # 过滤掉空字符串
        path_segments = [segment for segment in path_segments if segment.strip()]
        
        # logging.info(f"路径分割结果: {path_segments}")
        
        
        all_tags = []

        for i in range(extra_num):
            
            matched_tags = []
            # 只检查指定位置的路径段
            location_index=extra_location[i]
            
            if location_index > len(path_segments):
                logging.error(f"指定位置 {location_index} 超出路径段范围 (最大索引: {len(path_segments)})")
                continue
                
            current_segment = path_segments[location_index-1]
            # logging.info(f"检查位置 {location_index} 的路径段: '{current_segment}'")
            
            
            # 遍历每个关键词组
            for keyword_group in POSSIBLE_TAGS:
                
                # 在当前关键词组中查找匹配项（每组最多匹配一个）
                for keyword in keyword_group:
                    # 检查是否是正则表达式（包含正则特殊字符）
                    if re.search(r'[()[\]{}+*?^$|\\]', keyword):
                        # 正则表达式匹配
                        match = re.search(keyword, current_segment)
                        if match:
                            matched_text = match.group()
                            matched_tags.append((location_index, matched_text))
                            logging.info(f"正则匹配到标签: '{matched_text}' 位置: {location_index}")
                            
                            break
                    else:
                        # 普通字符串匹配
                        if keyword in current_segment:
                            matched_tags.append((location_index, keyword))
                            logging.info(f"字符串匹配到标签: '{keyword}' 位置: {location_index}")
                            
                            break
        
            # 按照在路径中出现的位置排序
            matched_tags.sort(key=lambda x: x[0])


            # 连接所有匹配的标签
            if matched_tags:
                result_temp = ''.join([tag[1] for tag in matched_tags])
                all_tags.append(result_temp)
                logging.info(f"提取的标签: '{result_temp}'")
                
            else:
                logging.error(f"在指定位置未找到任何匹配的标签")

        if len(all_tags) >1:
            result = '/'.join(all_tags)
        else:
            result = all_tags[0]

        
        return result
            
    except Exception as e:
        logging.error(f"标签提取异常: {e}, 文件路径: {filepath}")
        return None
        
if __name__ == "__main__":
    # 测试用例
    
    test_cases1 = [
        {
            "filepath": r"data\\input_data\\ppt\\宁波D586-2025.07.02\\广东鸿祺玩具实业有限公司.pptx",
            "extra_num": 1,
            "extra_location": [4]  # 检查"5月宁波到访工厂资料"
        },
        # {
        #     "filepath": r"input_files\\2024到访工厂打印资料\\9月\\武义艺佳休闲用品有限公司-2024.09.14.docx",
        #     "extra_num": 1,
        #     "extra_location": [2]  # 检查"5月宁波到访工厂资料"
        # },
        # {
        #     "filepath": r"data\\input_data\\pdf\\D50期双战略供应商\\宁波梅西照明电器有限公司.pdf",
        #     "extra_num": 1,
        #     "extra_location": [4]  # 检查"D50期双战略供应商"
        # },
        # {
        #     "filepath": r"data\\input_data\\excel\\2025到访工厂\\供应商交流会\\供应商交流会第五十六期名单（6.25）.xlsx",
        #     "extra_num": 2,
        #     "extra_location": [5, 6]  # 检查"供应商交流会"和文件名
        # }
    ]
    
    for test_case in test_cases1:
        filepath = test_case["filepath"]
        extra_num = test_case["extra_num"]
        extra_location = test_case["extra_location"]
        
        print(f"路径: {filepath}")
        print(f"检查位置: {extra_location}, 标签数量限制: {extra_num}")
        print(f"标签: {extract_tags(filepath, extra_num, extra_location)}")
        print("-" * 80)  