# 为指定目录下的所有JSON文件添加标签
# 功能：为指定目录下的所有JSON文件添加标签
# 支持多页处理、文本智能识别、AI模型解析、微信二维码提取、数据标准化转换


import os
import json
import sys
from pathlib import Path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from src.utils.extract_tags import extract_tags

def append_tags_to_all_json(search_directory: str = None):
    """
    为指定目录下的所有JSON文件添加标签
    
    参数：
        search_directory (str): 要搜索的目录路径
    
    返回：
        None
    
    """
    
    
    if not os.path.exists(search_directory):
        print(f"目录不存在: {search_directory}")
        return
    
    success_count = 0
    
    # 遍历所有JSON文件
    for root, dirs, files in os.walk(search_directory):
        for file in files:
            if file.endswith('.json'):
                json_path = os.path.join(root, file)
                
                try:
                    # 读取JSON文件
                    with open(json_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    # 检查是否有文件路径字段
                    if '文件路径' not in data or not data['文件路径']:
                        continue
                    
                    file_path = data['文件路径']
                    
                    # 自动推断检查位置
                    path_segments = os.path.normpath(file_path).split(os.sep)
                    path_segments = [seg for seg in path_segments if seg.strip()]
                    
                    extra_location = []
                    if len(path_segments) >= 2:
                        extra_location.append(len(path_segments) - 1)
                    if len(path_segments) >= 3:
                        extra_location.append(len(path_segments) - 2)
                    
                    # 提取标签
                    tags = extract_tags(file_path, 1, [4]) # 5:excel, 4:pdf pptx
                    
                    # 添加标签并保存
                    data['标签'] = tags
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=4)
                    
                    success_count += 1
                    print(f"已处理: {os.path.basename(json_path)} -> {tags}")
                    
                except Exception as e:
                    print(f"处理失败: {json_path} - {e}")
    
    print(f"完成！成功处理 {success_count} 个文件")

if __name__ == "__main__":
    search_directory = r"data\processed_data\data_tag\ppt"
    append_tags_to_all_json(search_directory) 