import os
import json
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def merge_json_files(input_path, output_file):
    """
    合并指定目录下所有JSON文件到一个总的JSON文件
    
    参数:
    input_path (str): 包含JSON文件的目录路径
    output_file (str): 输出的合并JSON文件名
    """

    
    if not os.path.exists(input_path):
        logging.error(f"输入路径不存在: {input_path}")
        return
    
    if not os.path.isdir(input_path):
        logging.error(f"输入路径不是目录: {input_path}")
        return
    
    # 查找所有JSON文件
    json_files = []
    for root, _, files in os.walk(input_path):
        for file in files:
            if file.lower().endswith('.json'):
                full_path = os.path.join(root, file)
                json_files.append(full_path)
    
    if not json_files:
        logging.warning(f"在目录 {input_path} 中未找到JSON文件")
        return
    
    logging.info(f"找到 {len(json_files)} 个JSON文件")
    
    # 初始化结果字典
    combined_data = []
    processed_files = 0
    skipped_files = 0
    
    # 处理每个JSON文件
    for file_path in json_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                file_data = json.load(f)
                
            # 使用相对路径作为键名
            
            combined_data.append(file_data)
            processed_files += 1
            logging.info(f"成功拼接文件")
            
        except json.JSONDecodeError:
            skipped_files += 1
            logging.error(f"JSON解析失败: {file_path}", exc_info=True)
        except Exception as e:
            skipped_files += 1
            logging.error(f"处理文件 {file_path} 时出错: {str(e)}", exc_info=True)
    
    # 写入合并后的JSON文件
    try:
        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            logging.info(f"创建输出目录: {output_dir}")

        with open(output_file, 'w', encoding='utf-8') as outfile:
            json.dump(combined_data, outfile, indent=2, ensure_ascii=False)
        
        logging.info(f"成功合并 {processed_files} 个文件到 {output_file}")
        logging.info(f"跳过 {skipped_files} 个无效文件")
        
    except Exception as e:
        logging.error(f"写入输出文件失败: {str(e)}", exc_info=True)

if __name__ == "__main__":
    # 示例用法
    input_directory = r"data\processed_data\data_tag"
    output_path = r"data\processed_data\combined\combined_tag.json"
    merge_json_files(input_directory,output_path)