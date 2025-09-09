# 判断验证模型返回的结果是否一致

import re
import json
import logging
from typing import List, Union, Dict, Any

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def normalize_text(text: Union[str, Any]) -> str:
    """
    标准化文本，去除符号、换行符等格式差异
    """
    if not isinstance(text, str):
        text = str(text)
    
    # 去除换行符、制表符等空白字符
    text = re.sub(r'\s+', '', text)
    # 去除标点符号（保留中文字符和数字）
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', text)
    return text.lower()

def parse_model_result(result: Union[str, Dict]) -> Union[Dict, None]:
    """
    解析模型返回的结果，处理各种可能的格式问题
    """
    if isinstance(result, dict):
        return result
    
    if not isinstance(result, str) or not result.strip():
        return None
    
    # 尝试多种解析方式
    try:
        # 方式1：直接解析JSON
        return json.loads(result)
    except json.JSONDecodeError:
        pass
    
    try:
        # 方式2：替换单引号为双引号后解析
        json_str = result.replace("'", '"')
        return json.loads(json_str)
    except json.JSONDecodeError:
        pass
    
    try:
        # 方式3：提取JSON部分（去除可能的markdown标记）
        # 查找 { 开始到 } 结束的部分
        json_match = re.search(r'\{.*\}', result, re.DOTALL)
        if json_match:
            json_str = json_match.group(0)
            return json.loads(json_str)
    except json.JSONDecodeError:
        pass
    
    try:
        # 方式4：处理可能的格式问题
        # 移除可能的markdown标记
        cleaned_result = result.replace('```json', '').replace('```', '').strip()
        # 替换单引号
        cleaned_result = cleaned_result.replace("'", '"')
        return json.loads(cleaned_result)
    except json.JSONDecodeError:
        pass
    
    # 如果所有方式都失败，记录原始内容并返回None
    logging.error(f"无法解析模型返回的JSON格式: {result[:200]}...")
    return None


def compare_results(results: List[Union[str, Dict]]) -> bool:
    """
    比较多个模型结果是否一致（只比较文本内容，忽略格式）
    
    输入：
        results (list): 包含多个结果的列表
        
    输出：
        bool: 如果所有结果文本内容一致返回True，否则返回False
    """
    try:
        if len(results) < 2:
            logging.error(f"结果数量不足，需要至少2个结果，实际只有{len(results)}个")
            return False
        
        # 解析结果为字典
        parsed_results = []
        for i, result in enumerate(results):
            parsed = parse_model_result(result)
            
            if parsed is None:
                logging.error(f"第{i+1}次结果解析失败")
                return False
            
            if not isinstance(parsed, dict):
                logging.error(f"第{i+1}次结果不是字典格式: {type(parsed)}")
                return False
            
            parsed_results.append(parsed)
        
        # 比较所有结果是否一致
        first_result = parsed_results[0]
        for i in range(1, len(parsed_results)):
            current_result = parsed_results[i]
            
            # 比较所有字段
            all_keys = set(first_result.keys()) | set(current_result.keys())
            for key in all_keys:
                first_value = first_result.get(key, '')
                current_value = current_result.get(key, '')
                
                if normalize_text(first_value) != normalize_text(current_value):
                    logging.info(f"第{i+1}次结果与第1次结果在字段 '{key}' 上不一致")
                    return False
        
        logging.info(f"{len(results)}次模型输出文本内容一致，验证通过")
        return True
        
    except Exception as e:
        logging.error(f"比较结果时发生错误: {e}")
        return False

def validate_and_get_result(results: List[Union[str, Dict]]) -> Union[Dict, None]:
    """
    验证模型输出的一致性并返回有效结果
    
    输入：
        results (list): 包含多个结果的列表
        
    输出：
        Union[Dict, None]: 验证通过返回解析后的字典，否则返回None
    """
    try:
        # 进行一致性验证
        if compare_results(results):
            # 返回第一个解析后的结果
            parsed_result = parse_model_result(results[0])
            if parsed_result is not None:
                return parsed_result
            else:
                logging.error("解析成功结果的JSON失败")
                return None
        else:
            logging.info("模型输出验证失败")
            return None
            
    except Exception as e:
        logging.error(f"验证模型输出时发生错误: {e}")
        return None
    



if __name__ == "__main__":
    results = [
        {'主销市场': '欧美、南美、东南亚', '备注': '工厂面积：8000平方米\n年  产  值：8000万\n员工人数：80人'},
        {'主销市场': '欧美、南美、东南亚', '备注': '工厂面积：8000平方米 年  产  值：8000万 员工人数：80人'}]
    print(validate_and_get_result(results))