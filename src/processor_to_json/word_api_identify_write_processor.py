# Word文档智能解析处理模块
# 功能：提取Word文档内容，调用AI模型识别工厂信息，转换为标准JSON格式
# 特性：支持DOC/DOCX格式、文本提取、电话号码验证、微信二维码提取、多轮验证机制

from docx import Document
import os
import re
import logging
import json
import sys

# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from src.utils.clean_factory_name import clean_factory_name
from src.utils.SaveImg_wechat_qr import extract_images_from_docx
from src.utils.convert_doc_docx import convert_doc_to_docx_and_replace
from src.utils.save_result_to_json import make_vendor_folder,save_result_to_vendor_folder
from src.processor_to_json.processor_rely.outmodel_results_validator import validate_and_get_result
from src.processor_to_json.processor_rely.model_word_identify import extract_word_text_info


# 日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#---------------------- 数据提取和验证工具函数 --------------------------------

# --- 文本行电话号码提取函数 ---
def extract_phone_numbers_from_lines(lines:list) -> list:
    """
    从文本行列表中提取11位电话号码

    提取逻辑：
    1. 使用正则表达式匹配11位数字序列
    2. 清理电话号码中的空格和短横线
    3. 去重并返回电话号码列表

    参数：
        lines (list): 文本行列表
        
    返回：
        list: 提取到的11位电话号码列表

    """
    phone_numbers = []
    phone_pattern = re.compile(r'(?<!\d)(?:\d[\s\-]*){11}(?!\d)')
    
    for line in lines:
        matches = phone_pattern.findall(line)
        for match in matches:
            # 清理电话号码（去除空格和短横线）
            clean_phone = re.sub(r'[\s\-]', '', match)
            if clean_phone not in phone_numbers:
                phone_numbers.append(clean_phone)
    
    return phone_numbers

# --- JSON字符串电话号码提取函数 ---
def extract_phone_numbers_from_json(json_str:str) -> list:
    """
    从JSON字符串中提取11位电话号码

    提取逻辑：
    1. 解析JSON字符串为Python对象
    2. 递归遍历JSON中的所有字符串值
    3. 使用正则表达式匹配11位电话号码
    4. 清理并去重电话号码

    参数：
        json_str (str): JSON字符串
        
    返回：
        list: 提取到的11位电话号码列表
    """
    try:
        # 解析JSON字符串
        if isinstance(json_str, str):
            data = json.loads(json_str)
        else:
            data = json_str
            
        phone_numbers = []
        phone_pattern = re.compile(r'(?<!\d)(?:\d[\s\-]*){11}(?!\d)')
        
        # 递归搜索JSON中的所有字符串值
        def search_phones(obj):
            if isinstance(obj, dict):
                for value in obj.values():
                    search_phones(value)
            elif isinstance(obj, list):
                for item in obj:
                    search_phones(item)
            elif isinstance(obj, str):
                matches = phone_pattern.findall(obj)
                for match in matches:
                    clean_phone = re.sub(r'[\s\-]', '', match)
                    if clean_phone not in phone_numbers:
                        phone_numbers.append(clean_phone)
        
        search_phones(data)
        return phone_numbers
        
    except (json.JSONDecodeError, TypeError) as e:
        logging.error(f"解析JSON时出错: {e}")
        return []


#---------------------- 文档内容提取和处理模块 --------------------------------

# --- Word文档文本提取函数 ---
def extract_text_info(file_path:str) -> list:
    """
    从DOCX文件中按行提取文本内容

    处理流程：
    1. 验证文件存在性和格式
    2. 自动转换DOC格式为DOCX格式
    3. 打开DOCX文档
    4. 遍历所有段落提取文本
    5. 过滤空行并返回文本行列表

    参数：
        file_path (str): Word文档文件路径
        
    返回：
        list: 包含所有文本行的列表，每个元素是一行文字

    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            logging.error(f"文件不存在: {file_path}")
        
        # 检查并转换文件格式
        if not file_path.lower().endswith('.docx'):
            if file_path.lower().endswith('.doc'):
                convert_doc_to_docx_and_replace(file_path)
            else:
                logging.error(f"文件格式错误，需要.docx文件: {file_path}")
                return []
        
        # 打开DOCX文档
        doc = Document(file_path)
        
        # 提取所有文本行
        text_lines = []
        
        # 遍历所有段落
        for paragraph in doc.paragraphs:
            # 获取段落文本并去除首尾空白
            text = paragraph.text.strip()
            # 如果段落不为空，添加到列表中
            if text:
                text_lines.append(text)
           
        return text_lines
        
    except Exception as e:
        logging.error(f"处理文件时发生错误: {e}")
        return []

# --- AI模型输出验证函数 ---
def verification_info(lines:list) -> dict:
    """
    AI模型多轮验证函数：确保模型输出一致性并验证电话号码准确性

    验证流程：
    1. 进行3轮验证，每轮调用3次模型
    2. 对比三次模型输出的文本内容一致性
    3. 提取文本行和模型输出中的电话号码
    4. 验证电话号码一致性
    5. 返回验证通过的结果或错误信息

    参数：
        lines (list): 文本行列表
        
    返回：
        dict: 验证通过返回解析后的字典，否则返回None

    验证机制：
        - 支持3轮重试机制
        - 严格的文本内容一致性检查
        - 电话号码验证确保数据准确性
        - 详细的日志记录和错误处理
    """
    try:
        # 进行3轮验证
        for round_num in range(3):
            logging.info(f"=== 第{round_num + 1}轮验证 ===")
            
            # 调用三次模型
            results = []
            for i in range(3):
                result = extract_word_text_info(lines)
                # 检查API调用是否成功
                if result is None or result == '':
                    logging.error(f"第{i+1}次API调用失败，返回空结果")
                    continue

                results.append(result)
                logging.info(f"第{i+1}次结果: {result}")
            
            # 使用通用工具验证并获取结果
            final_result = validate_and_get_result(results)
            
            if final_result is not None:
                # 验证电话号码
                lines_phones = extract_phone_numbers_from_lines(lines)
                model_phones = extract_phone_numbers_from_json(final_result)
                
                if set(lines_phones) != set(model_phones):
                    logging.error(f"电话号码不一致！lines: {lines_phones}, 模型: {model_phones}")
                    return None
                
                logging.info("模型分类成功")
                return final_result
            else:
                if round_num < 2:
                    logging.warning(f"第{round_num + 1}轮验证失败，准备下一轮...")
                    continue
                else:
                    logging.error("经过3轮验证，三次模型输出均不一致，返回第2个模型结果，请人工检查！")
                    
                    return json.loads(results[1]) if isinstance(results[1], str) else results[1]
        
    except Exception as e:
        logging.error(f"验证过程中发生错误: {e}")
        return None

# --- Word文档转JSON主处理函数 ---
def word_to_json(file_path:str,output_directory:str) -> bool:
    """
    Word文档转JSON完整处理流程

    处理流程：
    1. 从Word文档中提取文本行
    2. 调用AI模型验证函数处理文本并验证结果
    3. 清洗工厂名称并创建输出文件夹
    4. 提取微信二维码图片
    5. 保存JSON结果文件

    参数：
        file_path (str): Word文档路径
        output_directory (str): 输出目录路径
        
    返回：
        bool: 处理成功返回True，失败返回False

    处理特性：
        - 集成文本提取和模型验证功能
        - 自动创建厂商专属文件夹
        - 提取并保存微信二维码图片
        - 包含完整的错误处理机制
    """
    # 提取文档文本行
    lines = extract_text_info(file_path)
    
    # AI模型验证处理
    json_result = verification_info(lines)
    
    if json_result:
        # 清洗工厂名称并创建文件夹
        factory_name=clean_factory_name(json_result.get('厂商名称'))
        vendor_folder = make_vendor_folder(factory_name,output_directory)
        
        # 提取微信二维码图片

        img_path = extract_images_from_docx(file_path,vendor_folder)
        if img_path:
            json_result['微信'] = img_path
        else:
            logging.error(f"提取二维码失败")
            
        json_result['文件路径'] = file_path
        
        # 保存JSON结果
        outpath=save_result_to_vendor_folder(vendor_folder, json_result)

        logging.info(f"Word文档已转换为JSON格式,输出路径: {outpath}")
        return True
    else:
        logging.error("Word文档转换为JSON格式失败")
        return False

#---------------------- 程序测试入口 --------------------------------

if __name__ == "__main__":
    # 测试文件路径配置
    test_file = r"tests\word\惠州市隆青工艺品有限公司-2025.02.28.docx"
    output_directory = r"tests\word\惠州市隆青工艺品有限公司"

    # 执行转换测试
    if word_to_json(test_file,output_directory):
        print("转换成功")
    else:
        print("转换失败")
        
    