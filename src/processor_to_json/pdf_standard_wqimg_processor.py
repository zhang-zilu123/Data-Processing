import re
import sys
import os
import json
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *
import logging
from src.utils.save_result_to_json import make_vendor_folder, save_result_to_vendor_folder
from src.utils.SaveImg_wechat_qr import extract_images_from_pdf
from src.utils.clean_factory_name import clean_factory_name
from src.utils.extract_by_row import extract_text_lines_from_pdf

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')



#---------------------- 辅助功能函数 --------------------------------
# --- 函数1：日期文本格式识别 ---
def is_date_text(text:str) -> bool:
    """
    检查文本内容是否符合日期格式模式
    
    识别功能：
    通过config中DATE_PATTERNS配置的日期正则表达式模式，判断输入文本是否为日期信息。
    用于在文档解析过程中识别和分类日期相关的内容。
    
    识别流程：
    1. 基础文本有效性检查（长度和非空）
    2. 清理文本格式（去除首尾空白）
    3. 遍历配置的日期模式进行匹配
    4. 返回匹配结果
    
    参数：
        text (str): 待检查的文本内容
        
    返回：
        bool: True表示文本符合日期格式，False表示不符合

    """

    # 步骤1：基础有效性检查
    if not text or len(text) > 50: # 简单的长度过滤，避免检查过长的文本
        return False
        
    # 步骤2：清理文本格式
    cleaned_text = text.strip()
    
    # 步骤3：遍历配置的日期模式进行匹配
    for pattern in DATE_PATTERNS:
        if re.search(pattern, cleaned_text):
            return True
    return False


# --- 函数2：联系人和电话号码识别和提取 ---
def find_contact_line(text:str) -> bool:
    """
    识别文本中的联系方式信息（电话号码、职位、姓名）
    
    识别功能：
    使用正则表达式模式识别文本中的电话号码，支持多种格式的电话号码
    （包括带空格、短横线分隔的11位数字）。用于自动提取联系方式行。
    
    识别流程：
    1. 编译电话号码识别的正则表达式
    2. 在输入文本中搜索匹配模式
    3. 返回匹配结果或空值
    
    参数：
        text (str): 待搜索的文本内容
        
    返回：
        str: 如果找到电话号码则返回原始文本，否则返回None

    """
    # 匹配手机号的正则模式（11 位数字，允许有空格或短横线）
    phone_pattern = re.compile(r"(?<!\d)(?:\d[\s\-]*){11}(?!\d)")
    
    # 放宽的职位关键词匹配模式（支持空格和多种职位）
    position_keywords = ['经理', '负责人', '主管', '总监', '主任', '部长', '董事', '总','总经理','助理','秘书','代表','销售','业务']
    position_pattern = re.compile(r'.*(' + '|'.join(position_keywords) + r').*')

    # 匹配2-4个中文字的模式(联系方式规则化，格式基本为：姓名+职位+电话，而部分提取行中，会存在只有2-4个中文字的情况，所以不匹配)
    # chinese_name_pattern = re.compile(r'^[\u4e00-\u9fa5]{2,4}$')

    # 检查是否匹配手机号、职位关键词
    if phone_pattern.search(text) or position_pattern.search(text): # or chinese_name_pattern.search(text):
        return True
    else:
        return False
    

# --- 函数3：厂商名称格式识别 ---
def is_vendor_name(line:str) -> bool:
    """
    判断文本行是否为厂商名称格式
    
    识别功能：
    通过检查文本特征（无冒号且包含企业关键词），判断当前行
    是否为厂商名称信息。用于自动识别和分类厂商名称字段。
    
    识别流程：
    1. 检查文本中是否包含冒号
    2. 检查是否包含企业类型关键词
    3. 综合判断返回结果
    
    参数：
        line (str): 待判断的文本行
        
    返回：
        bool: True表示可能是厂商名称，False表示不是
        

    """
    # 步骤1：检查是否包含冒号（有冒号的通常是标签：值格式）
    if ':' in line or ':' in line:
        return False
        
    # 步骤2：检查是否包含企业类型关键词
    if any(kw in line for kw in ['公司', '厂', '集团', '有限', '企业' , '⼚' ]):
        return True
    return False

# --- 函数4：主销市场识别 ---
def extract_market_info(text: str) -> str:
    """
    从文本中提取市场相关信息行
    
    识别功能：
    通过检查文本是否包含市场相关关键词（主销市场/市场占比），
    提取包含完整市场信息的文本行。用于从非结构化文本中定位市场信息字段。
    
    识别流程：
    1. 检查文本是否包含预设的市场关键词（支持多种空格变体）
    2. 匹配到关键词则返回整行净化后的文本
    3. 未匹配到则返回空字符串
    
    参数：
        text (str): 待处理的原始文本行
        
    返回：
        str: 包含市场信息的净化文本行（去除首尾空格），若无匹配则返回空字符串
    """
    # 关键词列表（支持各种空格变体）
    keywords = [
        '主 销 市 场',
        '市 场 占 ⽐',
        '主销市场',
        '市场占比',
        '市场占⽐'
    ]
    
    # 检查是否包含任一关键词
    for kw in keywords:
        if kw in text:
            return text.strip()
    return ""


#---------------------- PDF文字按行提取并进行分类和清洗 --------------------------------
# --- 对提取文本进行分类 ---
def classify_pdf_text_lines(lines: list[str]) -> dict[str, list[str]]:
    """
    对PDF提取的文本行进行分类

    识别功能：
    1. 基于业务规则的智能识别：
       - 日期仅在最后一行检查
       - 厂商名称仅在前三行检查
       - 联系人信息使用正则匹配
       - 市场信息通过关键词识别
    2. 自动归类未识别文本
    
    处理流程：
    1. 初始化分类数据结构
    2. 按行遍历处理：
       a. 最后一行 → 日期检查（使用is_date_text）
       b. 前三行 → 厂商检查（使用is_vendor_name）
       c. 所有行 → 联系人检查（使用find_contact_line）
       d. 所有行 → 市场信息检查（使用extract_market_info）
       e. 未匹配 → 归入others分类
    3. 返回结构化分类结果

    参数:
        lines (list[str]): 从PDF中提取的文本行列表
        
    返回:
        dict: 分类后的结构化数据
    """
    classified_data = {
        'dates': [],
        'vendors': [],
        'contacts': [],
        'markets': [],
        'others': []
    }
    
        # 1. 检查是否为日期信息
    for i, line in enumerate(lines):
        # 只在最后一行检查日期
        if i == len(lines) - 1 and is_date_text(line):
            classified_data['dates'].append(line)
            continue
            
        # 2. 检查是否为厂商名称（仅在前三行检查）
        if i < 3 and is_vendor_name(line):
            classified_data['vendors'].append(line)
            continue
            
        # 3. 检查是否为联系人信息
        if i < 7 and find_contact_line(line):
            classified_data['contacts'].append(line)
            continue
            
        # 4. 检查是否为市场信息
        market_info = extract_market_info(line)
        if market_info:
            classified_data['markets'].append(line)
            continue
            
        # 5. 未分类的其他文本
        classified_data['others'].append(line)
    
    return classified_data



# ---------------------- 将文本文件转化为格式化JSON文件 --------------------------------
def convert_to_json_format(classified_data: dict[str, list[str]]) -> dict[str, str]:
    """
    将清洗后的分类数据转换为标准化的JSON格式输出

    核心功能：
    1. 数据结构转换：将分类数据字典转换为config定义的JSON格式
    2. 字段映射处理：根据UNCLASSIFIED_TEXT_MAPPING_PDF将未分类文本映射到标准字段
    3. 内容合并规则：
       - 厂商名称：使用"/"连接多个厂商
       - 联系方式和市场信息：使用换行符连接
       - 日期信息：使用换行符连接
    4. 智能分段处理：自动识别并处理备注等特殊字段内容

    识别流程：
    1. 初始化JSON数据结构（基于JSON_FORMAT模板）
    2. 处理已分类字段
    3. 处理未分类信息：
        a. 识别节标题（根据SECTION_MAPPING）
        b. 处理特殊字段（厂房面积/员工人数/年产值）
        c. 归类到相应字段
    4. 返回标准化JSON数据

    参数：
        cleaned_data (dict[str, list[str]]): 清洗后的分类数据
            
    返回：
        dict[str, str]: 标准化的JSON格式数据，符合config定义的输出格式
    """
    # 初始化JSON数据结构
    json_data = JSON_FORMAT.copy()
    
    # 1. 厂商名称 - 合并所有厂商名称
    if classified_data['vendors']:
        json_data['厂商名称'] = '/'.join(classified_data['vendors'])
    
    # 2. 联系方式 - 合并所有联系人信息
    if classified_data['contacts']:
        json_data['联系方式'] = '\n'.join(classified_data['contacts'])
    
    # 3. 主销市场 - 合并所有市场信息
    if classified_data.get('markets'):
        # 处理第一行（主销市场）
        part1 = classified_data['markets'][0].replace('主销市场', '') if classified_data['markets'] else ""
        # 处理第二行（市场占比）
        part2 = classified_data['markets'][1].replace('市场占比', '') if len(classified_data['markets']) > 1 else ""
        
        # 根据条件组合结果
        if part1 and part2:
            json_data['主销市场'] = f"{part1}\n市场占比：{part2}"
        elif part2:
            json_data['主销市场'] = f"市场占比：{part2}"
        else:
            json_data['主销市场'] = part1 or ""
        
    # 4. 日期 - 合并所有日期信息
    if classified_data['dates']:
        json_data['日期'] = ''.join(classified_data['dates'])
    
    # 5. 处理未分类信息
    current_section = '备注'  # 默认分类

    for line in classified_data['others']:
        line = line.strip()
        if not line:
            continue
        
        # 检查是否是新的节标题或数值字段
        section_found = False
        for section_key, section_value in UNCLASSIFIED_TEXT_MAPPING_PDF.items():
            if section_key in line:
                current_section = section_value
                section_found = True
                remaining_content = line.replace(section_key, '').strip()
                break
        
        # 处理数值字段的情况（整合特殊处理）
        if not section_found and any(keyword in line for keyword in ['厂房面积', '员工人数', '年产值']):
            current_section = '备注'
            section_found = True
            remaining_content = line  # 数值字段整行作为内容
        
        # 统一处理内容添加
        json_data.setdefault(current_section, '')
        if section_found:
            # 如果是新节且已有内容则添加换行
            if json_data[current_section]:
                json_data[current_section] += '\n'
            # 添加内容（标题行剩余内容或数值字段整行）
            json_data[current_section] += remaining_content
        else:
            # 普通内容直接追加
            json_data[current_section] += line

    return json_data


# ---------------------- PDF主处理函数 --------------------------------
def process_pdf_file(pdf_path: str, output_base_dir: str) -> bool:
    """
    完整的PDF文件处理，便于外界调用
    
    核心功能：
    1. 文本内容提取与结构化：
       - 提取PDF文本内容
       - 智能分类文本行
       - 清洗和标准化数据
       - 转换为标准JSON格式
    2. 图片提取：
       - 扫描并提取PDF中的微信二维码
    3. 结果组织与存储：
       - 创建厂商专属目录
       - 保存结构化数据到JSON文件
       - 保存二维码图片（如存在）
    
    处理流程：
    1. 输入验证：检查PDF文件是否存在
    2. 文本处理流水线：
        提取文本行 → 文本分类 → 数据清洗 → 格式转换
    3. 厂商信息处理：
       - 获取/清洗厂商名称
       - 创建厂商专属目录
    4. 二维码提取：
       - 仅扫描第一页
       - 保存到厂商目录
    5. 结果保存：
       - 更新JSON中的文件路径信息
       - 写回最终JSON文件
    
    参数：
        pdf_path (str): 输入的PDF文件路径
        output_base_dir (str): 输出目录基础路径
    
    返回：
        dict[str, Union[str, bool]]: 处理结果字典
    """
    
    try:
        # 检查输入文件是否存在
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        # 步骤1：提取文本内容并转换为JSON
        logging.info(f"开始处理PDF文件: {pdf_path}")
        # 1：提取文本行
        lines = extract_text_lines_from_pdf(pdf_path)

        # 2：分类文本内容
        classified_data = classify_pdf_text_lines(lines)

        # 4：转换为JSON格式
        json_data = convert_to_json_format(classified_data)

        # 5：添加文件信息
        json_data.update({'文件路径': pdf_path })

        # 6：获取并清洗厂商名称
        factory_name = json_data.get('厂商名称', os.path.splitext(os.path.basename(pdf_path))[0])
        if not factory_name or factory_name == '未知厂商':
            factory_name = os.path.splitext(os.path.basename(pdf_path))[0]
            json_data['厂商名称'] = factory_name
            logging.warning(f"厂商名称为空或未知，使用文件名作为厂商名: {factory_name}")
   
        
        # 7：创建厂商专属文件夹
        vendor_folder = make_vendor_folder(factory_name, output_base_dir)
        # logging.info(f"创建厂商文件夹: {vendor_folder}")

         # 8：提取PDF中的微信二维码(只处理第一页)
        qr_path = extract_images_from_pdf(pdf_path, vendor_folder, page_num=1)
        if qr_path:
            # 修改这里：保存完整路径而不仅仅是文件名
            json_data['微信'] = qr_path
            logging.info(f"找到并保存微信二维码: {qr_path}")
        else:
            logging.error(f"未找到微信二维码。")

        
        # 9：保存JSON结果到厂商文件夹
        json_path = save_result_to_vendor_folder(vendor_folder, json_data)
        if not json_path:
            logging.error("保存JSON文件失败")
            return False
        else:
            logging.info(f"PDF文件处理完成: {pdf_path}")
            return True
        
    except Exception as e:
        logging.error(f"处理PDF文件时出错: {pdf_path}", exc_info=True)
        return False



# 调用示例
if __name__ == "__main__":
    # 测试文件路径配置
    file_path = r"data\input_data\pdf\6月宁波到访工厂资料\龙泉市腾翔竹木有限责任公司1749544725.pdf"
    output_directory = r"tests\processed_data\pdf\曹县融悦木业有限公司"
    
    # 执行PDF处理测试
    if process_pdf_file(file_path,output_directory):
        print("处理成功")
    else:
        print("处理失败")
