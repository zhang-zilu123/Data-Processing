# PDF工厂信息综合处理模块
# 功能：从PDF文件中提取多工厂信息和产品图片，转换为标准化JSON格式
# 特性：支持多页处理、文字图片分离识别、图片智能过滤、工厂信息自动映射、批量处理

import fitz  # PyMuPDF
from pathlib import Path
import re
import logging
import os
import sys
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *
from src.utils.clean_factory_name import clean_factory_name
from src.utils.save_result_to_json import make_vendor_folder,save_result_to_vendor_folder

# 日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#---------------------- 文本和数据提取工具函数 --------------------------------

# --- PDF文本结构化字段提取函数 ---
def extract_fields(text: str) -> dict:
    """
    从PDF文本内容中提取结构化的工厂信息字段
    
    提取功能：
    解析PDF页面的文本内容，识别和提取工厂相关的各种信息字段。
    支持多种格式的文本结构，通过冒号分隔符识别键值对，
    并处理特殊情况如缺失冒号、多行内容合并等。
    
    提取流程：
    1. 标准化文本格式（中文冒号转英文冒号）
    2. 按行分割并处理无冒号行的合并
    3. 识别并跳过特殊标记行(如"第x家:")
    4. 提取键值对并构建结构化数据
    5. 处理第一行作为工厂名称的特殊情况
    
    参数：
        text (str): PDF页面提取的原始文本内容
        
    返回：
        dict: 结构化的工厂信息字典，键为字段名，值为对应内容
        
    文本处理特点：
        - 自动标准化冒号格式（英文转中文）
        - 智能合并无冒号的续行内容
        - 跳过"第x家:"格式的分组标记
        - 自动识别第一行作为工厂名称
        
    字段识别规则：
        - 通过英文冒号":"分隔键值对
        - 自动过滤空值字段
        - 支持工厂名称的多种表达方式

    """
    # 标准化文本格式 - 将中文冒号统一转换为英文冒号
    text = text.replace('：',':')
    
    # 按行分割文本并处理无冒号行的合并
    lines = text.splitlines()
    processed_lines = []
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
            
        # 如果当前行没有冒号，且不是第一行，且上一行存在，则合并到上一行
        if ':' not in line and i > 0 and processed_lines:
            processed_lines[-1] += ' ' + line
        else:
            processed_lines.append(line)
    
    # 初始化结果字典
    result = {}

    # 遍历处理后的文本行，提取字段信息
    for i, line in enumerate(processed_lines):
        # 将包含"第x家："的标记行作为厂商名称
        if re.match(r'^第.*家:', line):
                      
            result['工厂名称'] = re.sub(r'^第.*家:\s*', '', line)
            continue

        # 查找冒号位置进行键值分离
        colon_pos = line.find(':')
        
        # 跳过处理第一行作为工厂名称的情况
        if i == 0 and colon_pos == -1 and '工厂名称' not in result:
            continue
        elif colon_pos == -1:
            continue
            
        # 提取键值对
        key = line[:colon_pos].strip()
        value = line[colon_pos + 1:].strip()
        
        # 跳过没有实际内容的键值对
        if not value:
            continue
            
        result[key] = value
    
    return result




# --- PDF页面产品图片提取函数 ---
def extract_images_from_pdf(page, doc, output_dir:str, factory_name:str, page_num: int) -> str:
    """
    从PDF的指定页面中提取所有有效图片并保存到本地文件夹
    
    提取功能：
    遍历PDF页面中的所有图像对象，根据尺寸和质量筛选有效的产品图片，
    过滤掉背景图片、装饰图标等无关图像，按标准格式命名保存。
    
    提取流程：
    1. 获取页面中的所有图像对象列表
    2. 遍历每个图像并提取图像数据
    3. 检查图像尺寸并应用过滤规则
    4. 生成标准化的图片文件名
    5. 保存符合条件的图片到指定目录
    6. 统计并返回保存结果
    
    参数：
        page: PyMuPDF页面对象，包含图像数据
        doc: PyMuPDF文档对象，用于提取图像
        output_dir (str): 图片保存的目标目录路径
        factory_name (str): 工厂名称，用于图片文件命名
        page_num (int): 页面编号，用于图片文件命名
        
    返回：
        str: 成功保存图片时返回保存目录路径，无有效图片时返回None
        
    图片过滤规则：
        - 跳过分辨率>=4005x2251的背景图片
        - 跳过分辨率<100x100的小图标
        - 保留中等尺寸的产品展示图片
        
    文件命名格式：
        - 格式：{工厂名称}_第{页面号}页_图片{序号}.{扩展名}
        - 支持原始图片格式（jpg、png等）
        
    
    """
    try:
        # 获取页面中的所有图像对象
        img_list = page.get_images()
        
        if not img_list:
            logging.info(f"第{page_num}页没有找到图像")
            return None
            
        # 初始化保存计数器
        saved_count = 0
        
        # 遍历处理每个图像对象
        for img_index, img in enumerate(img_list):
            # 获取图像的引用ID
            xref = img[0]
            
            try:
                # 提取图像的原始数据
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # 获取图像尺寸信息用于过滤
                width = base_image.get("width", 0)
                height = base_image.get("height", 0)
                
                # 应用图片过滤规则
                # 跳过分辨率过大的背景图片
                if width >= 4005 and height >= 2251:
                    logging.info(f"跳过分辨率为 {width}x{height} 的背景图片")
                    continue
                
                # 跳过过小的图片（可能是图标或装饰性图片）
                if width < 100 or height < 100:
                    logging.info(f"跳过分辨率过小的图片 {width}x{height}")
                    continue
                
                # 生成标准化的图片文件名
                img_filename = f"{factory_name}_第{page_num}页_图片{img_index + 1}.{image_ext}"
                img_path = os.path.join(output_dir, img_filename)
                
                # 保存图片到本地文件
                with open(img_path, "wb") as img_file:
                    img_file.write(image_bytes)
                
                saved_count += 1
                
            except Exception as e:
                logging.error(f"提取第{img_index + 1}个图像时出错: {str(e)}")
                continue
        
        # 返回处理结果
        if saved_count > 0:
            logging.info(f"第{page_num}页成功保存了{saved_count}张图片")
            return output_dir
        else:
            logging.info(f"第{page_num}页没有保存任何图片")
            return None
        
    except Exception as e:
        logging.error(f"提取第{page_num}页图片时出错: {str(e)}")
        return None



# --- 信息映射到标准JSON格式函数 ---
def map_to_standard_json(info_dict: dict, pdf_path: str, img_folder_path: str = None) -> dict:
    """
    将PDF提取信息映射转换为标准JSON格式
    
    映射功能：
    根据预定义的字段映射配置，将PDF中提取的各种格式的字段名和内容
    转换为统一的标准JSON格式。支持字段别名映射、前缀处理、未匹配字段的自动归类
    
    映射流程：
    1. 初始化结果字典和匹配跟踪
    2. 遍历标准JSON格式的所有字段
    3. 使用配置的关键词进行模糊匹配
    4. 应用字段前缀和格式化规则
    5. 处理未匹配字段并归类到备注
    6. 添加文件路径等额外信息
    
    参数：
        info_dict (dict): 从PDF中提取的原始字段字典
        pdf_path (str): PDF文件的完整路径
        img_folder_path (str, optional): 图片文件夹路径，默认为None
        
    返回：
        dict: 标准化的JSON格式工厂信息，符合JSON_FORMAT格式
        
    映射规则：
        - 使用TEXT_LABELS_pdf配置进行字段名映射
        - 支持一对多的字段别名匹配
        - 模糊匹配允许部分字段名包含关键词
        
    格式化处理：
        - 特定字段（联系方式、主销市场等）添加原字段名前缀
        - 多个匹配值使用分号连接
        - 空值字段填充""
        
    未匹配字段处理：
        - 自动收集所有未映射的字段
        - 统一归类到"备注"字段中
        - 保持原有格式便于后续人工审核
 

    """
    # 初始化结果字典和匹配跟踪
    result = {}
    matched_keys = set()  # 记录已匹配的原始字段名
    
    # 定义需要添加原字段名前缀的特殊字段
    prefix_fields = {'联系方式', '主销市场', '合作情况', '备注'}
    
    # 遍历标准JSON格式的所有字段进行映射
    for json_key in JSON_FORMAT.keys():
        # 获取当前字段的映射关键词列表
        mapped_keywords = TEXT_LABELS_pdf.get(json_key, [json_key])
        
        found_values = []
        
        # 在原始数据中查找匹配的字段
        for info_key, info_value in info_dict.items():
            if info_key in matched_keys:
                continue
                
            # 检查是否包含任何映射关键词
            for keyword in mapped_keywords:
                if keyword in info_key and info_value.strip():
                    # 根据字段类型应用格式化规则
                    if json_key in prefix_fields:
                        # 需要添加原字段名前缀的字段
                        formatted_value = f"{info_key}:{info_value}"
                    else:
                        # 直接使用原值的字段
                        formatted_value = info_value
                    
                    found_values.append(formatted_value)
                    matched_keys.add(info_key)
                    break
        
        # 设置字段值（多个匹配值用分号连接）
        if found_values:
            result[json_key] = '; '.join(found_values)
        else:
            result[json_key] = ''
    
    # 处理未匹配的字段，添加到备注中
    unmatched_items = []
    for info_key, info_value in info_dict.items():
        if info_key not in matched_keys and info_value.strip():
            unmatched_items.append(f"{info_key}:{info_value}")
    
    # 将未匹配项目合并到备注字段
    if unmatched_items:
        if result['备注'] == '':
            result['备注'] = '; '.join(unmatched_items)
        else:
            result['备注'] += '; ' + '; '.join(unmatched_items)
    
    # 添加文件路径等额外字段
    result['文件路径'] = pdf_path
    result['图片文件夹路径'] = img_folder_path if img_folder_path else ''
    
    return result

#---------------------- PDF文件综合处理主模块 --------------------------------

# --- PDF文件综合处理主函数 ---
def process_pdf(pdf_path: str, output_dir: str ) -> bool:
    """
    PDF工厂信息文件的完整处理主函数
    
    处理功能：
    统一处理包含多个工厂信息的PDF文件，自动识别文字页和图片页，
    提取工厂基础信息和产品图片，转换为标准化的JSON格式数据。
    支持多工厂信息的批量处理和智能分页识别。
    
    处理流程：
    1. 打开PDF文件
    2. 逐页分析页面类型（文字页或图片页）
    3. 文字页：提取工厂信息并创建图片文件夹
    4. 图片页：提取产品图片到对应工厂文件夹
    5. 智能判断工厂信息完整性并生成JSON记录
    6. 清理资源并返回所有工厂的处理结果
    
    参数：
        pdf_path (str): 待处理的PDF文件完整路径
        output_dir (str): 图片和数据的输出根目录路径
        
    返回：
        bool: 处理成功返回True，失败返回False
        
    页面识别机制：
        - 文字页：文本内容长度>10个字符
        - 图片页：文本内容长度<=10个字符
        - 自动适应不同的PDF布局格式
        
        
    图片处理策略：
        - 图片页关联到最近的文字页工厂信息
        - 支持一个工厂对应多个图片页的情况
        - 自动过滤无效或装饰性图片
        

    
    """
    try:
        # 打开PDF文件
        doc = fitz.open(pdf_path)
        
        # 初始化处理状态变量
        page_index = 0             # 当前页面索引
        total_pages = len(doc)     # PDF总页数
        
        current_factory_info = None    # 当前工厂的信息字典
        current_factory_name = None    # 当前工厂的清洗后名称
        current_img_folder = None      # 当前工厂的图片文件夹路径

        old_factory_name = None    # 用于标记是否是新工厂

        success_count = 0           # 转换成功页数
        failed_count = 0            # 转换失败页数

        logging.info(f"开始处理PDF文件: {pdf_path}，共 {total_pages} 页")

        # 逐页处理PDF内容
        while page_index < total_pages:
            page = doc[page_index]
            text = page.get_text().strip()

            # 判断页面类型（文字页 vs 图片页）
            if len(text) > 10:
                # 文字页：提取工厂信息
                info = extract_fields(text)
                
                # 识别和处理工厂名称
                for key in ['工厂名称', '厂商名称', '公司名称']:
                    if info.get(key):
                        factory_name = info[key]
                        cleaned_name = clean_factory_name(factory_name)
                        if cleaned_name != old_factory_name:
                            info[key] = cleaned_name
                            break
                        break
                else:
                    # 如果没有找到明确的工厂名称，使用页面编号生成
                    factory_name = f"工厂_{page_index + 1}"
                    cleaned_name = factory_name
                
                # 创建工厂名称文件夹
                vendor_folder = make_vendor_folder(cleaned_name,output_dir)

                # 创建工厂专属的图片保存文件夹
                img_save_folder = os.path.join(vendor_folder, f"{cleaned_name}_产品图片")
                os.makedirs(img_save_folder, exist_ok=True)
                
                # 更新当前工厂的处理状态
                current_factory_info = info
                current_factory_name = cleaned_name
                current_img_folder = img_save_folder  # 设置图片文件夹路径
                
                logging.info(f"处理文字页，工厂: {current_factory_name}")
                
            else:
                # 图片页：提取产品图片
                if current_factory_name and current_img_folder:
                    try:
                        # 从图片页提取并保存图片
                        saved_folder_path = extract_images_from_pdf(page, doc, current_img_folder, current_factory_name, page_index + 1)
                        
                        if saved_folder_path:
                            logging.info(f"从第 {page_index + 1} 页提取并保存了图片到: {saved_folder_path}")
                        else:
                            logging.info(f"第 {page_index + 1} 页没有提取到有效图片")
                                
                    except Exception as e:
                        logging.error(f"处理第 {page_index + 1} 页图片时出错: {str(e)}")
                else:
                    logging.warning(f"第 {page_index + 1} 页是图片页，但没有对应的工厂信息")
            
            # 判断是否需要输出当前工厂的JSON记录
            # 条件：下一页是文字页（新工厂开始）或已到最后一页
            next_is_text_page = False
            if page_index + 1 < total_pages:
                next_text = doc[page_index + 1].get_text().strip()
                next_is_text_page = len(next_text) > 10
                
            if next_is_text_page or page_index + 1 == total_pages:
                if current_factory_info:
                    # 生成标准JSON格式并保存到记录列表
                    json_record = map_to_standard_json(current_factory_info, pdf_path, current_img_folder)
                    if json_record:
                        save_result_to_vendor_folder(vendor_folder, json_record)
                        logging.info(f"保存厂信息成功: {json_record.get('厂商名称')}")
                        success_count += 1
                    else:
                        failed_count+=1
                        logging.error(f"保存厂信息失败: {json_record.get('厂商名称')}")

            page_index += 1

        logging.info(f"多页数PDF文档处理完成: 总数：{total_pages}个, 成功：{success_count}个, 失败：{failed_count}个")
        return True

    except Exception as e:
        logging.error(f"处理PDF文件时出错: {str(e)}")
        return False
    finally:
        # 清理资源
        doc.close()

#---------------------- 程序测试入口 --------------------------------

if __name__ == "__main__":
    # 测试文件路径配置
    pdf_file = r"tests\pdf\柬埔寨工厂.pdf"
    output_dir = r"tests\processed_data\pdf\柬埔寨工厂"
    
    # 执行PDF文件处理测试
    results = process_pdf(pdf_file, output_dir)
