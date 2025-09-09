# 将包含多个工厂信息的Excel文件转换为标准化的JSON格式数据
# 支持多工作表处理，自动解析工厂信息字段，提取产品图片，生成规范化输出

import xlwings as xw
import json
import logging
import os
import sys
import re
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *
from processor_rely.parse_factory_info import parse_factory_info
from processor_rely.excel_extract_product_img import extract_product_images

from src.utils.clean_factory_name import clean_factory_name
from src.utils.save_result_to_json import make_vendor_folder,save_result_to_vendor_folder

# 配置日志格式
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')



#主营产品清洗
def clean_product_category(product_category: str) -> str:
    """
    清洗主营产品字段，去除重复信息
    
    参数：
        product_category (str): 原始主营产品字段
    
    返回：
        (str): 清洗后的主营产品数据

    处理逻辑：
        1. 按换行符分割字符串
        2. 对每个片段使用多种分隔符进行分割
        3. 去除重复项并合并成最终结果
        
    """
    if not product_category:
        return ""
    
    # 第一步：按换行符分割
    segments = product_category.split('\n')
    
    # 结果列表，用于存储所有产品
    result_list = []
    
    # 第二步：对每个片段使用分隔符进行分割
    for segment in segments:
        segment = segment.strip()
        if not segment:
            continue
        
        # 使用所有分隔符逐一分割
        current_items = [segment]
        for separator in FIELD_SEPARATORS:
            if separator == '\n':  # 跳过换行符，已经处理过了
                continue
            
            new_items = []
            for item in current_items:
                new_items.extend([part.strip() for part in item.split(separator) if part.strip()])
            current_items = new_items
        
        # 将分割后的结果添加到总列表
        result_list.extend(current_items)
    
    # 第三步：使用set()去重并保持顺序
    unique_items = []
    seen = set()
    for item in result_list:
        if item and item not in seen:
            unique_items.append(item)
            seen.add(item)
    
    # 第四步：转换为字符串
    return ','.join(unique_items)


#---------------------- 核心转换处理函数 --------------------------------

# --- Excel转JSON主处理函数 ---
def excel_standard_allftys_map_to_json(input_path: str, output_dir: str, header_row: int) -> bool:
    """
    将Excel文件中的工厂信息转换为标准化JSON格式并按厂商分类存储
    
    核心处理流程：
    1. 初始化Excel应用和工作簿
    2. 遍历所有工作表进行数据提取
    3. 解析表头并建立字段映射关系
    4. 逐行处理数据：
        a. 执行常规字段映射转换
        b. 特殊处理"工厂信息"复合字段
        c. 解析后更新主数据结构
    5. 为每个工厂创建专属文件夹和JSON文件
    6. 提取并保存产品图片资源
    
    参数：
        input_path (str): 输入Excel文件路径
        output_dir (str): 输出文件夹路径
        header_row (int): 表头所在行号(从1开始)
    
    返回：
        bool: 处理成功返回True，失败返回False
    """
    
    # 初始化Excel应用实例
    app = None
    
    try:
        # 步骤1: 启动Excel应用程序
        app = xw.App(visible=False)
        wb = app.books.open(input_path)
        
        # 步骤2: 遍历工作簿中的所有工作表
        for sheet in wb.sheets:
            sheet_name = sheet.name
            
            # 检查工作表数据有效性
            if not sheet.used_range or sheet.used_range.last_cell.row <= header_row:
                logging.warning(f"工作表 '{sheet_name}' 无有效数据，跳过处理")
                continue
                
            # 步骤3: 解析表头信息
            last_col = sheet.used_range.last_cell.column
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            headers = []
            
            # 获取并清理表头数据
            for cell in header_range:
                header_value = cell.value
                # 处理空表头和不同数据类型
                if header_value is None:
                    headers.append("")
                elif isinstance(header_value, str):
                    headers.append(header_value.strip())
                else:
                    headers.append(str(header_value).strip())
            
            # 步骤4: 建立字段映射关系
            # 映射结构: {列索引: 目标字段名}
            column_mapping = {}
            for col_idx, header in enumerate(headers):
                # 跳过空表头
                if not header:
                    continue
                
                # 在配置映射中查找匹配项
                for field, aliases in TEXT_LABELS_excel_all_factory.items():
                    if header in aliases:
                        column_mapping[col_idx] = field
                        break
            
            # 获取产品图片起始列索引
            product_img_start_col_idx = None
            for idx, header in enumerate(headers):
                if header == "产品图片1":
                    product_img_start_col_idx = idx
                    break

            # 成功处理计数器
            success_count = 0
            
            # 步骤5: 处理数据行内容
            data_start_row = header_row + 1
            last_row = sheet.used_range.last_cell.row
            
            for row_idx in range(data_start_row, last_row + 1):
                # 读取当前行所有单元格数据
                row_values = [sheet.cells(row_idx, col + 1).value for col in range(len(headers))]
                # 检查是否为空行，如是则终止处理
                if all(cell is None or str(cell).strip() == "" for cell in row_values):
                    break
                
                # 初始化工厂数据结构(基于JSON模板)
                factory_data = JSON_FORMAT.copy()
                factory_info_raw = None  # 存储原始工厂信息字段
                
                # 步骤6: 遍历当前行的所有列数据
                for col_idx in range(len(headers)):
                    # 获取当前列的表头名称
                    header_name = headers[col_idx]
                    if not header_name:  # 跳过空表头列
                        continue
                    
                    # 获取单元格数值
                    cell_value = sheet.cells(row_idx, col_idx + 1).value
                    
                    # 处理空值情况
                    if cell_value is None:
                        continue
                    
                    # 转换为字符串并清理空白
                    str_value = str(cell_value).strip()
                    
                    # 根据表头名称进行字段分类处理（特殊字段）
                    if header_name == "工厂信息":
                        # 特殊处理复合工厂信息字段
                        factory_info_raw = str_value
                    elif col_idx in column_mapping:
                        # 处理已建立映射的标准字段
                        field_name = column_mapping[col_idx]
                        
                        # 处理多值字段合并
                        current_value = factory_data[field_name]
                        if current_value:
                            factory_data[field_name] += f"\n{str_value}"
                        else:
                            factory_data[field_name] = str_value
                    else:
                        # 记录未映射的字段信息
                        logging.warning(f"未映射的列: {header_name} = {str_value}")
                
                # 步骤7: 解析复合工厂信息字段（特殊字段）
                if factory_info_raw:
                    parsed_info = parse_factory_info(factory_info_raw)
                    factory_data.update(parsed_info)

                # 步骤8: 处理工厂名称并生成文件路径
                factory_name = factory_data["厂商名称"]
            
                if not factory_name:
                    logging.warning(f"第 {row_idx} 行缺少工厂名称，跳过")
                    continue
                factory_name=clean_factory_name(factory_name)
                
                # 步骤9: 处理主营产品字段清洗
                if factory_data.get("主营产品") and "\n" in factory_data["主营产品"]:
                    cleaned_products = clean_product_category(factory_data["主营产品"])
                    factory_data["主营产品"] = cleaned_products
                    logging.debug(f"主营产品清洗: {factory_data['主营产品'][:50]}{'...' if len(factory_data['主营产品']) > 50 else ''}")
                
                # 创建厂商专属文件夹
                vendor_folder = make_vendor_folder(factory_name, output_dir)
                
                # 提取并保存产品图片
                if product_img_start_col_idx is not None:
                    img_folder = extract_product_images(sheet, row_idx, product_img_start_col_idx, vendor_folder, factory_name)
                    factory_data["图片文件夹路径"] = img_folder
                
                # 记录源文件路径
                factory_data["文件路径"] = input_path
                
                # 保存处理结果到厂商文件夹
                save_result_to_vendor_folder(vendor_folder, factory_data)
                
                success_count += 1
        
        logging.info(f"所有工作表处理完成   输出目录: {output_dir}  成功处理 {success_count} 个工厂数据")
        return True
    
    except Exception as e:
        logging.error(f"处理过程中发生错误: {str(e)}", exc_info=True)
        return False
    
    finally:
        # 清理Excel应用资源
        try:
            if app is not None:
                wb.close()
                app.quit()
                logging.info("Excel资源已释放")
        except Exception as e:
            logging.warning(f"资源清理时出错: {str(e)}")


#---------------------- 程序执行入口 --------------------------------

# 主程序入口 - 用于测试和批量处理
if __name__ == "__main__":
    # 配置输入输出路径
    # input_file = r"tests\excel\供应商交流会第五十六期名单（6.25）.xlsx"
    # input_file = r"tests\excel\供应商交流会测试数据\供应商交流会第三十八期（1.10）.xlsx"
    input_file = r"tests\excel\test.xlsx"
    output_dir = r"tests\processed_data\前缀测试_1"

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    logging.info(f"输出目录已创建: {output_dir}")

    # 执行Excel转JSON处理
    result_data_count = excel_standard_allftys_map_to_json(
        input_path=input_file,
        header_row=1,  # 指定表头行
        output_dir=output_dir
    )
    
    