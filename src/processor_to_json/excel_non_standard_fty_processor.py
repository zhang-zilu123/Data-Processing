# 非标准Excel工厂信息处理模块
# 功能：处理工厂情况信息表Excel文件，提取工厂概况数据和产品图片，转换为标准JSON格式
# 特性：支持2种模板格式、sheet图片提取、数据标准化转换

import xlwings as xw
import os
import sys
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from setting.config import *
import logging
import time
from PIL import ImageGrab
from processor_rely.excel_convert_data_json import json_from_factory_data
from src.utils.save_result_to_json import make_vendor_folder,save_result_to_vendor_folder
from src.utils.clean_factory_name import clean_factory_name


#---------------------- 图片和数据提取工具函数 --------------------------------

# --- 产品图片提取处理函数 ---
def process_product_images(excel_path:str, output_dir: str) -> str:
    """
    从Excel文件中提取主要产品图片并保存到指定目录
    
    提取流程：
    1. 检查并创建输出目录
    2. 打开Excel文件并初始化应用程序
    3. 访问"主要产品图片"sheet工作表
    4. 遍历工作表中的所有图形对象
    5. 识别图片类型并复制到剪贴板
    6. 从剪贴板获取图片数据并保存
    7. 按序号命名并统计保存结果
    
    参数：
        excel_path (str): Excel文件完整路径
        output_dir (str): 图片保存的根目录路径
        
    返回：
        str: 图片输出文件夹的完整路径，失败时返回None
        
    保存格式：
        - 文件夹命名：产品图片
        - 文件命名：产品图片_{序号}.png
        - 图片格式：PNG格式

    """
    try:
        # 步骤1：检查并创建输出根目录
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 步骤2：初始化Excel应用程序
        app = xw.App(visible=False)
        wb = app.books.open(excel_path)
        
        # 步骤3：访问"主要产品图片"工作表
        if "主要产品图片"  in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets["主要产品图片"]
        else:
            logging.warning(f"主要产品图片sheet不存在")
            return None
        
        # 步骤4：创建工厂专属图片文件夹
        img_output_folder = os.path.join(output_dir, "产品图片")
        os.makedirs(img_output_folder, exist_ok=True)

        # 步骤5：初始化图片计数器
        img_count = 0
        
        # 步骤6：遍历工作表中的所有图形对象
        for shape in sheet.shapes:
            # 检查是否为图片类型
            if shape.type == 'picture':
                # 复制图片到系统剪贴板
                shape.api.Copy()
                time.sleep(0.2)  # 添加延迟确保剪贴板数据就绪

                # 从剪贴板获取图片数据
                img = ImageGrab.grabclipboard()
                if img is not None:
                    # 生成图片文件路径并保存
                    img_count += 1
                    img_path = os.path.join(img_output_folder, f"产品图片_{img_count}.png")
                    img.save(img_path)
                else:
                    logging.error(f"未能从剪贴板获取图片")
        
        # 步骤7：记录处理结果
        logging.info(f"图片文件夹已保存:{img_output_folder}")
        return img_output_folder

    except Exception as e:
        logging.error(f"处理主要产品图片sheet时出错: {str(e)}") 
    finally:
        # Excel资源清理
        try:
            if 'wb' in locals() and wb is not None:
                wb.close()
        except Exception as e:
            logging.warning(f'关闭wb时出错: {e}')

        try:
            if 'app' in locals() and app is not None:
                app.quit()
        except Exception as e:
            logging.warning(f'关闭app时出错: {e}')




# --- Excel模板类型检测函数 ---
def detect_template(sheet: xw.Sheet) -> dict:
    """
    通过关键单元格内容检测Excel文件使用的模板格式
    
    检测逻辑：
    分析工厂概况工作表中的特定单元格内容，判断Excel文件使用的是哪种预定义模板。
    不同模板对应不同的字段位置和数据结构配置。
    
    检测流程：
    1. 读取模板识别关键单元格（A38）的内容
    2. 与预设的模板特征值进行匹配
    3. 返回对应的模板配置对象
    4. 如果无法识别则返回默认模板
    
    参数：
        sheet: xlwings工作表对象，指向"工厂概况"工作表
        
    返回：
        dict: 模板配置对象，包含字段映射关系，无法识别时返回None
        
    模板识别标准：
        - 模板1：A38单元格内容为"合作的贸易公司及合作情况"
        - 模板2：A38单元格内容为其他值（默认模板）
        
    """
    try:
        # 步骤1：读取模板识别关键单元格
        cell_value = sheet.range('A38').value
        
        # 步骤2：根据单元格内容选择对应模板
        if cell_value == "合作的贸易公司及合作情况":
            # 识别为第一种模板格式
            return EXCEL_FORMATE_FTY_1
        else:
            # 识别为第二种模板格式（默认）
            return EXCEL_FORMATE_FTY_2
    except Exception as e:
        logging.error(f"检测模板类型时出错: {str(e)}")
        return None


# --- 模板数据提取函数 ---
def extract_data_by_template(sheet: xw.Sheet, template: dict) -> dict:
    """
    根据模板配置从Excel工作表中提取工厂数据
    
    提取功能：
    遍历模板配置中定义的所有字段，按照配置的单元格位置读取数据值，
    构建标准化的工厂数据字典。支持字段验证和数据清洗。
    
    提取流程：
    1. 初始化结果数据字典
    2. 遍历模板配置中的所有字段定义
    3. 读取关键词单元格验证字段位置（可选）
    4. 从值单元格提取实际数据
    5. 构建完整的工厂数据结果
    
    参数：
        sheet: xlwings工作表对象，指向包含工厂数据的工作表
        template (dict): 模板配置对象，包含字段到单元格的映射关系
        
    返回：
        dict: 提取的工厂数据字典，处理失败时返回None
        
    模板配置结构：
        - 每个字段包含keyword_cell（关键词位置）和value_cell（数据位置）
        - expected_keyword: 用于验证字段位置的预期关键词
        - 支持跳过非字典类型的配置项
        
    """
    try:
        # 步骤1：初始化结果数据字典
        result = {}
        
        # 步骤2：遍历模板配置中的所有字段
        for field, config in template.items():
            # 步骤3：跳过非字典类型的配置项
            if not isinstance(config, dict):
                continue
            
            # 步骤4：获取关键词单元格（用于判断模板位置是否与实际相符，当前已禁用，后续如果需要调试可打开）
            keyword_cell = sheet.range(config["keyword_cell"])
            
            # # 检查单元格值是否为None
            # if keyword_cell.value:
            #     keyword_cell_value = keyword_cell.value.replace('\n', '').replace(' ', '')
            #
            # print(f"keyword_cell_value: {keyword_cell_value}")
            # if keyword_cell_value != config["expected_keyword"]:
            #     logging.error(
            #         f"字段 {field} 的关键词不匹配",
            #         {
            #             "expected": config["expected_keyword"],
            #             "actual": keyword_cell.value
            #         }
            #     )
            #     continue

            # 步骤5：从值单元格提取实际数据
            value_cell = sheet.range(config["value_cell"]).value
            result[field] = value_cell
        
        # 步骤6：返回完整的工厂数据
        return result
        
    except Exception as e:
        logging.error(f"提取数据时出错: {str(e)}")
        return None



# --- 工厂概况工作表处理函数 ---
def process_factory_overview(wb: xw.Book) -> tuple[dict, dict]:
    """
    处理工厂概况工作表，自动识别模板并提取工厂基础数据
    
    处理功能：
    检查工作簿中是否存在"工厂概况"工作表，自动检测模板类型，
    并根据模板配置提取所有工厂相关的基础数据信息。
    
    处理流程：
    1. 验证"工厂概况"工作表是否存在
    2. 获取工厂概况工作表对象
    3. 自动检测使用的模板类型
    4. 根据模板配置提取工厂数据
    5. 返回数据和模板配置供后续处理
    
    参数：
        wb: xlwings工作簿对象，已打开的Excel文件
        
    返回：
        tuple: (工厂数据字典, 模板配置对象)
               任一处理失败时返回 (None, None)             
        
    """
    try:
        # 步骤1：检查"工厂概况"工作表是否存在
        if "工厂概况" not in [sheet.name for sheet in wb.sheets]:
            logging.error(f"工厂概况sheet不存在")
            return None,None

        # 步骤2：获取工厂概况工作表对象
        sheet = wb.sheets["工厂概况"]
        
        # 步骤3：自动检测模板类型
        template = detect_template(sheet)
        if not template:
            return None,None
            
        # 步骤4：根据模板配置提取工厂数据
        dic_data=extract_data_by_template(sheet,template)
        
        # 步骤5：返回数据和模板配置
        return dic_data,template
        
    except Exception as e:
        logging.error(f"处理工厂概况sheet时出错: {str(e)}")
        return None,None


# --- Excel转换成json(dict)函数 ---
def process_excel(excel_path:str) -> dict:
    """
    Excel工厂信息文件的完整处理主函数
    
    处理功能：
    统一处理包含工厂概况和产品图片的Excel文件，提取所有相关数据，
    转换为标准化的JSON格式，并保存产品图片到指定目录。
    
    处理流程：
    1. 打开Excel文件并初始化应用程序
    2. 处理工厂概况工作表获取基础数据
    3. 验证工厂名称等关键字段
    4. 处理主要产品图片工作表（如果存在）
    5. 将工厂数据转换为标准JSON格式
    6. 清理资源并关闭Excel应用程序
    
    参数：
        excel_path (str): 待处理的Excel文件完整路径
        
    返回：
        dict: 标准化的工厂数据JSON对象，处理失败时返回None
        
    """
    try:
        # 步骤1：初始化Excel应用程序（无界面模式）
        app = xw.App(visible=False)
        wb = app.books.open(excel_path)
        img_output_folder_path = None
        
        # 步骤2：处理工厂概况工作表，获取基础数据和模板
        factory_data ,template= process_factory_overview(wb)
        if not factory_data:
            logging.error("未能正确获取工厂数据")
            return None
        
        # 步骤3：验证关键字段 - 工厂名称
        if "factory_name" not in factory_data:
            logging.error("缺少factory_name字段")
            return None

        # 步骤4：将工厂数据转换为标准JSON格式
        data_json=json_from_factory_data(factory_data,template,excel_path,img_output_folder_path)
        
        return data_json
        
    except Exception as e:
        logging.error(f"处理Excel文件 {excel_path} 时出错: {str(e)}")
    
    finally:
        # Excel资源清理
        try:
            if 'wb' in locals() and wb is not None:
                wb.close()
        except Exception as e:
            logging.warning(f'关闭wb时出错: {e}')

        try:
            if 'app' in locals() and app is not None:
                app.quit()
        except Exception as e:
            logging.warning(f'关闭app时出错: {e}')


#---------------------- Excel文件处理主模块 --------------------------------

# --- 非标准Excel文件保存JSON主函数 ---
def non_standard_excel_save_json(file_path:str, output_dir:str) -> bool:
    """
    非标准Excel工厂信息表转JSON完整处理流程
    
    处理功能：
    处理工厂情况信息表Excel文件，提取工厂数据和产品图片，
    转换为标准JSON格式并保存到指定目录。
    
    处理流程：
    1. 调用Excel处理函数提取工厂数据
    2. 清洗工厂名称并创建输出文件夹
    3. 提取并保存产品图片
    4. 保存JSON结果文件
    
    参数：
        file_path (str): Excel文件路径
        output_dir (str): 输出目录路径
        
    返回：
        bool: 处理成功返回True，失败返回False
    """
        
    # 步骤1：处理Excel文件提取数据
    json_result = process_excel(file_path)

    if json_result:
        # 步骤2：清洗工厂名称并创建文件夹
        factory_name = clean_factory_name(json_result.get('厂商名称'))
        vendor_folder = make_vendor_folder(factory_name,output_dir)
        
        # 步骤3：提取并保存产品图片
        img_product_dir = process_product_images(file_path, vendor_folder)
        if img_product_dir:
            json_result['图片文件夹路径'] = img_product_dir
        else:
            logging.warning(f"不存在产品图片文件夹")
        
        # 步骤4：保存JSON结果文件
        outpath=save_result_to_vendor_folder(vendor_folder, json_result)
        if outpath:
            logging.info(f"Excel文档已转换为JSON格式")
            return True
        else:
            logging.error("Excel文档转换为JSON格式失败")
            return False
    
    return False



#---------------------- 程序测试入口 --------------------------------

if __name__ == "__main__":
    # 测试文件路径配置
    excel_path=r"tests\excel\曹县春军工艺品-工厂情况信息表.xlsx"
    output_dir=r"tests\processed_data\fty_excel"
    
    # 执行Excel文件处理测试
    non_standard_excel_save_json(excel_path,output_dir)

    
