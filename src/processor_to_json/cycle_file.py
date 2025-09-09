# 文档批量处理循环处理模块
# 功能：批量处理多种格式文档(Word/Excel/PPTX)，转换为标准化JSON格式
# 支持：标准Excel、非标准Excel、Word文档、PPTX演示文稿的自动化处理

import os
import logging
import sys

# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from src.utils.json_logger import setup_json_logger
from src.utils.folder_img_save import process_image_folders
from src.utils.clean_factory_name import clean_factory_name
from src.utils.save_result_to_json import make_vendor_folder,save_result_to_vendor_folder
from src.processor_to_json.excel_standard_allftys_map_processor import excel_standard_allftys_map_to_json
from src.processor_to_json.word_api_identify_write_processor import word_to_json
from src.processor_to_json.excel_non_standard_fty_processor import non_standard_excel_save_json,process_excel
from src.processor_to_json.pdf_allftys_imgs_processor import process_pdf
from src.processor_to_json.pptx_processor import process_pptx_file
from src.processor_to_json.pdf_standard_wqimg_processor import process_pdf_file


# --- Word文档批量处理函数 ---
def module_word(input_directory:str, output_directory:str) -> None:
    """
    批量处理目录中的Word格式文档
    
    处理功能：
    遍历输入目录中的所有DOC和DOCX文件，提取文本内容和微信二维码图片，
    转换为标准JSON格式并保存到输出目录。
    
    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
        
    返回：
        None
        
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    

    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            try:
                # 筛选Word文件格式
                if filename.endswith(('.doc', '.docx')):
                    total_count += 1
                    file_path = os.path.join(root, filename)

                # 执行Word转JSON处理
                result_bool=word_to_json(file_path,output_directory)
                
                # 统计处理结果
                if result_bool:
                    success_count += 1           
                else:
                    failure_count += 1
            except Exception as e:
                logging.error(f"处理文档异常: {filename}, 错误: {str(e)}")
                failure_count += 1
    
    # 输出处理统计结果
    logging.info(f"Word文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")
    

# --- PPTX演示文稿批量处理函数 ---
def module_ppt(input_directory:str, output_directory:str) -> None:
    """
    批量处理目录中的PPTX格式工厂信息文档
    
    处理功能：
    遍历输入目录中的所有PPTX文件，提取每个工厂的基础数据和微信二维码图片，
    支持批量处理多个工厂信息。
    
    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
        
    返回：
        None
       
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    
    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            try:
                # 筛选PPTX文件格式
                if filename.lower().endswith('.pptx'):
                    file_path = os.path.join(root, filename)
                    total_count += 1
                    
                    # 执行PPTX处理
                    file_success = process_pptx_file(file_path,output_directory)
                        
                    # 统计处理结果
                    if file_success:
                        success_count += 1
                    else:
                        failure_count += 1
            except Exception as e:
                logging.error(f"处理文档异常: {filename}, 错误: {str(e)}")
                failure_count += 1
    
    # 输出处理统计结果
    logging.info(f"PPT文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")


#---------------------- Excel文档处理模块 --------------------------------

# --- 标准Excel表格批量处理函数 ---
def module_standard_excel(input_directory:str, output_directory:str,header_row:int) -> None:
    """
    批量处理标准Excel格式工厂信息表(供应商交流会格式)

    处理功能：
    处理具有标准表头结构的Excel文件，自动识别字段映射关系，
    支持多工作表处理和产品图片提取。

    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
        header_row (int): 表头所在的行号
    返回：
        None
    
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    
    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_directory):
       for filename in files:
            try:
                # 筛选Excel文件格式
                if filename.endswith(('.xls', '.xlsx')):
                    total_count += 1
                    file_path = os.path.join(root, filename)

                # 执行标准Excel转JSON处理
                result_bool = excel_standard_allftys_map_to_json(file_path, output_directory, header_row)
                    
                # 统计处理结果
                if result_bool:
                    success_count += 1
                else:
                    failure_count += 1

            except Exception as e:
                logging.error(f"处理文档异常: {filename}, 错误: {str(e)}")
                failure_count += 1
        
    # 输出处理统计结果
    logging.info(f"Excel文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")



# --- 非标准Excel表格批量处理函数 ---
def module_non_standard_excel(input_directory: str, output_directory: str) -> None:
    """
    批量处理非标准Excel格式工厂信息表(工厂信息表格式)

    处理功能：
    处理包含"工厂情况信息表"或"工厂信息情况表"的Excel文件，
    自动识别工厂文件夹中的产品图片，支持单个或多个文件处理。

    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
       
    返回：
        None
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    
    # 遍历输入目录中的所有文件夹
    for root, dirs, files in os.walk(input_directory):
        # 在当前目录中查找符合条件的Excel文件
        target_files = []
        for filename in files:
            if filename.endswith(('.xls', '.xlsx')):
                if '工厂情况信息表' in filename or '工厂信息情况表' in filename:
                    file_path = os.path.join(root, filename)
                    target_files.append(file_path)
                else:
                    logging.warning(f"跳过非工厂情况信息表的Excel文件: {filename}")
        
        # 根据找到的目标文件数量进行分类处理
        if len(target_files) == 0:
            # 没有找到符合条件的Excel文件
            continue
        
        total_count += 1
        
        if len(target_files) > 1:
            # 处理多个符合条件的Excel文件
            try:
                for file_path in target_files:
                    succss_bool = non_standard_excel_save_json(file_path, output_directory)
                    if succss_bool:
                        success_count += 1
                    else:
                        failure_count += 1
                    
            except Exception as e:
                failure_count += 1
                logging.error(f"处理多个工厂信息表Excel文件失败: {e}")
        else:
            # 处理单个Excel文件（含产品图片文件夹）
            try:
                file_path = target_files[0]
                
                # Excel文件数据提取
                json_result = process_excel(file_path)
                factory_name=json_result.get('厂商名称')
                factory_name=clean_factory_name(factory_name)

                if json_result:
                    # 创建厂商文件夹
                    vendor_folder = make_vendor_folder(factory_name,output_directory)
                    
                    # 处理产品图片文件夹
                    img_folder_path = process_image_folders(
                        root,  # 工厂文件夹路径
                        vendor_folder,  # 输出路径
                        factory_name  # 工厂名称
                    )
                    
                    # 添加图片路径信息
                    if img_folder_path:
                        json_result['图片文件夹路径'] = img_folder_path
                    else:
                        logging.warning(f"不存在产品图片文件夹")
                    
                    # 保存JSON结果文件
                    outpath=save_result_to_vendor_folder(vendor_folder, json_result)
                    if outpath:
                        logging.info(f"Excel文档已转换为JSON格式")
                        success_count += 1
                    else:
                        logging.error("Excel文档转换为JSON格式失败")
                        failure_count += 1
                        

            except Exception as e:
                failure_count += 1
                logging.error(f"处理文件失败: {file_path}, 错误: {e}")

    # 输出处理统计结果
    logging.info(f"Excel文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")



#---------------------- PDF文档处理模块 --------------------------------


# --- 标准模板含有微信二维码PDF文档批量处理函数 ---
def module_standard_qwimg_pdf(input_directory: str, output_directory: str) -> None:
    """
    批量处理目录中的PDF格式工厂信息文档
    
    处理功能：
    遍历输入目录中的所有PDF文件，提取工厂信息和二维码图片，
    转换为标准JSON格式并保存到输出目录。
    
    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
        
    返回：
        None
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    
    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            try:
                # 筛选PDF文件格式
                if filename.lower().endswith('.pdf'):
                    file_path = os.path.join(root, filename)
                    total_count += 1
                    
                    # 执行PDF处理
                    file_success = process_pdf_file(file_path, output_directory)
                        
                    # 统计处理结果
                    if file_success:
                        success_count += 1
                    else:
                        failure_count += 1
            except Exception as e:
                logging.error(f"处理文档异常: {filename}, 错误: {str(e)}")
                failure_count += 1
    
    # 输出处理统计结果
    logging.info(f"PDF文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")


# --- 多家工厂信息PDF文档批量处理函数 ---
def module_allftys_imgs_pdf(input_directory: str, output_directory: str) -> None:
    """
    批量处理目录中的PDF格式工厂信息文档
    
    处理功能：
    遍历输入目录中的所有PDF文件，提取工厂信息和产品图片，
    转换为标准JSON格式并保存到输出目录。
    
    参数：
        input_directory (str): 输入文件目录路径
        output_directory (str): 输出结果目录路径
        
    返回：
        None
    """
    # 处理统计计数器
    total_count = 0
    success_count = 0
    failure_count = 0
    
    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            try:
                # 筛选PDF文件格式
                if filename.lower().endswith('.pdf'):
                    file_path = os.path.join(root, filename)
                    total_count += 1
                    
                    # 执行PDF处理
                    file_success = process_pdf(file_path, output_directory)
                        
                    # 统计处理结果
                    if file_success:
                        success_count += 1
                    else:
                        failure_count += 1
            except Exception as e:
                logging.error(f"处理文档异常: {filename}, 错误: {str(e)}")
                failure_count += 1
    
    # 输出处理统计结果
    logging.info(f"PDF文档处理完成: 总数：{total_count}个, 成功：{success_count}个, 失败：{failure_count}个")



#---------------------- 程序执行入口 --------------------------------

if __name__ == "__main__":
    # 日志系统配置
    log_file = 'pdf_processor.log'
    logger = setup_json_logger(log_dir='logs', log_file=log_file)
    

    # # Word文档批量处理测试
    # input_directory = r'tests\word'
    # output_directory = r'tests\processed_data\word'
    # module_word(input_directory, output_directory)

    # 标准Excel文档批量处理
    # input_directory = r'data\input_data\excel\供应商交流会'
    # output_directory = r'data\processed_data\excel\供应商交流会_日期'
    # header_row = 2
    # module_standard_excel(input_directory, output_directory, header_row)



    # # 非标准Excel文档批量处理测试
    # input_directory = r'tests\excel\非标准excel测试'
    # output_directory = r'tests\processed_data\non_standard_excel'
    # module_non_standard_excel(input_directory, output_directory)


    # PPTX文档批量处理测试
    input_directory = r'data\input_data\ppt'
    output_directory = r'data\processed_data\ppt'
    module_ppt(input_directory, output_directory)


    # 标准模板含有微信二维码PDF文档批量处理测试
    # input_directory = r'data\input_data\pdf'
    # output_directory = r'data\processed_data\pdf'
    # module_standard_qwimg_pdf(input_directory, output_directory)
    
    # 多家工厂信息PDF文档批量处理测试
    # input_directory = r'tests\pdf\东南亚工厂'
    # output_directory = r'tests\processed_data\pdf\东南亚工厂'
    # module_allftys_imgs_pdf(input_directory, output_directory)
    