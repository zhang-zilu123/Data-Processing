# JSON转Excel图片标签处理模块，将工厂信息JSON数据转换为包含图片和标签的Excel表格

import os
import json
import re
import xlwings as xw
import logging
from PIL import Image
import sys
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *  # 导入配置模块
from src.utils.json_logger import setup_json_logger
from src.utils.extract_tags import extract_tags

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#---------------------- 辅助功能函数 --------------------------------


# --- 函数：获取文件夹中的排序图片列表 ---
def get_sorted_images(folder_path, limit=5):
    """
    获取指定文件夹中的图片文件列表，按文件名排序
    
    
    处理流程：
    1. 验证文件夹路径的有效性
    2. 扫描文件夹中的所有文件
    3. 过滤有效的图片格式文件
    4. 转换为绝对路径
    5. 按文件名排序并限制数量
    6. 返回排序后的图片路径列表
    
    参数：
        folder_path (str): 图片文件夹的路径
        limit (int): 最多返回的图片数量，默认为5
        
    返回：
        list: 图片路径列表（绝对路径），失败时返回空列表
        

    """
    try:
        # 步骤1：定义支持的图片格式
        valid_exts = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
        
        # 步骤2：验证文件夹路径的有效性
        if not os.path.exists(folder_path):
            logging.error(f"图片文件夹不存在: {folder_path}")
            return []
        if not os.path.isdir(folder_path):
            logging.error(f"图片路径不是文件夹: {folder_path}")
            return []
        
        # 步骤3：扫描文件夹中的所有文件
        images = []
        for fname in os.listdir(folder_path):
            fpath = os.path.join(folder_path, fname)
            if os.path.isfile(fpath):
                # 步骤4：过滤有效的图片格式文件
                ext = os.path.splitext(fname)[1].lower()
                if ext in valid_exts:
                    # 步骤5：转换为绝对路径
                    abs_path = os.path.abspath(fpath)
                    images.append(abs_path)
        
        # 步骤6：按文件名排序并限制数量
        sorted_images = sorted(images)[:limit]
        logging.info(f"在文件夹 {folder_path} 中找到 {len(sorted_images)} 张图片")
        return sorted_images
    
    except Exception as e:
        logging.error(f"获取图片列表时出错: {e} | 文件夹路径: {folder_path}")
        return []

#---------------------- 主要转换功能 --------------------------------

# --- 函数3：JSON转Excel主函数 ---
def json_to_excel(json_path, output_excel_path):
    """
    将工厂信息JSON文件转换为包含图片和标签的Excel表格
    
    功能说明：
    读取工厂信息的JSON数据，创建Excel工作簿，批量写入数据，
    插入微信二维码图片、产品图片和分类标签，生成完整的工厂信息表格。
    
    处理流程：
    1. 读取和验证JSON数据文件
    2. 创建Excel工作簿和工作表
    3. 写入表头和设置列宽
    4. 批量写入基础数据（每100行一批）
    5. 处理特殊字段（图片和标签）
    6. 设置行高和图片布局
    7. 保存Excel文件并清理资源
    
    参数：
        json_path (str): 输入的JSON文件路径
        output_excel_path (str): 输出的Excel文件路径
        
    返回：
        None（通过日志记录处理结果）
        
    """
    logging.info(f"开始处理JSON文件: {json_path}")
    logging.info(f"输出Excel路径: {output_excel_path}")

    # 步骤1：创建输出目录
    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)  
    
    # 步骤2：读取JSON数据
    try:
        logging.info("正在读取JSON文件...")
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        logging.info(f"成功读取JSON文件，包含 {len(data)} 条记录")
    except Exception as e:
        logging.error(f"读取JSON文件失败: {e}")
        return
    
    # 步骤3：创建Excel工作簿
    app = None
    wb = None
    try:
        logging.info("正在创建Excel工作簿...")
        app = xw.App(visible=False)  # 不显示Excel界面
        wb = app.books.add()
        sheet = wb.sheets[0]
        logging.info("Excel工作簿创建成功")
    except Exception as e:
        logging.error(f"创建Excel工作簿失败: {e}")
        return
    
    # 步骤4：写入表头
    try:
        sheet.range('A1').value = [EXCEL_HEADERS]
        logging.info(f"已写入表头: {EXCEL_HEADERS}")
    except Exception as e:
        logging.error(f"写入表头失败: {e}")
    
    # 步骤5：设置列宽（特别是图片列）
    try:
        for i, header in enumerate(EXCEL_HEADERS):
            if header.startswith('图片') or header == '微信':
                sheet.range((1, i+1)).column_width = 20  # 设置图片列宽
        logging.info("已设置图片列宽度")
    except Exception as e:
        logging.error(f"设置列宽失败: {e}")
    
    # 步骤6：批量准备数据
    logging.info(f"开始处理 {len(data)} 条记录...")
    success_count = 0
    error_count = 0
    
    # 批量写入数据（每100行一批）
    batch_size = 100
    for batch_start in range(0, len(data), batch_size):
        batch_end = min(batch_start + batch_size, len(data))
        batch_data = data[batch_start:batch_end]
        
        logging.info(f"正在处理第 {batch_start+1}-{batch_end} 条记录...")
        
        # 步骤7：准备批量数据
        batch_values = []
        for item in batch_data:
            row_data = []
            for col_name in EXCEL_HEADERS:
                if col_name in ['微信', '图片1', '图片2', '图片3', '图片4', '图片5', '厂商信息拼接', '图片描述', '拼接结果']:
                    row_data.append('')  # 占位符，稍后处理
                elif col_name == '数据来源':
                    row_data.append('集团内部：产品开发部')
                else:
                    value = item.get(col_name, '')
                    # 将列表类型数据转换为逗号连接的字符串（除了合并来源字段）
                    if isinstance(value, list) and col_name != '合并来源' and col_name != '标签':
                        value = ','.join(str(v) for v in value)
                    elif isinstance(value, list) and col_name == '标签':
                        value = '\n'.join(str(v) for v in value)
                    
                    if value in ('', '无', None):
                        value = '暂无记录'
                    row_data.append(value)
            batch_values.append(row_data)
        
        # 步骤8：批量写入数据
        try:
            start_row = batch_start + 2  # 从第2行开始
            end_row = batch_start + len(batch_data) + 1
            range_address = f'A{start_row}:{chr(ord("A") + len(EXCEL_HEADERS) - 1)}{end_row}'
            sheet.range(range_address).value = batch_values
            logging.info(f"批量写入成功: {range_address}")
        except Exception as e:
            logging.error(f"批量写入失败: {e}")
            # 如果批量写入失败，尝试逐行写入
            for i, item in enumerate(batch_data):
                row_idx = batch_start + i + 2
                try:
                    for col_idx, col_name in enumerate(EXCEL_HEADERS, start=1):
                        if col_name not in ['微信', '图片1', '图片2', '图片3', '图片4', '图片5', '厂商信息拼接', '图片描述', '拼接结果']:
                            if col_name == '数据来源':
                                value = '集团内部：产品开发部'
                            else:
                                value = item.get(col_name, '')
                                # 将列表类型数据转换为逗号连接的字符串（除了合并来源字段）
                                if isinstance(value, list) and col_name != '合并来源' and col_name != '标签':
                                    value = ','.join(str(v) for v in value)
                                elif isinstance(value, list) and col_name == '标签':
                                    value = '\n'.join(str(v) for v in value)
                                
                                if value in ('', '无', None):
                                    value = '暂无记录'
                            sheet.range(f'{chr(ord("A") + col_idx - 1)}{row_idx}').value = value
                    success_count += 1
                except Exception as row_e:
                    error_count += 1
                    logging.error(f"第 {row_idx} 行写入失败: {row_e}")
            continue
        
        # 步骤9：设置行高
        try:
            for i in range(len(batch_data)):
                row_idx = batch_start + i + 2
                sheet.range(f'A{row_idx}').row_height = 100
        except Exception as e:
            logging.error(f"设置行高失败: {e}")
        
        # 步骤10：处理特殊字段（图片和标签）
        for i, item in enumerate(batch_data):
            row_idx = batch_start + i + 2
            item_log_prefix = f"第 {row_idx-1} 行数据 [厂商: {item.get('厂商名称', '未知')}] - "
            
            try:
                # 步骤11：处理微信字段（插入图片）
                wechat_path = item.get('微信', '')
                
                if wechat_path and wechat_path != "暂无记录":
                    # 标准化路径分隔符
                    wechat_path = wechat_path.replace('/', '\\')
                
                    
                    # 转换为绝对路径
                    if not os.path.isabs(wechat_path):
                        wechat_path = os.path.abspath(wechat_path)
                    


                        if os.path.exists(wechat_path):
                            try:
                                col_name = "微信"
                                col_idx = EXCEL_HEADERS.index(col_name) + 1
                                cell = sheet.range(f'{chr(ord("A") + col_idx - 1)}{row_idx}')
                                
                                # 插入图片
                                pic = sheet.pictures.add(wechat_path, 
                                                        left=cell.left,
                                                        top=cell.top)
                                
                                # ---动态适应单元格的缩放逻辑---

                                # 1. 获取图片的原始尺寸（单位：点）
                                original_width = pic.width
                                original_height = pic.height

                                # 2. 获取单元格的实际动态尺寸（单位：点）
                                #    这是关键，我们使用 cell.height 而不是固定值
                                cell_height = cell.height
                                cell_width = cell.width

                                # 3. 计算能适应单元格的正确缩放比例
                                scale_ratio = 0.7  # 默认缩放比例
                                if original_width > 250 and original_height > 250:
                                    width_ratio = cell_width / original_width
                                    height_ratio = cell_height / original_height
                                    scale_ratio = min(width_ratio, height_ratio)

                                # 4. 设置最小缩放比例为 0.13
                                scale_ratio = max(scale_ratio, 0.13)

                                # 5. 应用等比例缩放
                                pic.width = original_width * scale_ratio
                                pic.height = original_height * scale_ratio

                                # 6. 图片在单元格内左上角对齐
                                pic.left = cell.left
                                pic.top = cell.top


                                try:
                                    # 1. 获取缩放后图片的最终尺寸
                                    final_height = pic.height
                                    final_width = pic.width

                                    # 2. 调整行高以适应图片高度
                                    #    - 行高单位是“点”，与图片高度单位一致，直接使用
                                    #    - 限制最大行高为409.5，这是Excel的上限
                                    required_row_height = min(final_height, 409.5)
                                    #    - 只在需要时增大行高，避免缩小
                                    if required_row_height > cell.row_height:
                                        cell.row_height = required_row_height

                                    # 3. 调整列宽以适应图片宽度
                                    #    - 列宽单位是“字符数”，需要从“点”转换（约7点/字符）
                                    required_col_width = final_width / 7
                                    #    - 只在需要时增大列宽
                                    if required_col_width > cell.column_width:
                                        cell.column_width = required_col_width
                                    
                                    logging.info(f"成功调整单元格尺寸以适应图片：行高={cell.row_height:.1f}点, 列宽={cell.column_width:.1f}字符")

                                except Exception as e:
                                    logging.warning(f"动态调整单元格尺寸失败: {e}")
                                
                                
                                logging.info(f"{item_log_prefix}微信图片插入成功")
                            except Exception as e:
                                logging.error(f"{item_log_prefix}插入微信图片失败: {e} | 路径: {wechat_path}")
                        else:
                            logging.error(f"{item_log_prefix}微信路径不存在: {wechat_path}")
                else:
                    logging.info(f"微信路径为空或为'暂无记录'，跳过处理")
                
                 # 步骤12：处理产品图片字段
                img_folder = item.get('图片文件夹路径', '')
                if img_folder and os.path.exists(img_folder):
                    images = get_sorted_images(img_folder, 5)
                    for j, img_path in enumerate(images):
                        if j >= 5:
                            break
                        try:
                            col_name = f'图片{j+1}'
                            col_idx = EXCEL_HEADERS.index(col_name) + 1
                            cell = sheet.range(f'{chr(ord("A") + col_idx - 1)}{row_idx}')
                            pic = sheet.pictures.add(img_path, 
                                                   left=cell.left,
                                                   top=cell.top,
                                                   width=cell.width,
                                                   height=cell.height)
                            pic.api.Placement = 1
                            logging.info(f"{item_log_prefix}图片{j+1}插入成功")
                        except Exception as e:
                            logging.error(f"{item_log_prefix}插入图片{j+1}失败: {e}")
                
                success_count += 1
                
            except Exception as e:
                error_count += 1
                logging.error(f"{item_log_prefix}处理特殊字段失败: {e}")
        
        # 步骤14：每批处理后保存一次，避免数据丢失
        try:
            wb.save(output_excel_path)
            logging.info(f"第 {batch_start+1}-{batch_end} 批数据已保存")
        except Exception as e:
            logging.error(f"保存失败: {e}")
    
    # 步骤15：最终保存
    try:
        logging.info(f"正在保存Excel文件: {output_excel_path}")
        wb.save(output_excel_path)
        logging.info(f"Excel文件已成功保存至: {output_excel_path}")
        logging.info(f"处理完成: 成功 {success_count} 条, 失败 {error_count} 条, 总计 {len(data)} 条")
    except Exception as e:
        logging.error(f"保存Excel失败: {e}")
    
    finally:
        # 步骤16：清理资源
        try:
            if wb:
                wb.close()
            if app:
                app.quit()
            logging.info("已关闭Excel应用")
        except Exception as e:
            logging.error(f"关闭Excel应用时出错: {e}")

#---------------------- 主程序入口 --------------------------------

if __name__ == "__main__":
    # 配置处理参数
    input_json = r"data\processed_data\merged_factories\merged_factories_tag.json"  # 输入JSON文件路径
    output_excel = r"data\output_data\excel_img_tag.xlsx"  # 输出Excel文件路径
    log_file = "json_to_excel_img_tag.log"
    
    # 设置日志系统
    logger = setup_json_logger(log_dir='logs', log_file=log_file)
    
    # 执行JSON转Excel处理
    json_to_excel(input_json, output_excel)
    
    logging.info("程序执行完毕")