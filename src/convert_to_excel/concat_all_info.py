# 将所有信息拼接到"拼接结果"列

import xlwings as xw
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def process_workbook(input_path, output_path, header_row, columns_to_concatenate, new_column_name):
    try:
        # 启动Excel应用程序
        app = xw.App(visible=False, add_book=False)
        # 打开工作簿
        wb = app.books.open(input_path)
        
        # 遍历工作簿中的每个工作表
        for sheet in wb.sheets:
            logging.info(f"开始处理工作表: {sheet.name}")
            
            # 读取数据区域
            df = sheet.range(f'A{header_row}').expand().value
            headers = df[0]  # 表头
            data = df[1:]  # 数据行

            # 清理表头，去除空格
            headers = [header.strip() if isinstance(header, str) else header for header in headers]
            logging.info(f"表头: {headers}")

            # 找到需要拼接的列索引
            concat_columns = [i for i, col in enumerate(headers) if col in columns_to_concatenate]
            exclude_columns = [i for i, col in enumerate(headers) if col not in columns_to_concatenate]
            
            # 拼接数据
            concatenated_data = []
            for row in data:
                concat_values = []
                for col_idx in concat_columns:
                    value = row[col_idx]
                    if value is None:
                        concat_values.append('')
                    else:
                        concat_values.append(str(value))

                # 按照指定格式拼接数据
                concatenated_string = '。'.join(concat_values)
                concatenated_data.append([concatenated_string])

            # 查找目标列索引，如果不存在则添加
            if new_column_name in headers:
                target_col_index = headers.index(new_column_name) + 1  # xlwings使用1基索引
            else:
                headers.append(new_column_name)
                target_col_index = len(headers)  # 新列位置
            
            # 写入新列标题（如果是新列）和数据
            sheet.range(header_row, target_col_index).value = new_column_name
            sheet.range(header_row + 1, target_col_index).value = concatenated_data

            logging.info(f"工作表 {sheet.name} 处理完成，已添加新列: {new_column_name}")

        # 保存并关闭工作簿
        wb.save(output_path)
        logging.info(f"文件已保存: {output_path}")

    except Exception as e:
        logging.error(f"处理过程中发生错误: {e}")
    finally:
        # 关闭工作簿
        if 'wb' in locals():
            wb.close()
        # 退出应用
        if 'app' in locals():
            app.quit()





# 调用示例
if __name__ == "__main__":
    input_path = r'data\output_data\merged_manufacturers_final_concat.xlsx'  # 输入文件路径
    output_path = r'data\output_data\final_allcow_output.xlsx'  # 输出文件路径
    header_row = 1  # 表头所在行
    header_list = ['厂商信息拼接', '图片描述']  # 需要拼接的表头
    new_column_name = '拼接结果'  # 新列的名字

    # 调用函数
    process_workbook(input_path, output_path, header_row, header_list, new_column_name)
