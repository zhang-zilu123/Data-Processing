#拼接指定信息到"厂商信息拼接"列

import xlwings as xw
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def process_excel(input_file_path, output_file_path, header_row=2, exclude_columns=None, column_name='厂商信息拼接'):
    """
    处理Excel文件，将每个sheet中的数据拼接成新的字符串，并保存到新的Excel文件中。

    参数:
    input_file_path (str): 输入Excel文件的路径。
    output_file_path (str): 输出Excel文件的路径。
    header_row (int): 标题行的索引，从1开始计数，默认为2。
    exclude_columns (list): 需要排除在拼接之外的列名列表，默认为空列表。
    column_name (str): 新列的名称，用于存放拼接结果，默认为'厂商信息拼接'。
    """
    if exclude_columns is None:
        exclude_columns = []

    app = None
    wb = None
    try:
        # 启动Excel应用程序
        app = xw.App(visible=False, add_book=False)
        # 打开工作簿
        wb = app.books.open(input_file_path)
        # 遍历工作簿中的每个工作表
        for sheet in wb.sheets:
            logging.info(f"开始处理sheet: {sheet.name}")
            # 读取数据
            df = sheet.range(f'A{header_row}').expand().value
            headers = df[0]
            data = df[1:]
            # 找到需要拼接的列索引
            concat_columns = [i for i, col in enumerate(headers) if col not in exclude_columns]
            # 查找目标列索引，如果不存在则添加
            if column_name in headers:
                target_column_index = headers.index(column_name)
            else:
                headers.append(column_name)
                target_column_index = len(headers) - 1
                # 为每行数据添加空白列以匹配新的列数
                for row in data:
                    while len(row) < len(headers):
                        row.append('')
            
            # 拼接数据并直接写入目标列
            for row in data:
                parts = []
                company_name = ''
                for i in concat_columns:
                    value = row[i] if row[i] not in (None, '', 'nan') else '暂无记录'
                    if headers[i] == '厂商名称':
                        company_name = value
                    else:
                        parts.append(f"{headers[i]}为{value}")
                concat_str = f"名称为{company_name}的厂商，其信息如下：{'，'.join(parts)}"
                row[target_column_index] = concat_str
            # 写入数据
            sheet.range(f'A{header_row}').value = [headers] + data
            logging.info(f"完成处理sheet: {sheet.name}")
    except Exception as e:
        logging.error(f"处理Excel文件时发生错误: {e}")
    finally:
        # 确保工作簿和Excel应用程序被正确关闭
        if wb:
            wb.save(output_file_path)
            wb.close()
        if app:
            app.quit()






if __name__ == '__main__':
    input_file = r'data\output_data\excel_img_tag.xlsx'  # 输入文件路径
    output_file = r'data\output_data\merged_manufacturers_final_concat.xlsx'  # 输出文件路径
    header_row_number = 1  # 假设表头在第一行
    exclude_columns = ['图片1', '图片2', '图片3', '图片4', '图片5', '标签','厂商信息拼接','图片描述','拼接结果']  # 需要排除的列名
    column_name = "厂商信息拼接"

    process_excel(input_file, output_file, header_row_number, exclude_columns,column_name)
