# 提取Excel中的"产品图片"列的图片，并保存到指定文件夹

import os
import xlwings as xw

def extract_product_images(sheet:xw.Sheet, row_idx:int, start_col_idx:int, output_dir:str, factory_name:str) -> str:
    """
    提取指定行的“产品图片1”及其后4列（共5列）单元格的所有图片，
    保存到 output_dir/img_product/{factory_name}/ 下。
    
    参数：
        sheet(xw.Sheet): 工作表对象
        row_idx(int): 目标行索引
        start_col_idx(int): 产品图片列的起始列索引
        output_dir(str): 输出文件夹路径
        factory_name(str): 工厂名称

    返回：
        str: 图片文件夹路径（无图片时返回空字符串）
        
    处理逻辑：
        1. 创建图片文件夹
        2. 遍历目标单元格
        3. 判断图片左上角是否在该单元格内
        4. 保存图片
        
    """

    img_folder = os.path.join(output_dir, "img_product", factory_name)
    os.makedirs(img_folder, exist_ok=True)
    img_count = 0

    # 目标单元格对象列表（xlwings的行列都是1基准）
    target_cells = [sheet.cells(row_idx, start_col_idx + 1 + i) for i in range(5)]

    for cell in target_cells:
        cell_left, cell_top = cell.left, cell.top
        cell_right = cell_left + cell.width
        cell_bottom = cell_top + cell.height

        for picture in sheet.pictures:
            pic_left, pic_top = picture.left, picture.top
            # 判断图片左上角是否在该单元格内
            if (cell_left <= pic_left < cell_right) and (cell_top <= pic_top < cell_bottom):
                img_path = os.path.join(img_folder, f"{factory_name}_img{img_count+1}.png")
                picture.api.Copy()
                try:
                    from PIL import ImageGrab
                    img = ImageGrab.grabclipboard()
                    if img:
                        img.save(img_path)
                        img_count += 1
                except Exception as e:
                    print(f"图片保存失败: {e}")

    return img_folder if img_count > 0 else ""
