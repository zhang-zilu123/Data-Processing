# 保存名为产品图片的文件夹路径
import os
import logging
import shutil
import sys

# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# --- 保存名为产品图片的文件夹路径 ---
def process_image_folders(folder_path:str, OUTPUT_PATH:str, factory_name:str) -> str:
    """
    处理产品图片文件夹
    Args:
        folder_path (str): 工厂文件夹路径
        OUTPUT_PATH (str): 输出路径
        factory_name (str): 工厂名称
    Returns:
        output_folder (str): 图片输出文件夹路径（如无图片则返回None）
    """
    try:
        # 输出路径
        output_folder = os.path.join(OUTPUT_PATH, f"{factory_name}_产品图片")
        found_image = False

        # 遍历所有可能的图片文件夹名称
        for folder_name in IMAGE_FOLDER_NAMES:
            image_folder = os.path.join(folder_path, folder_name)
            # 检查文件夹是否存在
            if os.path.exists(image_folder) and os.path.isdir(image_folder):
                found_image = True
                # 确保输出目录存在
                os.makedirs(output_folder, exist_ok=True)
                
                # 复制整个文件夹到输出路径
                try:
                    shutil.copytree(image_folder, output_folder, dirs_exist_ok=True)
                    # logging.info(f"成功复制图片文件夹: {folder_name} -> {output_folder}")
                    break  # 找到第一个匹配的文件夹就停止
                except Exception as e:
                    logging.error(f"复制图片文件夹时出错: {str(e)}")
                    continue
            
                
        if found_image:
            logging.info(f"图片文件夹已保存到: {output_folder}")
            return output_folder
        else:
            logging.info(f"未找到图片文件夹:{folder_path}")
            return None

    except Exception as e:
        logging.error(f"处理图片文件夹时出错: {str(e)}")
        return None


if __name__ == "__main__":
    # 测试用例
    test_folder_path = r"tests\曹县华悦工艺品有限公司"
    test_output_path = r"tests\folder_img_save"
    test_factory_name = "曹县华悦工艺品有限公司"
    process_image_folders(test_folder_path, test_output_path, test_factory_name)