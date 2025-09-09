# 将单个文件的处理结果保存到厂商专属文件夹中
import os
import json
import logging
import sys
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from src.utils.clean_factory_name import clean_factory_name


def get_vendor_folder_name(vendor_name:str) -> str:
    """
    处理厂商名称，获取文件夹名称
    
    参数：
        vendor_name (str): 原始厂商名称
        
    返回：
        str: 处理后的文件夹名称
    """
    if not vendor_name:
        return '未知厂商'
    
    # 如果包含"/"，取第一个
    if '/' in vendor_name:
        vendor_name = vendor_name.split('/')[0]
    
    # 去除前后空格
    vendor_name = vendor_name.strip()
    
    # 如果处理后为空，使用默认名称
    if not vendor_name:
        return '未知厂商'
    
    return vendor_name

def get_unique_folder_name(base_name:str, output_path:str) -> str:
    """
    获取唯一的文件夹名称，如果存在重复则添加后缀
    
    参数：
        base_name (str): 基础文件夹名称
        output_path (str): 输出路径
        
    返回：
        str: 唯一的文件夹名称
    """
    folder_name = base_name
    counter = 2
    
    # 检查文件夹是否存在，如果存在则添加后缀
    while os.path.exists(os.path.join(output_path, folder_name)):
        folder_name = f"{base_name}_{counter}"
        counter += 1
    
    return folder_name

def make_vendor_folder(factory_name:str,output_path:str) -> str:
    """
    创建厂商专属文件夹

    参数：
        factory_name (str): 厂商名称
        output_path (str): 输出路径
        
    返回：
        str: 厂商专属文件夹路径
    """

    # 获取厂商文件夹名称
    vendor_name = clean_factory_name(factory_name)
    
    # 获取唯一的文件夹名称
    unique_folder_name = get_unique_folder_name(vendor_name, output_path)
    
    # 创建厂商专属文件夹
    vendor_folder = os.path.join(output_path, unique_folder_name)
    os.makedirs(vendor_folder, exist_ok=True)
    return vendor_folder


def save_result_to_vendor_folder(vendor_folder:str, result:dict) -> str:
    """
    将单个文件的处理结果保存到厂商专属文件夹中
    
    参数：
        vendor_folder (str): 厂商专属文件夹路径
        result (dict): 单个文件的处理结果字典
        
    返回：
        str: 保存的JSON文件路径
    """
    try:
        factory_name = get_vendor_folder_name(result.get('厂商名称', ''))
        vendor_name = clean_factory_name(factory_name)
        
        # 保存JSON文件到厂商文件夹
        json_filename = f"{vendor_name}_信息.json"
        json_file_path = os.path.join(vendor_folder, json_filename)
        
        
        # 直接保存result到JSON文件
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
        

        logging.info(f"厂商信息json已保存到: {json_file_path}")
        return json_file_path
        
    except Exception as e:
        logging.error(f"保存结果时出错: {str(e)}", exc_info=True)
        return None



#-------------------------------- 测试主函数 --------------------------------
if __name__ == "__main__":
    result = {
        "厂商名称": "江苏东塑休闲用品有限公司/浙江科泽户外用品有限公司",
        "主营产品": "塑料折叠桌、折叠凳、折叠椅、户外储物箱、野餐桌等吹塑家具休闲产品",
        "联系方式": "林忠巧 总 经 理 手机：13795202769\n董宏伟 业务代表 手机：18251027703\n地址：江苏东塑：江苏扬州仪征市月塘镇工业区\n地址：浙江科泽：浙江省湖州市长兴县林城镇工业集中区志远路16号",
        "验厂/认证": "BSCI",
        "合作情况": "合作公司：易佰、豪雅、旗奥、安徽轻工、FDW、ALPEMUSA\n合作客户：家乐福",
        "是否供样": "",
        "网址": "",
    }
    output_path = r"tests\processed_data\word"
    vendor_folder = make_vendor_folder("江苏东塑休闲用品有限公司",output_path)
    print(f"厂商文件夹：{vendor_folder}")
    save_result_to_vendor_folder(vendor_folder, result)