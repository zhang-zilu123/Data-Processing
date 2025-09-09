# 将Excel数据转换为JSON数据(字段值格式化特殊化处理、验厂/认证信息处理、映射字段添加前缀)

import sys
import os
# 添加项目根目录到路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
sys.path.insert(0, project_root)

from setting.config import *
import logging  # 导入日志模块

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#---------------------- 数据格式化功能函数 --------------------------------

# --- 函数：字段值格式化特殊化处理 ---
def format_value_with_prefix(key:str, value, template_config:dict) -> str:
    """
    根据模板配置对字段值进行格式化处理，包括特殊字段处理和前缀添加
    
    功能说明：
    对从Excel中提取的原始字段值进行标准化处理，支持多种数据类型的转换，
    处理特殊格式的字段值，并根据配置添加相应的前缀。
    
    处理流程：
    1. 特殊字段类型检测和处理（手机号、日期等）
    2. 无效值过滤和清理
    3. 数据类型转换（确保为字符串）
    4. 根据模板配置添加前缀
    5. 返回格式化后的值
    
    参数：
        key (str): 字段键名，用于在模板配置中查找对应的处理规则
        value(Any): 字段值，需要格式化的原始数据（可能是各种类型）
        template_config (dict): 模板配置字典，包含字段的前缀和格式化规则
        
    返回：
        str/None: 格式化后的字符串值，无效值返回None
        
    特殊处理字段：
        - factory_phone: 手机号去除小数点
        - vat_invoice_count: 税点字段清理
        - season_capacity: 季节产能字段验证
        - export_port: 外发字段验证
        - establish_time: 日期格式标准化
        
        - 根据template_config中的prefix配置自动添加前缀
        - 空值不添加前缀
        

    """
    # 步骤1：特殊字段处理 - 手机号小数点问题
    if key == 'factory_phone' and isinstance(value, float) and value.is_integer():
        value = str(int(value))

    # 步骤2：特殊字段值过滤和清理    
       
    # 处理税点字段：移除表单模板文本，只保留实际值
    if key == 'vat_invoice_count' and '是□否□' in str(value):
        value = value.replace('是□否□', '').strip()
        # 如果只剩下"税点："标签，则认为无有效值
        if value.rstrip() == '税点：' or value.rstrip() == '税点:':
            return None

    # 处理季节产能字段：过滤表单模板文本
    elif key == 'season_capacity' and value.rstrip() == '□淡季Low Season\n□旺季Peak Season':
        return None  # 直接返回None，这样这个字段会被跳过
    
    # 处理外发字段：过滤表单模板文本
    elif key == 'export_port' and value.rstrip() == '是否外发\n是□否□':
        return None  # 直接返回None，这样这个字段会被跳过
    
    # 处理产能相关字段：过滤无意义的单位
    elif key in ['Used production capacity per month', 'Total production capacity per month', 'Spare capacity per month'] and value == '个':
        return None
    
    # 步骤3：日期字段特殊处理
    # 确保日期格式的统一性（YYYY-MM-DD）
    elif key == 'establish_time' and value:
        if hasattr(value, 'strftime'):
            # 如果是datetime对象，格式化为标准日期字符串
            value = value.strftime('%Y-%m-%d')
        else:
            # 如果是字符串，只取前10位确保格式正确
            value = str(value)[:10]

    # 步骤4：数据类型统一 - 确保所有值都转换为字符串
    if value is not None:
        value = str(value)

    prefix = template_config.get(key, {}).get('prefix', '')
    # 如果存在前缀且值非空，则添加前缀
    if prefix and value:
        value = f"{prefix}{value}"

    # 步骤5：返回格式化后的值
    return value


#---------------------- 认证信息处理功能函数 --------------------------------

# --- 函数：验厂/认证信息处理 ---
def process_certificates(factory_data:dict, certificate_fields:list, template_config:dict) -> str:
    """
    处理工厂的验厂/认证信息，将多个认证字段合并为标准格式的认证字符串
    
    功能说明：
    从工厂数据中提取各种认证信息，包括标准认证和自定义认证，
    根据模板配置转换为显示名称，并合并为统一的认证字符串。
    
    处理流程：
    1. 遍历所有预定义的认证字段
    2. 检查每个认证字段是否存在非空值，若存在，则返回标准认证名称
    3. 自定义认证（Other字段）使用原始值
    4. 合并所有有效认证为字符串
    
    参数：
        factory_data (dict): 工厂数据字典，包含各种认证字段
        certificate_fields (list): 认证字段名称列表，定义需要检查的认证类型
        template_config (dict): 模板配置，包含认证字段的显示名称映射
        
    返回：
        str: 合并后的认证字符串，用顿号分隔；无认证时返回''


    """
    # 步骤1：初始化认证名称列表
    cert_names = []
    
    # 步骤2：遍历所有认证字段，提取有效认证
    for cert in certificate_fields:
        value = factory_data.get(cert)
        
        # 步骤3：特殊处理自定义认证字段（Other）
        if cert == 'Other':
            # 确保Other字段的值不为None且非空（支持非字符串类型）
            if value is not None and (not isinstance(value, str) or value.strip()):
                cert_names.append(str(value))
        else:
            # 步骤4：处理标准认证字段
            # 检查认证值是否存在且非空
            if value is not None and (not isinstance(value, str) or value.strip()):
                # 从模板配置中获取认证的标准显示名称
                cert_names.append(template_config[cert]['expected_keyword'])

    # 步骤5：合并认证信息
    # 如果有认证信息则用中文顿号连接，否则返回默认提示
    return '、'.join(cert_names) if cert_names else ''




#---------------------- 字段映射处理功能函数 --------------------------------

# --- 函数：映射字段添加前缀 ---
def process_mapped_fields(factory_data:dict, mapped_keys:list, template_config:dict) -> list:
    """
    根据配置模板中的设定，为存在prefix的字段添加前缀，并返回格式化后的字段值列表
    
    处理流程：
    1. 遍历映射字段列表
    2. 从工厂数据中获取字段值 
    3. 调用格式化函数处理字段值
    4. 返回格式化后的字段值列表
    
    参数：
        factory_data (dict): 工厂原始数据，包含从Excel提取的所有字段
        mapped_keys (list): 映射的键列表，定义当前JSON字段对应的Excel字段
        template_config (dict): 模板配置，包含字段的格式化规则和前缀信息
        
    返回：
        list: 处理后的有效字段值列表，每个值都经过格式化和验证

        
    格式化处理：
        - 调用format_value_with_prefix进行值格式化
        - 自动添加配置的前缀
        - 处理特殊字段类型
        
    """
    # 步骤1：初始化有效值列表
    valid_items = []
    
    # 步骤2：遍历所有映射字段
    for key in mapped_keys:
        # 步骤3：从工厂数据中获取字段值
        value = factory_data.get(key)
        
        # 步骤4：验证字段值的有效性
        # 检查字段值是否存在且非空（对字符串进行trim检查）
        if value is not None and (not isinstance(value, str) or value.strip()):
            
            # 步骤5：格式化字段值，添加前缀以及特殊处理一些字段
            formatted_value = format_value_with_prefix(key, value, template_config)
            
            # 步骤6：过滤格式化后的无效值
            # 如果格式化函数返回None，则跳过该字段
            if formatted_value is not None:
                valid_items.append(formatted_value)
            
    return valid_items

#---------------------- 主要转换功能函数 --------------------------------

# --- 函数：JSON数据生成主函数 ---
def json_from_factory_data(factory_data:dict, template_config:dict, input_path:str, img_output_folder_path:str) -> dict:
    """
    功能说明：
    根据JSON_excel_FORMAT模板和TEXT_LABELS_xlsx_fty映射规则，将从Excel文件中
    提取的工厂原始数据转换为标准化的JSON数据结构，支持各种字段类型的处理。
    
    处理流程：
    1. 初始化结果字典和配置变量
    2. 遍历JSON格式模板中的所有字段
    3. 对验厂/认证字段进行特殊处理
    4. 对普通字段进行映射和格式化处理
    5. 根据映射类型决定字段值的组合方式
    6. 添加文件路径和图片路径信息
    7. 返回完整的标准化JSON数据
    
    参数：
        factory_data (dict): 工厂原始数据，从Excel文件中提取的键值对
        template_config (dict): 模板配置（如EXCEL_FORMATE_FTY_1），定义字段格式和前缀
        input_path (str): 输入文件路径，用于追溯数据来源
        img_output_folder_path (str): 图片输出文件夹路径，用于关联图片资源
        
    返回：
        dict: 标准化的JSON数据字典，失败时返回None
        
        
    字段处理策略：
        - 单字段映射：直接使用格式化后的值
        - 多字段映射：用分号（；）连接多个值
        - 无值字段：填充"暂无记录"作为默认值
        - 认证字段：特殊处理，用顿号（、）连接        
 

    """
    try:
        # 步骤1：初始化结果字典和配置
        result = {}
        # 定义需要特殊处理的认证字段列表
        certificate_fields = ['ISO9001', 'BSCI', 'Sedex', 'Disney FAMA', 'Walmart', 'Target', 'Other']

        # 步骤2：遍历JSON模板中定义的所有字段
        for json_key in JSON_FORMAT.keys():
            # 步骤3：特殊处理"验厂/认证"字段
            # 认证字段需要合并多个认证类型，使用专门的处理函数
            if json_key == '验厂/认证':
                result[json_key] = process_certificates(factory_data, certificate_fields, template_config)
                continue

            # 步骤4：处理普通字段的映射和转换
            # 从映射配置中获取当前JSON字段对应的Excel字段列表
            mapped_keys = TEXT_LABELS_xlsx_fty.get(json_key, [json_key])
            # 处理映射字段，获取格式化后的有效值列表
            valid_items = process_mapped_fields(factory_data, mapped_keys, template_config)
            
            # 步骤5：根据映射类型决定字段值的组合方式
            if valid_items:
                if len(mapped_keys) > 1:
                    # 多字段映射：用中文分号连接多个值
                    result[json_key] = '；'.join(valid_items)
                else:
                    # 单字段映射：直接使用第一个（也是唯一一个）值
                    result[json_key] = valid_items[0]
            else:
                # 步骤6：无有效值时填充默认值
                result[json_key] = ''
        
        # 步骤7：添加元数据信息

        # 记录数据来源文件路径，便于追溯和审计
        result['文件路径'] = input_path
        # 记录关联的图片文件夹路径，便于图片资源管理
        result['图片文件夹路径'] = img_output_folder_path

        return result
        
    except Exception as e:
        logging.error(f"生成JSON数据时出错: {str(e)}")
        return None



#---------------------- 主程序入口 --------------------------------

# 主程序入口 - 用于测试和调试
if __name__ == "__main__":
    # 定义测试用的输入输出路径
    excel_path = r"test_file\第二次走访\临沭县兴隆五金工具有限公司-工厂情况信息表.xlsx"
    output_path = r"test_output"
    
    # 处理Excel文件的示例代码
    # 实际使用时需要先调用process_excel函数获取工厂数据
    # factory_data, img_output_folder_path = process_excel(excel_path, output_path)
    # print(f'factory_data:{factory_data}')

    img_output_folder_path = r"test_output\临沭县兴隆五金工具有限公司_产品图片"

    factory_data={'factory_name': '曹县舍得工艺有限公司', 
    'factory_contact': '李锦锋', 
    'factory_address': '山东省菏泽市曹县青岗集镇胡王庄', 
    'factory_phone': '15865175077', 
    'factory_legal_representative': '王小明', 
    'factory_mony': '100万', 
    'product_category': '各种木制工艺品', 
    'establish_time': '2014年9月', 
    'annual_sales': '1200万', 
    'employee_count': '24个', 
    'factory_website': 'www.shedecraft.com', 
    'vat_invoice_count': '是□否□     税点：', 
    'factory_area': 12000.0, 
    'warehouse_area': None, 
    'dormitory_area': '200平方',
    'canteen_area': '100平方', 'production_process': '雕刻，组装', 
    'export_port': '是否外发\n否', 'season_capacity': '□淡季Low Season\n□旺季Peak Season', 
    'main_customer': '美国，欧洲', 'usa_share': 50.0, 'eu_share': 40.0, 
    'others_share': 10.0, 'domestic_share': 10.0, 'export_share': 90.0, 
    'ISO9001': None, 'BSCI': None, 'Sedex': None, 'Disney FAMA': None, 'Walmart': None, 'Target': None,
    }
    
    # 生成标准JSON数据的示例代码
    result = json_from_factory_data(factory_data, EXCEL_FORMATE_FTY_1,excel_path,img_output_folder_path)
    print(result)