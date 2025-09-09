# 合并JSON文件中相同厂商的记录，解决重名厂商数据重复问题

import json
from datetime import datetime
import logging
import os
from collections import defaultdict


import sys
# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from setting.config import *
from src.utils.json_logger import setup_json_logger
from src.utils.clean_factory_name import clean_factory_name

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#---------------------- 辅助功能函数 --------------------------------

# 检查记录是否有有效的日期字段
def has_valid_date(record:dict) -> bool:
    """
    检查记录是否有有效的日期字段

    参数：
        record  工厂记录字典
    
    返回：
        bool  是否有有效日期
    """
    date_val = record.get('日期', '')
    # 有效日期判断：非空且不是"无"
    return date_val not in ['', None, '无']


# 处理主营产品/验厂认证 字段，将值添加到产品集合
def process_product_field(value:str, product_set:set) -> None:
    """
    处理主营产品字段，将值添加到产品集合

    参数：
        value: 主营产品字段值
        product_set: 用于收集产品的集合

    返回：
        None
    """

    if not value:
        return
    
    # 处理不同类型的产品数据
    if isinstance(value, list):
        # 列表类型：直接添加到集合
        for item in value:
            if item:  # 跳过空项
                product_set.add(item.strip())
    elif isinstance(value, str):
        # 字符串类型：选择最佳分隔符进行分割
        best_separator = None
        max_splits = 0
        
        # 找出能产生最多有效分割的分隔符
        for separator in FIELD_SEPARATORS:
            if separator in value:
                splits = value.split(separator)
                valid_splits = [item.strip() for item in splits if item.strip()]
                if len(valid_splits) > max_splits:
                    max_splits = len(valid_splits)
                    best_separator = separator
        
        if best_separator:
            # 使用最佳分隔符进行分割
            for item in value.split(best_separator):
                if item.strip():
                    product_set.add(item.strip())
        else:
            # 如果没有找到任何分隔符，添加整个字符串
            if value.strip():
                product_set.add(value.strip())

def choose_better_field(value1:str, value2:str) -> str:
    """
    选择两个字段值中更好的那个
    规则：无的用另一个补充，都有值的选择更长的
    
    参数：
        value1: 第一个字段值
        value2: 第二个字段值
    
    返回：
        更好的字段值
    """
    # 将None转为空字符串处理
    v1 = value1 if value1 is not None else ""
    v2 = value2 if value2 is not None else ""
    
    # 判断是否为"无"
    is_empty_v1 = v1 == "" or v1 == "无"
    is_empty_v2 = v2 == "" or v2 == "无"
    
    # 如果v1为空，选择v2
    if is_empty_v1 and not is_empty_v2:
        return v2
    # 如果v2为空，选择v1  
    elif is_empty_v2 and not is_empty_v1:
        return v1
    # 如果都为空，返回空字符串
    elif is_empty_v1 and is_empty_v2:
        return ""
    # 如果都有值，选择更长的
    else:
        
        return v1 if len(str(v1)) >= len(str(v2)) else v2

def merge_same_date_records(records:list, name:str) -> dict:
    """
    合并相同日期的多条记录，互相补充
    
    参数：
        records(list): 相同日期的记录列表
        name(str): 厂商名称（用于日志）
    
    返回：
        dict: 合并后的记录
    """
    if len(records) == 1:
        return records[0]
    
    logging.info(f"厂商 '{name}' 有 {len(records)} 条相同日期记录，进行互相补充")
    
    # 以第一条记录为基础
    merged_record = records[0].copy()

    # 需要进行比较长度的字段
    length_comparison_fields = ['主销市场', '备注','合作情况','联系方式']
    # 只需要取第一个有效值的字段
    first_value_fields = ['微信', '网址', '图片文件夹路径']

    for record in records[1:]:
        # 处理需要比较长度的字段
        for field in length_comparison_fields:
            merged_record[field] = choose_better_field(merged_record.get(field), record.get(field))

        # 处理只需要取第一个有效值的字段
        for field in first_value_fields:
            current_val = merged_record.get(field)
            if not current_val or current_val == '无':
                merged_record[field] = record.get(field)

        # 主营产品和验厂/认证需要特殊处理（合并去重）
        all_products = set()
        cert_set = set()
        
        # 处理主营产品
        product = record.get('主营产品', '')
        if product and product != '无':
            process_product_field(product, all_products)
        # 处理验厂/认证
        certification = record.get('验厂/认证', '')
        if certification and certification != '无':
            process_product_field(certification, cert_set)
            
    
    merged_record['主营产品'] = list(all_products) if all_products else ""
    merged_record['验厂/认证'] = list(cert_set) if cert_set else ""

    # 标签合并
    tag_list = set()
    
    for record in records:
        tag = record.get('标签', '')
        if tag and tag != '无':
            if isinstance(tag, list):
                for t in tag:
                    tag_list.add(t)
            else:
                tag_list.add(tag)
                
    merged_record['标签'] = list(tag_list) if tag_list else ""
    
    logging.info(f"厂商 '{name}' 相同日期记录补充完成")
    return merged_record




#-------------------------------- 主合并函数 --------------------------------
def merge_factories(data:list)->list:
    """
    主合并函数：处理工厂数据并合并重复项

    参数：
        data: 原始工厂数据列表

    返回：
        合并后的工厂数据列表
    """

    # 1. 按厂商名称分组
    groups = defaultdict(list)
    for factory in data:
        # 使用厂商名称作为分组键
        name = factory.get('厂商名称', '')
        if name:  # 确保名称不为空
            clean_name = clean_factory_name(name)
            groups[clean_name].append(factory)
    
    logging.info(f"分组完成，共有 {len(groups)} 个不同的厂商")
    
    # 统计需要合并的分组
    need_merge_groups = 0
    for clean_name, factories in groups.items():
        if len(factories) > 1:
            need_merge_groups += 1
            logging.info(f"厂商 '{clean_name}' 有 {len(factories)} 条记录需要合并")
    
    logging.info(f"需要合并的厂商数量: {need_merge_groups}")
    
    result = []
    
    # 2. 处理每个分组
    for clean_name, factories in groups.items():
        # 2.1 单条记录直接添加到结果
        if len(factories) == 1:
            result.append(factories[0])
            continue
        
        logging.info(f"开始合并厂商 '{clean_name}' 的 {len(factories)} 条记录")
        
        # 2.2 收集合并来源（文件路径）
        sources = set()
        for f in factories:
            file_path = f.get('文件路径', '')
            if file_path:
                sources.add(file_path)
        merged_source = list(sources)
        logging.info(f"厂商 '{clean_name}' 的数据来源: {merged_source}")
        
        # 2.3 判断日期情况
        all_have_dates = all(has_valid_date(f) for f in factories)
        logging.info(f"厂商 '{clean_name}' 所有记录都有有效日期: {all_have_dates}")
        
        # 2.4 初始化合并记录
        merged_record = {}
        
        if all_have_dates:
            # 场景1：所有记录都有有效日期
            logging.info(f"厂商 '{clean_name}' 使用场景1合并：所有记录都有有效日期")
            
            # 找到最新日期
            max_date = max(factory['日期'] for factory in factories)
            # 找到所有与最新日期相同或在一个月内的记录
            recent_records = []
            for factory in factories:
                factory_date = factory['日期']
                
                # 检查是否是相同日期
                if factory_date == max_date:
                    recent_records.append(factory)
                else:

                    # 检查是否在一个月内（30天）
                    try:
                        date_formats = ['%Y/%m/%d', '%Y-%m-%d', '%Y.%m.%d', '%Y年%m月%d日']
                        d1 = None
                        d2 = None
                        
                        # 尝试解析日期
                        for fmt in date_formats:
                            try:
                                d1 = datetime.strptime(factory_date, fmt)
                                d2 = datetime.strptime(max_date, fmt)
                                break
                            except ValueError:
                                continue
                        
                        # 如果解析成功且在30天内，加入近期记录
                        if d1 is not None and d2 is not None:
                            diff = abs((d1 - d2).days)
                            if diff <= 30:
                                recent_records.append(factory)
                                # logging.info(f"厂商 '{clean_name}' 记录日期 {factory_date} 与最新日期 {max_date} 相差 {diff} 天，纳入近期记录")
                    except Exception:
                        # 日期解析失败，只使用相同日期的记录
                        pass
            
            logging.info(f"厂商 '{clean_name}' 最新日期: {max_date}, 近期记录数量: {len(recent_records)}")
            
            if len(recent_records) == 1:
                # 1、日期不相同，且超过30天，只保留最新日期
                logging.info(f"厂商 '{clean_name}' 使用场景1-1合并：日期不相同，且超过30天，只保留最新日期")
                base_record = recent_records[0]
                merged_record = base_record.copy()
            
                
                # 合并所有记录的主营产品
                all_products = set()
                process_product_field(base_record.get('主营产品', ''), all_products)
                original_products_count = len(all_products)
                
                for other in factories:
                    if other != base_record:
                        process_product_field(other.get('主营产品', ''), all_products)
                
                merged_record['主营产品'] = list(all_products) if all_products else ""

                all_tag = set()
                for record in recent_records:
                    tag = record.get('标签', '')
                    if tag and tag != '无':
                        if isinstance(tag, list):
                            for t in tag:
                                all_tag.add(t)
                # logging.info(f"厂商 '{clean_name}' 主营产品合并：原始 {original_products_count} -> 合并后 {all_products}")

                
            else:
                # 2、有多条相同日期的最新记录，进行互相补充
                merged_record = merge_same_date_records(recent_records, clean_name)
                
                
        else:
            # 场景2：存在无日期记录
            logging.info(f"厂商 '{clean_name}' 使用场景2合并：存在无日期记录")

            # 初始化字段收集容器
            all_products = set()  # 主营产品
            contact_list = []    # 联系方式
            wechat_list = []      # 微信
            website_list = []     # 网址
            merged_img_list = []       # 图片文件夹路径
            cert_set = set()      # 验厂/认证（去重）
            cooperation_situation_list =[] # 合作情况
            market_candidates = []  # 主销市场候选
            remark_candidates = []  # 备注候选
            tag_list = set()  # 标签
            
            # 2.4.4 遍历所有记录收集字段值
            for i, factory in enumerate(factories):
                logging.info(f"厂商 '{clean_name}' 处理第 {i+1} 条记录，日期: {factory.get('日期', '无日期')}")
                
                # 处理主营产品
                product = factory.get('主营产品', '')
                if product and product != '无':
                    original_product_count = len(all_products)
                    process_product_field(product, all_products)
                    

                # 处理联系方式
                contact = factory.get('联系方式', '')
                if contact and contact != '无':
                    contact_list.append(contact)
                    
                
                # 处理微信（合并）
                wechat = factory.get('微信', '')
                if wechat and wechat != '无':
                    wechat_list.append(wechat)
                    logging.info(f"厂商 '{clean_name}' 第 {i+1} 条记录添加了微信信息")
                
                # 处理网址
                website = factory.get('网址', '')
                if website and website != '无':
                    website_list.append(website)
         

                # 处理合作情况
                cooperation_situation = factory.get('合作情况', '')
                if cooperation_situation and cooperation_situation != '无':
                    cooperation_situation_list.append(cooperation_situation)
         
                
                # 处理验厂/认证
                certification = factory.get('验厂/认证', '')
                if certification and certification != '无':
                    process_product_field(certification, cert_set)
                    
         

                #处理文件图片
                img_field = factory.get('图片文件夹路径', '')
                if img_field and img_field != '无':
                        merged_img_list.append(img_field)
                        

                
                # 收集主销市场候选
                market = factory.get('主销市场', '')
                if market and market != '无':
                    market_candidates.append(market)
                
                # 收集备注候选
                remark = factory.get('备注', '')
                if remark and remark != '无':
                    remark_candidates.append(remark)
                
                # 收集标签
                tag = factory.get('标签', '')
                if tag and tag != '无':
                    if isinstance(tag, list):
                        for t in tag:
                            tag_list.add(t)
                    else:
                        tag_list.add(tag)

            
            # 2.4.6 复制其他字段（使用第一条记录的值）
            for key in factories[0].keys():
                if key not in merged_record:
                    merged_record[key] = factories[0][key]
            
            # 2.4.7 设置合并后的多值字段
            merged_record['联系方式'] = contact_list if contact_list else ""
            logging.info(f"厂商 '{clean_name}' 联系方式合并：{merged_record['联系方式']}")

            if len(wechat_list) > 1:
                logging.error(f"厂商 '{clean_name}' 有两个微信合并，请检查")

            merged_record['微信'] = wechat_list if wechat_list else ""
            logging.info(f"厂商 '{clean_name}' 微信合并：{merged_record['微信']}")
            merged_record['网址'] = website_list if website_list else ""
            logging.info(f"厂商 '{clean_name}' 网址合并：{merged_record['网址']}")
            merged_record['验厂/认证'] = list(cert_set) if cert_set else ""
            logging.info(f"厂商 '{clean_name}' 验厂/认证合并：{merged_record['验厂/认证']}")
            merged_record['合作情况'] = cooperation_situation_list if cooperation_situation_list else ""
            logging.info(f"厂商 '{clean_name}' 合作情况合并：{merged_record['合作情况']}")
            merged_record['图片文件夹路径'] = merged_img_list[0] if merged_img_list else ""
            logging.info(f"厂商 '{clean_name}' 图片文件夹路径合并：{merged_record['图片文件夹路径']}")
            merged_record['标签'] = list(tag_list) if tag_list else ""

            # 2.4.8 选择最优单值字段
            merged_record['主销市场'] = max(market_candidates, key=len) if market_candidates else ""
            logging.info(f"厂商 '{clean_name}' 主销市场合并：{merged_record['主销市场']}")
            merged_record['备注'] = max(remark_candidates, key=len) if remark_candidates else ""
            logging.info(f"厂商 '{clean_name}' 备注合并：{merged_record['备注']}")
        
        # 2.5 添加合并来源字段
        merged_record['合并来源'] = merged_source
        result.append(merged_record)
        logging.info(f"厂商 '{clean_name}' 合并完成，{len(factories)} 条记录合并为 1 条")
    
    logging.info(f"所有厂商合并完成，最终记录数: {len(result)}")
    return result


#-------------------------------- 主函数 --------------------------------
def merge_unique_factory_json(input_file:str, output_file:str) -> None:
    """
    读取输入文件，处理数据，保存结果

    参数：
        input_file: 输入JSON文件路径
        output_file: 输出JSON文件路径

    返回：
        None
    """
    logging.info(f"开始处理文件: {input_file}")
    
    # 1. 读取输入数据
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    logging.info(f"成功读取输入文件，包含 {len(data)} 条记录")
    
    # 2. 处理数据
    merged_data = merge_factories(data)

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # 3. 保存结果
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(merged_data, f, ensure_ascii=False, indent=2)
    
    logging.info(f"成功保存结果文件")
    logging.info(f"处理完成！原始记录数: {len(data)}, 合并后记录数: {len(merged_data)}")
    logging.info(f"记录减少: {len(data) - len(merged_data)} 条")



#-------------------------------- 测试主函数 --------------------------------
if __name__ == "__main__":

    # 配置输入输出文件路径
    input_json = r"data\processed_data\combined\combined_tag.json"
    output_json = r"data\processed_data\merged_factories\merged_factories_tag111.json"
    
    log_file = 'merge_processor_tag111.log'
    logger = setup_json_logger(log_dir='logs', log_file=log_file)
    
    # 执行主程序
    merge_unique_factory_json(input_json, output_json)


    

