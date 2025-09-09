# 解析工厂信息字段，分别提取主销市场和备注信息

import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def parse_factory_info(raw_text: str) -> dict:
    """
    解析工厂信息字段的专用函数
    
    处理逻辑：
    1. 将原始文本按换行符分割为多行
    2. 逐行匹配关键词并分类
    3. 主销市场类信息：外贸占比、跨境占比、主营市场
    4. 备注类信息：工厂面积、员工人数、年 产 值
    5. 同类别信息用换行符连接
    
    参数：
        raw_text (str): 原始工厂信息文本
    
    返回：
        dict: 包含解析后数据的字典，格式为：
              {"主销市场": "解析后的市场信息", "备注": "解析后的备注信息"}
    """
    # 初始化结果容器
    result = {"主销市场": "", "备注": ""}
    
    # 关键词映射配置
    MARKET_KEYWORDS = ["外贸占比", "跨境占比", "主营市场"]
    REMARK_KEYWORDS = ["工厂面积", "员工人数", "年 产 值"]
    
    # 空值检查
    if not raw_text or not isinstance(raw_text, str):
        return result
    
    try:
        # 分割文本为多行
        lines = raw_text.split('\n')
        
        # 遍历每一行文本
        for line in lines:
            # 清理空白字符并跳过空行
            clean_line = line.strip()
            if not clean_line:
                continue
            
            # 检查是否为市场信息
            if any(kw in clean_line for kw in MARKET_KEYWORDS):
                if result["主销市场"]:
                    result["主销市场"] += "\n"  # 多行信息用换行符分隔
                result["主销市场"] += clean_line
            
            # 检查是否为备注信息
            elif any(kw in clean_line for kw in REMARK_KEYWORDS):
                if result["备注"]:
                    result["备注"] += "\n"
                result["备注"] += clean_line
        
        return result
    
    except Exception as e:
        logging.error(f"解析工厂信息时出错: {str(e)}", exc_info=True)
        return result
    


if __name__ == "__main__":
    raw_text = "外贸占比：95%\n主营市场：欧洲占比50%\n亚洲占比30% \n美国占比15%\n其他占比5%\n工厂面积：3000㎡\n员工人数：110名\n年 产 值：8000万人民币"
    result = parse_factory_info(raw_text)
    print(result)