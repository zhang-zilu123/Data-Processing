# 清洗厂商名称（只返回第一个）
import re

# ---清洗厂商名称（只返回第一个）---
def clean_factory_name(factory_name: str) -> str:
    """
    厂商名称清洗
    处理逻辑：1、遇到'\'或'/'或 '\n' 或 多个空格 分隔，只返回第一个
             2、返回不含英文的部分和末尾的括号以内容
             
    args:
        factory_name: 厂商名称
    return:
        str: 清洗后的厂商名称
    """
    
    if not factory_name:
        return ''
    
    factory_name = factory_name.replace('（','(').replace('）',')')

    # 步骤1：按 \ 或 / 或 \n 或 多个空格分割，只取第一个
    if re.search(r'[\\/\n]|\s{3,}', factory_name):
        first_part = re.split(r'[\\/\n]|\s{3,}', factory_name)[0].strip()
        factory_name = first_part
    
    # 步骤2：去除末尾括号内容
    content_without_bracket = re.sub(r'\([^)]+\)$', '', factory_name).strip()
    factory_name = content_without_bracket
    
    # 步骤3：去除英文字符
    if not re.match(r'^[A-Za-z]+$', factory_name):
    # 如果不是纯英文，删除末尾的英文字符
        factory_name = re.sub(r'\s+[A-Za-z].*$', '', factory_name).strip()
    
    return factory_name



if __name__ == "__main__":
    # 测试用例
    test_names = [
        "广东昕晟实业有限公司/柬埔寨工厂:R&G环球家居用品有限公司",
        "宁波源乾日用品有限公司\越南亿源家居用品有限公司",
        "柬埔寨易宏箱包有限公司\n(易宏xxxxxx)",
        "柬埔寨恒丰(实业)有限公司",
        "台州市黄岩品信橡塑科技有限公司\n（宁波市翔升生活用品有限责任公司）",
        "越南兆荣家具有限公司Asdfbbferb /n 安吉悦信家具有限公司",
        "越南兆荣家具ascd有限公司",
        "兆荣家具有限公司（阿斯顿）   悦信家具有限公司       字节跳动",
        "柬埔寨怡荣生手袋实业有限公司 T&L(Cambodia) HANDBAGS INDUSTRIAL CO., LTD.",

    ]

    for name in test_names:
        cleaned_name = clean_factory_name(name)
        print(f"原始名称: {name} -> 清理后名称: {cleaned_name} \n")