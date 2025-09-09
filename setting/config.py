#配置文件：存放日志文件名、文本标签映射等常量


#------------------特殊列值配置-----------------------
DATA_SOURCE='集团内部：产品开发部'

#“标签”列 根据输入文件的路径进行提取。未匹配上则跳过，每个子list中最多匹配一个标签，最后将所有子标签拼接成一个字符串。
POSSIBLE_TAGS = [
    ['2021','2022','2023','2024','2025'],
    ["宁波","义乌","潮汕","青岛"],
    [r"D([1-9]|[1-9][0-9]|[1-9][0-9][0-9])期",
     r"第[一二三四五六七八九十百千零]{1,10}期",
     r"(?:0?[1-9]|1[0-2])月",
     r"D([1-9][0-9]|[1-9][0-9][0-9])",
     r"\d+\.\d+"],
    ['双战略供应商','到访工厂','走访供应商','供应商交流会','供应商资料表','双战略','推介会','越南工厂','柬埔寨工厂'],
]



# 需要识别为图片的文件扩展名列表 (不区分大小写)
IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff']

# 中文字符到英文字符的映射
CHINESE_TO_ENGLISH_MAP = {'，': ',', '；': ';', '：': ':','（':'(','）':')'}
 
# 需要识别为产品图片保存路径的文件扩展名  （应用于folder_img_save）
IMAGE_FOLDER_NAMES = ['产品图片', '产品照片']

# 重要字段列表  （应用于set_same_name  拼接厂商信息时，优先拼接这些字段）
IMPORTANT_FIELDS = ["主营产品", "联系方式", "主销市场", "合作情况", "备注", "验厂/认证"]

# 用于识别日期文本的关键词或正则表达式模式
DATE_PATTERNS = [
    r'\d{4}[年/\-.]\d{1,2}[月/\-.]\d{1,2}日?', # 匹配 2023年12月31日, 2023/12/31, 2023-12-31 等
    r'\d{1,2}[月/\-.]\d{1,2}[日]?,\s*\d{4}', # 匹配 12/31, 2023, 12-31-2023 等
    r'\d{4}年\d{1,2}月', # 匹配 2023年12月
    r'\d{1,2}/\d{4}', # 匹配 12/2023
    r'(一|二|三|四|五|六|七|八|九|十|元)月', # 匹配中文月份
    r'(周一|周二|周三|周四|周五|周六|周日|星期一|星期二|星期三|星期四|星期五|星期六|星期日)' # 匹配星期
    
]


# 字段分隔符
FIELD_SEPARATORS = [',',';',' ', '、', '/', '|','。','\n','\t','，','：','；','\r']


#------------------目标表格----------------------------

# 最终Excel中所有列的顺序
EXCEL_HEADERS = [
    '厂商名称', '主营产品', '联系方式', '微信', '主销市场', '验厂/认证',
    '合作情况', '是否供样', '网址', '备注', '图片1', '图片2', '图片3',
    '图片4', '图片5', '数据来源', '厂商信息拼接', '图片描述',
    '拼接结果', '标签'
]

JSON_FORMAT = {
    '厂商名称': '',
    '主营产品': '',
    '联系方式': '',
    '微信': '',
    '主销市场': '',
    '验厂/认证': '',
    '合作情况': '',
    '是否供样': '',
    '网址': '',
    '备注': '',
    '日期': '',
    '图片文件夹路径': '',
    '文件路径': '',
    '附件': '',  
    
}





#----------------------------------------------文本解析的标签映射----------------------------------------------

#----------------------------------------------ppt文本解析的标签映射----------------------------------------------

# 用于ppt (期望的Excel列名 -> PPTX中可能出现的标签列表)
TEXT_LABELS_pptx = {
    '主营产品': ['Main Products','Main Products 主营产品','主营产品'],
    # '联系方式': ['联系方式', 'Tel:','Tel：', '电话:','电话','手机','手机：','手机号码：'], # 电话标签更重要，联系方式作为列名
    # '主销市场': ['市场占比', '主营市场','主销市场', '主要市场', 'Market Share', 'Main Markets'], # Found within '公司信息'
    '验厂/认证': ['Factory Audit Certification 验厂认证','验厂/认证','Factory Audit Certification','验厂认证'],
    '合作情况': ['Cooperative Customers 合作客户','Partner Company 合作公司','合作公司','合作客户','Cooperative Customers','Partner Company'],
    '备注': ['Company Information 公司信息','公司信息','Company Information'] # 用于定位相关细节的区域
}



#----------------------------------------------pdf文本解析的标签映射----------------------------------------------
# 字段映射模糊匹配 配置
TEXT_LABELS_pdf = {
    '厂商名称': ['工厂名称', '公司名称', '厂商名称'],
    '主营产品': ['主营产品', '主要产品'],
    '联系方式': ['经理', '联系人', '负责人','号码','地址','手机','电话'],
    '验厂/认证': ['认证', '验厂', '认证情况','证书'],
    '主销市场': ['主营业','主销','占比','市场','主销市场占比','市场占比'],
    '合作情况': ['合作客户',  '合作公司','合作情况'],
    '备注': ['面积','员工','年产值','人数','公司信息']
}

# 未分类文本到标准字段的映射关系
UNCLASSIFIED_TEXT_MAPPING_PDF = {
    'Main Products 主营产品': '主营产品',
    'Cooperative Customers 合作客户': '合作情况',
    'Partner Company 合作公司': '合作情况',
    'Factory Audit Certification 验厂认证': '验厂/认证',
    'Company Information 公司信息': '备注',


}

#-------------------------------------------------------------------------------
#用于word文本解析的标签映射 （完全匹配映射）
TEXT_LABELS_word={
    # '厂商名称': ['有限公司'],
    '主营产品': ['主营产品','主要产品','公司主营','主营','主管产品'],
    '主销市场': ['主销市场','主营市场'],
    '验厂/认证': ['验厂认证', '验厂', '认证', '证书', '验厂证书', '工厂证书', '认证证书', '验厂/认证', '出口认证'], 
    '联系方式':['QQ','qq','邮箱',"地址", '公司地址','厂商地址', '工厂地址'], # 电话和联系方式采取另一种方式映射
    '合作情况':['合作公司','合作客户','合作品牌', '合作'],
    '网址': ['官网'],
    '备注': ['员工人数','工厂面积', '企业信息','工厂面积','年产值','工厂信息','公司信息'],

}


#-------------------------------------------------------------------------------
#用于excel文本解析的标签映射 (包含多个工厂信息的工厂资料) (完全匹配映射)
TEXT_LABELS_excel_all_factory={
    '厂商名称': ['工厂名称',],
    '主营产品': ['主打产品','具体产品（范围）','主要产品名称'],
    '主销市场': ['适合市场'],
    '验厂/认证': ['验厂/认证'], 
    '是否供样': ['是否供样'],
    '联系方式':["联系方式"],
    '合作情况':['合作情况'],
    '网址': ['网站'],
    '备注': ['备注'],
    '日期': ['日期'],
    
}





#-------------------------------------------------------------------------------
# 用于xlsx_fty（工厂信息表）标签对应关系映射
TEXT_LABELS_xlsx_fty = {
    '厂商名称': ['factory_name'],
    '主营产品': ['product_category'],
    '联系方式': ['factory_contact','factory_phone','factory_address'], # 电话标签更重要，联系方式作为列名
    '主销市场': ['usa_share','eu_share', 'others_share','export_share','domestic_share','main_customer'], # Found within '公司信息'
    '验厂/认证': ['ISO9001','BSCI','Sedex','Disney FAMA','Walmart','Target','Other'],
    '是否供样':['sample_provided'],
    '合作情况': ['trade_company_cooperation', 'domestic_market_cooperation', 'group_cooperation'],
    '网址':['factory_website'],
    '备注': ['factory_mony','establish_time','annual_sales','employee_count','vat_invoice_count','factory_area','warehouse_area','dormitory_area','canteen_area','production_process','export_port','season_capacity','Production Facility_Name','Production Facility_Quantity','Production Facility_Age','Packing Facility_Name','Packing Facility_Quantity','Packing Facility_Age','Total production capacity per month','Used production capacity per month','Spare capacity per month'] ,# 用于定位相关细节的区域
    
    #'图片1', '图片2', '图片3', '图片4', '图片5' # 图片列将在后续处理中确定
}

# QUALITY_COMFORM=['✔','有','√','☑']

#----------------------------------------------xlsx_fty各字段的对应位置解析，保存数据信息----------------------------------------------


EXCEL_FORMATE_FTY_1 = {
    "factory_name": {
        "keyword_cell": "A4",
        "value_cell": "B4",
        "expected_keyword": "工厂名称"
    },
    "factory_contact": {
        "keyword_cell": "K4",
        "value_cell": "O4",
        "expected_keyword": "联系人",
        "prefix": "联系人："
    },
    "factory_address": {
        "keyword_cell": "A5",
        "value_cell": "B5",
        "expected_keyword": "工厂地址",
        "prefix": "工厂地址："
    },
    "factory_phone": {
        "keyword_cell": "K5",
        "value_cell": "O5",
        "expected_keyword": "手机",
        "prefix": "手机："
    },
    "factory_legal_representative": {
        "keyword_cell": "A6",
        "value_cell": "B6",
        "expected_keyword": "法人代表",
        "prefix": "法人代表："
    },
    "factory_mony": {
        "keyword_cell": "K6",
        "value_cell": "O6",
        "expected_keyword": "注册资金",
        "prefix": "注册资金："
    },
     "product_category": {
        "keyword_cell": "A7",
        "value_cell": "B7",
        "expected_keyword": "产品类别",
        
    },
    "establish_time": {
        "keyword_cell": "K7",
        "value_cell": "O7",
        "expected_keyword": "成立时间（年/月)",
        "prefix": "成立时间："
    },
    "annual_sales": {
        "keyword_cell": "A8",
        "value_cell": "B8",
        "expected_keyword": "年销售额",
        "prefix": "年销售额："
    },
    "employee_count": {
        "keyword_cell": "K8",
        "value_cell": "O8",
        "expected_keyword": "员工人数",
        "prefix": "员工人数："
    },
    "factory_website": {
        "keyword_cell": "A9",
        "value_cell": "B9",
        "expected_keyword": "工厂网站"
    },
    "vat_invoice_count": {
        "keyword_cell": "K9",
        "value_cell": "O9",
        "expected_keyword": "增值税开票点数",
        "prefix": "增值税开票点数："
    },
    

   "factory_area": {
        "keyword_cell": "B10",
        "value_cell": "B11",
        "expected_keyword": "厂房(办公区/车间)",
        "prefix": "厂房(办公区/车间)面积："
        },
    "warehouse_area": {
        "keyword_cell": "G10",
        "value_cell": "G11",
        "expected_keyword": "独立仓库(不适用,如与车间处于相同建筑)",
        "prefix": "独立仓库(不适用, 如与车间处于相同建筑)面积："
    },
    
    "dormitory_area": {
        "keyword_cell": "K10",
        "value_cell": "K11",
        "expected_keyword": "宿舍",
        "prefix": "宿舍面积："
    },
    "canteen_area": {
        "keyword_cell": "O10",
        "value_cell": "O11",
        "expected_keyword": "食堂",
        "prefix": "食堂面积："
    },
    
    "production_process": {
        "keyword_cell": "A12",
        "value_cell": "B12",
        "expected_keyword": "生产工艺",
        "prefix": "生产工艺："
    },
    "export_port": {
        "keyword_cell": "G12",
        "value_cell": "G12",
        "expected_keyword": "是否外发是□否□",
        "prefix": "是否外发："
    },
    "season_capacity": {
        "keyword_cell": "K12",
        "value_cell": "K12",
        "expected_keyword": "□淡季LowSeason□旺季PeakSeason",
        "prefix": "淡季旺季："
    },

    "main_customer": {
        "keyword_cell": "A13",
        "value_cell": "B13",
        "expected_keyword": "MainCustomer(nameandcountry)主要客户(名称及国家)",
        "prefix": "主要客户(名称及国家)："
    },
    
    "usa_share": {
        "keyword_cell": "K13",
        "value_cell": "M13",
        "expected_keyword": "USA%美国占比",
        "prefix": "美国占比："
    },
    "eu_share": {
        "keyword_cell": "K14",
        "value_cell": "M14",
        "expected_keyword": "EU%欧洲占比",
        "prefix": "欧洲占比："
    },
    "others_share": {
        "keyword_cell": "K17",
        "value_cell": "M17",
        "expected_keyword": "Others%其他占比",
        "prefix": "其他占比："
    },
    "domestic_share": {
        "keyword_cell": "O13",
        "value_cell": "R13",
        "expected_keyword": "Domestic%内销占比",
        "prefix": "内销占比："
    },
    "export_share": {
        "keyword_cell": "O15",
        "value_cell": "R15",
        "expected_keyword": "Export%外销占比",
        "prefix": "外销占比："
    },
    "ISO9001":{
        "keyword_cell": "B19",
        "value_cell": "B20",
        "expected_keyword": "ISO9001"
    },
    "BSCI":{
        "keyword_cell": "D19",
        "value_cell": "D20",
        "expected_keyword": "BSCI"
    },
    "Sedex":{
        "keyword_cell": "G19",
        "value_cell": "G20",
        "expected_keyword": "Sedex"
    },
    "Disney FAMA":{
        "keyword_cell": "I19",
        "value_cell": "I20",
        "expected_keyword": "DisneyFAMA"
    },
    "Walmart":{
        "keyword_cell": "K19",
        "value_cell": "K20",
        "expected_keyword": "Wal-mart"
    },
    "Target":{
        "keyword_cell": "M19",
        "value_cell": "M20",
        "expected_keyword": "Target"
    },
    "Other":{
        "keyword_cell": "O19",
        "value_cell": "O20",
        "expected_keyword": "Others"
    },
    "More_Certificates":{
        "keyword_cell": "A21",
        "value_cell": "G21",
        "expected_keyword": "其他认证证书或客户审核,有效时间"
    },
    "Production Facility_Name":{
        "keyword_cell": "G31",
        "value_cell": "J31",
        "expected_keyword": "Name名称",
        "prefix": "生产设备名称："
    },
    "Production Facility_Quantity":{
        "keyword_cell": "G32",
        "value_cell": "J32",
        "expected_keyword": "Quantity数量",
        "prefix": "生产设备数量："
    },
    "Production Facility_Age":{
        "keyword_cell": "G33",
        "value_cell": "J33",
        "expected_keyword": "Age使用年限",
        "prefix": "生产设备使用年限："
    },
    "Packing Facility_Name":{
        "keyword_cell": "G34",
        "value_cell": "J34",
        "expected_keyword": "Name名称",
        "prefix": "包装设备名称："
    },
    "Packing Facility_Quantity":{
        "keyword_cell": "G35",
        "value_cell": "J35",
        "expected_keyword": "Quantity数量",
        "prefix": "包装设备数量："
    },
    "Packing Facility_Age":{
        "keyword_cell": "G36",
        "value_cell": "J36",
        "expected_keyword": "Age使用年限",
        "prefix": "包装设备使用年限："
    },
    "Total production capacity per month":{
        "keyword_cell": "A37",
        "value_cell": "B37",
        "expected_keyword": "Totalproductioncapacitypermonth总产能/月",
        "prefix": "总产能/月："
    },
    "Used production capacity per month":{
        "keyword_cell": "G37",
        "value_cell": "J37",
        "expected_keyword": "Usedproductioncapacitypermonth当前已使用产能/月",
        "prefix": "当前已使用产能/月："
    },
    "Spare capacity per month":{
        "keyword_cell": "M37",
        "value_cell": "Q37",
        "expected_keyword": "Sparecapacitypermonth剩余可用产能/月",
        "prefix": "剩余可用产能/月："
    },



    "trade_company_cooperation": {
        "keyword_cell": "A38",
        "value_cell": "B38",
        "expected_keyword": "合作的贸易公司及合作情况",
        "prefix": "合作的贸易公司及合作情况："
    },
    "domestic_market_cooperation": {
        "keyword_cell": "A39",
        "value_cell": "B39",
        "expected_keyword": "合作的内销市场及合作情况",
        "prefix": "合作的内销市场及合作情况："
    },
    "group_cooperation": {
        "keyword_cell": "A40",
        "value_cell": "B40",
        "expected_keyword": "有无与本集团合作及合作情况",
        "prefix": "有无与本集团合作及合作情况："
    },
    "sample_provided": {
        "keyword_cell": "A41",
        "value_cell": "B41",
        "expected_keyword": "是否可以提供样品"
    },
    "sample_date": {
        "keyword_cell": "M41",
        "value_cell": "Q41",
        "expected_keyword": "Date供样日期"
    }

}



EXCEL_FORMATE_FTY_2 = {
    "factory_name": {
        "keyword_cell": "A4",
        "value_cell": "B4",
        "expected_keyword": "工厂名称"
    },
    "factory_contact": {
        "keyword_cell": "K4",
        "value_cell": "O4",
        "expected_keyword": "联系人",
        "prefix": "联系人："
    },
    "factory_address": {
        "keyword_cell": "A5",
        "value_cell": "B5",
        "expected_keyword": "工厂地址",
        "prefix": "工厂地址："
    },
    "factory_phone": {
        "keyword_cell": "K5",
        "value_cell": "O5",
        "expected_keyword": "手机",
        "prefix": "手机："
    },
    "factory_legal_representative": {
        "keyword_cell": "A7",
        "value_cell": "B7",
        "expected_keyword": "法人代表",
        "prefix": "法人代表："
    },
    "factory_mony": {
        "keyword_cell": "K7",
        "value_cell": "O7",
        "expected_keyword": "注册资金",
        "prefix": "注册资金："
    },
     "product_category": {
        "keyword_cell": "A9",
        "value_cell": "B9",
        "expected_keyword": "产品类别",
        
    },
    "establish_time": {
        "keyword_cell": "K9",
        "value_cell": "O9",
        "expected_keyword": "成立时间（年/月)",
        "prefix": "成立时间："
    },
    "annual_sales": {
        "keyword_cell": "A10",
        "value_cell": "B10",
        "expected_keyword": "年销售额",
        "prefix": "年销售额："
    },
    "employee_count": {
        "keyword_cell": "K10",
        "value_cell": "O10",
        "expected_keyword": "员工人数",
        "prefix": "员工人数："
    },
    "factory_website": {
        "keyword_cell": "A12",
        "value_cell": "B12",
        "expected_keyword": "工厂网站"
    },
    "vat_invoice_count": {
        "keyword_cell": "K12",
        "value_cell": "O12",
        "expected_keyword": "增值税开票点数",
        "prefix": "增值税开票点数："
    },
    

   "factory_area": {
        "keyword_cell": "B14",
        "value_cell": "B15",
        "expected_keyword": "厂房(办公区/车间)",
        "prefix": "厂房(办公区/车间)面积："
    },
    "warehouse_area": {
        "keyword_cell": "G14",
        "value_cell": "G15",
        "expected_keyword": "独立仓库(不适用,如与车间处于相同建筑)",
        "prefix": "独立仓库(不适用, 如与车间处于相同建筑)面积："
    },
    
    "dormitory_area": {
        "keyword_cell": "K14",
        "value_cell": "K15",
        "expected_keyword": "宿舍",
        "prefix": "宿舍面积："
    },
    "canteen_area": {
        "keyword_cell": "O14",
        "value_cell": "O15",
        "expected_keyword": "食堂",
        "prefix": "食堂面积："
    },
    
    "production_process": {
        "keyword_cell": "A16",
        "value_cell": "B16",
        "expected_keyword": "生产工艺",
        "prefix": "生产工艺："
    },
    "export_port": {
        "keyword_cell": "G16",
        "value_cell": "G16",
        "expected_keyword": "是否外发是□否□",
        "prefix": "是否外发："
    },
    "season_capacity": {
        "keyword_cell": "K16",
        "value_cell": "K16",
        "expected_keyword": "□淡季LowSeason□旺季PeakSeason",
        "prefix": "淡季旺季："
    },

    "main_customer": {
        "keyword_cell": "A17",
        "value_cell": "B17",
        "expected_keyword": "MainCustomer(nameandcountry)主要客户(名称及国家)",
        "prefix": "主要客户(名称及国家)："
    },
    
    "usa_share": {
        "keyword_cell": "K17",
        "value_cell": "M17",
        "expected_keyword": "USA%美国占比",
        "prefix": "美国占比："
    },
    "eu_share": {
        "keyword_cell": "K18",
        "value_cell": "M18",
        "expected_keyword": "EU%欧洲占比",
        "prefix": "欧洲占比："
    },
    "others_share": {
        "keyword_cell": "K21",
        "value_cell": "M21",
        "expected_keyword": "Others%其他占比",
        "prefix": "其他占比："
    },
    "domestic_share": {
        "keyword_cell": "O17",
        "value_cell": "R17",
        "expected_keyword": "Domestic%内销占比",
        "prefix": "内销占比："
    },
    "export_share": {
        "keyword_cell": "O19",
        "value_cell": "R19",
        "expected_keyword": "Export%外销占比",
        "prefix": "外销占比："
    },
    "ISO9001":{
        "keyword_cell": "B23",
        "value_cell": "B24",
        "expected_keyword": "ISO9001"
    },
    "BSCI":{
        "keyword_cell": "D23",
        "value_cell": "D24",
        "expected_keyword": "BSCI"
    },
    "Sedex":{
        "keyword_cell": "G23",
        "value_cell": "G24",
        "expected_keyword": "Sedex"
    },
    "Disney FAMA":{
        "keyword_cell": "I23",
        "value_cell": "I24",
        "expected_keyword": "DisneyFAMA"
    },
    "Walmart":{
        "keyword_cell": "K23",
        "value_cell": "K24",
        "expected_keyword": "Wal-mart"
    },
    "Target":{
        "keyword_cell": "M23",
        "value_cell": "M24",
        "expected_keyword": "Target"
    },
    "Other":{
        "keyword_cell": "O23",
        "value_cell": "O24",
        "expected_keyword": "Others"
    },

    "More_Certificates":{
        "keyword_cell": "A25",
        "value_cell": "G23",
        "expected_keyword": "其他认证证书或客户审核,有效时间"
    },
    "Production Facility_Name":{
        "keyword_cell": "Y3",
        "value_cell": "AB3",
        "expected_keyword": "Name名称",
        "prefix": "生产设备名称："
    },
    "Production Facility_Quantity":{
        "keyword_cell": "Y4",
        "value_cell": "AB4",
        "expected_keyword": "Quantity数量",
        "prefix": "生产设备数量："
    },
    "Production Facility_Age":{
        "keyword_cell": "Y5",
        "value_cell": "AB5",
        "expected_keyword": "Age使用年限",
        "prefix": "生产设备使用年限："
    },
    "Packing Facility_Name":{
        "keyword_cell": "Y7",
        "value_cell": "AB7",
        "expected_keyword": "Name名称",
        "prefix": "包装设备名称："
    },
    "Packing Facility_Quantity":{
        "keyword_cell": "Y9",
        "value_cell": "AB9",
        "expected_keyword": "Quantity数量",
        "prefix": "包装设备数量："
    },
    "Packing Facility_Age":{
        "keyword_cell": "Y11",
        "value_cell": "AB11",
        "expected_keyword": "Age使用年限",
        "prefix": "包装设备使用年限："
    },
    "Total production capacity per month":{
        "keyword_cell": "S14",
        "value_cell": "T14",
        "expected_keyword": "Totalproductioncapacitypermonth总产能/月",
        "prefix": "总产能/月："
    },
    "Used production capacity per month":{
        "keyword_cell": "Y14",
        "value_cell": "AB14",
        "expected_keyword": "Usedproductioncapacitypermonth当前已使用产能/月",
        "prefix": "当前已使用产能/月："
    },
    "Spare capacity per month":{
        "keyword_cell": "AE14",
        "value_cell": "AI14",
        "expected_keyword": "Sparecapacitypermonth剩余可用产能/月",
        "prefix": "剩余可用产能/月："
    },

    "trade_company_cooperation": {
        "keyword_cell": "S16",
        "value_cell": "T16",
        "expected_keyword": "合作的贸易公司及合作情况",
        "prefix": "合作的贸易公司及合作情况："
    },
    "domestic_market_cooperation": {
        "keyword_cell": "S17",
        "value_cell": "T17",
        "expected_keyword": "合作的内销市场及合作情况",
        "prefix": "合作的内销市场及合作情况："
    },
    "group_cooperation": {
        "keyword_cell": "S19",
        "value_cell": "T19",
        "expected_keyword": "有无与本集团合作及合作情况",
        "prefix": "有无与本集团合作及合作情况："
    },
    "sample_provided": {
        "keyword_cell": "S23",
        "value_cell": "T23",
        "expected_keyword": "是否可以提供样品"
    },
    "sample_date": {
        "keyword_cell": "AE23",
        "value_cell": "AI23",
        "expected_keyword": "Date供样日期"
    }


}



