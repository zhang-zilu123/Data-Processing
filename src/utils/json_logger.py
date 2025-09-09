import logging
import json
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler

class JsonFormatter(logging.Formatter):
    """自定义JSON格式的日志格式化器"""
    
    def format(self, record):
        """将日志记录转换为JSON格式"""
        log_data = {
            'timestamp': datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S'),
            'level': record.levelname,
            'message': record.getMessage(),
            'module': record.module,
            'function': record.funcName,
            'line': record.lineno
        }
        
        # 如果有异常信息，添加到日志中
        if record.exc_info:
            log_data['exception'] = self.formatException(record.exc_info)
            
        return json.dumps(log_data, ensure_ascii=False)

def setup_json_logger(log_dir='logs', log_file='app.log', max_bytes=10*1024*1024, backup_count=5):
    """
    设置JSON格式的日志处理器
    
    Args:
        log_dir (str): 日志文件目录
        log_file (str): 日志文件名
        max_bytes (int): 单个日志文件的最大大小（字节）
        backup_count (int): 保留的备份文件数量
    """
    # 确保日志目录存在
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    # 创建日志文件路径
    log_path = os.path.join(log_dir, log_file)
    
    # 创建日志处理器
    handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    
    # 设置JSON格式化器
    handler.setFormatter(JsonFormatter())
    
    # 获取根日志记录器
    logger = logging.getLogger()
    
    # 设置日志级别
    logger.setLevel(logging.INFO)
    
    # 添加处理器
    logger.addHandler(handler)
    
    return logger

if __name__ == "__main__":
    # 测试日志记录
    logger = setup_json_logger()
    logger.info("这是一条测试信息")
    logger.error("这是一条错误信息")
    try:
        1/0
    except Exception as e:
        logger.exception("发生异常") 