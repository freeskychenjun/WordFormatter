import logging
import os
from datetime import datetime
from enum import Enum


class LogLevel(Enum):
    """日志级别枚举"""
    DEBUG = logging.DEBUG
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    CRITICAL = logging.CRITICAL


class Logger:
    """增强的日志记录器"""
    
    def __init__(self, name="WordFormatter", log_file=None, level=LogLevel.INFO):
        """
        初始化日志记录器
        
        Args:
            name (str): 日志记录器名称
            log_file (str): 日志文件路径，如果为None则只输出到控制台
            level (LogLevel): 日志级别
        """
        self.logger = logging.getLogger(name)
        self.logger.setLevel(level.value)
        
        # 避免重复添加处理器
        if not self.logger.handlers:
            # 创建格式化器
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            
            # 创建控制台处理器
            console_handler = logging.StreamHandler()
            console_handler.setLevel(level.value)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
            
            # 如果指定了日志文件，创建文件处理器
            if log_file:
                # 确保日志目录存在
                log_dir = os.path.dirname(log_file)
                if log_dir and not os.path.exists(log_dir):
                    os.makedirs(log_dir)
                
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setLevel(level.value)
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
    
    def debug(self, message):
        """记录调试信息"""
        self.logger.debug(message)
    
    def info(self, message):
        """记录一般信息"""
        self.logger.info(message)
    
    def warning(self, message):
        """记录警告信息"""
        self.logger.warning(message)
    
    def error(self, message):
        """记录错误信息"""
        self.logger.error(message)
    
    def critical(self, message):
        """记录严重错误信息"""
        self.logger.critical(message)
    
    def log_progress(self, current, total, description="处理进度"):
        """
        记录进度信息
        
        Args:
            current (int): 当前进度
            total (int): 总数
            description (str): 进度描述
        """
        percentage = (current / total) * 100 if total > 0 else 0
        self.info(f"{description}: {current}/{total} ({percentage:.1f}%)")
    
    def log_exception(self, exception, message="发生异常"):
        """
        记录异常信息
        
        Args:
            exception (Exception): 异常对象
            message (str): 额外的描述信息
        """
        self.logger.exception(f"{message}: {str(exception)}")
    
    def log_file_operation(self, operation, file_path, status="开始"):
        """
        记录文件操作日志
        
        Args:
            operation (str): 操作类型（如"读取"、"写入"、"删除"等）
            file_path (str): 文件路径
            status (str): 操作状态（"开始"、"完成"、"失败"等）
        """
        self.info(f"文件{operation} {status}: {file_path}")
    
    def log_document_processing(self, doc_name, operation, details=""):
        """
        记录文档处理日志
        
        Args:
            doc_name (str): 文档名称
            operation (str): 操作类型
            details (str): 详细信息
        """
        msg = f"文档 '{doc_name}' {operation}"
        if details:
            msg += f" - {details}"
        self.info(msg)


class LogCollector:
    """日志收集器，用于收集GUI界面中的日志信息"""
    
    def __init__(self):
        self.logs = []
    
    def add_log(self, level, message):
        """
        添加日志条目
        
        Args:
            level (str): 日志级别
            message (str): 日志消息
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = {
            'timestamp': timestamp,
            'level': level,
            'message': message
        }
        self.logs.append(log_entry)
    
    def get_logs(self):
        """
        获取所有日志条目
        
        Returns:
            list: 日志条目列表
        """
        return self.logs.copy()
    
    def clear_logs(self):
        """清空日志"""
        self.logs.clear()
    
    def get_logs_by_level(self, level):
        """
        根据级别获取日志条目
        
        Args:
            level (str): 日志级别
            
        Returns:
            list: 指定级别的日志条目列表
        """
        return [log for log in self.logs if log['level'] == level]
    
    def save_to_file(self, file_path):
        """
        将日志保存到文件
        
        Args:
            file_path (str): 文件路径
        """
        with open(file_path, 'w', encoding='utf-8') as f:
            for log in self.logs:
                f.write(f"[{log['timestamp']}] {log['level']}: {log['message']}\n")


# 全局日志记录器实例
global_logger = Logger("WordFormatter")


def get_logger(name="WordFormatter"):
    """
    获取日志记录器实例
    
    Args:
        name (str): 日志记录器名称
        
    Returns:
        Logger: 日志记录器实例
    """
    return Logger(name)