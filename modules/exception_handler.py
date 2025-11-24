import logging
import traceback
from functools import wraps


class WordFormatterError(Exception):
    """报告自动排版工具的基类异常"""
    pass


class FileProcessingError(WordFormatterError):
    """文件处理相关的异常"""
    def __init__(self, message, original_error=None):
        super().__init__(message)
        self.original_error = original_error


class DocumentFormatError(WordFormatterError):
    """文档格式化相关的异常"""
    def __init__(self, message, original_error=None):
        super().__init__(message)
        self.original_error = original_error


class ApplicationError(WordFormatterError):
    """应用程序相关的异常"""
    def __init__(self, message, original_error=None):
        super().__init__(message)
        self.original_error = original_error


class ConfigError(WordFormatterError):
    """配置相关的异常"""
    def __init__(self, message, original_error=None):
        super().__init__(message)
        self.original_error = original_error


class ExceptionHandler:
    """异常处理器，用于统一处理应用程序中的异常"""
    
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
    
    def handle_exception(self, exception, context=""):
        """
        处理异常并记录日志
        
        Args:
            exception (Exception): 捕获的异常
            context (str): 异常发生的上下文信息
            
        Returns:
            str: 用户友好的错误消息
        """
        error_msg = f"在{context}中发生错误: {str(exception)}"
        
        # 记录详细错误信息
        self.logger.error(error_msg)
        self.logger.debug(f"异常详情: {traceback.format_exc()}")
        
        # 根据异常类型返回用户友好的消息
        if isinstance(exception, FileProcessingError):
            return f"文件处理错误: {str(exception)}"
        elif isinstance(exception, DocumentFormatError):
            return f"文档格式化错误: {str(exception)}"
        elif isinstance(exception, ApplicationError):
            return f"应用程序错误: {str(exception)}"
        elif isinstance(exception, ConfigError):
            return f"配置错误: {str(exception)}"
        elif isinstance(exception, FileNotFoundError):
            return f"文件未找到: {str(exception)}"
        elif isinstance(exception, PermissionError):
            return f"权限不足: {str(exception)}"
        elif isinstance(exception, ValueError):
            return f"参数错误: {str(exception)}"
        else:
            return f"发生未知错误: {str(exception)}"
    
    def safe_execute(self, func, *args, **kwargs):
        """
        安全执行函数，捕获并处理异常
        
        Args:
            func (callable): 要执行的函数
            *args: 函数参数
            **kwargs: 函数关键字参数
            
        Returns:
            tuple: (success, result_or_error_message)
        """
        try:
            result = func(*args, **kwargs)
            return True, result
        except Exception as e:
            error_msg = self.handle_exception(e, func.__name__)
            return False, error_msg


def handle_exceptions(logger=None):
    """
    装饰器：用于自动处理函数中的异常
    
    Args:
        logger (logging.Logger): 日志记录器
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            handler = ExceptionHandler(logger)
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_msg = handler.handle_exception(e, func.__name__)
                # 重新抛出异常，或者根据需要返回默认值
                raise type(e)(error_msg) from e
        return wrapper
    return decorator


# 全局异常处理器实例
global_exception_handler = ExceptionHandler()