import json
import os
import logging


class ConfigManager:
    """配置管理器，用于处理应用程序的配置加载、保存和验证"""
    
    def __init__(self, default_config_path="default_config.json", update_config_path="update_config.json"):
        self.default_config_path = default_config_path
        self.update_config_path = update_config_path
        self.logger = logging.getLogger(__name__)
        
        # 存储加载后的配置
        self.format_config = None
        self.update_config = None
        
        # 默认排版配置参数
        self.default_format_params = {
            'page_number_align': '奇偶分页', 'line_spacing': 28,
            'margin_top': 3.7, 'margin_bottom': 3.5, 
            'margin_left': 2.8, 'margin_right': 2.6,
            'h1_font': '黑体', 'h2_font': '楷体_GB2312', 'h3_font': '宋体', 'body_font': '仿宋_GB2312',
            'page_number_font': '宋体', 'table_caption_font': '黑体', 'figure_caption_font': '黑体',
            'h1_size': 16, 'h1_space_before': 24, 'h1_space_after': 24,
            'h2_size': 16, 'h2_space_before': 24, 'h2_space_after': 24,
            'h3_size': 12, 'h3_space_before': 24, 'h3_space_after': 24,
            'body_size': 16, 'page_number_size': 14,
            'table_caption_size': 14, 'figure_caption_size': 14,
            # 添加表格标题和图表标题的大纲级别设置，默认为6级
            'table_caption_outline_level': 8, 'figure_caption_outline_level': 6,
            'set_outline': True,
            # 添加标题粗体设置
            'h1_bold': False,  # 一级标题默认不加粗
            'h2_bold': True,   # 二级标题默认加粗
            'h3_bold': False,  # 三级标题默认不加粗
            'table_caption_bold': False,  # 表格标题默认不加粗
            'figure_caption_bold': False,  # 图形标题默认不加粗
            'body_use_times_roman': True,  # 正文默认使用Times New Roman
            'table_use_times_roman': True  # 表格默认使用Times New Roman
        }
        
        # 默认自动更新配置参数
        self.default_update_params = {
            'auto_update': True,  # 默认启用自动更新
            'update_check_url': 'http://172.14.60.197/update.xml'  # 默认更新检查地址
        }
        
        # 合并所有默认参数（用于兼容旧代码）
        self.default_params = {
            **self.default_format_params,
            **self.default_update_params
        }
        
        # 字体选项
        self.font_options = {
            'h1': ['黑体', '方正黑体_GBK', '方正黑体简体', '华文黑体', '宋体', '仿宋', '仿宋_GB2312'],
            'h2': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体', '宋体', '仿宋', '仿宋_GB2312'],
            'h3': ['宋体', '仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋'],
            'body': ['仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋', '宋体'], 
            'table_caption': ['黑体', '宋体', '仿宋_GB2312', '仿宋'], 
            'figure_caption': ['黑体', '宋体', '仿宋_GB2312', '仿宋']
        }
        
        # 字号映射
        self.font_size_map = {
            '一号 (26pt)': 26, '小一 (24pt)': 24, '二号 (22pt)': 22, '小二 (18pt)': 18,
            '三号 (16pt)': 16, '小三 (15pt)': 15, '四号 (14pt)': 14, '小四 (12pt)': 12,
            '五号 (10.5pt)': 10.5, '小五 (9pt)': 9
        }
        self.font_size_map_rev = {v: k for k, v in self.font_size_map.items()}

    def load_config(self, config_path=None):
        """
        加载排版配置文件
        
        Args:
            config_path (str): 配置文件路径，如果为None则使用默认路径
            
        Returns:
            dict: 排版配置字典
        """
        if config_path is None:
            config_path = self.default_config_path
            
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.logger.info(f"排版配置文件 '{config_path}' 加载成功")
                self.format_config = self._validate_format_config(config)
                return self.format_config
            except Exception as e:
                self.logger.error(f"加载排版配置文件 '{config_path}' 失败: {e}")
                self.format_config = self.default_format_params.copy()
                return self.format_config
        else:
            self.logger.warning(f"排版配置文件 '{config_path}' 不存在，使用默认配置")
            self.format_config = self.default_format_params.copy()
            return self.format_config

    def load_update_config(self, config_path=None):
        """
        加载自动更新配置文件
        
        Args:
            config_path (str): 更新配置文件路径，如果为None则使用默认路径
            
        Returns:
            dict: 自动更新配置字典
        """
        if config_path is None:
            config_path = self.update_config_path
            
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.logger.info(f"自动更新配置文件 '{config_path}' 加载成功")
                self.update_config = self._validate_update_config(config)
                return self.update_config
            except Exception as e:
                self.logger.error(f"加载自动更新配置文件 '{config_path}' 失败: {e}")
                self.update_config = self.default_update_params.copy()
                return self.update_config
        else:
            self.logger.warning(f"自动更新配置文件 '{config_path}' 不存在，使用默认配置")
            # 注意：这里不设置默认配置，而是返回None，让调用者知道文件不存在
            self.update_config = None
            return None

    def save_config(self, config, config_path=None):
        """
        保存排版配置到文件
        
        Args:
            config (dict): 排版配置字典
            config_path (str): 配置文件路径，如果为None则使用默认路径
        """
        if config_path is None:
            config_path = self.default_config_path
            
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            self.logger.info(f"排版配置已保存到 '{config_path}'")
        except Exception as e:
            self.logger.error(f"保存排版配置到 '{config_path}' 失败: {e}")
            raise

    def save_update_config(self, config, config_path=None):
        """
        保存自动更新配置到文件
        
        Args:
            config (dict): 自动更新配置字典
            config_path (str): 更新配置文件路径，如果为None则使用默认路径
        """
        if config_path is None:
            config_path = self.update_config_path
            
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            self.logger.info(f"自动更新配置已保存到 '{config_path}'")
        except Exception as e:
            self.logger.error(f"保存自动更新配置到 '{config_path}' 失败: {e}")
            raise

    def _validate_format_config(self, config):
        """
        验证排版配置参数的有效性
        
        Args:
            config (dict): 排版配置字典
            
        Returns:
            dict: 验证后的排版配置字典
        """
        validated_config = self.default_format_params.copy()
        
        for key, value in config.items():
            if key in validated_config:
                # 验证数值类型参数
                if key in ['line_spacing', 'margin_top', 'margin_bottom', 'margin_left', 'margin_right',
                          'h1_size', 'h1_space_before', 'h1_space_after', 'h2_size', 'h2_space_before', 
                          'h2_space_after', 'h3_size', 'h3_space_before', 'h3_space_after', 'body_size', 
                          'page_number_size', 'table_caption_size', 'figure_caption_size']:
                    try:
                        validated_config[key] = float(value)
                    except (ValueError, TypeError):
                        self.logger.warning(f"无效的数值参数 '{key}': {value}，使用默认值")
                # 验证布尔类型参数
                elif key in ['set_outline', 'h1_bold', 'h2_bold', 'h3_bold', 'table_caption_bold', 'figure_caption_bold', 'body_use_times_roman', 'table_use_times_roman']:
                    validated_config[key] = bool(value)
                # 验证大纲级别参数
                elif key in ['table_caption_outline_level', 'figure_caption_outline_level']:
                    if value == '无' or value == '':
                        validated_config[key] = value
                    else:
                        try:
                            level = int(value)
                            if 1 <= level <= 9:
                                validated_config[key] = level
                            else:
                                self.logger.warning(f"无效的大纲级别 '{key}': {value}，使用默认值")
                        except (ValueError, TypeError):
                            self.logger.warning(f"无效的大纲级别 '{key}': {value}，使用默认值")
                # 其他参数直接使用
                else:
                    validated_config[key] = value
                    
        return validated_config

    def _validate_update_config(self, config):
        """
        验证自动更新配置参数的有效性
        
        Args:
            config (dict): 自动更新配置字典
            
        Returns:
            dict: 验证后的自动更新配置字典
        """
        validated_config = self.default_update_params.copy()
        
        for key, value in config.items():
            if key in validated_config:
                # 验证布尔类型参数
                if key == 'auto_update':
                    validated_config[key] = bool(value)
                # 更新检查URL
                elif key == 'update_check_url':
                    validated_config[key] = str(value)
                    
        return validated_config

    def _validate_config(self, config):
        """
        验证所有配置参数的有效性（兼容旧代码）
        
        Args:
            config (dict): 配置字典
            
        Returns:
            dict: 验证后的配置字典
        """
        # 分别验证排版配置和更新配置
        format_config = {k: v for k, v in config.items() if k in self.default_format_params}
        update_config = {k: v for k, v in config.items() if k in self.default_update_params}
        
        validated_format_config = self._validate_format_config(format_config)
        validated_update_config = self._validate_update_config(update_config)
        
        # 合并验证后的配置（兼容旧代码）
        return {
            **validated_format_config,
            **validated_update_config
        }

    def get_font_options(self, font_type):
        """
        获取指定类型的字体选项
        
        Args:
            font_type (str): 字体类型 ('h1', 'h2', 'h3', 'body', 'table_caption', 'figure_caption')
            
        Returns:
            list: 字体选项列表
        """
        return self.font_options.get(font_type, [])

    def get_font_size_options(self):
        """
        获取字号选项
        
        Returns:
            list: 字号选项列表
        """
        return list(self.font_size_map.keys())

    def font_size_to_pt(self, size_key):
        """
        将字号键转换为磅值
        
        Args:
            size_key (str): 字号键（如 '三号 (16pt)'）
            
        Returns:
            float: 字号磅值
        """
        return self.font_size_map.get(size_key, 16)  # 默认16pt

    def pt_to_font_size(self, pt_value):
        """
        将磅值转换为字号键
        
        Args:
            pt_value (float): 字号磅值
            
        Returns:
            str: 字号键
        """
        return self.font_size_map_rev.get(pt_value, '三号 (16pt)')  # 默认三号

    def get_default_config(self):
        """
        获取默认配置的副本（兼容旧代码）
        
        Returns:
            dict: 默认配置字典的副本
        """
        return self.default_params.copy()

    def get_default_format_config(self):
        """
        获取默认排版配置的副本
        
        Returns:
            dict: 默认排版配置字典的副本
        """
        return self.default_format_params.copy()

    def get_default_update_config(self):
        """
        获取默认自动更新配置的副本
        
        Returns:
            dict: 默认自动更新配置字典的副本
        """
        return self.default_update_params.copy()