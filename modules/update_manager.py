import json
import os
import requests
import logging
import subprocess
import tempfile
import shutil
from datetime import datetime, timedelta
import sys
import zipfile


class UpdateManager:
    """更新管理器，用于处理应用程序的自动更新"""
    
    def __init__(self, config, log_callback):
        """
        初始化更新管理器
        
        Args:
            config (dict): 应用程序配置
            log_callback (callable): 日志回调函数
        """
        self.config = config
        self.log_callback = log_callback
        self.current_version = "1.0.3"  # 当前应用版本
        self.update_check_url = self.config.get('update_check_url')  # 从配置中获取更新检查地址
        self.last_check_time = None
        self.logger = logging.getLogger(__name__)
        
        # 从配置中获取自动更新设置
        self.auto_update = config.get('auto_update', True)
        # 删除定期检查更新的间隔设置，因为我们只需要启动时检查一次
    
    def check_for_updates(self):
        """
        检查是否有可用更新（仅在程序启动时调用一次）
        
        Returns:
            bool: 是否有更新可用
        """
        if not self.auto_update:
            self.log_callback("自动更新已关闭")
            return False
        
        self.log_callback("正在检查更新...")
        
        try:
            # 发送请求获取最新版本信息
            response = requests.get(self.update_check_url, timeout=10, verify=False)  # 跳过SSL验证
            response.raise_for_status()
            
            # 确保响应内容使用正确的编码
            response.encoding = 'utf-8'
            
            # 解析XML格式的更新信息
            import xml.etree.ElementTree as ET
            root = ET.fromstring(response.text)
            
            # 提取版本号和更新信息，使用更安全的方式
            version_element = root.find('version')
            url_element = root.find('url')
            notes_element = root.find('notes')
            
            latest_version = version_element.text.strip() if version_element is not None and version_element.text else '1.0.0'
            update_url = url_element.text.strip() if url_element is not None and url_element.text else ''
            update_notes = notes_element.text.strip() if notes_element is not None and notes_element.text else '无'
            
            # 清理URL中的特殊字符（如反引号）
            update_url = update_url.replace('`', '')
            
            self.last_check_time = datetime.now()
            
            # 比较版本号
            if self._is_newer_version(latest_version, self.current_version):
                self.log_callback(f"发现新版本: v{latest_version}")
                self.log_callback(f"更新说明: {update_notes}")
                
                # 构建release_info字典，包含下载URL和版本信息
                # 清理URL并从URL中提取文件名，而不是硬编码
                import os
                # 确保URL有效
                if not update_url.startswith(('http://', 'https://')):
                    self.log_callback(f"无效的更新URL: {update_url}")
                    return False, self.current_version, None
                    
                filename = os.path.basename(update_url)
                
                release_info = {
                    'tag_name': f'v{latest_version}',
                    'body': update_notes,
                    'assets': [{
                        'name': filename,
                        'browser_download_url': update_url
                    }]
                }
                
                return True, latest_version, release_info
            else:
                self.log_callback(f"当前已是最新版本: V{self.current_version}")
                return False, self.current_version, None
        
        except requests.exceptions.RequestException as e:
            self.log_callback(f"更新检查失败: {e}")
            self.logger.error(f"更新检查失败: {e}")
            return False, self.current_version, None
        except Exception as e:
            self.log_callback(f"更新检查时发生错误: {e}")
            self.logger.error(f"更新检查时发生错误: {e}", exc_info=True)
            return False, self.current_version, None
    
    def download_update(self, release_info):
        """
        下载更新包
        
        Args:
            release_info (dict): 更新信息字典
        
        Returns:
            str: 下载的更新包路径
        """
        try:
            # 查找更新包（支持exe和zip格式）
            assets = release_info.get('assets', [])
            update_asset = None
            
            for asset in assets:
                asset_name = asset.get('name', '')
                if asset_name.endswith(('.exe', '.zip')):
                    update_asset = asset
                    break
            
            if not update_asset:
                self.log_callback("未找到适合的更新包")
                return None
            
            download_url = update_asset.get('browser_download_url')
            if not download_url:
                self.log_callback("更新包下载地址无效")
                return None
            
            self.log_callback(f"开始下载更新包: {update_asset.get('name')}")
            
            # 创建临时文件保存更新包
            temp_dir = tempfile.gettempdir()
            update_file_path = os.path.join(temp_dir, update_asset.get('name'))
            
            # 下载更新包
            response = requests.get(download_url, stream=True, timeout=30, verify=False)  # 跳过SSL验证
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded_size = 0
            
            with open(update_file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        
                        # 计算下载进度
                        if total_size > 0:
                            progress = (downloaded_size / total_size) * 100
                            self.log_callback(f"下载进度: {progress:.1f}%")
            
            self.log_callback(f"更新包下载完成: {update_file_path}")
            return update_file_path
            
        except requests.exceptions.RequestException as e:
            self.log_callback(f"更新包下载失败: {e}")
            self.logger.error(f"更新包下载失败: {e}")
            return None
        except Exception as e:
            self.log_callback(f"下载更新包时发生错误: {e}")
            self.logger.error(f"下载更新包时发生错误: {e}", exc_info=True)
            return None
    
    def install_update(self, update_file_path):
        """
        安装更新
        
        Args:
            update_file_path (str): 更新包路径
        """
        try:
            self.log_callback(f"开始安装更新: {update_file_path}")
            
            # 根据文件类型选择安装方式
            if update_file_path.endswith('.exe'):
                # 启动exe更新程序
                subprocess.Popen([update_file_path, '--update'], shell=True)
            elif update_file_path.endswith('.zip'):
                # 处理zip格式更新包
                self._install_zip_update(update_file_path)
            
            # 关闭当前程序
            sys.exit(0)
            
        except Exception as e:
            self.log_callback(f"安装更新时发生错误: {e}")
            self.logger.error(f"安装更新时发生错误: {e}", exc_info=True)
            return False
    
    def _install_zip_update(self, update_file_path):
        """
        安装zip格式的更新包
        
        Args:
            update_file_path (str): 更新包路径
        """
        try:
            self.log_callback(f"开始处理zip更新包: {update_file_path}")
            
            # 获取当前程序所在目录
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe文件
                current_dir = os.path.dirname(sys.executable)
                current_exe_path = sys.executable
                is_frozen = True
            else:
                # 如果是Python脚本
                current_dir = os.getcwd()
                current_exe_path = os.path.join(current_dir, 'Wordformatter.exe')  # 假设目标exe文件名
                is_frozen = False
            
            self.log_callback(f"当前程序目录: {current_dir}")
            self.log_callback(f"当前程序路径: {current_exe_path}")
            self.log_callback(f"运行模式: {'打包exe' if is_frozen else 'Python脚本'}")
            
            # 创建临时解压目录
            temp_dir = tempfile.gettempdir()
            extract_dir = os.path.join(temp_dir, 'wordformatter_update')
            
            # 确保解压目录存在并清空
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir)
            
            # 解压更新包
            with zipfile.ZipFile(update_file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            self.log_callback(f"更新包已解压到: {extract_dir}")
            
            if is_frozen:
                # 生产环境：处理exe文件更新
                # 查找解压后的exe文件
                exe_files = []
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.endswith('.exe'):
                            exe_files.append(os.path.join(root, file))
                
                if not exe_files:
                    self.log_callback("错误: 在更新包中未找到exe文件")
                    return False
                
                # 假设第一个exe文件是主程序
                new_exe_path = exe_files[0]
                self.log_callback(f"找到新的程序文件: {new_exe_path}")
                
                # 创建更新脚本
                update_script_path = os.path.join(temp_dir, 'update_wordformatter.bat')
                with open(update_script_path, 'w', encoding='utf-8') as f:
                    f.write('@echo off\n')
                    f.write('chcp 65001 >nul\n')  # 设置UTF-8编码
                    f.write('timeout /t 3 /nobreak >nul\n')  # 等待3秒确保主程序完全退出
                    f.write(f'echo 正在更新程序...\n')
                    f.write(f'copy /Y "{new_exe_path}" "{current_exe_path}"\n')  # 覆盖原文件
                    f.write(f'if errorlevel 1 echo 文件复制失败\n')
                    f.write(f'del /Q "{update_file_path}" 2>nul\n')  # 删除更新包
                    f.write(f'rmdir /S /Q "{extract_dir}" 2>nul\n')  # 删除解压目录
                    f.write(f'del /Q "{update_script_path}" 2>nul\n')  # 删除自身
                    f.write(f'echo 启动新版本程序...\n')
                    f.write(f'start "" "{current_exe_path}"\n')  # 重新启动程序
                
                self.log_callback(f"更新脚本已创建: {update_script_path}")
                
                # 先关闭当前程序的日志系统，然后再启动更新脚本
                try:
                    logging.shutdown()
                except:
                    pass
                
                # 使用Windows ShellExecute启动更新脚本，避免权限问题
                import ctypes
                ctypes.windll.shell32.ShellExecuteW(
                    None, 
                    "runas", 
                    update_script_path, 
                    None, 
                    None, 
                    1
                )
            else:
                # 开发环境：复制所有文件到当前目录
                self.log_callback("开发环境模式：直接复制文件")
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        src_path = os.path.join(root, file)
                        # 计算相对路径
                        rel_path = os.path.relpath(src_path, extract_dir)
                        dst_path = os.path.join(current_dir, rel_path)
                        
                        # 确保目标目录存在
                        dst_dir = os.path.dirname(dst_path)
                        if not os.path.exists(dst_dir):
                            os.makedirs(dst_dir)
                        
                        # 复制文件
                        shutil.copy2(src_path, dst_path)
                        self.log_callback(f"已复制: {rel_path}")
                
                # 清理临时文件
                try:
                    os.remove(update_file_path)
                except:
                    pass
                try:
                    shutil.rmtree(extract_dir)
                except:
                    pass
                
                # 重新启动程序（如果是Python脚本）
                self.log_callback("更新完成，重新启动程序...")
                # 先关闭当前程序的日志系统
                try:
                    logging.shutdown()
                except:
                    pass
                subprocess.Popen([sys.executable, os.path.join(current_dir, 'WordFormatter.py')], close_fds=True)
            
            return True
            
        except Exception as e:
            self.log_callback(f"处理zip更新包时发生错误: {e}")
            self.logger.error(f"处理zip更新包时发生错误: {e}", exc_info=True)
            return False
    
    def _is_newer_version(self, latest, current):
        """
        比较版本号，判断是否为新版本
        
        Args:
            latest (str): 最新版本号
            current (str): 当前版本号
        
        Returns:
            bool: latest是否比current新
        """
        try:
            latest_parts = list(map(int, latest.split('.')))
            current_parts = list(map(int, current.split('.')))
            
            # 确保版本号部分长度一致
            max_length = max(len(latest_parts), len(current_parts))
            latest_parts += [0] * (max_length - len(latest_parts))
            current_parts += [0] * (max_length - len(current_parts))
            
            # 逐位比较版本号
            for latest_part, current_part in zip(latest_parts, current_parts):
                if latest_part > current_part:
                    return True
                elif latest_part < current_part:
                    return False
            
            return False  # 版本号相同
            
        except ValueError:
            # 版本号格式不正确，默认不是新版本
            return False
    
    def set_auto_update(self, auto_update):
        """
        设置是否启用自动更新
        
        Args:
            auto_update (bool): 是否启用自动更新
        """
        self.auto_update = auto_update
        self.config['auto_update'] = auto_update
    
    def set_check_interval(self, interval):
        """
        设置更新检查间隔
        
        Args:
            interval (int): 检查间隔（秒）
        """
        self.check_interval = interval
        self.config['check_update_interval'] = interval