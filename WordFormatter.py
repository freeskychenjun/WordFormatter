import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, Menu
import json
import os
import logging

from tkinterdnd2 import DND_FILES, TkinterDnD

# 从模块导入WordProcessor
from modules.word_processor import WordProcessor
from modules.update_manager import UpdateManager
from modules.config_manager import ConfigManager
from gui.settings_window import SettingsWindow

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class WordFormatterGUI:
    def __init__(self, master):
        self.master = master
        master.title("报告自动排版工具_JXSLY V1.0.3")
        # 增加窗体尺寸：宽度增加7%，高度再增加5%
        # 原始尺寸：1320x813，调整后约为1412x942
        master.geometry("1412x942")
        master.minsize(1200, 700)  # 设置最小窗口大小
        
        # 使程序启动时界面位于屏幕中央
        # 先更新窗口任务，确保窗口尺寸已应用
        master.update_idletasks()
        # 获取屏幕尺寸
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        # 获取窗口尺寸
        window_width = 1412
        window_height = 942
        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        # 设置窗口位置
        master.geometry(f'{window_width}x{window_height}+{x}+{y}')

        self.set_outline_var = tk.BooleanVar(value=True)
        
        self.default_config_path = "default_config.json"
        
        self.create_menu()
        self.create_widgets()

        # 初始化配置管理器
        self.config_manager = ConfigManager(self.default_config_path)
        self.config_manager.load_config()
        
        # 初始化更新配置管理器
        update_config_path = os.path.join(os.path.dirname(self.default_config_path), "update_config.json")
        update_config = self.config_manager.load_update_config(update_config_path)
        
        # 检查update_config.json文件是否存在
        if update_config is None:
            self.log_to_debug_window("警告: 缺少update_config.json文件，将使用默认更新设置")
        
        # 加载初始配置
        self.load_initial_config()
        
        # 初始化更新管理器
        # 如果update_config为None，使用默认更新配置
        if update_config is None:
            update_config = self.config_manager.get_default_update_config()
        self.update_manager = UpdateManager(update_config, self.log_to_debug_window)
        
        # 程序启动时检查更新
        self.master.after(1000, self.check_for_updates_once)

    # set_initial_pane_position方法已移除，因为不再使用分割面板

    def create_menu(self):
        menubar = Menu(self.master)
        
        # 文件菜单
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="退出", command=self.master.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 设置菜单
        settings_menu = Menu(menubar, tearoff=0)
        settings_menu.add_command(label="参数设置", command=self.open_settings_window)
        menubar.add_cascade(label="设置", menu=settings_menu)
        
        self.master.config(menu=menubar)

    def create_widgets(self):
        # 创建主容器，使用垂直布局
        content_frame = ttk.Frame(self.master)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建文件处理区域
        file_frame = ttk.LabelFrame(content_frame, text="文件处理", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 文件列表区域
        list_frame = ttk.LabelFrame(file_frame, text="待处理文件列表（可拖拽文件或文件夹）", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件列表和滚动条
        list_inner_frame = ttk.Frame(list_frame)
        list_inner_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_inner_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(list_inner_frame, yscrollcommand=scrollbar.set, selectmode=tk.EXTENDED)
        scrollbar.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 5))
        
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
        self.placeholder_label = ttk.Label(self.file_listbox, text="可以拖拽文件或文件夹到这里", foreground="grey")
        
        # 文件操作按钮区域
        file_button_frame = ttk.Frame(file_frame)
        file_button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 使用网格布局优化按钮排列
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="添加文件夹", command=self.add_folder).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="移除文件", command=self.remove_files).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_list).grid(row=1, column=1, sticky='ew', padx=2, pady=2)
        file_button_frame.columnconfigure(0, weight=1)
        file_button_frame.columnconfigure(1, weight=1)
        
        # 控制按钮区域
        control_frame = ttk.Frame(content_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 开始排版按钮 - 使用pack布局并填充整个可用宽度
        style = ttk.Style()
        style.configure('Success.TButton', font=('Helvetica', 11, 'bold'))
        self.start_button = ttk.Button(control_frame, text="开始排版", style='Success.TButton', command=self.start_processing)
        self.start_button.pack(fill=tk.X, padx=5, ipady=8)  # 使用fill=tk.X使按钮水平填充整个空间，增加内边距使按钮更高
        
        # 在主面板下方创建处理日志区域
        log_frame = ttk.LabelFrame(content_frame, text="处理日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.debug_text = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', wrap=tk.WORD)
        self.debug_text.pack(fill=tk.BOTH, expand=True)
        
        self._update_listbox_placeholder()
    
    def log_to_debug_window(self, message):
        self.master.update_idletasks()
        self.debug_text.config(state='normal')
        self.debug_text.insert(tk.END, message + '\n')
        self.debug_text.config(state='disabled')
        self.debug_text.see(tk.END)
    
    # _apply_default_spacing_values方法已移除，因为不再需要

    def load_initial_config(self):
        # 使用配置管理器加载排版配置
        if not self.config_manager.format_config:
            self.config_manager.load_config()
    
    # 原有的参数配置方法已移除，因为已转移到SettingsWindow类中

    def _update_listbox_placeholder(self):
        if self.file_listbox.size() == 0:
            self.placeholder_label.place(in_=self.file_listbox, relx=0.5, rely=0.5, anchor=tk.CENTER)
        else:
            self.placeholder_label.place_forget()

    def handle_drop(self, event):
        paths = self.master.tk.splitlist(event.data)
        self._add_paths_to_listbox(paths)

    def _add_paths_to_listbox(self, paths):
        current_files = set(self.file_listbox.get(0, tk.END))
        added_count = 0
        
        for path in paths:
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(('.docx', '.doc', '.wps', '.txt')):
                            full_path = os.path.join(root, f)
                            if full_path not in current_files:
                                self.file_listbox.insert(tk.END, full_path)
                                current_files.add(full_path)
                                added_count += 1
            elif os.path.isfile(path):
                if path.lower().endswith(('.docx', '.doc', '.wps', '.txt')):
                    if path not in current_files:
                        self.file_listbox.insert(tk.END, path)
                        current_files.add(path)
                        added_count += 1
        
        if added_count > 0:
            self.log_to_debug_window(f"通过按钮或拖拽添加了 {added_count} 个新文件。")
        
        self._update_listbox_placeholder()

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("所有支持的文件", "*.docx;*.doc;*.wps;*.txt"), ("Word 文档", "*.docx;*.doc"), ("WPS 文档", "*.wps"), ("纯文本", "*.txt")])
        if files:
            self._add_paths_to_listbox(files)
        
    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self._add_paths_to_listbox([folder])

    def remove_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先在列表中选择要移除的文件。")
            return
        for index in sorted(selected_indices, reverse=True):
            self.file_listbox.delete(index)
        self._update_listbox_placeholder()

    def clear_list(self): 
        self.file_listbox.delete(0, tk.END)
        self._update_listbox_placeholder()
    
    def open_settings_window(self):
        """打开参数设置窗体"""
        settings_window = SettingsWindow(self.master, self.config_manager, self.log_to_debug_window)
    
    def on_settings_updated(self, config):
        """当设置更新时调用的回调函数"""
        self.log_to_debug_window("参数设置已更新，当前处理将使用新参数")
    
    def check_for_updates_once(self):
        """
        程序启动时检查更新（仅检查一次）
        """
        try:
            # 调用更新管理器检查更新
            result = self.update_manager.check_for_updates()
            
            # 处理返回结果
            if isinstance(result, tuple) and len(result) == 3:
                has_update, version, release_info = result
                if has_update:
                    # 询问用户是否更新
                    self.log_to_debug_window(f"发现新版本 v{version}，是否立即更新？")
                    response = messagebox.askyesno("更新提示", f"发现新版本 v{version}\n\n是否立即更新？")
                    if response:
                        self.log_to_debug_window("用户选择更新，开始下载...")
                        # 下载更新
                        update_file = self.update_manager.download_update(release_info)
                        if update_file:
                            # 安装更新
                            self.update_manager.install_update(update_file)
                else:
                    # 更新管理器已记录日志，此处不再重复输出
                    pass
            else:
                self.log_to_debug_window("未检查到更新")
        except Exception as e:
            self.log_to_debug_window(f"更新检查失败: {e}")
            logging.error(f"更新检查失败: {e}", exc_info=True)



    def start_processing(self):
        warning_title = "处理前重要提示"
        warning_message = (
            "为了防止数据丢失，请在继续前关闭所有已打开的Word和WPS文档（包括wps、表格、PPT等所有文档）。\n\n"
            "本程序在转换文件格式时需要调用Word/WPS程序，这可能会导致您未保存的工作被强制关闭。\n\n"
            "您确定要继续吗？"
        )
        if not messagebox.askokcancel(warning_title, warning_message):
            self.log_to_debug_window("用户已取消操作。")
            return
            
        self.debug_text.config(state='normal'); self.debug_text.delete('1.0', tk.END); self.debug_text.config(state='disabled')
        
        # 确保配置已加载，如果没有则加载默认配置
        if self.config_manager.format_config is None:
            self.config_manager.load_config()
        
        processor = WordProcessor(self.config_manager.format_config, self.log_to_debug_window)

        try:
            file_list = self.file_listbox.get(0, tk.END)
            if not file_list:
                messagebox.showwarning("警告", "文件列表为空，请先添加文件！"); return
            output_dir = filedialog.askdirectory(title="请选择一个文件夹用于存放处理后的文件")
            if not output_dir: return

            success_count, fail_count = 0, 0
            for i, input_path in enumerate(file_list):
                try:
                    self.log_to_debug_window(f"\n--- 开始处理文件 {i+1}/{len(file_list)}: {os.path.basename(input_path)} ---")
                    base_name = os.path.splitext(os.path.basename(input_path))[0]
                    output_path = os.path.join(output_dir, f"{base_name}_formatted.docx")
                    processor.format_document(input_path, output_path)
                    self.log_to_debug_window(f"✅ 文件处理成功，已保存至: {output_path}")
                    success_count += 1
                except Exception as e:
                    logging.error(f"处理文件失败: {input_path}\n{e}", exc_info=True)
                    self.log_to_debug_window(f"\n❌ 处理文件 {os.path.basename(input_path)} 时发生严重错误：\n{e}")
                    fail_count += 1
                finally:
                    processor._cleanup_temp_files()
            
            summary_message = f"批量处理完成！\n\n成功: {success_count}个\n失败: {fail_count}个"
            if fail_count > 0: summary_message += "\n\n失败详情请查看日志窗口。"
            messagebox.showinfo("完成", summary_message)
            self.log_to_debug_window(f"\n🎉 {summary_message}")
            self.log_to_debug_window("\n💡 提示：处理完成的文件可能正在被系统占用，请稍等几秒后再打开。")

    
        except Exception as e:
            logging.error(f"处理过程中发生严重错误: {e}", exc_info=True)
            self.log_to_debug_window(f"\n❌ 处理过程中发生严重错误：\n{e}")
            messagebox.showerror("错误", f"处理过程中发生错误：\n{e}")
        finally:
            processor.quit_com_app()
            self.log_to_debug_window("\n💡 所有任务完成，WPS/Word应用已关闭，现在可以安全地打开处理后的文件了。")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = WordFormatterGUI(root)
    root.mainloop()