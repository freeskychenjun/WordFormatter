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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class WordFormatterGUI:
    def __init__(self, master):
        self.master = master
        master.title("报告自动排版工具_JXSLY V1.0.2")
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

        self.font_size_map = {
            '一号 (26pt)': 26, '小一 (24pt)': 24, '二号 (22pt)': 22, '小二 (18pt)': 18,
            '三号 (16pt)': 16, '小三 (15pt)': 15, '四号 (14pt)': 14, '小四 (12pt)': 12,
            '五号 (10.5pt)': 10.5, '小五 (9pt)': 9
        }
        self.font_size_map_rev = {v: k for k, v in self.font_size_map.items()}
        
        self.default_params = {
            'page_number_align': '奇偶分页', 'line_spacing': 28,
            'margin_top': 3.7, 'margin_bottom': 3.5, 
            'margin_left': 2.8, 'margin_right': 2.6,
            'h1_font': '黑体', 'h2_font': '楷体_GB2312', 'h3_font': '宋体', 'body_font': '宋体',
            'page_number_font': '宋体', 'table_caption_font': '黑体', 'figure_caption_font': '黑体',
            'h1_size': 18, 'h1_space_before': 24, 'h1_space_after': 24,
            'h2_size': 12, 'h2_space_before': 24, 'h2_space_after': 24,
            'h3_size': 12, 'h3_space_before': 24, 'h3_space_after': 24,
            'body_size': 12, 'page_number_size': 14,
            'table_caption_size': 10.5, 'figure_caption_size': 10.5,
            # 添加表格标题和图表标题的大纲级别设置，默认为6级
            'table_caption_outline_level': 8, 'figure_caption_outline_level': 6,
            'set_outline': True,
            # 添加标题粗体设置
            'h1_bold': False,  # 一级标题默认不加粗
            'h2_bold': True,   # 二级标题默认加粗
            'h3_bold': False,  # 三级标题默认不加粗
            'table_caption_bold': False,  # 表格标题默认不加粗
            'figure_caption_bold': False,  # 图形标题默认不加粗
            # 自动更新默认设置
            'auto_update': True  # 默认启用自动更新
        }
        self.font_options = {
            'h1': ['黑体', '方正黑体_GBK', '方正黑体简体', '华文黑体', '宋体', '仿宋', '仿宋_GB2312'],
            'h2': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体', '宋体', '仿宋', '仿宋_GB2312'],
            'h3': ['宋体', '仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋'],
            'body': ['仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋', '宋体'], 
            'table_caption': ['黑体', '宋体', '仿宋_GB2312', '仿宋'], 'figure_caption': ['黑体', '宋体', '仿宋_GB2312', '仿宋']
        }
        self.set_outline_var = tk.BooleanVar(value=self.default_params['set_outline'])

        self.entries = {}
        self.checkboxes = {}  # 存储复选框变量
        
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
        
        self.master.after(250, self.set_initial_pane_position)
        # 程序启动时检查更新
        self.master.after(1000, self.check_for_updates_once)

    def set_initial_pane_position(self):
        # 获取窗口总宽度，设置左侧占约30%
        total_width = self.master.winfo_width()
        
        if total_width > 100:  # 确保窗口已经渲染
            left_width = int(total_width * 0.3)  # 左侧占30%
            # 使用保存的main_pane引用直接设置位置
            try:
                if hasattr(self, 'main_pane'):
                    self.main_pane.sashpos(0, left_width)
            except Exception as e:
                # 如果直接设置失败，回退到原方法
                for widget in self.master.winfo_children():
                    if isinstance(widget, ttk.PanedWindow):
                        widget.sashpos(0, left_width)
                        break

    def create_menu(self):
        menubar = Menu(self.master)
        # 删除帮助菜单
        self.master.config(menu=menubar)

    def create_widgets(self):
        # 创建主容器，使用垂直布局
        content_frame = ttk.Frame(self.master)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建水平分割的主面板（上方部分）
        main_pane = ttk.PanedWindow(content_frame, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        # 保存main_pane引用，便于后续访问
        self.main_pane = main_pane

        # 左侧文件处理区域
        left_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, weight=3)

        notebook = ttk.Notebook(left_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook = notebook

        file_tab = ttk.Frame(notebook)
        notebook.add(file_tab, text=' 文件批量处理 ')
        
        # 创建统一的内容区域，优化布局减少空白
        left_content_frame = ttk.Frame(file_tab)
        left_content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 文件列表区域
        list_frame = ttk.LabelFrame(left_content_frame, text="待处理文件列表（可拖拽文件或文件夹）", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件列表和滚动条
        list_inner_frame = ttk.Frame(list_frame)
        list_inner_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_inner_frame, orient=tk.VERTICAL)
        # 为文件列表设置固定高度，避免占用过多空间
        self.file_listbox = tk.Listbox(list_inner_frame, yscrollcommand=scrollbar.set, selectmode=tk.EXTENDED)
        scrollbar.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 5))
        
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
        self.placeholder_label = ttk.Label(self.file_listbox, text="可以拖拽文件或文件夹到这里", foreground="grey")
        
        # 文件操作按钮区域
        file_button_frame = ttk.Frame(left_content_frame)
        file_button_frame.pack(fill=tk.X, pady=(5, 0))
        
        # 使用网格布局优化按钮排列
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="添加文件夹", command=self.add_folder).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="移除文件", command=self.remove_files).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_list).grid(row=1, column=1, sticky='ew', padx=2, pady=2)
        
        file_button_frame.columnconfigure(0, weight=1)
        file_button_frame.columnconfigure(1, weight=1)

        # 右侧参数设置区域
        right_frame = ttk.Frame(main_pane, padding=(5, 0, 0, 0))
        main_pane.add(right_frame, weight=7)
        
        # 在主面板下方创建处理日志区域
        log_frame = ttk.LabelFrame(content_frame, text="处理日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=False)
        # 确保调试日志文本框能完全拉伸至窗体边缘
        # 限制调试日志面板高度，仅显示必要内容
        self.debug_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', wrap=tk.WORD)
        self.debug_text.pack(fill=tk.BOTH, expand=True)
        
        # 创建统一的右侧内容区域，与左侧面板结构保持一致
        right_content_frame = ttk.Frame(right_frame)
        right_content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建带滚动条的参数设置区域
        canvas = tk.Canvas(right_content_frame)
        v_scrollbar = ttk.Scrollbar(right_content_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=v_scrollbar.set)
        
        # 创建参数容器
        params_container = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=params_container, anchor='nw', width=right_content_frame.winfo_width()-20)
        
        # 参数设置框架
        params_frame = ttk.LabelFrame(params_container, text="参数设置", padding=10)
        params_frame.pack(fill=tk.BOTH, expand=True)
        params_frame.columnconfigure(1, weight=1)
        params_frame.columnconfigure(3, weight=1)
        params_frame.columnconfigure(5, weight=1)

        # Helper functions for creating widgets
        def create_entry(label, var_name, r, c, width=12):
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=5, pady=3)
            entry = ttk.Entry(params_frame, width=width)
            entry.grid(row=r, column=c+1, sticky=tk.EW, padx=5, pady=3)
            self.entries[var_name] = entry
            return entry
        
        def create_combo(label, var_name, opts, r, c, readonly=True, width=15): 
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=5, pady=3)
            state = 'readonly' if readonly else 'normal'
            combo = ttk.Combobox(params_frame, values=opts, state=state, width=width)
            combo.grid(row=r, column=c+1, sticky=tk.EW, padx=5, pady=3)
            self.entries[var_name] = combo
            return combo

        def create_font_size_combo(label, var_name, r, c, width=15):
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=5, pady=3)
            combo = ttk.Combobox(params_frame, values=list(self.font_size_map.keys()), width=width)
            combo.grid(row=r, column=c+1, sticky=tk.EW, padx=5, pady=3)
            self.entries[var_name] = combo
            return combo
        
        def create_checkbox(label, var_name, r, c, default_value=False):
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=5, pady=3)
            checkbox_var = tk.BooleanVar(value=default_value)
            checkbox = ttk.Checkbutton(params_frame, variable=checkbox_var)
            checkbox.grid(row=r, column=c+1, sticky=tk.W, padx=5, pady=3)
            self.checkboxes[var_name] = checkbox_var
            return checkbox_var
        
        def create_section_header(text, help_text, r):
            header_frame = ttk.Frame(params_frame)
            header_frame.grid(row=r, column=0, columnspan=6, sticky='ew', pady=(15, 5))
            ttk.Label(header_frame, text=text, font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT)
            # 删除帮助提示功能
            ttk.Separator(params_frame, orient='horizontal').grid(row=r+1, column=0, columnspan=6, sticky='ew', pady=(5, 10))
            return r + 2

        row = 0
        
        # Section: Page Layout
        row = create_section_header("页面设置", None, row)
        create_entry("上边距(cm)", 'margin_top', row, 0, width=15)
        create_entry("下边距(cm)", 'margin_bottom', row, 2, width=15)
        row += 1
        create_entry("左边距(cm)", 'margin_left', row, 0, width=15)
        create_entry("右边距(cm)", 'margin_right', row, 2, width=15)
        row += 1

        # Section: Document Title

        # Section: Body and Headings
        row = create_section_header("正文与层级", None, row)
        create_combo("一级标题字体", 'h1_font', self.font_options['h1'], row, 0, readonly=False, width=18)
        create_font_size_combo("一级标题字号", 'h1_size', row, 2, width=18)
        create_checkbox("一级标题加粗", 'h1_bold', row, 4, default_value=False)  # 一级标题默认不加粗
        row += 1
        create_entry("一级段前(磅)", 'h1_space_before', row, 0, width=15)
        create_entry("一级段后(磅)", 'h1_space_after', row, 2, width=15)
        row += 1
        create_combo("二级标题字体", 'h2_font', self.font_options['h2'], row, 0, readonly=False, width=18)
        create_font_size_combo("二级标题字号", 'h2_size', row, 2, width=18)
        create_checkbox("二级标题加粗", 'h2_bold', row, 4, default_value=True)  # 二级标题默认加粗
        row += 1
        create_entry("二级段前(磅)", 'h2_space_before', row, 0, width=15)
        create_entry("二级段后(磅)", 'h2_space_after', row, 2, width=15)
        row += 1
        create_combo("三级标题字体", 'h3_font', self.font_options['h3'], row, 0, readonly=False, width=18)
        create_font_size_combo("三级标题字号", 'h3_size', row, 2, width=18)
        create_checkbox("三级标题加粗", 'h3_bold', row, 4, default_value=False)  # 三级标题默认不加粗
        row += 1
        create_entry("三级段前(磅)", 'h3_space_before', row, 0, width=15)
        create_entry("三级段后(磅)", 'h3_space_after', row, 2, width=15)
        row += 1
        create_combo("正文/四级字体", 'body_font', self.font_options['body'], row, 0, readonly=False, width=18)
        create_font_size_combo("正文/四级字号", 'body_size', row, 2, width=18)
        create_entry("正文行距(磅)", 'line_spacing', row, 4, width=15)
        row += 1
        # 在同一行添加正文Times New Roman复选框和表格标题加粗复选框
        create_checkbox("正文英文/数字使用Times New Roman", 'body_use_times_roman', row, 0, default_value=True)  # 默认启用
        # 添加表格标题加粗复选框（放在同一行的右侧）
        ttk.Label(params_frame, text="表格标题加粗").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        table_bold_var = tk.BooleanVar(value=False)  # 默认为不加粗
        table_bold_checkbox = ttk.Checkbutton(params_frame, variable=table_bold_var)
        table_bold_checkbox.grid(row=row, column=5, sticky=tk.W, padx=5, pady=3)
        self.checkboxes['table_caption_bold'] = table_bold_var
        row += 1
        
        # Section: Other Elements
        row = create_section_header("其他元素", None, row)
        create_combo("表格标题字体", 'table_caption_font', self.font_options['table_caption'], row, 0, readonly=False, width=18)
        create_font_size_combo("表格标题字号", 'table_caption_size', row, 2, width=18)
        # 添加表格标题大纲级别（移到同一行）
        ttk.Label(params_frame, text="表格标题大纲级别").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        table_outline_combo = ttk.Combobox(params_frame, values=['无', '1', '2', '3', '4', '5', '6', '7', '8', '9'], width=18)
        table_outline_combo.grid(row=row, column=5, sticky=tk.EW, padx=5, pady=3)
        table_outline_combo.set('8')  # 默认为8级
        self.entries['table_caption_outline_level'] = table_outline_combo
        row += 1
        create_checkbox("表格内容英文/数字使用Times New Roman", 'table_use_times_roman', row, 0, default_value=True)  # 默认启用
        row += 1
        
        create_combo("图形标题字体", 'figure_caption_font', self.font_options['figure_caption'], row, 0, readonly=False, width=18)
        create_font_size_combo("图形标题字号", 'figure_caption_size', row, 2, width=18)
        # 添加图形标题大纲级别（移到同一行）
        ttk.Label(params_frame, text="图形标题大纲级别").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        figure_outline_combo = ttk.Combobox(params_frame, values=['无', '1', '2', '3', '4', '5', '6', '7', '8', '9'], width=18)
        figure_outline_combo.grid(row=row, column=5, sticky=tk.EW, padx=5, pady=3)
        figure_outline_combo.set('6')  # 默认为6级
        self.entries['figure_caption_outline_level'] = figure_outline_combo
        row += 1
        # 添加图形标题加粗复选框（放在大纲级别控件下方）
        ttk.Label(params_frame, text="图形标题加粗").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        figure_bold_var = tk.BooleanVar(value=False)  # 默认为不加粗
        figure_bold_checkbox = ttk.Checkbutton(params_frame, variable=figure_bold_var)
        figure_bold_checkbox.grid(row=row, column=5, sticky=tk.W, padx=5, pady=3)
        self.checkboxes['figure_caption_bold'] = figure_bold_var
        row += 1


        
        # Section: Global Options
        ttk.Separator(params_frame, orient='horizontal').grid(row=row, column=0, columnspan=6, sticky='ew', pady=10)
        row += 1

        # 按钮区域
        button_frame = ttk.Frame(params_container, padding=(0, 10, 0, 10))
        button_frame.pack(fill=tk.X)
        
        # 配置按钮 - 2x2布局
        config_buttons = ttk.LabelFrame(button_frame, text="参数管理", padding=10)
        config_buttons.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(config_buttons, text="加载参数", command=self.load_config).grid(row=0, column=0, sticky='ew', padx=5, pady=5)
        ttk.Button(config_buttons, text="保存参数", command=self.save_config).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(config_buttons, text="保存为默认", command=self.save_default_config).grid(row=1, column=0, sticky='ew', padx=5, pady=5)
        ttk.Button(config_buttons, text="恢复内置默认", command=self.load_defaults).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        config_buttons.columnconfigure(0, weight=1)
        config_buttons.columnconfigure(1, weight=1)

        # 开始排版按钮
        style = ttk.Style()
        style.configure('Success.TButton', font=('Helvetica', 11, 'bold'))
        start_button_frame = ttk.Frame(button_frame)
        # 向下移动1cm（约38像素）
        start_button_frame.pack(fill=tk.X, pady=(38, 0))
        ttk.Button(start_button_frame, text="开始排版", style='Success.TButton', command=self.start_processing).pack(fill=tk.X, ipady=10)

        # 配置Canvas滚动
        def on_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 调整Canvas内容宽度以适应Canvas
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width-20)

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind('<Configure>', on_canvas_configure)
        params_container.bind('<Configure>', on_frame_configure)
        
        # 添加鼠标滚轮支持
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        params_container.bind_all("<MouseWheel>", on_mousewheel)
        
        # 布局Canvas和滚动条
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._update_listbox_placeholder()
        
        # 添加定时器，延迟一小段时间后再次应用默认配置，确保UI控件完全创建
        self.master.after(100, self._apply_default_spacing_values)
    
    def log_to_debug_window(self, message):
        self.master.update_idletasks()
        self.debug_text.config(state='normal')
        self.debug_text.insert(tk.END, message + '\n')
        self.debug_text.config(state='disabled')
        self.debug_text.see(tk.END)
    
    def _apply_default_spacing_values(self):
        # 直接设置标题字体和字号
        if 'h3_font' in self.entries:
            self.entries['h3_font'].set(self.default_params['h3_font'])
        if 'h3_size' in self.entries:
            display_val = self.font_size_map_rev.get(self.default_params['h3_size'], str(self.default_params['h3_size']))
            self.entries['h3_size'].set(display_val)
        
        # 直接设置标题间距输入框的值
        if 'h1_space_before' in self.entries:
            self.entries['h1_space_before'].delete(0, tk.END)
            self.entries['h1_space_before'].insert(0, str(self.default_params['h1_space_before']))
        if 'h1_space_after' in self.entries:
            self.entries['h1_space_after'].delete(0, tk.END)
            self.entries['h1_space_after'].insert(0, str(self.default_params['h1_space_after']))
        if 'h2_space_before' in self.entries:
            self.entries['h2_space_before'].delete(0, tk.END)
            self.entries['h2_space_before'].insert(0, str(self.default_params['h2_space_before']))
        if 'h2_space_after' in self.entries:
            self.entries['h2_space_after'].delete(0, tk.END)
            self.entries['h2_space_after'].insert(0, str(self.default_params['h2_space_after']))
        if 'h3_space_before' in self.entries:
            self.entries['h3_space_before'].delete(0, tk.END)
            self.entries['h3_space_before'].insert(0, str(self.default_params['h3_space_before']))
        if 'h3_space_after' in self.entries:
            self.entries['h3_space_after'].delete(0, tk.END)
            self.entries['h3_space_after'].insert(0, str(self.default_params['h3_space_after']))
        
        # 确认已设置的值 - 不再输出到日志窗口
        # self.log_to_debug_window("标题间距值已设置到输入框:")
        # for key in ['h1_space_before', 'h1_space_after', 'h2_space_before', 'h2_space_after', 'h3_space_before', 'h3_space_after']:
        #     if key in self.entries:
        #         self.log_to_debug_window(f"{key}: {self.entries[key].get()}")
        # 直接设置标题字体和字号
        if 'h3_font' in self.entries:
            self.entries['h3_font'].set(self.default_params['h3_font'])
        if 'h3_size' in self.entries:
            display_val = self.font_size_map_rev.get(self.default_params['h3_size'], str(self.default_params['h3_size']))
            self.entries['h3_size'].set(display_val)
        
        # 直接设置标题间距输入框的值
        if 'h1_space_before' in self.entries:
            self.entries['h1_space_before'].delete(0, tk.END)
            self.entries['h1_space_before'].insert(0, str(self.default_params['h1_space_before']))
        if 'h1_space_after' in self.entries:
            self.entries['h1_space_after'].delete(0, tk.END)
            self.entries['h1_space_after'].insert(0, str(self.default_params['h1_space_after']))
        if 'h2_space_before' in self.entries:
            self.entries['h2_space_before'].delete(0, tk.END)
            self.entries['h2_space_before'].insert(0, str(self.default_params['h2_space_before']))
        if 'h2_space_after' in self.entries:
            self.entries['h2_space_after'].delete(0, tk.END)
            self.entries['h2_space_after'].insert(0, str(self.default_params['h2_space_after']))
        if 'h3_space_before' in self.entries:
            self.entries['h3_space_before'].delete(0, tk.END)
            self.entries['h3_space_before'].insert(0, str(self.default_params['h3_space_before']))
        if 'h3_space_after' in self.entries:
            self.entries['h3_space_after'].delete(0, tk.END)
            self.entries['h3_space_after'].insert(0, str(self.default_params['h3_space_after']))
        
        # 确认已设置的值 - 不再输出到日志窗口
        # self.log_to_debug_window("标题间距值已设置到输入框:")
        # for key in ['h1_space_before', 'h1_space_after', 'h2_space_before', 'h2_space_after', 'h3_space_before', 'h3_space_after']:
        #     if key in self.entries:
        #         self.log_to_debug_window(f"{key}: {self.entries[key].get()}")

    def load_initial_config(self):
        # 使用配置管理器加载排版配置
        if not self.config_manager.format_config:
            self.load_defaults()
        else:
            self._apply_config(self.config_manager.format_config)
        
        # 添加定时器，延迟一小段时间后再次应用默认配置，确保UI控件完全创建
        self.master.after(100, self._apply_default_spacing_values)
    
    def _apply_config(self, loaded_config):
        self.set_outline_var.set(loaded_config.get('set_outline', True))
        for key, value in loaded_config.items():
            if key in ['set_outline', 'auto_update']: continue
            
            # 处理输入框和下拉框的值
            widget = self.entries.get(key)
            if widget:
                if "_size" in key:
                    display_val = self.font_size_map_rev.get(value, str(value))
                    widget.set(display_val)
                elif isinstance(widget, ttk.Combobox):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(value))
            
            # 处理复选框的值（仅标题粗体设置）
            checkbox_var = self.checkboxes.get(key)
            if checkbox_var is not None:
                checkbox_var.set(bool(value))

    def load_defaults(self):
        self._apply_config(self.default_params)
    
    def collect_config(self):
        config = {}
        # 收集输入框和下拉框的值
        for key, widget in self.entries.items():
            value = widget.get().strip()
            if "_size" in key:
                if value in self.font_size_map:
                    config[key] = self.font_size_map[value]
                else:
                    try: config[key] = float(value)
                    except (ValueError, TypeError):
                        self.log_to_debug_window(f"警告: 无效的字号值 '{value}' for '{key}'. 使用默认值 16pt。")
                        config[key] = 16
            else:
                try: config[key] = float(value) if '.' in value else int(value)
                except (ValueError, TypeError): config[key] = value
        # 收集复选框的值（标题粗体设置）
        for key, checkbox_var in self.checkboxes.items():
            config[key] = checkbox_var.get()
        # 添加自动更新的默认配置
        config['auto_update'] = self.default_params['auto_update']
        config['set_outline'] = self.set_outline_var.get()
        return config

    def save_config(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f: json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"配置已保存至 {file_path}")
    
    def save_default_config(self):
        try:
            with open(self.default_config_path, 'w', encoding='utf-8') as f:
                json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"当前配置已保存为默认配置。\n下次启动软件时将自动加载。")
        except Exception as e:
            messagebox.showerror("错误", f"保存默认配置失败: {e}")

    def load_config(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                self._apply_config(loaded_config)
                messagebox.showinfo("成功", "配置已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载参数文件失败: {e}")

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
        
        processor = WordProcessor(self.collect_config(), self.log_to_debug_window)
        active_tab_index = self.notebook.index(self.notebook.select())

        try:
            if active_tab_index == 0:
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