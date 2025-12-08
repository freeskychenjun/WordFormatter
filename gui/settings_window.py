import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os

class SettingsWindow(tk.Toplevel):
    """参数设置窗体"""
    
    def __init__(self, parent, config_manager, log_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.config_manager = config_manager
        self.log_callback = log_callback
        
        # 窗体设置
        self.title("参数设置")
        self.geometry("900x700")
        self.resizable(True, True)
        
        # 使窗体位于父窗口中心
        self.transient(parent)
        self.grab_set()
        parent.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (900 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (700 // 2)
        self.geometry(f"900x700+{x}+{y}")
        
        # 控件存储
        self.entries = {}
        self.checkboxes = {}
        
        # 创建界面
        self.create_widgets()
        
        # 加载当前配置
        self.load_current_config()
    
    def create_widgets(self):
        """创建界面控件"""
        # 创建主容器
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建带滚动条的参数设置区域
        canvas = tk.Canvas(main_frame)
        v_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=v_scrollbar.set)
        
        # 创建参数容器
        params_container = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=params_container, anchor='nw', width=850)
        
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
            combo = ttk.Combobox(params_frame, values=list(self.config_manager.font_size_map.keys()), width=width)
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
        
        def create_section_header(text, r):
            header_frame = ttk.Frame(params_frame)
            header_frame.grid(row=r, column=0, columnspan=6, sticky='ew', pady=(15, 5))
            ttk.Label(header_frame, text=text, font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT)
            ttk.Separator(params_frame, orient='horizontal').grid(row=r+1, column=0, columnspan=6, sticky='ew', pady=(5, 10))
            return r + 2

        row = 0
        
        # Section: Page Layout
        row = create_section_header("页面设置", row)
        create_entry("上边距(cm)", 'margin_top', row, 0, width=15)
        create_entry("下边距(cm)", 'margin_bottom', row, 2, width=15)
        row += 1
        create_entry("左边距(cm)", 'margin_left', row, 0, width=15)
        create_entry("右边距(cm)", 'margin_right', row, 2, width=15)
        row += 1

        # Section: Document Title

        # Section: Body and Headings
        row = create_section_header("正文与层级", row)
        create_combo("一级标题字体", 'h1_font', self.config_manager.get_font_options('h1'), row, 0, readonly=False, width=18)
        create_font_size_combo("一级标题字号", 'h1_size', row, 2, width=18)
        create_checkbox("一级标题加粗", 'h1_bold', row, 4, default_value=False)
        row += 1
        create_entry("一级段前(磅)", 'h1_space_before', row, 0, width=15)
        create_entry("一级段后(磅)", 'h1_space_after', row, 2, width=15)
        row += 1
        create_combo("二级标题字体", 'h2_font', self.config_manager.get_font_options('h2'), row, 0, readonly=False, width=18)
        create_font_size_combo("二级标题字号", 'h2_size', row, 2, width=18)
        create_checkbox("二级标题加粗", 'h2_bold', row, 4, default_value=True)
        row += 1
        create_entry("二级段前(磅)", 'h2_space_before', row, 0, width=15)
        create_entry("二级段后(磅)", 'h2_space_after', row, 2, width=15)
        row += 1
        create_combo("三级标题字体", 'h3_font', self.config_manager.get_font_options('h3'), row, 0, readonly=False, width=18)
        create_font_size_combo("三级标题字号", 'h3_size', row, 2, width=18)
        create_checkbox("三级标题加粗", 'h3_bold', row, 4, default_value=False)
        row += 1
        create_entry("三级段前(磅)", 'h3_space_before', row, 0, width=15)
        create_entry("三级段后(磅)", 'h3_space_after', row, 2, width=15)
        row += 1
        create_combo("正文/四级字体", 'body_font', self.config_manager.get_font_options('body'), row, 0, readonly=False, width=18)
        create_font_size_combo("正文/四级字号", 'body_size', row, 2, width=18)
        create_entry("正文行距(磅)", 'line_spacing', row, 4, width=15)
        row += 1
        create_checkbox("正文英文/数字使用Times New Roman", 'body_use_times_roman', row, 0, default_value=True)
        row += 1
        
        # Section: Other Elements
        row = create_section_header("其他元素", row)
        create_combo("表格标题字体", 'table_caption_font', self.config_manager.get_font_options('table_caption'), row, 0, readonly=False, width=18)
        create_font_size_combo("表格标题字号", 'table_caption_size', row, 2, width=18)
        ttk.Label(params_frame, text="表格标题大纲级别").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        table_outline_combo = ttk.Combobox(params_frame, values=['无', '1', '2', '3', '4', '5', '6', '7', '8', '9'], width=18)
        table_outline_combo.grid(row=row, column=5, sticky=tk.EW, padx=5, pady=3)
        table_outline_combo.set('8')
        self.entries['table_caption_outline_level'] = table_outline_combo
        row += 1
        create_checkbox("表格内容英文/数字使用Times New Roman", 'table_use_times_roman', row, 0, default_value=True)
        ttk.Label(params_frame, text="表格标题加粗").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        table_bold_var = tk.BooleanVar(value=False)
        table_bold_checkbox = ttk.Checkbutton(params_frame, variable=table_bold_var)
        table_bold_checkbox.grid(row=row, column=5, sticky=tk.W, padx=5, pady=3)
        self.checkboxes['table_caption_bold'] = table_bold_var
        row += 1
        
        create_combo("图形标题字体", 'figure_caption_font', self.config_manager.get_font_options('figure_caption'), row, 0, readonly=False, width=18)
        create_font_size_combo("图形标题字号", 'figure_caption_size', row, 2, width=18)
        ttk.Label(params_frame, text="图形标题大纲级别").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        figure_outline_combo = ttk.Combobox(params_frame, values=['无', '1', '2', '3', '4', '5', '6', '7', '8', '9'], width=18)
        figure_outline_combo.grid(row=row, column=5, sticky=tk.EW, padx=5, pady=3)
        figure_outline_combo.set('6')
        self.entries['figure_caption_outline_level'] = figure_outline_combo
        row += 1
        ttk.Label(params_frame, text="图形标题加粗").grid(row=row, column=4, sticky=tk.W, padx=5, pady=3)
        figure_bold_var = tk.BooleanVar(value=False)
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
    
    def load_current_config(self):
        """加载当前配置到控件"""
        # 使用配置管理器中的当前配置
        if self.config_manager.format_config:
            self._apply_config(self.config_manager.format_config)
        else:
            self.load_defaults()
    
    def load_defaults(self):
        """加载默认配置"""
        default_config = self.config_manager.get_default_format_config()
        self._apply_config(default_config)
    
    def _apply_config(self, config):
        """应用配置到控件"""
        for key, value in config.items():
            # 处理输入框和下拉框的值
            widget = self.entries.get(key)
            if widget:
                if "_size" in key:
                    display_val = self.config_manager.pt_to_font_size(value)
                    widget.set(display_val)
                elif isinstance(widget, ttk.Combobox):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(value))
            
            # 处理复选框的值
            checkbox_var = self.checkboxes.get(key)
            if checkbox_var is not None:
                checkbox_var.set(bool(value))
    
    def collect_config(self):
        """收集控件中的配置"""
        config = {}
        # 收集输入框和下拉框的值
        for key, widget in self.entries.items():
            value = widget.get().strip()
            if "_size" in key:
                if value in self.config_manager.font_size_map:
                    config[key] = self.config_manager.font_size_map[value]
                else:
                    try: 
                        config[key] = float(value)
                    except (ValueError, TypeError):
                        if self.log_callback:
                            self.log_callback(f"警告: 无效的字号值 '{value}' for '{key}'. 使用默认值 16pt。")
                        config[key] = 16
            else:
                try: 
                    config[key] = float(value) if '.' in value else int(value)
                except (ValueError, TypeError): 
                    config[key] = value
        
        # 收集复选框的值
        for key, checkbox_var in self.checkboxes.items():
            config[key] = checkbox_var.get()
        
        return config
    
    def load_config(self):
        """加载配置文件"""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                validated_config = self.config_manager._validate_format_config(loaded_config)
                self._apply_config(validated_config)
                messagebox.showinfo("成功", "配置已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载参数文件失败: {e}")
    
    def save_config(self):
        """保存配置到文件"""
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                config = self.collect_config()
                validated_config = self.config_manager._validate_format_config(config)
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(validated_config, f, ensure_ascii=False, indent=4)
                
                # 更新配置管理器中的配置
                self.config_manager.format_config = validated_config
                
                # 通知父窗口配置已更新
                if hasattr(self.parent, 'on_settings_updated'):
                    self.parent.on_settings_updated(validated_config)
                
                messagebox.showinfo("成功", f"配置已保存至 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {e}")
    
    def save_default_config(self):
        """保存为默认配置"""
        try:
            config = self.collect_config()
            validated_config = self.config_manager._validate_format_config(config)
            self.config_manager.save_config(validated_config)
            
            # 更新配置管理器中的配置
            self.config_manager.format_config = validated_config
            
            # 通知父窗口配置已更新
            if hasattr(self.parent, 'on_settings_updated'):
                self.parent.on_settings_updated(validated_config)
            
            messagebox.showinfo("成功", "当前配置已保存为默认配置。\n下次启动软件时将自动加载。")
        except Exception as e:
            messagebox.showerror("错误", f"保存默认配置失败: {e}")
    
    # confirm_settings方法已移除，因为不再需要确认按钮