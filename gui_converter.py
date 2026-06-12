"""
PDF 转 DOCX 转换工具 GUI 应用
提供直观的文件选择和批量转换功能
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import List
import threading
import os
import sys
from pdf_to_docx_converter import convert_pdf_to_docx, PDFToDOCXError, get_pdf_converter_info


class PDFToDOCXConverterGUI:
    """PDF 转 DOCX 转换 GUI 应用，支持批量处理"""
    
    def __init__(self, root):
        """初始化 GUI 应用"""
        self.root = root
        self.root.title("PDF 转 DOCX 转换器")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 配置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 应用程序状态
        self.selected_files: List[str] = []
        self.output_directory: str = ""
        self.conversion_method = tk.StringVar(value="auto")
        self.is_converting = False
        
        # 创建 GUI 元素
        self._create_widgets()
        self._load_available_methods()
        
    def _create_widgets(self):
        """创建所有 GUI 小部件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # ═══ 标题 ═══
        title_label = ttk.Label(
            main_frame,
            text="PDF 转 DOCX 转换器",
            font=("微软雅黑", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=tk.W)
        
        # ═══ 文件选择区域 ═══
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        file_frame.columnconfigure(1, weight=1)
        
        # 单文件按钮
        single_btn = ttk.Button(
            file_frame,
            text="📄 选择单个 PDF",
            command=self._select_single_file
        )
        single_btn.grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # 文件夹按钮
        folder_btn = ttk.Button(
            file_frame,
            text="📁 选择文件夹",
            command=self._select_folder
        )
        folder_btn.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # 清空按钮
        clear_btn = ttk.Button(
            file_frame,
            text="🗑️ 清空选择",
            command=self._clear_selection
        )
        clear_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # 已选文件标签
        ttk.Label(file_frame, text="已选文件：", font=("微软雅黑", 10, "bold")).grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 5)
        )
        
        # 文件列表框
        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.files_listbox = tk.Listbox(
            list_frame,
            height=6,
            yscrollcommand=scrollbar.set,
            font=("Courier", 9)
        )
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.config(command=self.files_listbox.yview)
        
        # 文件数量标签
        self.file_count_label = ttk.Label(
            file_frame,
            text="总共文件：0",
            font=("微软雅黑", 9)
        )
        self.file_count_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # ═══ 输出目录区域 ═══
        output_frame = ttk.LabelFrame(main_frame, text="输出目录", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        output_frame.columnconfigure(1, weight=1)
        
        # 输出目录标签
        ttk.Label(output_frame, text="输出路径：").grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # 输出目录显示
        self.output_label = ttk.Label(
            output_frame,
            text="（与源文件相同目录）",
            font=("Courier", 9),
            foreground="gray"
        )
        self.output_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # 浏览输出按钮
        output_btn = ttk.Button(
            output_frame,
            text="📁 浏览",
            command=self._select_output_directory
        )
        output_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # ═══ 转换设置区域 ═══
        settings_frame = ttk.LabelFrame(main_frame, text="转换设置", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        settings_frame.columnconfigure(1, weight=1)
        
        # 转换方法选择
        ttk.Label(settings_frame, text="转换方法：").grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # 创建方法单选按钮
        methods_frame = ttk.Frame(settings_frame)
        methods_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5)
        
        self.method_radios = {}
        method_names = {
            'auto': '自动',
            'pdfplumber': 'pdfplumber',
            'ocr': 'OCR识别',
            'libreoffice': 'LibreOffice',
            'pymupdf': 'PyMuPDF'
        }
        
        for idx, (method, display_name) in enumerate(method_names.items()):
            radio = ttk.Radiobutton(
                methods_frame,
                text=display_name,
                variable=self.conversion_method,
                value=method
            )
            radio.grid(row=0, column=idx, padx=5)
            self.method_radios[method] = radio
        
        # 可用方法信息
        ttk.Label(settings_frame, text="可用方法：").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        self.available_methods_label = ttk.Label(
            settings_frame,
            text="",
            font=("Courier", 9),
            foreground="green"
        )
        self.available_methods_label.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=(10, 0))
        
        # 进度区域
        progress_frame = ttk.LabelFrame(settings_frame, text="进度", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        progress_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 进度标签
        self.progress_label = ttk.Label(
            progress_frame,
            text="准备转换",
            font=("微软雅黑", 9)
        )
        self.progress_label.grid(row=1, column=0, sticky=tk.W, padx=5)
        
        # ═══ 操作按钮 ═══
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        button_frame.columnconfigure(1, weight=1)
        
        # 转换按钮
        self.convert_btn = ttk.Button(
            button_frame,
            text="🚀 开始转换",
            command=self._start_conversion
        )
        self.convert_btn.grid(row=0, column=0, padx=5)
        
        # 停止按钮
        self.stop_btn = ttk.Button(
            button_frame,
            text="⏹️ 停止",
            command=self._stop_conversion,
            state=tk.DISABLED
        )
        self.stop_btn.grid(row=0, column=1, padx=5)
        
        # 打开输出按钮
        open_btn = ttk.Button(
            button_frame,
            text="📂 打开输出文件夹",
            command=self._open_output_folder
        )
        open_btn.grid(row=0, column=2, sticky=tk.E, padx=5)
        
        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=1, pady=1)
        self.root.columnconfigure(0, weight=1)
    
    def _load_available_methods(self):
        """加载并显示可用的转换方法"""
        try:
            info = get_pdf_converter_info()
            available = [method for method, available in info.items() if available]
            
            if available:
                method_display = {
                    'pdfplumber': 'pdfplumber',
                    'ocr': 'OCR识别',
                    'libreoffice': 'LibreOffice',
                    'pymupdf': 'PyMuPDF'
                }
                display_names = [method_display.get(m, m) for m in available]
                self.available_methods_label.config(
                    text=f"✅ {', '.join(display_names)}",
                    foreground="green"
                )
            else:
                self.available_methods_label.config(
                    text="❌ 无可用方法 - 请安装依赖",
                    foreground="red"
                )
                self.convert_btn.config(state=tk.DISABLED)
                
        except Exception as e:
            self.available_methods_label.config(
                text=f"❌ 检查方法出错: {str(e)}",
                foreground="red"
            )
    
    def _select_single_file(self):
        """选择单个 PDF 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.selected_files = [file_path]
            self._update_file_list()
    
    def _select_folder(self):
        """选择文件夹并添加其中所有 PDF"""
        folder_path = filedialog.askdirectory(
            title="选择包含 PDF 文件的文件夹"
        )
        
        if folder_path:
            folder = Path(folder_path)
            pdf_files = list(folder.glob("*.pdf"))
            
            if not pdf_files:
                messagebox.showwarning(
                    "未找到 PDF",
                    f"在以下目录中未找到 PDF 文件：\n{folder_path}"
                )
                return
            
            self.selected_files = [str(f) for f in sorted(pdf_files)]
            self._update_file_list()
    
    def _clear_selection(self):
        """清空所有选择的文件"""
        self.selected_files = []
        self.files_listbox.delete(0, tk.END)
        self.file_count_label.config(text="总共文件：0")
        self.status_var.set("选择已清空")
    
    def _select_output_directory(self):
        """选择输出目录"""
        folder_path = filedialog.askdirectory(
            title="选择输出目录"
        )
        
        if folder_path:
            self.output_directory = folder_path
            self.output_label.config(
                text=folder_path,
                foreground="black"
            )
            self.status_var.set(f"输出目录：{folder_path}")
    
    def _update_file_list(self):
        """更新文件列表框"""
        self.files_listbox.delete(0, tk.END)
        
        for file_path in self.selected_files:
            display_path = Path(file_path).name
            self.files_listbox.insert(tk.END, display_path)
        
        count = len(self.selected_files)
        self.file_count_label.config(text=f"总共文件：{count}")
        self.status_var.set(f"已选择 {count} 个文件")
    
    def _start_conversion(self):
        """启动转换过程"""
        if not self.selected_files:
            messagebox.showwarning("无文件", "请选择要转换的 PDF 文件")
            return
        
        # 禁用按钮
        self.convert_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.is_converting = True
        
        # 在后台线程中启动转换
        thread = threading.Thread(target=self._conversion_worker)
        thread.daemon = True
        thread.start()
    
    def _stop_conversion(self):
        """停止转换过程"""
        self.is_converting = False
        self.convert_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("用户停止转换")
        self.progress_var.set(0)
    
    def _conversion_worker(self):
        """批量转换工作线程"""
        method = self.conversion_method.get()
        total = len(self.selected_files)
        successful = 0
        failed = 0
        errors = []
        
        try:
            for idx, pdf_file in enumerate(self.selected_files, 1):
                if not self.is_converting:
                    break
                
                try:
                    # 确定输出路径
                    pdf_path = Path(pdf_file)
                    if self.output_directory:
                        output_path = Path(self.output_directory) / (pdf_path.stem + '.docx')
                    else:
                        output_path = pdf_path.parent / (pdf_path.stem + '.docx')
                    
                    # 更新进度
                    progress = int((idx - 1) / total * 100)
                    self.progress_var.set(progress)
                    self.progress_label.config(
                        text=f"正在转换 {idx}/{total}：{pdf_path.name}"
                    )
                    self.status_var.set(f"正在转换：{pdf_path.name}")
                    self.root.update()
                    
                    # 执行转换
                    result = convert_pdf_to_docx(
                        pdf_file,
                        str(output_path),
                        method=method,
                        timeout=300
                    )
                    
                    successful += 1
                    
                except PDFToDOCXError as e:
                    failed += 1
                    error_msg = f"{Path(pdf_file).name}：{str(e)}"
                    errors.append(error_msg)
                except Exception as e:
                    failed += 1
                    error_msg = f"{Path(pdf_file).name}：{str(e)}"
                    errors.append(error_msg)
            
            # 转换完成
            self.progress_var.set(100)
            self.progress_label.config(
                text=f"转换完成：{successful} 个成功，{failed} 个失败"
            )
            self.status_var.set("转换完成")
            
            # 显示结果对话框
            self.root.after(0, self._show_conversion_results, successful, failed, errors)
            
        except Exception as e:
            self.status_var.set(f"错误：{str(e)}")
            messagebox.showerror("转换错误", f"意外错误：{str(e)}")
        
        finally:
            self.convert_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_converting = False
    
    def _show_conversion_results(self, successful, failed, errors):
        """显示转换结果对话框"""
        total = successful + failed
        
        if failed == 0:
            messagebox.showinfo(
                "✅ 转换完成",
                f"成功转换了 {successful} 个 PDF 文件为 DOCX！"
            )
        else:
            error_details = "\n".join(errors[:5])  # 显示前5个错误
            if len(errors) > 5:
                error_details += f"\n... 还有 {len(errors) - 5} 个错误"
            
            messagebox.showwarning(
                "⚠️ 转换完成（有错误）",
                f"成功：{successful}/{total}\n"
                f"失败：{failed}/{total}\n\n"
                f"错误：\n{error_details}"
            )
    
    def _open_output_folder(self):
        """在文件管理器中打开输出文件夹"""
        folder = self.output_directory if self.output_directory else str(Path.home())
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(folder)
            elif os.name == 'posix':  # macOS 和 Linux
                os.system(f'open "{folder}"' if sys.platform == 'darwin' else f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：{str(e)}")


def main():
    """GUI 应用的主入口"""
    root = tk.Tk()
    app = PDFToDOCXConverterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
