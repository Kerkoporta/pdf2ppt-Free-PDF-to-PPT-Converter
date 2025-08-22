import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pdf_to_ppt_core import pdf_to_ppt
import threading

class PDFtoPPTConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 转 PPT 转换器")
        self.root.geometry("600x300")
        self.root.resizable(False, False)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure('TLabel', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('TEntry', font=('Arial', 10))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="PDF 转 PPT 转换器", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PDF 文件选择
        ttk.Label(main_frame, text="PDF 文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.pdf_path = tk.StringVar()
        pdf_entry = ttk.Entry(main_frame, textvariable=self.pdf_path, width=50)
        pdf_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_pdf).grid(row=1, column=2, pady=5)
        
        # PPTX 文件保存位置
        ttk.Label(main_frame, text="保存位置:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.pptx_path = tk.StringVar()
        pptx_entry = ttk.Entry(main_frame, textvariable=self.pptx_path, width=50)
        pptx_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_pptx).grid(row=2, column=2, pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)
        
        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        ttk.Button(button_frame, text="开始转换", command=self.start_conversion).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=root.quit).pack(side=tk.LEFT, padx=5)
        
        # 配置列权重
        main_frame.columnconfigure(1, weight=1)
        
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            # 自动生成 PPTX 文件名
            if not self.pptx_path.get():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                dir_name = os.path.dirname(file_path)
                self.pptx_path.set(os.path.join(dir_name, f"{base_name}.pptx"))
    
    def browse_pptx(self):
        file_path = filedialog.asksaveasfilename(
            title="保存 PPTX 文件",
            defaultextension=".pptx",
            filetypes=[("PowerPoint 文件", "*.pptx")]
        )
        if file_path:
            self.pptx_path.set(file_path)
    
    def start_conversion(self):
        pdf_path = self.pdf_path.get()
        pptx_path = self.pptx_path.get()
        
        if not pdf_path:
            messagebox.showerror("错误", "请选择 PDF 文件")
            return
            
        if not pptx_path:
            messagebox.showerror("错误", "请指定 PPTX 保存路径")
            return
            
        # 启动转换线程
        self.progress.start()
        self.status_label.config(text="转换中...")
        
        thread = threading.Thread(target=self.convert, args=(pdf_path, pptx_path))
        thread.daemon = True
        thread.start()
    
    def convert(self, pdf_path, pptx_path):
        try:
            success = pdf_to_ppt(pdf_path, pptx_path)
            self.root.after(0, self.conversion_complete, success, pptx_path)
        except Exception as e:
            self.root.after(0, self.conversion_error, str(e))
    
    def conversion_complete(self, success, pptx_path):
        self.progress.stop()
        if success:
            self.status_label.config(text="转换完成!")
            messagebox.showinfo("成功", f"文件已成功转换并保存到:\n{pptx_path}")
        else:
            self.status_label.config(text="转换失败!")
            messagebox.showerror("错误", "转换失败，请检查PDF文件是否有效")
    
    def conversion_error(self, error_msg):
        self.progress.stop()
        self.status_label.config(text="转换出错!")
        messagebox.showerror("错误", f"转换过程中出错:\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoPPTConverter(root)
    root.mainloop()
