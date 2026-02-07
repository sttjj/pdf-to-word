import os
import argparse
import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter

def pdf_to_word(input_pdf, output_docx=None):
    """
    将PDF文件转换为Word文档
    
    Args:
        input_pdf (str): 输入PDF文件路径
        output_docx (str): 输出Word文件路径（可选）
        
    Returns:
        str: 转换后的Word文件路径
    """
    # 验证输入文件是否存在
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"输入文件不存在: {input_pdf}")
    
    # 验证输入文件是否为PDF
    if not input_pdf.lower().endswith('.pdf'):
        raise ValueError("输入文件必须是PDF格式")
    
    # 如果未指定输出路径，使用与输入文件相同的名称和位置
    if output_docx is None:
        output_docx = os.path.splitext(input_pdf)[0] + '.docx'
    
    try:
        print(f"正在转换: {input_pdf} -> {output_docx}")
        print("正在提取PDF内容和格式...")
        
        # 创建转换器实例
        cv = Converter(input_pdf)
        # 转换整个PDF，保留格式
        cv.convert(output_docx, start=0, end=None)
        # 关闭转换器
        cv.close()
        print(f"转换成功: {output_docx}")
        return output_docx
    except Exception as e:
        print(f"转换失败: {str(e)}")
        raise

def select_file():
    """
    打开文件选择对话框，让用户选择PDF文件
    
    Returns:
        str: 选择的PDF文件路径
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    file_path = filedialog.askopenfilename(
        title="选择PDF文件",
        filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
    )
    
    return file_path

def select_output_path(default_path):
    """
    打开文件保存对话框，让用户选择输出Word文件路径
    
    Args:
        default_path (str): 默认输出路径
        
    Returns:
        str: 选择的输出Word文件路径
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    file_path = filedialog.asksaveasfilename(
        title="选择输出Word文件路径",
        defaultextension=".docx",
        filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
        initialfile=os.path.basename(default_path)
    )
    
    return file_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF转Word转换器")
    parser.add_argument("input", nargs="?", help="输入PDF文件路径")
    parser.add_argument("-o", "--output", help="输出Word文件路径（可选）")
    
    args = parser.parse_args()
    
    try:
        # 如果命令行没有提供输入文件路径，则打开文件选择对话框
        if not args.input:
            input_pdf = select_file()
            if not input_pdf:
                print("未选择文件，程序退出")
                exit(0)
        else:
            input_pdf = args.input
        
        # 如果命令行没有提供输出路径，则打开保存对话框
        if not args.output:
            default_output = os.path.splitext(input_pdf)[0] + '.docx'
            output_docx = select_output_path(default_output)
            if not output_docx:
                print("未选择输出路径，程序退出")
                exit(0)
        else:
            output_docx = args.output
        
        result = pdf_to_word(input_pdf, output_docx)
        print(f"转换完成: {result}")
    except Exception as e:
        print(f"错误: {str(e)}")
        exit(1)