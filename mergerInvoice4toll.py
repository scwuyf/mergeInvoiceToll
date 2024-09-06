import os, io, sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import send2trash
import pandas as pd
import pdfplumber
import fitz
import pypdf
from pypdf import PdfMerger, PdfReader, PdfWriter, PageObject
import re
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from datetime import datetime  # 导入 datetime 模块
import configparser
import glob
import subprocess
import platform

def create_config_file():
    # 创建ConfigParser对象
    config = configparser.ConfigParser()
    
    config.add_section('config')
    # 添加配置节
    config.set('config', 'Binding_Position', '1')
    config.set('config', 'summary_page_position', '1')
    config.set('config', 'header_or_footer', '1')
    config.set('config', 'system_font_path', r'c:\windows\Fonts\simhei.ttf')
        
    # 写入配置文件
    with open('settingtoll.ini', 'w') as configfile:
        config.write(configfile)
    
def read_config_file():
    # 创建ConfigParser对象
    config = configparser.ConfigParser()
    if not os.path.exists('settingtoll.ini'):
        create_config_file()
    config.read('settingtoll.ini')
        
    try:
        # 从配置文件中获取全局变量
        global Binding_Position, summary_page_position, header_or_footer, system_font_path
        Binding_Position = config.getint('config', 'Binding_Position')
        summary_page_position = config.getint('config', 'summary_page_position')
        header_or_footer = config.getint('config', 'header_or_footer')
        system_font_path = config.get('config', 'system_font_path')
        
        # 检查设置项的合法性
        if Binding_Position not in [1, 2]:
            raise ValueError("Binding_Position 的值必须是 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
        if summary_page_position not in [1, 2]:
            raise ValueError("summary_page_position 的值必须是 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
        if header_or_footer not in [1, 2]:
            raise ValueError("header_or_footer 的值必须是 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        show_error_message(str(e))
        sys.exit()
        
def show_error_message(message):
    """显示错误消息并退出程序"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    messagebox.showerror("配置文件错误", message)
    root.destroy()

    if not os.path.exists('settingtoll.ini'):
        create_config_file()

# 全局变量来收集所有处理过的汇总单号
summary_numbers = []

# 检查文件是否存在，如果不存在则创建
if not os.path.exists('settingtoll.ini'):
    create_config_file()
# 读取配置文件
read_config_file()

def select_folder():
    """选择文件夹并开始处理"""
    global folder_path
    folder_path = filedialog.askdirectory()

    if not folder_path:  # 用户点击了取消
       return  # 退出函数
   	    
    # 检查文件夹内是否存在tempfolder
    tempfolder_path = os.path.join(folder_path, 'tempfolder')
    if os.path.exists(tempfolder_path):
        answer = messagebox.askyesno("警告", "当前文件夹内存在临时文件夹tempfolder，请先检查是否有未保存的文件，不清空临时文件夹将造成合并发票错误\n\n点\"是（Y）\"将清空临时文件夹，点\"否（N）\"取消操作\n\n")
        if answer:  # 用户点击确定
            # 使用绝对路径
            tempfolder_abs_path = os.path.abspath(tempfolder_path)
            try:
                send2trash.send2trash(tempfolder_abs_path)
            except OSError as e:
                messagebox.showwarning("警告", "临时文件夹中的某些文件可能正在被其他应用打开，请关闭后再试。")
                return  # 返回文件夹选择界面
        else:  # 用户点击否
            return
    global output_folder
    output_folder = os.path.join(folder_path, 'tempfolder').replace("\\", "/")  # 使用正斜杠作为路径分隔符
    check_files(folder_path)
    update_button_state()  # 直接调用更新按钮状态

def check_files(folder_path):
    summary_pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf') and '通行费电子票据汇总单(票据)' in f]
    if not summary_pdf_files:
        messagebox.showwarning("警告", "当前文件夹中没找到通行费电子票据汇总单，请检查发票文件是否齐全")
        select_folder()
    else:
        process_files(folder_path, summary_pdf_files)

def calculate_table_to_page_ratio(page):
    """计算表格高度与页面高度的比例"""
    total_table_height = 0
    for block in page.get_text("dict")["blocks"]:
        if block["type"] == 1:  # 表格类型
            top, bottom = block["bbox"][1], block["bbox"][3]
            total_table_height += abs(bottom - top)
    page_height = page.rect.height
    ratio = total_table_height / page_height if page_height > 0 else 0
    return ratio

def process_summarysheet(pdf_path, output_path, summary_number):
    """处理 PDF 文件"""
    doc = fitz.open(pdf_path)
    new_doc = fitz.open()

    a4_width, a4_height = 595, 842  # A4 像素点尺寸
    top_margin = 60  # 上边距 60 点

    for page in doc:  # 处理整个文档中的所有页面
        ratio = calculate_table_to_page_ratio(page)
        page_width = page.rect.width
        page_height = page.rect.height
        
        if ratio >= 0.8:
            # 如果表格占比大于等于80%，则按原页面尺寸的82%缩放
            scale = 0.82
        else:
            # 否则，保持原尺寸不变
            scale = min(a4_width / page_width, (a4_height-top_margin) / page_height)

        # 创建新的 PDF 页面
        new_page = new_doc.new_page(width=a4_width, height=a4_height)
        
        # 计算变换矩阵
        mat = fitz.Matrix(scale, scale)
        # 设置新的页面内容位置
        rect = fitz.Rect(0, top_margin, a4_width, a4_height)
        new_page.show_pdf_page(rect, doc, page.number)
    
    # 保存处理后的文件
    #new_pdf_path = os.path.join(output_path, f"{summary_number}_{summary_page_position}_票据汇总单_temp4prt.pdf")
    # 获取原始文件名，并在其基础上添加 _temp4prt 后缀
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    if '1piece' in base_name:
        new_pdf_path = os.path.join(output_path, f"{base_name}_temp4prt.pdf")
    else:
        new_pdf_path = os.path.join(output_path, f"{summary_number}_{summary_page_position}_票据汇总单_temp4prt.pdf")
    new_doc.save(new_pdf_path)
    doc.close()
    new_doc.close()
    #print(f"已将文件 {base_name} 处理保存为 {new_pdf_path}")

    # 检测文件名是否包含 "1piece"
    if '1piece' in os.path.basename(pdf_path):
        # 合并具有相同summary_number且文件名以ipiece结尾的全部文件
        merged_pdf_path = merge_1piece_files(output_folder, summary_number, summary_page_position)
        #if merged_pdf_path:
            # 调整页面到A4尺寸
            #adjusted_pdf_path = adjust_pages_to_a4(merged_pdf_path, output_folder, summary_number, binding_position)
            #print(f"已调整 {merged_pdf_path} 至 A4 尺寸，保存为 {adjusted_pdf_path}")
        #else:
            #print(f"文件名 {os.path.basename(pdf_path)} 不包含 ipiece，不进行额外处理。")
    
    # 清理原始文件
    #shutil.move(pdf_path, send2trash.send2trash(pdf_path))

def merge_1piece_files(output_folder, summary_number, summary_page_position):
    # 寻找以 _1piece_temp4prt 结尾的文件
    temp_files = [f for f in os.listdir(output_folder) 
                  if f.endswith('_1piece_temp4prt.pdf')]
    if not temp_files:
        print(f"没有找到任何以 {summary_number} 开头且以 _temp4prt 结尾的 PDF 文件")
        return
    # 假设只有一个这样的文件
    temp_file = temp_files[0]
    temp_pdf_path = os.path.join(output_folder, temp_file)
    
    # 打开临时文件
    temp_doc = fitz.open(temp_pdf_path)

    # 正则表达式匹配8位数字
    invoice_number_pattern = r'\d{8}'
        
    # 寻找以 summary_number 开头的 invoice.pdf 文件
    summary_invoice_files = [f for f in os.listdir(output_folder) 
                             if f.startswith(summary_number) and f.endswith('第一次临时合并.pdf')]
    
    if not summary_invoice_files:
        print(f"没有找到任何以 {summary_number} 开头且以 invoice_number 结尾的 PDF 文件")
        temp_doc.close()
        return
    
    # 假设只有一个这样的文件
    summary_invoice_file = summary_invoice_files[0]
    summary_invoice_path = os.path.join(output_folder, summary_invoice_file)
    
    # 打开 summary_invoice 文件
    summary_invoice_doc = fitz.open(summary_invoice_path)
    
    # 创建一个新的 PDF 文档
    merged_doc = fitz.open()
    
    # 遍历临时文件中的每一页
    for temp_page in temp_doc:
        # 获取临时页面的高度
        temp_page_height = temp_page.rect.height
        
        # 遍历 summary_invoice 文件
        for summary_invoice_file in summary_invoice_files:
            summary_invoice_path = os.path.join(output_folder, summary_invoice_file)
            
            # 打开 summary_invoice 文件
            summary_invoice_doc = fitz.open(summary_invoice_path)
            
            # 获取 summary_invoice 页面
            summary_invoice_page = summary_invoice_doc[0]
            
            # 计算 summary_invoice 页面的高度
            invoice_page_height = summary_invoice_page.rect.height
            
            # 创建新的 PDF 页面
            new_page = merged_doc.new_page(width=temp_page.rect.width, height=temp_page.rect.height)
            
            # 显示临时页面
            new_page.show_pdf_page(fitz.Rect(0, 0, temp_page.rect.width, temp_page_height), temp_doc, temp_page.number)
            
            # 计算底部对齐的位置
            bottom_offset = temp_page.rect.height - invoice_page_height - 30
            
            # 创建缩放矩阵
            scale_matrix = fitz.Matrix(0.93, 0.93)
            
            # 显示 summary_invoice 页面
            new_page.show_pdf_page(
                fitz.Rect(0, bottom_offset, temp_page.rect.width - 40, bottom_offset + invoice_page_height * 0.93),
                summary_invoice_doc, summary_invoice_page.number,
                clip=fitz.Rect(0, 0, temp_page.rect.width, invoice_page_height)                
            )
            
            summary_invoice_doc.close()
    
    # 保存最终合并的文件
    final_pdf_path = os.path.join(output_folder, f"{summary_number}_{summary_page_position}_票据汇总单_temp4prt.pdf")
    merged_doc.save(final_pdf_path)
    
    temp_doc.close()    
    merged_doc.close()
    return final_pdf_path

def process_files(folder_path, summary_pdf_files):
    for pdf_file in summary_pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        new_pdf_path=os.path.join(folder_path, 'tempfolder')
        try:
            with pdfplumber.open(pdf_path) as pdf:
                summary_number = extract_summary_number(pdf)
                if summary_number:
                   tables = extract_tables_from_pdf(pdf)
                   process_tables(tables, summary_number, folder_path)
                   for i, df in enumerate(tables):
                       #print(len(df))
                       if len(df) <= 4:
                          new_file_name = f"{summary_number}_票据汇总单_1piece.pdf"
                       else:
                          new_file_name = f"票据汇总单_{summary_number}.pdf"
                   new_file_path = os.path.join(output_folder, new_file_name)
                   shutil.copy(pdf_path, new_file_path)
                   #print(f"已将文件 {pdf_file} 复制并重命名为 {new_file_name}")

                   # 处理PDF文件
                   process_summarysheet(new_file_path, new_pdf_path, summary_number)                   
                
        except Exception as e:
            messagebox.showerror("错误", f"处理文件 {pdf_file} 时发生错误: {e}")

def extract_summary_number(pdf):
    first_page_text = pdf.pages[0].extract_text()
    lines = first_page_text.splitlines()
    if len(lines) >= 3:
        third_line_text = lines[2]
        match = re.search(r'汇总单号:\s+(\d{16})', third_line_text)
        if match:
            summary_number = match.group(1)
            return summary_number

def extract_tables_from_pdf(pdf):
    tables = []
    is_first_page = True  # 标记是否为第一页
    for page in pdf.pages:
        try:
            table = page.extract_table()
            if table:
                # 确保表头和数据分开
                #header = table[0]
                data = table[0:]
                
                # 创建DataFrame
                #df = pd.DataFrame(data, columns=header)
                df = pd.DataFrame(data)
                
                # 如果是第一页，删除前4行
                if is_first_page:
                    df = df.iloc[4:]
                    is_first_page = False
                
                tables.append(df)
        except Exception as e:
            print(f"提取页面 {page.page_number} 的表格时发生错误: {e}")
    return tables
    
def process_tables(tables, summary_number, folder_path):
    all_data = []
    is_first_page = True
    total_pages = len(tables)
    
    for page_number, df in enumerate(tables, start=1):
        try:
            # 确保DataFrame至少有一行数据
            if df.empty:
                print(f"页面 {page_number} 中的表格为空，跳过此页。")
                continue
            
            # 第一页特殊处理
            if is_first_page:
                # 删除前4行
                #df = df.iloc[4:]

                # 将第4列复制到第3列
                df[2] = df[3]
                is_first_page = False

            # 删除第4到最后一个列
            df = df.drop(df.columns[3:], axis=1)
            
            # 如果是最后一页，删除最后3行
            if page_number == total_pages:
                df = df.iloc[:-3]
            
            # 替换空白单元格为 NaN
            df.replace('', pd.NA, inplace=True)
            
            # 删除所有空行
            df.dropna(how='all', inplace=True)
            
            # 重置列名
            df.columns = ['票据序号', '发票代码', '发票号码']
            
            # 收集数据
            all_data.append(df)
        except Exception as e:
            print(f"处理表格时发生错误: {e}")
    
    # 合并所有数据
    if all_data:  # 确保all_data非空
        combined_df = pd.concat(all_data, ignore_index=True)
    else:
        print("没有有效的表格数据可合并。")
        return  # 返回，不执行后续操作
    
    # 输出Excel文件
    output_folder = os.path.join(folder_path, 'tempfolder')
    os.makedirs(output_folder, exist_ok=True)
    #excel_path = os.path.join(output_folder, f"{summary_number}.xlsx")
    #combined_df.to_excel(excel_path, index=False)

    # 进一步处理
    match_invoices(combined_df['发票号码'], folder_path, summary_number, output_folder)
    
    # 添加汇总单号到全局列表
    summary_numbers.append(summary_number)

def merge_all_print_versions(output_folder):
    """合并所有打印版文件"""
    # 初始化PDF Writer
    writer = PdfWriter()
    
    # 构建目标文件名
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")  # 当前时间戳
    target_pdf_name = f"通行费发票按汇总单号合并的打印版本_{current_time}.pdf"
    target_pdf_path = os.path.join(output_folder, target_pdf_name)
    
    # 遍历所有汇总单号
    for summary_number in summary_numbers:
        # 查找所有符合要求的PDF文件
        pdf_files = [f for f in os.listdir(output_folder) if f.startswith(summary_number) and f.endswith('temp4prt.pdf') and ('1piece') not in f]
        if not pdf_files:
            print(f"没有找到任何以 {summary_number}开头的 temp4prt.PDF 文件")
            continue
        
        # 合并所有找到的PDF文件
        for pdf_file in pdf_files:
            pdf_path = os.path.join(output_folder, pdf_file)
            with open(pdf_path, "rb") as fin:
                reader = PdfReader(fin)
                for page in reader.pages:
                    writer.add_page(page)
    
    # 写入合并后的PDF文件
    with open(target_pdf_path, "wb") as fout:
        writer.write(fout)
    
    print(f"所有打印版文件已合并完成，文件已保存为: {os.path.join(output_folder, target_pdf_name)}")
    return target_pdf_path

def draw_progress_bar(current, total, length=50):
    percent = int(current / total * 100)
    filled_length = int(length * current // total)
    bar = '▮' * filled_length + '-' * (length - filled_length)
    sys.stdout.write(f'\r[{bar}] {percent}% Complete')
    sys.stdout.flush()

def match_invoices(invoice_numbers, folder_path, summary_number, output_folder):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf') and "汇总单" not in f]
    total_files = len(pdf_files)
    
    for index, invoice_number in enumerate(invoice_numbers):
        if len(invoice_number) == 8 and invoice_number.isdigit():
            for pdf_file in pdf_files:
                pdf_path = os.path.join(folder_path, pdf_file)
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        pdf_invoice_number = extract_invoice_number(pdf)
                        if pdf_invoice_number is None:
                            print(f"无法从文件 {pdf_file} 中提取发票号码，这张发票不是汇总单号{summary_number}里的发票")
                            continue
                        if pdf_invoice_number == invoice_number:
                            new_file_name = f"{summary_number}_{invoice_number}.pdf"
                            new_file_path = os.path.join(output_folder, new_file_name)
                            shutil.copy(pdf_path, new_file_path)
                except Exception as e:
                    messagebox.showerror("错误", f"处理文件 {pdf_file} 时发生错误: {e}")

        # 更新进度条
        draw_progress_bar(index + 1, len(invoice_numbers))

    merged_pdf_path = merge_pdfs(summary_number, output_folder)
    append_blank_page_if_needed(merged_pdf_path, output_folder, summary_number)  # 添加空白页使页数为偶数
    final_pdf_path = adjust_pages_to_a4(merged_pdf_path, output_folder, summary_number)  # 调整页面大小并合并到A4
    print("汇总单号：",summary_number," 所有文件已处理完毕")

def extract_invoice_number(pdf):
    first_page_text = pdf.pages[0].extract_text()
    lines = first_page_text.splitlines()
    if len(lines) >= 2:
        second_line_text = lines[1]
        match = re.search(r'发票号码:\s*(\d{8})', second_line_text)
        if match:
            invoice_number = match.group(1)
            return invoice_number

def merge_pdfs(summary_number, output_folder):
    # 使用 PdfWriter 替换 PdfMerger
    writer = PdfWriter()
    pdf_files = [f for f in os.listdir(output_folder) if f.startswith(summary_number) and f.endswith('.pdf')]
    if not pdf_files:
        raise ValueError(f"没有找到任何以 {summary_number} 开头的 PDF 文件")
    for pdf_file in pdf_files:
        pdf_path = os.path.join(output_folder, pdf_file)
        with open(pdf_path, "rb") as fin:
            reader = PdfReader(fin)
            for page in reader.pages:
                writer.add_page(page)
    
    merged_pdf_path = os.path.join(output_folder, f"{summary_number}_第一次临时合并.pdf")
    with open(merged_pdf_path, "wb") as fout:
        writer.write(fout)
    
    return merged_pdf_path

def create_blank_page(output_folder, summary_number):
    # 使用reportlab创建一个空白页
    
    blank_page_path = os.path.join(output_folder, f"blank_page_{summary_number}.pdf")
    c = canvas.Canvas(blank_page_path, pagesize=(610,394))  # 普通发票处理后尺寸是215.9mm x 139.7mm，换算成像素点等于 612 x 396
    c.showPage()
    c.save()
    return blank_page_path

def append_blank_page_if_needed(merged_pdf_path, output_folder, summary_number):
    # 检查合并后的PDF文件的页数
    with open(merged_pdf_path, "rb") as fin:
        reader = PdfReader(fin)
        num_pages = len(reader.pages)
        if num_pages == 1:
           return
        # 如果页数为奇数，则添加一个空白页
        if num_pages % 2 != 0:
            blank_page_path = create_blank_page(output_folder, summary_number)
            
            # 使用PdfWriter来添加空白页
            writer = PdfWriter()
            with open(merged_pdf_path, "rb") as fin:
                for page in PdfReader(fin).pages:
                    writer.add_page(page)
            with open(blank_page_path, "rb") as fin:
                writer.add_page(PdfReader(fin).pages[0])
            
            # 保存修改后的PDF文件
            with open(merged_pdf_path, "wb") as fout:
                writer.write(fout)
            
def create_pdf_with_headerfooter(text, width, height, font_size=8, margin=28.346645669):
    """生成一个带有页眉或页脚的 PDF 页面，支持中文"""
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))
    
    # 注册中文字体
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    
    pdfmetrics.registerFont(TTFont('ChineseFont', system_font_path))
    
    # 设置页脚字体
    can.setFont("ChineseFont", font_size)
    
    # 添加中文页眉页脚
    if header_or_footer == 1:
        can.drawString(margin, height-margin, text)
    else:
        can.drawString(margin, margin, text)
    
    # 保存并关闭
    can.save()
    
    # 读取生成的 PDF 内容
    packet.seek(0)
    
    return PdfReader(packet) 

def adjust_pages_to_a4(merged_pdf_path, output_folder, summary_number):
    # 创建一个新的PDF Writer对象
    writer = PdfWriter()
    width, height = A4
    print("    ")
    
    if Binding_Position == 1:
       #print(f"装订选项 = 短边装订（A4 横向）")
       #短边装订边距
       top_margin = 70.86614173  # 上边距
       bottom_margin = 28.34645669  # 下边距
       left_margin = 7.65354331  # 左边距
       right_margin = 26.36220472  # 右边距
   
       # 计算实际的可打印区域
       effective_width = width - left_margin - right_margin
       effective_height = height - top_margin - bottom_margin
       #print(f" A4 页面有效宽度= {effective_width}, 有效高度= {effective_height}")
       # 获取原始发票页面的宽度和高度
       original_pdf = PdfReader(merged_pdf_path)
       original_page1_widths = [page.mediabox.width for page in original_pdf.pages[::2]]
       original_page2_widths = [page.mediabox.width for page in original_pdf.pages[1::2]]
   
       original_page1_heights = [page.mediabox.height for page in original_pdf.pages[::2]]
       original_page2_heights = [page.mediabox.height for page in original_pdf.pages[1::2]]
       #print(f" 发票原始页面1宽度= {original_page1_widths}, 高度= {original_page1_heights}")
       #print(f" 发票原始页面2宽度= {original_page2_widths}, 高度= {original_page2_heights}")
       
       # 打开合并后的PDF文件
       with open(merged_pdf_path, "rb") as fin:
           reader = PdfReader(fin)
           if len(reader.pages) == 1:
               return
               
           # 每次处理两个页面
           for i in range(0, len(reader.pages), 2):
               new_page = PageObject.create_blank_page(width=width, height=height)
               
               # 第一个页面
               if i < len(reader.pages):
                   page1 = reader.pages[i]
                   # 第二个页面
                   page2 = reader.pages[i + 1] if i + 1 < len(reader.pages) else None
                   
                   # 计算两个页面合并后的有效宽度和高度，使用原始发票页面的高度来计算
                   total_height = original_page1_heights[i // 2] + (original_page2_heights[i // 2] if page2 else 0)
                   #print(f" 合并2张发票页面高度=页面1高度 {original_page1_heights[i // 2]} + 页面2高度 {original_page2_heights[i // 2]} = {total_height}")
                   scale = min(effective_width / float(original_pdf.pages[i].mediabox.width), effective_height / float(total_height))
                   
                   # 缩放第一个页面
                   page1.scale_by(scale)
                   # 调整位置
                   x_offset = effective_width - page1.mediabox.width
                   #print(f" 发票第一组1页X偏移点=A4有效宽度{effective_width} - 发票页缩小宽度 {page1.mediabox.width} = {x_offset}")
                   y_offset = effective_height - page1.mediabox.height
                   #print(f" 发票第一组1页Y偏移点=A4有效高度{effective_height} - 发票页缩小高度 {page1.mediabox.height} = {y_offset}")
                   new_page.merge_page(page1)
                   new_page.add_transformation((1, 0, 0, 1, x_offset, y_offset))
                   #new_page.add_transformation((1, 0, 0, 1, 0, 395))
                   
                   # 如果存在第二个页面, 则缩放并调整位置
                   if page2:
                       page2.scale_by(scale)
                       # 调整位置
                       x_offset2 = effective_width - page2.mediabox.width
                       #print(f" 发票第一组2页X偏移点=A4有效宽度{effective_width} - 发票页缩小宽度 {page2.mediabox.width} = {x_offset2}")
                       y_offset2 = effective_height - (page1.mediabox.height + page2.mediabox.height)
                       #print(f" 发票第一组2页Y偏移点=A4有效高度{effective_height} - 发票页缩小高度 {page2.mediabox.height} = {y_offset2}")
                       new_page.merge_page(page2)
                       new_page.add_transformation((1, 0, 0, 1, x_offset2, y_offset2))
                       #new_page.add_transformation((1, 0, 0, 1, 0, 25))

                   print(f"处理完成页面 pages {i+1} and {i+2 if page2 else 'N/A'}:")                       
                   #print(f"  缩放比例: {scale}")
                   #print(f"  发票第一组1页 X offset: {x_offset}, Y offset: {y_offset}")
                   #if page2:
                   #    print(f"  发票第一组2页 X offset: {x_offset2}, Y offset: {y_offset2}")

                   # 创建带有页脚的 PDF 页面
                   text = (f"本页发票所属通行费电子票据汇总单号：{summary_number}")
                   #font_path = system_font_path
                   pdf_reader = create_pdf_with_headerfooter(text, 595, 842, font_size=8, margin=28.346645669)
                   
                   # 合并页脚
                   new_page.merge_page(pdf_reader.pages[0])
                   writer.add_page(new_page)
                
    else:       
       #print(f"装订选项 = 长边装订（A4 纵向）")
       # 长边装订边距
       top_margin = 42.51968504  # 上边距
       bottom_margin = 28.34645669  # 下边距
       left_margin = 64.34645669  # 左边距
       right_margin = -16.15748031  # 右边距

       # 计算实际的可打印区域
       effective_width = width - left_margin - right_margin
       effective_height = height - top_margin - bottom_margin
       #print(f" A4 页面有效宽度= {effective_width}, 有效高度= {effective_height}")
       # 获取原始发票页面的高度
       original_pdf = PdfReader(merged_pdf_path)
       original_page1_widths = [page.mediabox.width for page in original_pdf.pages[::2]]
       original_page2_widths = [page.mediabox.width for page in original_pdf.pages[1::2]]
      
       original_page1_heights = [page.mediabox.height for page in original_pdf.pages[::2]]
       original_page2_heights = [page.mediabox.height for page in original_pdf.pages[1::2]]
       #print(f" 发票原始页面1宽度= {original_page1_widths}, 高度= {original_page1_heights}")
       #print(f" 发票原始页面2宽度= {original_page2_widths}, 高度= {original_page2_heights}")
       
       # 打开合并后的PDF文件
       with open(merged_pdf_path, "rb") as fin:
           reader = PdfReader(fin)
           if len(reader.pages) == 1:
              return
           
           # 每次处理两个页面
           for i in range(0, len(reader.pages), 2):
               new_page = PageObject.create_blank_page(width=width, height=height)
               
               # 第一个页面
               if i < len(reader.pages):
                   page1 = reader.pages[i]
                   # 第二个页面
                   page2 = reader.pages[i + 1] if i + 1 < len(reader.pages) else None
                   
                   # 计算两个页面合并后的有效宽度和高度
                   # 使用原始发票页面的高度来计算
                   #print(page1)   #  显示页面信息, 主要是看mediabox中的width\height
                   total_height = original_page1_heights[i // 2] + (original_page2_heights[i // 2] if page2 else 0)
                   #print(f" 合并2张发票页面高度=页面1高度 {original_page1_heights[i // 2]} + 页面2高度 {original_page2_heights[i // 2]} = {total_height}")
                   scale = min(effective_width / float(original_pdf.pages[i].mediabox.width), effective_height / float(total_height))
                   # scale = 0.93
                   
                   # 缩放第一个页面
                   page1.scale_by(scale)
                   # 调整位置
                   x_offset =  (effective_width - page2.mediabox.width * scale)
                   #print(f" 发票第一组1页X偏移点=A4有效宽度{effective_width} - 发票页缩小宽度 {page1.mediabox.width} = {x_offset}")
                   y_offset = effective_height - page1.mediabox.height
                   #print(f" 发票第一组1页Y偏移点=A4有效高度{effective_height} - 发票页缩小高度 {page1.mediabox.height} = {y_offset}")
                   new_page.merge_page(page1)
                   new_page.add_transformation((1, 0, 0, 1, x_offset, y_offset))
                   #new_page.add_transformation((1, 0, 0, 1, 0, 395))
                   
                   # 如果存在第二个页面, 则缩放并调整位置
                   if page2:
                       page2.scale_by(scale)
                       # 调整位置
                       x_offset2 = (effective_width - page2.mediabox.width * scale)/2
                       #print(f" 发票第一组2页X偏移点=A4有效宽度{effective_width} - 发票页缩小宽度 {page2.mediabox.width} = {x_offset2}")
                       y_offset2 = bottom_margin
                       #y_offset2 = effective_height - (page1.mediabox.height + page2.mediabox.height)
                       #y_offset2 = height - (top_margin + page1.mediabox.height * scale + page2.mediabox.height * scale)
                       #print(f" 发票第一组2页Y偏移点=A4有效高度{effective_height} - 发票页缩小高度 {page2.mediabox.height} = {y_offset2}")
                       new_page.merge_page(page2)
                       new_page.add_transformation((1, 0, 0, 1, x_offset2, y_offset2))
                       #new_page.add_transformation((1, 0, 0, 1, 0, 25))
                       
                   # 打印关键信息
                   #print(f"处理完成页面 pages {i+1} and {i+2 if page2 else 'N/A'}:")
                   #print(f"  缩放比例: {scale}")
                   #print(f"  发票第一组1页 X offset: {x_offset}, Y offset: {y_offset}")
                   #if page2:
                   #    print(f"  发票第一组2页 X offset: {x_offset2}, Y offset: {y_offset2}")

                   # 创建带有页眉或页脚的 PDF 页面
                   text = (f"本页发票所属通行费电子票据汇总单号：{summary_number}")
                   #font_path = system_font_path
                   pdf_reader = create_pdf_with_headerfooter(text, 595, 842, font_size=8, margin=28.346645669)
                   
                   # 合并页脚
                   new_page.merge_page(pdf_reader.pages[0])
                   writer.add_page(new_page)
    
    # 保存调整后的PDF文件
    if summary_page_position == 1:
       output_path = os.path.join(output_folder, f"{summary_number}_2_temp4prt.pdf")
    else:
       output_path = os.path.join(output_folder, f"{summary_number}_1_temp4prt.pdf")
       
    with open(output_path, 'wb') as f:
        writer.write(f)
    return output_path

def open_folder(folder_path):
    """根据不同的操作系统打开文件夹"""
    if not os.path.exists(folder_path):
        #print("文件夹不存在:", folder_path)
        return

    # 规范化路径
    folder_path = os.path.normpath(folder_path)
    
    system = platform.system()
    
    try:
        if system == 'Windows':
            # Windows
            os.startfile(folder_path)
        elif system == 'Darwin':
            # macOS
            subprocess.run(['open', folder_path], check=True)
        elif system == 'Linux':
            # Linux
            subprocess.run(['xdg-open', folder_path], check=True)
        else:
            print("不支持的操作系统:", system)
    except Exception as e:
        print("无法打开文件夹:", e)

def exit_program(root):
    # 退出程序前打开选择的文件夹
    if 'folder_path' in globals():
        open_folder(folder_path)
    root.quit()

def clean_temp_files():
    """清理过程文件"""
    if os.path.exists(output_folder):
        # 将合并后的PDF文件移动到上一级文件夹
        target_pdf_pattern = os.path.join(output_folder, "通行费发票按汇总单号合并的打印版本_*.pdf")
        matching_files = glob.glob(target_pdf_pattern)
        if matching_files:
            latest_file = max(matching_files, key=os.path.getctime)
            target_pdf_name = os.path.basename(latest_file)
            target_pdf_path_in_upper = os.path.join(folder_path, target_pdf_name)
            
            # 检查上一级文件夹中是否存在同名文件
            if os.path.exists(target_pdf_path_in_upper):
                messagebox.showwarning("警告", f"文件 {target_pdf_name} 已经存在于上一级文件夹中，跳过移动。")
            else:
                shutil.move(latest_file, folder_path)
        
        # 清理 tempfolder 文件夹内的文件
        tempfolder_abs_path = os.path.abspath(output_folder)
        
        # 删除名字中包含 temp4prt 和 临时合并 的文件
        temp4prt_pattern = os.path.join(output_folder, "*temp4prt*")
        temp_merge_pattern = os.path.join(output_folder, "*临时合并*")
        blank_pattern = os.path.join(output_folder, "*blank_page*")
        for pattern in [temp4prt_pattern, temp_merge_pattern, blank_pattern]:
            for file in glob.glob(pattern):
                os.remove(file)
        
        # 处理名字以 _1piece.pdf 结尾的文件
        one_piece_pattern = os.path.join(output_folder, "*_1piece.pdf")
        for file in glob.glob(one_piece_pattern):
            new_filename = os.path.splitext(file)[0].replace("_1piece", "") + ".pdf"
            os.rename(file, new_filename)
        
        # 处理以 票据汇总单_summary_number.pdf 形式命名的文件
        summary_pattern = os.path.join(output_folder, "票据汇总单_*.pdf")
        for file in glob.glob(summary_pattern):
            
            parts = os.path.basename(file).split("_")[1]
            summarynumber = parts.split(".pdf")[0]
            new_filename = f"{summarynumber}_通行费电子票据汇总单.pdf"
            os.rename(file, os.path.join(output_folder, new_filename))
        
        # 将 tempfolder 文件夹更名为 “汇总单发票整理”
        new_folder_name = "汇总单发票整理"
        new_folder_path = os.path.join(os.path.dirname(output_folder), new_folder_name)
        
        # 如果目标文件夹已经存在，则将文件移动到该文件夹
        if os.path.exists(new_folder_path):
            # 移动文件到目标文件夹
            for file in os.listdir(output_folder):
                src_file = os.path.join(output_folder, file)
                dst_file = os.path.join(new_folder_path, file)
                
        # 如果目标文件夹已经存在，则将文件移动到该文件夹
        if os.path.exists(new_folder_path):
            # 移动文件到目标文件夹
            for file in os.listdir(output_folder):
                src_file = os.path.join(output_folder, file)
                dst_file = os.path.join(new_folder_path, file)
                
                if os.path.exists(dst_file):
                    answer = messagebox.askyesno("文件已存在", f"文件 {file} 已经存在，是否覆盖？")
                    if answer:
                        os.remove(dst_file)
                        shutil.move(src_file, dst_file)
                    else:
                        os.remove(src_file)
                else:
                    shutil.move(src_file, dst_file)
        else:
            # 创建新文件夹并移动文件
            os.makedirs(new_folder_path)
            for file in os.listdir(output_folder):
                src_file = os.path.join(output_folder, file)
                dst_file = os.path.join(new_folder_path, file)
                shutil.move(src_file, dst_file)
        
        # 删除空的 tempfolder 文件夹
        os.rmdir(output_folder)
        
        messagebox.showinfo("清理完成", "过程文件已清理。")
    else:
        messagebox.showwarning("警告", "没有找到临时文件夹。")
        
def update_button_state():
    """更新按钮的状态"""
    if summary_numbers and os.path.exists(output_folder):
        btn_clean.config(state=tk.NORMAL)
        btn_complete.config(state=tk.NORMAL)
    else:
        btn_clean.config(state=tk.DISABLED)
        btn_complete.config(state=tk.DISABLED)
                

def main():
    global output_folder

    root = tk.Tk()
    root.title("通行费发票合并处理")
    # 设置窗口大小
    window_width = 300
    window_height = 360
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.resizable(False, False)

    # 设置按钮的大小
    button_width = 18
    button_height = 2
    
    btn_select_folder = tk.Button(root, text="选择文件夹", width=button_width, height=button_height, command=select_folder)
    btn_select_folder.pack(pady=20)

    # 添加一个“生成打印文件”按钮
    global btn_complete
    btn_complete = tk.Button(root, text="生成打印文件", width=button_width, height=button_height, command=lambda: merge_all_print_versions(output_folder) if summary_numbers else messagebox.showwarning("警告", "没有汇总单号可供处理"), state=tk.DISABLED)
    btn_complete.pack(pady=20)

    # 添加一个“清理过程文件”按钮
    global btn_clean
    btn_clean = tk.Button(root, text="清理过程文件", width=button_width, height=button_height, command=lambda: clean_temp_files() if summary_numbers else messagebox.showwarning("警告", "没有汇总单号可供处理"), state=tk.DISABLED)
    btn_clean.pack(pady=20)

    # 添加一个“退出”按钮
    btn_exit = tk.Button(root, text="退        出", width=button_width, height=button_height, command=lambda: exit_program(root))
    btn_exit.pack(pady=20)

    # 调用更新按钮状态
    update_button_state()

    root.mainloop()

if __name__ == "__main__":
    main()
    
#### mix by immt ####
