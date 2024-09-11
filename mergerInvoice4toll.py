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
from datetime import datetime
import configparser
import glob
import subprocess
import platform
import random

def create_config_file():
    system = platform.system()

    config = configparser.ConfigParser()
    
    config.add_section('config')
    config.set('config', 'Binding_Position', '1')
    config.set('config', 'summary_page_position', '1')
    config.set('config', 'header_or_footer', '1')
    if system == 'Windows':
       config.set('config', 'system_font_path', r'c:\windows\Fonts\simhei.ttf')
    elif system == 'Darwin':
       config.set('config', 'system_font_path', r'/System/Library/Fonts/STHeiti Light.ttc')
        
    with open('settingtoll.ini', 'w') as configfile:
        # 添加注释
        configfile.write("# 配置文件设置\n")        
        configfile.write("# Binding_Position: 1 横向装订, 2 纵向装订\n")
        configfile.write("# summary_page_position: 1 汇总单在发票前, 2 汇总单在发票后\n")
        configfile.write("# header_or_footer: 1 汇总单号打印在发票页顶部（横向装订后被遮挡）, 2 汇总单号打印在发票页底部, 0 不打印\n")
        configfile.write("# system_font_path: 系统字体路径，Linux系统可根据实际情况自行修改\n")
        config.write(configfile)
    
def read_config_file():
    config = configparser.ConfigParser()
    if not os.path.exists('settingtoll.ini'):
        create_config_file()
    config.read('settingtoll.ini')
        
    try:
        global Binding_Position, summary_page_position, header_or_footer, system_font_path
        Binding_Position = config.getint('config', 'Binding_Position')
        summary_page_position = config.getint('config', 'summary_page_position')
        header_or_footer = config.getint('config', 'header_or_footer')
        system_font_path = config.get('config', 'system_font_path')
        
        if Binding_Position not in [1, 2]:
            raise ValueError("Binding_Position 的值必须是 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
        if summary_page_position not in [1, 2]:
            raise ValueError("summary_page_position 的值必须是 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
        if header_or_footer not in [0, 1, 2]:
            raise ValueError("header_or_footer 的值必须是0、 1 或 2\n\n请修改程序所在文件夹下configtoll.ini文件有关设置项")
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        show_error_message(str(e))
        sys.exit()
        
def show_error_message(message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("配置文件错误", message)
    root.destroy()

    if not os.path.exists('settingtoll.ini'):
        create_config_file()

# 初始化全局变量
summary_numbers = []
source_files_list = pd.DataFrame(columns=['原文件'])

if not os.path.exists('settingtoll.ini'):
    create_config_file()
read_config_file()

def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    if not folder_path:
       return

    tempfolder_path = os.path.join(folder_path, 'tempfolder')
    if os.path.exists(tempfolder_path):
        answer = messagebox.askyesno("警告", "当前文件夹内存在临时文件夹tempfolder 或 处理过的原文件保存位置“已处理原文件”\n\n请先检查是否有未保存的文件，不清空临时文件夹将造成合并发票错误\n\n点\"是（Y）\"将清空临时文件夹，点\"否（N）\"取消操作\n\n")
        if answer:
            tempfolder_abs_path = os.path.abspath(tempfolder_path)
            try:
                send2trash.send2trash(tempfolder_abs_path)
            except OSError as e:
                messagebox.showwarning("警告", "临时文件夹中的某些文件可能正在被其他应用打开，请关闭后再试。")
                return
        else:
            return
    global output_folder
    output_folder = os.path.join(folder_path, 'tempfolder').replace("\\", "/")
    check_files(folder_path)
    update_button_state()

def check_files(folder_path):
    global source_files_list
    summary_pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf') and '通行费电子票据汇总单(票据)' in f or 'apply' in f]
    if not summary_pdf_files:
        messagebox.showwarning("警告", "当前文件夹中没找到通行费电子票据汇总单，请检查发票文件是否齐全")
        select_folder()
    else:
        # 创建空的DataFrame，并赋值给全局变量source_files_list
        source_files_list = pd.DataFrame(columns=['原文件'])
        # 执行process_files
        process_files(folder_path, summary_pdf_files)

def calculate_table_to_page_ratio(page):
    total_table_height = 0
    for block in page.get_text("dict")["blocks"]:
        if block["type"] == 1:
            top, bottom = block["bbox"][1], block["bbox"][3]
            total_table_height += abs(bottom - top)
    page_height = page.rect.height
    ratio = total_table_height / page_height if page_height > 0 else 0
    return ratio

def process_summarysheet(pdf_path, output_path, summary_number):
    doc = fitz.open(pdf_path)
    new_doc = fitz.open()

    a4_width, a4_height = 595, 842
    top_margin = 60  # 上边距 60 点

    for page in doc:
        ratio = calculate_table_to_page_ratio(page)
        page_width = page.rect.width
        page_height = page.rect.height

        if ratio >= 0.8:
            scale = 0.82
        else:
            scale = min(a4_width / page_width, (a4_height-top_margin) / page_height)
        new_page = new_doc.new_page(width=a4_width, height=a4_height)
        
        mat = fitz.Matrix(scale, scale)
        rect = fitz.Rect(0, top_margin, a4_width, a4_height)
        new_page.show_pdf_page(rect, doc, page.number)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    if '1piece' in base_name:
        new_pdf_path = os.path.join(output_path, f"{base_name}_temp4prt.pdf")
    else:
        new_pdf_path = os.path.join(output_path, f"{summary_number}_{summary_page_position}_票据汇总单_temp4prt.pdf")
    new_doc.save(new_pdf_path)
    doc.close()
    new_doc.close()
    #print(f"已将文件 {base_name} 处理保存为 {new_pdf_path}")

    if '1piece' in os.path.basename(pdf_path):
        merged_pdf_path = merge_1piece_files(output_folder, summary_number, summary_page_position)

def merge_1piece_files(output_folder, summary_number, summary_page_position):
    temp_files = [f for f in os.listdir(output_folder) 
                  if f.endswith('_1piece_temp4prt.pdf')]
    if not temp_files:
        print(f"没有找到任何以 {summary_number} 开头且以 _temp4prt 结尾的 PDF 文件")
        return
    temp_file = temp_files[0]
    temp_pdf_path = os.path.join(output_folder, temp_file)
    
    temp_doc = fitz.open(temp_pdf_path)
    invoice_number_pattern = r'\d{8}'
    summary_invoice_files = [f for f in os.listdir(output_folder) 
                             if f.startswith(summary_number) and f.endswith('第一次临时合并.pdf')]
    if not summary_invoice_files:
        print(f"没有找到任何以 {summary_number} 开头且以 invoice_number 结尾的 PDF 文件")
        temp_doc.close()
        return
    summary_invoice_file = summary_invoice_files[0]
    summary_invoice_path = os.path.join(output_folder, summary_invoice_file)
    summary_invoice_doc = fitz.open(summary_invoice_path)

    merged_doc = fitz.open()

    for temp_page in temp_doc:
        temp_page_height = temp_page.rect.height
        for summary_invoice_file in summary_invoice_files:
            summary_invoice_path = os.path.join(output_folder, summary_invoice_file)
            summary_invoice_doc = fitz.open(summary_invoice_path)
            summary_invoice_page = summary_invoice_doc[0]
            invoice_page_height = summary_invoice_page.rect.height

            new_page = merged_doc.new_page(width=temp_page.rect.width, height=temp_page.rect.height)
            new_page.show_pdf_page(fitz.Rect(0, 0, temp_page.rect.width, temp_page_height), temp_doc, temp_page.number)
            bottom_offset = temp_page.rect.height - invoice_page_height - 30
            scale_matrix = fitz.Matrix(0.93, 0.93)

            new_page.show_pdf_page(
                fitz.Rect(0, bottom_offset, temp_page.rect.width - 40, bottom_offset + invoice_page_height * 0.93),
                summary_invoice_doc, summary_invoice_page.number,
                clip=fitz.Rect(0, 0, temp_page.rect.width, invoice_page_height)                
            )
            summary_invoice_doc.close()

    final_pdf_path = os.path.join(output_folder, f"{summary_number}_{summary_page_position}_票据汇总单_temp4prt.pdf")
    merged_doc.save(final_pdf_path)
    
    temp_doc.close()    
    merged_doc.close()
    return final_pdf_path

def process_files(folder_path, summary_pdf_files):
    global source_files_list
    for pdf_file in summary_pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        new_pdf_path=os.path.join(folder_path, 'tempfolder')
        try:
            with pdfplumber.open(pdf_path) as pdf:
                summary_number = extract_summary_number(pdf, pdf_path)
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

                   # 将被复制的文件名添加到source_files_list
                   new_row = pd.DataFrame({'原文件': [pdf_file]})
                   source_files_list = pd.concat([source_files_list, new_row], ignore_index=True)

                   process_summarysheet(new_file_path, new_pdf_path, summary_number)

        except Exception as e:
            messagebox.showerror("错误", f"处理文件 {pdf_file} 时发生错误: {e}")

def extract_summary_number(pdf, pdf_path):
    first_page_text = pdf.pages[0].extract_text()
    lines = first_page_text.splitlines()
    if len(lines) >= 3:
        third_line_text = lines[2]
        match = re.search(r'汇总单号:\s+(\d{16})', third_line_text)
        if match:
            summary_number = match.group(1)
            return summary_number
        else:
            # 显示yes/no对话框
            answer = messagebox.askyesno("警告", f"文件 {os.path.basename(pdf_path)} 似乎没有汇总单号, 请检查文件\n\n继续请点 yes，程序给文件自动生成一个随机号码，取消请点 No")

            if answer:
                # 如果用户点击Yes，则生成一个以“无汇总单号”加上6位随机数字的字符串
                random_suffix = ''.join([str(random.randint(0, 9)) for _ in range(6)])
                summary_number = "无汇总单号" + random_suffix
                return summary_number
            else:
                # 如果用户点击No，则直接返回None
                return None
    else:
        # 如果文本行数不足3行，则直接返回None
        return None        

def extract_tables_from_pdf(pdf):
    tables = []
    is_first_page = True
    for page in pdf.pages:
        try:
            table = page.extract_table()
            if table:
                data = table[0:]
                df = pd.DataFrame(data)

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
            if df.empty:
                print(f"页面 {page_number} 中的表格为空，跳过此页。")
                continue
            if is_first_page:
                df[2] = df[3]
                is_first_page = False
            df = df.drop(df.columns[3:], axis=1)
            if page_number == total_pages:
                df = df.iloc[:-3]
            df.replace('', pd.NA, inplace=True)
            df.dropna(how='all', inplace=True)
            df.columns = ['票据序号', '发票代码', '发票号码']

            all_data.append(df)
        except Exception as e:
            print(f"处理表格时发生错误: {e}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
    else:
        print("没有有效的表格数据可合并。")
        return
    output_folder = os.path.join(folder_path, 'tempfolder')
    os.makedirs(output_folder, exist_ok=True)

    match_invoices(combined_df['发票号码'], folder_path, summary_number, output_folder)
    summary_numbers.append(summary_number)

def merge_all_print_versions(output_folder):
    writer = PdfWriter()

    current_time = datetime.now().strftime("%Y%m%d%H%M%S")
    target_pdf_name = f"通行费发票按汇总单号合并的打印版本_{current_time}.pdf"
    target_pdf_path = os.path.join(output_folder, target_pdf_name)

    for summary_number in summary_numbers:
        pdf_files = [f for f in os.listdir(output_folder) if f.startswith(summary_number) and f.endswith('temp4prt.pdf') and ('1piece') not in f]
        if not pdf_files:
            print(f"没有找到任何以 {summary_number}开头的 temp4prt.PDF 文件")
            continue

        for pdf_file in pdf_files:
            pdf_path = os.path.join(output_folder, pdf_file)
            with open(pdf_path, "rb") as fin:
                reader = PdfReader(fin)
                for page in reader.pages:
                    writer.add_page(page)

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
    global source_files_list
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf') and "汇总单" not in f]
    total_files = len(pdf_files)
    matched_count = 0   # 记录匹配到的发票数量
    for index, invoice_number in enumerate(invoice_numbers):
        if len(invoice_number) == 8 and invoice_number.isdigit():
            matched = False
            for pdf_file in pdf_files:
                pdf_path = os.path.join(folder_path, pdf_file)
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        pdf_invoice_number = extract_invoice_number(pdf)
                        if pdf_invoice_number is None:
                            print(f"无法从文件 {pdf_file} 中提取发票号码，这张发票不是汇总单号{summary_number}里的发票")
                            continue
                        if pdf_invoice_number == invoice_number:
                            matched = True
                            matched_count += 1   # 记录匹配数量
                            new_file_name = f"{summary_number}_{invoice_number}.pdf"
                            new_file_path = os.path.join(output_folder, new_file_name)
                            shutil.copy(pdf_path, new_file_path)
                            
                            # 将被复制的文件名添加到source_files_list
                            new_row = pd.DataFrame({'原文件': [pdf_file]})
                            source_files_list = pd.concat([source_files_list, new_row], ignore_index=True)
                                        
                except Exception as e:
                    messagebox.showerror("错误", f"处理文件 {pdf_file} 时发生错误: {e}")
            if not matched:
                messagebox.showwarning("警告", f"发票号码 {invoice_number} 没找到对应的发票文件, 请核实你的发票文件\n\n并且不要使用本次生成的打印文件，准备好完整的发票文件后重新生成\n\n请在全部检查结束后直接退出或重新选择文件夹")

        draw_progress_bar(index + 1, len(invoice_numbers))
        print("\n")        

    if matched_count == 0:
       print(f"汇总单 {summary_number} 没有找到任何发票文件，请检查！")
    elif matched_count < len(invoice_numbers):
       print(f"汇总单 {summary_number} 中有部分发票号码未找到发票文件，汇总单共有 {len(invoice_numbers)} 个发票号码，仅找到 {matched_count} 个发票文件。")
         


    merged_pdf_path = merge_pdfs(summary_number, output_folder)
    append_blank_page_if_needed(merged_pdf_path, output_folder, summary_number)
    final_pdf_path = adjust_pages_to_a4(merged_pdf_path, output_folder, summary_number)
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
    blank_page_path = os.path.join(output_folder, f"blank_page_{summary_number}.pdf")
    c = canvas.Canvas(blank_page_path, pagesize=(610,394))
    c.showPage()
    c.save()
    return blank_page_path

def append_blank_page_if_needed(merged_pdf_path, output_folder, summary_number):
    with open(merged_pdf_path, "rb") as fin:
        reader = PdfReader(fin)
        num_pages = len(reader.pages)
        if num_pages == 1:
           return
        if num_pages % 2 != 0:
            blank_page_path = create_blank_page(output_folder, summary_number)

            writer = PdfWriter()
            with open(merged_pdf_path, "rb") as fin:
                for page in PdfReader(fin).pages:
                    writer.add_page(page)
            with open(blank_page_path, "rb") as fin:
                writer.add_page(PdfReader(fin).pages[0])

            with open(merged_pdf_path, "wb") as fout:
                writer.write(fout)
            
def create_pdf_with_headerfooter(text, width, height, font_size=8, margin=28.346645669):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))

    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    
    pdfmetrics.registerFont(TTFont('ChineseFont', system_font_path))

    can.setFont("ChineseFont", font_size)

    if header_or_footer == 1:
        can.drawString(margin, height-margin, text)
    else:
        can.drawString(margin, margin, text)

    can.save()
    packet.seek(0)
    return PdfReader(packet) 

def adjust_pages_to_a4(merged_pdf_path, output_folder, summary_number):
    writer = PdfWriter()
    width, height = A4
    print("    ")
    
    if Binding_Position == 1:
       top_margin = 70.86614173  # 上边距
       bottom_margin = 28.34645669  # 下边距
       left_margin = 7.65354331  # 左边距
       right_margin = 26.36220472  # 右边距

       effective_width = width - left_margin - right_margin
       effective_height = height - top_margin - bottom_margin

       original_pdf = PdfReader(merged_pdf_path)
       original_page1_widths = [page.mediabox.width for page in original_pdf.pages[::2]]
       original_page2_widths = [page.mediabox.width for page in original_pdf.pages[1::2]]
   
       original_page1_heights = [page.mediabox.height for page in original_pdf.pages[::2]]
       original_page2_heights = [page.mediabox.height for page in original_pdf.pages[1::2]]

       with open(merged_pdf_path, "rb") as fin:
           reader = PdfReader(fin)
           if len(reader.pages) == 1:
               return

           for i in range(0, len(reader.pages), 2):
               new_page = PageObject.create_blank_page(width=width, height=height)

               if i < len(reader.pages):
                   page1 = reader.pages[i]
                   page2 = reader.pages[i + 1] if i + 1 < len(reader.pages) else None

                   total_height = original_page1_heights[i // 2] + (original_page2_heights[i // 2] if page2 else 0)
                   scale = min(effective_width / float(original_pdf.pages[i].mediabox.width), effective_height / float(total_height))

                   page1.scale_by(scale)
                   x_offset = effective_width - page1.mediabox.width
                   y_offset = effective_height - page1.mediabox.height
                   new_page.merge_page(page1)
                   new_page.add_transformation((1, 0, 0, 1, x_offset, y_offset))

                   if page2:
                       page2.scale_by(scale)
                       x_offset2 = effective_width - page2.mediabox.width
                       y_offset2 = effective_height - (page1.mediabox.height + page2.mediabox.height)
                       new_page.merge_page(page2)
                       new_page.add_transformation((1, 0, 0, 1, x_offset2, y_offset2))

                   #print(f"处理完成页面 pages {i+1} and {i+2 if page2 else 'N/A'}:")                       

                   if header_or_footer == 0:
                      text = (" ")
                   else:
                      text = (f"本页发票所属通行费电子票据汇总单号：{summary_number}")
                   pdf_reader = create_pdf_with_headerfooter(text, 595, 842, font_size=8, margin=28.346645669)

                   new_page.merge_page(pdf_reader.pages[0])
                   writer.add_page(new_page)
                
    else:       
       top_margin = 42.51968504  # 上边距
       bottom_margin = 28.34645669  # 下边距
       left_margin = 64.34645669  # 左边距
       right_margin = -16.15748031  # 右边距

       effective_width = width - left_margin - right_margin
       effective_height = height - top_margin - bottom_margin

       original_pdf = PdfReader(merged_pdf_path)
       original_page1_widths = [page.mediabox.width for page in original_pdf.pages[::2]]
       original_page2_widths = [page.mediabox.width for page in original_pdf.pages[1::2]]
      
       original_page1_heights = [page.mediabox.height for page in original_pdf.pages[::2]]
       original_page2_heights = [page.mediabox.height for page in original_pdf.pages[1::2]]

       with open(merged_pdf_path, "rb") as fin:
           reader = PdfReader(fin)
           if len(reader.pages) == 1:
              return

           for i in range(0, len(reader.pages), 2):
               new_page = PageObject.create_blank_page(width=width, height=height)

               if i < len(reader.pages):
                   page1 = reader.pages[i]
                   page2 = reader.pages[i + 1] if i + 1 < len(reader.pages) else None

                   total_height = original_page1_heights[i // 2] + (original_page2_heights[i // 2] if page2 else 0)
                   scale = min(effective_width / float(original_pdf.pages[i].mediabox.width), effective_height / float(total_height))

                   page1.scale_by(scale)
                   x_offset =  (effective_width - page2.mediabox.width * scale)
                   y_offset = effective_height - page1.mediabox.height
                   new_page.merge_page(page1)
                   new_page.add_transformation((1, 0, 0, 1, x_offset, y_offset))
                   
                   if page2:
                       page2.scale_by(scale)
                       x_offset2 = (effective_width - page2.mediabox.width * scale)/2
                       y_offset2 = bottom_margin

                       new_page.merge_page(page2)
                       new_page.add_transformation((1, 0, 0, 1, x_offset2, y_offset2))

                   if header_or_footer == 0:
                      text = (" ")
                   else:
                      text = (f"本页发票所属通行费电子票据汇总单号：{summary_number}")

                   new_page.merge_page(pdf_reader.pages[0])
                   writer.add_page(new_page)

    if summary_page_position == 1:
       output_path = os.path.join(output_folder, f"{summary_number}_2_temp4prt.pdf")
    else:
       output_path = os.path.join(output_folder, f"{summary_number}_1_temp4prt.pdf")
       
    with open(output_path, 'wb') as f:
        writer.write(f)
    return output_path

def open_folder(folder_path):
    if not os.path.exists(folder_path):
        return
    folder_path = os.path.normpath(folder_path)
    system = platform.system()
    try:
        if system == 'Windows':
            os.startfile(folder_path)
        elif system == 'Darwin':
            subprocess.run(['open', folder_path], check=True)
        elif system == 'Linux':
            subprocess.run(['xdg-open', folder_path], check=True)
        else:
            print("不支持的操作系统:", system)
    except Exception as e:
        print("无法打开文件夹:", e)

def exit_program(root):

    if 'folder_path' in globals():
        open_folder(folder_path)
    root.quit()

def clean_temp_files():
    global source_files_list

    # 在源文件夹下创建"已处理原文件"文件夹
    processed_folder = os.path.join(folder_path, "已处理原文件")
    # 确保已处理原文件文件夹存在
    if not os.path.exists(processed_folder):
        os.makedirs(processed_folder)
    # 遍历source_files_list中的所有文件
    for index, row in source_files_list.iterrows():
        original_file = row['原文件']
        original_file_path = os.path.join(folder_path, original_file)
        # 检查文件是否存在于源文件夹
        if os.path.exists(original_file_path):
            # 移动文件到已处理原文件文件夹
            shutil.move(original_file_path, processed_folder)
            #print(f"文件 {original_file} 已移动到 {processed_folder}")
        else:
            print(f"文件 {original_file} 不存在，无法移动")

    if os.path.exists(output_folder):
        target_pdf_pattern = os.path.join(output_folder, "通行费发票按汇总单号合并的打印版本_*.pdf")
        matching_files = glob.glob(target_pdf_pattern)
        if matching_files:
            latest_file = max(matching_files, key=os.path.getctime)
            target_pdf_name = os.path.basename(latest_file)
            target_pdf_path_in_upper = os.path.join(folder_path, target_pdf_name)

            if os.path.exists(target_pdf_path_in_upper):
                messagebox.showwarning("警告", f"文件 {target_pdf_name} 已经存在于上一级文件夹中，跳过移动。")
            else:
                shutil.move(latest_file, folder_path)
        #else:
              #tempfolder_abs_path = os.path.abspath(output_folder)
              #send2trash.send2trash(tempfolder_abs_path)

        tempfolder_abs_path = os.path.abspath(output_folder)

        temp4prt_pattern = os.path.join(output_folder, "*temp4prt*")
        temp_merge_pattern = os.path.join(output_folder, "*临时合并*")
        blank_pattern = os.path.join(output_folder, "*blank_page*")
        for pattern in [temp4prt_pattern, temp_merge_pattern, blank_pattern]:
            for file in glob.glob(pattern):
                os.remove(file)

        one_piece_pattern = os.path.join(output_folder, "*票据汇总单_1piece.pdf")
        for file in glob.glob(one_piece_pattern):
            new_filename = os.path.splitext(file)[0].replace("票据汇总单_1piece", "通行费电子票据汇总单") + ".pdf"
            os.rename(file, new_filename)

        summary_pattern = os.path.join(output_folder, "票据汇总单_*.pdf")
        for file in glob.glob(summary_pattern):
            parts = os.path.basename(file).split("_")[1]
            summarynumber = parts.split(".pdf")[0]
            new_filename = f"{summarynumber}_通行费电子票据汇总单.pdf"
            os.rename(file, os.path.join(output_folder, new_filename))
        
        new_folder_name = "通行费电子票据汇总单和发票整理"
        new_folder_path = os.path.join(os.path.dirname(output_folder), new_folder_name)
        
        if os.path.exists(new_folder_path):
            for file in os.listdir(output_folder):
                src_file = os.path.join(output_folder, file)
                dst_file = os.path.join(new_folder_path, file)

        if os.path.exists(new_folder_path):
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
            os.makedirs(new_folder_path)
            for file in os.listdir(output_folder):
                src_file = os.path.join(output_folder, file)
                dst_file = os.path.join(new_folder_path, file)
                shutil.move(src_file, dst_file)
        
        os.rmdir(output_folder)
        
        messagebox.showinfo("清理完成", "过程文件已清理。")
    else:
        messagebox.showwarning("警告", "没有找到临时文件夹。")
        
def update_button_state():
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

    window_width = 300
    window_height = 360
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.resizable(False, False)

    button_width = 18
    button_height = 2
    
    btn_select_folder = tk.Button(root, text="选择文件夹", width=button_width, height=button_height, command=select_folder)
    btn_select_folder.pack(pady=20)

    global btn_complete
    btn_complete = tk.Button(root, text="生成打印文件", width=button_width, height=button_height, command=lambda: merge_all_print_versions(output_folder) if summary_numbers else messagebox.showwarning("警告", "没有汇总单号可供处理"), state=tk.DISABLED)
    btn_complete.pack(pady=20)

    global btn_clean
    btn_clean = tk.Button(root, text="清理过程文件", width=button_width, height=button_height, command=lambda: clean_temp_files() if summary_numbers else messagebox.showwarning("警告", "没有汇总单号可供处理"), state=tk.DISABLED)
    btn_clean.pack(pady=20)

    btn_exit = tk.Button(root, text="退        出", width=button_width, height=button_height, command=lambda: exit_program(root))
    btn_exit.pack(pady=20)

    update_button_state()
    root.mainloop()

if __name__ == "__main__":
    main()
    
#### mix by immt ####
