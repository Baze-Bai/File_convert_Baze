import streamlit as st
import PyPDF2
import pandas as pd
import os
from io import BytesIO
from utils.common import return_to_main, cleanup_temp_dirs
import base64
import tempfile
import shutil
import io
# 添加OCR相关库
import pytesseract
from pdf2image import convert_from_path, convert_from_bytes
from PIL import Image
import cv2
import numpy as np

def extract_pdf_text(pdf_file):
    # 保存上传的文件到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_path = temp_pdf.name
    
    try:
        # 将PDF转换为图像
        images = convert_from_bytes(open(temp_path, 'rb').read())
        
        all_tables = []
        
        for i, image in enumerate(images):
            # 将PIL图像转换为OpenCV格式
            opencv_image = np.array(image)
            opencv_image = opencv_image[:, :, ::-1].copy()  # RGB to BGR转换
            
            # 转灰度图
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # 二值化处理
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
            
            # 检测水平线和垂直线
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
            
            horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
            vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=1)
            
            # 合并水平线和垂直线
            table_borders = cv2.addWeighted(horizontal_lines, 0.5, vertical_lines, 0.5, 0)
            
            # 找到表格轮廓
            contours, _ = cv2.findContours(table_borders, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 如果找到表格
            if contours:
                # 根据轮廓面积排序，找出最大的轮廓（假设是表格）
                contours = sorted(contours, key=cv2.contourArea, reverse=True)
                
                for contour in contours[:3]:  # 处理最大的3个轮廓
                    # 获取轮廓的坐标
                    x, y, w, h = cv2.boundingRect(contour)
                    
                    # 从原图中提取表格区域
                    table_roi = gray[y:y+h, x:x+w]
                    
                    # 使用pytesseract进行OCR识别
                    custom_config = r'--oem 3 --psm 6'
                    table_text = pytesseract.image_to_string(table_roi, lang='chi_sim+eng', config=custom_config)
                    
                    # 改进：更智能地处理表格数据
                    rows = [row.strip() for row in table_text.split('\n') if row.strip()]
                    
                    # 改进：通过分析所有行来决定最佳的列分隔策略
                    data = []
                    max_cols = 0
                    
                    # 第一步：确定可能的列数量和分隔符
                    for row in rows:
                        # 尝试不同的分隔符和方法
                        spaces_split = row.split()
                        tabs_split = row.split('\t')
                        
                        # 选择分割出的列最多的方法
                        best_split = spaces_split if len(spaces_split) > len(tabs_split) else tabs_split
                        
                        if len(best_split) > max_cols:
                            max_cols = len(best_split)
                    
                    # 第二步：使用确定的列数量重新处理每一行
                    if max_cols > 0:
                        for row in rows:
                            row_data = []
                            
                            # 先尝试按空格分割
                            spaces_split = row.split()
                            
                            # 如果分割后的列数与预期差距太大，尝试其他方法
                            if len(spaces_split) >= max_cols * 0.7:  # 允许一些容错
                                row_data = spaces_split
                            else:
                                # 尝试按制表符分割
                                tabs_split = row.split('\t')
                                if len(tabs_split) >= max_cols * 0.7:
                                    row_data = tabs_split
                                else:
                                    # 如果还不行，使用启发式方法
                                    # 将连续的数字和文字分组
                                    parts = []
                                    current_part = ""
                                    
                                    for char in row:
                                        if char.strip():  # 忽略空白字符
                                            current_part += char
                                        elif current_part:  # 如果有空白并且current_part不为空
                                            parts.append(current_part)
                                            current_part = ""
                                    
                                    if current_part:  # 添加最后一部分
                                        parts.append(current_part)
                                    
                                    row_data = parts
                            
                            # 填充不足的列
                            while len(row_data) < max_cols:
                                row_data.append("")
                                
                            # 截断多余的列
                            row_data = row_data[:max_cols]
                            
                            data.append(row_data)
                    
                    # 如果成功解析了数据
                    if data and len(data) > 0:
                        # 创建DataFrame
                        if len(data) > 1:
                            # 假设第一行是表头
                            df = pd.DataFrame(data[1:], columns=data[0])
                        else:
                            df = pd.DataFrame([data[0]])
                        
                        all_tables.append(df)
            
            # 如果没有成功通过表格轮廓提取，尝试直接从整个页面OCR提取表格
            if not all_tables:
                # 使用pytesseract的OSD功能检测文本方向
                osd_config = r'--oem 3 --psm 0'
                try:
                    osd = pytesseract.image_to_osd(image, config=osd_config)
                    # 根据OSD结果选择合适的PSM模式
                    psm_mode = 6  # 默认：假设是单个文本块
                    if "Script: 1" in osd:  # 拉丁文
                        psm_mode = 6
                    else:  # 可能是中文或其他文字
                        psm_mode = 3  # 尝试自动页面分割
                except:
                    psm_mode = 3  # 出错时默认使用自动分割
                
                custom_config = f'--oem 3 --psm {psm_mode}'
                page_text = pytesseract.image_to_string(image, lang='chi_sim+eng', config=custom_config)
                
                # 分析提取出的文本
                rows = [row.strip() for row in page_text.split('\n') if row.strip()]
                
                # 用更智能的方法处理表格数据
                # 计算每行的平均字符数和单词数
                avg_chars = sum(len(row) for row in rows) / len(rows) if rows else 0
                avg_words = sum(len(row.split()) for row in rows) / len(rows) if rows else 0
                
                # 策略：根据文本特征判断如何分割
                if avg_words > 5:  # 可能是有多列的表格
                    # 尝试找出表头行
                    header_row_index = -1
                    max_score = 0
                    
                    for i, row in enumerate(rows):
                        # 计算这行成为表头的可能性
                        words = row.split()
                        score = len(words)  # 列数越多越可能是表头
                        
                        # 表头通常不会太长也不会太短
                        if 3 <= len(words) <= 15:
                            score += 3
                        
                        # 表头通常不含长数字
                        if not any(len(w) > 3 and w.isdigit() for w in words):
                            score += 2
                        
                        if score > max_score:
                            max_score = score
                            header_row_index = i
                    
                    # 分析所有行找出最可能的列数
                    col_counts = {}
                    for row in rows:
                        col_count = len(row.split())
                        col_counts[col_count] = col_counts.get(col_count, 0) + 1
                    
                    # 找出最常见的列数
                    most_common_cols = 0
                    max_count = 0
                    for col_count, count in col_counts.items():
                        if count > max_count:
                            max_count = count
                            most_common_cols = col_count
                    
                    # 构建数据
                    data = []
                    for i, row in enumerate(rows):
                        cols = row.split()
                        # 确保列数统一
                        while len(cols) < most_common_cols:
                            cols.append("")
                        data.append(cols[:most_common_cols])
                    
                    # 创建DataFrame
                    if header_row_index >= 0 and header_row_index < len(data):
                        # 如果找到了表头
                        header = data[header_row_index]
                        data_rows = [row for i, row in enumerate(data) if i != header_row_index]
                        df = pd.DataFrame(data_rows, columns=header)
                    else:
                        # 如果没找到表头
                        df = pd.DataFrame(data)
                    
                    all_tables.append(df)
                else:
                    # 可能是单列文本或非表格内容
                    # 尝试寻找数据行的特征（例如人名/编号+数字的模式）
                    data_rows = []
                    for row in rows:
                        parts = row.split()
                        if len(parts) >= 2 and any(p.isdigit() for p in parts):
                            data_rows.append(parts)
                    
                    if data_rows:
                        # 找出最大的列数
                        max_cols = max(len(row) for row in data_rows)
                        # 统一列数
                        for row in data_rows:
                            while len(row) < max_cols:
                                row.append("")
                        
                        df = pd.DataFrame(data_rows)
                        all_tables.append(df)
                    else:
                        # 如果没找到明显的数据行，就把整个文本作为一个表格
                        df = pd.DataFrame({"内容": rows})
                        all_tables.append(df)
        
        # 返回提取的表格
        if all_tables:
            return all_tables
        else:
            # 回退到文本提取
            pdf_file.seek(0)
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()
            
            # 尝试将文本分割成表格结构
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            if lines:
                # 找出可能的列数
                words_per_line = [len(line.split()) for line in lines]
                most_common = max(set(words_per_line), key=words_per_line.count)
                
                # 构建数据
                data = []
                for line in lines:
                    parts = line.split()
                    if len(parts) > 0:
                        # 调整为统一列数
                        while len(parts) < most_common:
                            parts.append("")
                        data.append(parts[:most_common])
                
                if data:
                    # 假设第一行是表头
                    if len(data) > 1:
                        df = pd.DataFrame(data[1:], columns=data[0])
                    else:
                        df = pd.DataFrame([data[0]])
                    return [df]
            
            # 如果无法构建表格，返回原始文本
            return pd.DataFrame({"内容": [text]})
    except Exception as e:
        st.warning(f"处理PDF时出错: {str(e)}")
        # 回退到文本提取
        pdf_file.seek(0)
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return pd.DataFrame({"内容": [text]})
    finally:
        # 清理临时文件
        if os.path.exists(temp_path):
            os.unlink(temp_path)

def pdf_to_excel():
    return_to_main()
    st.title("批量PDF转Excel工具")
    st.write("请上传一个或多个PDF文件")

    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
    

    # 文件上传
    uploaded_files = st.file_uploader("选择PDF文件", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        if st.button("开始转换"):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            # 创建一个字典来存储所有PDF的内容
            all_pdf_contents = []
            
            # 处理每个上传的文件
            for uploaded_file in uploaded_files:
                file_name = uploaded_file.name
                try:
                    # 提取PDF表格
                    result = extract_pdf_text(uploaded_file)
                    
                    # 如果结果是DataFrame列表（表格提取成功）
                    if isinstance(result, list):
                        # 为每个表创建单独的工作表
                        with pd.ExcelWriter(os.path.join(temp_dir, file_name.replace('.pdf', '.xlsx'))) as writer:
                            for i, table_df in enumerate(result):
                                table_df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
                        
                        # 读取生成的Excel文件
                        with open(os.path.join(temp_dir, file_name.replace('.pdf', '.xlsx')), 'rb') as f:
                            excel_data = f.read()
                            
                        # 添加到列表
                        all_pdf_contents.append({
                            "文件名": file_name,
                            "Excel数据": excel_data
                        })
                    else:
                        # 单个DataFrame情况
                        all_pdf_contents.append({
                            "文件名": file_name,
                            "内容": result
                        })
                    
                except Exception as e:
                    st.error(f"处理文件 {file_name} 时出错: {str(e)}")

            if all_pdf_contents:
                # 创建一个合并的Excel文件
                output_excel_path = os.path.join(temp_dir, "合并结果.xlsx")
                with pd.ExcelWriter(output_excel_path) as writer:
                    for i, pdf_content in enumerate(all_pdf_contents):
                        if "Excel数据" in pdf_content:
                            # 如果有Excel数据，读取并写入
                            temp_excel = os.path.join(temp_dir, f"temp_{i}.xlsx")
                            with open(temp_excel, 'wb') as f:
                                f.write(pdf_content["Excel数据"])
                            
                            # 读取所有工作表并写入主Excel
                            excel_file = pd.ExcelFile(temp_excel)
                            for sheet_name in excel_file.sheet_names:
                                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                                new_sheet_name = f"{pdf_content['文件名']}_{sheet_name}"
                                df.to_excel(writer, sheet_name=new_sheet_name[:31], index=False)  # Excel限制工作表名31字符
                        else:
                            # 如果是文本内容
                            if isinstance(pdf_content["内容"], pd.DataFrame):
                                pdf_content["内容"].to_excel(writer, sheet_name=pdf_content["文件名"][:31], index=False)
                            else:
                                pd.DataFrame({"内容": [pdf_content["内容"]]}).to_excel(
                                    writer, sheet_name=pdf_content["文件名"][:31], index=False
                                )
                
                # 读取合并后的Excel
                with open(output_excel_path, 'rb') as f:
                    excel_buffer = BytesIO(f.read())
                
                # 添加CSS样式
                st.markdown("""
                <style>
                .download-button {
                    display: inline-block;
                    padding: 8px 16px;
                    background-color: #4CAF50;
                    color: white !important;
                    text-align: center;
                    text-decoration: none;
                    font-size: 16px;
                    margin: 10px 5px;
                    border-radius: 4px;
                    cursor: pointer;
                    transition: background-color 0.3s;
                }
                .download-button:hover {
                    background-color: #45a049;
                }
                </style>
                """, unsafe_allow_html=True)
                
                # 准备文件名
                if len(uploaded_files) == 1:
                    # 如果只有一个文件，使用原始文件名（去掉.pdf后缀，添加.xlsx）
                    file_name = os.path.splitext(uploaded_files[0].name)[0] + ".xlsx"
                else:
                    # 多个文件时使用默认名称
                    file_name = "PDF内容汇总.xlsx"
                
                # 创建base64编码的Excel下载链接
                b64_excel = base64.b64encode(excel_buffer.getvalue()).decode()
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                href = f'<a href="data:{mime_type};base64,{b64_excel}" download="{file_name}" class="download-button">下载Excel文件</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("转换完成！")
