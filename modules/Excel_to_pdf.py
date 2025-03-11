import streamlit as st
import pandas as pd
from fpdf import FPDF
import base64
import io
import os
from utils.common import return_to_main
import tempfile
import shutil
from utils.common import cleanup_temp_dirs

def excel_to_pdf():
    # 设置页面标题
    return_to_main()
    # 页面标题
    st.title("批量Excel转PDF转换工具") 
    st.markdown("上传Excel文件并将其转换为PDF格式")
    
    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
    


    # 文件上传功能
    uploaded_files = st.file_uploader("选择一个或多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)

    def excel_to_pdf(excel_file):
        # 读取Excel文件
        xls = pd.ExcelFile(excel_file)
        
        # 创建PDF对象 - 添加unicode支持
        pdf = FPDF()
        pdf.add_font('DejaVu', '', os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                               'fonts', 'DejaVuSansCondensed.ttf'), uni=True)
        pdf.set_font('DejaVu', '', 14)
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # 遍历Excel中的每个工作表
        for sheet_name in xls.sheet_names:
            # 读取工作表数据
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # 添加新页面
            pdf.add_page()
            
            # 添加工作表名称作为标题
            pdf.cell(0, 10, f"工作表: {sheet_name}", ln=True, align="C")
            pdf.ln(10)
            
            # 设置表格字体
            pdf.set_font('DejaVu', '', 10)
            
            # 获取列名并计算列宽
            columns = df.columns.tolist()
            col_width = pdf.w / len(columns)
            
            # 添加表头
            for col in columns:
                pdf.cell(col_width, 10, str(col), border=1, align="C")
            pdf.ln()
            
            # 设置数据字体
            pdf.set_font('DejaVu', '', 10)
            
            # 添加数据行
            for _, row in df.iterrows():
                for item in row:
                    pdf.cell(col_width, 10, str(item), border=1, align="C")
                pdf.ln()
        
        # 返回PDF内容 - 确保返回bytes类型
        output = pdf.output(dest="S").encode('latin1')
        return output

    # 当用户上传文件时
    if uploaded_files:
        # 显示文件预览信息
        for uploaded_file in uploaded_files:
            # 显示文件信息
            st.subheader(f"处理文件: {uploaded_file.name}")
            file_details = {"文件名": uploaded_file.name, "文件大小": f"{uploaded_file.size / 1024:.2f} KB"}
            st.write(file_details)
            
            # 显示Excel预览
            st.write("Excel预览")
            df = pd.read_excel(uploaded_file)
            st.dataframe(df)
        
        # 转换按钮
        if st.button("批量转换为PDF"):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            with st.spinner("正在转换所有文件..."):
                # 添加下载按钮的CSS样式
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
                
                # 创建一个容器来在页面底部存储所有PDF下载链接
                st.markdown("---")
                download_container = st.container()
                
                # 检查是否只有一个文件
                if len(uploaded_files) == 1:
                    # 重置文件指针
                    uploaded_files[0].seek(0)
                    
                    # 转换为PDF
                    pdf_data = excel_to_pdf(uploaded_files[0])
                    
                    # 创建下载链接
                    b64_pdf = base64.b64encode(pdf_data).decode()
                    pdf_filename = f"{os.path.splitext(uploaded_files[0].name)[0]}.pdf"
                    href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_filename}" class="download-button">下载 {pdf_filename}</a>'
                    
                    st.success(f"文件转换成功！点击上方链接下载 {pdf_filename}。")
                    
                    # 在下载容器中添加下载链接
                    with download_container:
                        st.markdown(href, unsafe_allow_html=True)
                else:
                    # 处理多个文件的情况
                    all_hrefs = []
                    for uploaded_file in uploaded_files:
                        # 重置文件指针
                        uploaded_file.seek(0)
                        
                        # 转换为PDF
                        pdf_data = excel_to_pdf(uploaded_file)
                        
                        # 创建下载链接
                        b64_pdf = base64.b64encode(pdf_data).decode()
                        pdf_filename = f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
                        href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_filename}" class="download-button">下载 {pdf_filename}</a>'
                        all_hrefs.append(href)
                    
                    st.success("所有文件转换成功！点击上方链接下载PDF文件。")
                    
                    # 在下载容器中添加所有下载链接
                    with download_container:
                        for href in all_hrefs:
                            st.markdown(href, unsafe_allow_html=True)

    # 添加页脚
    st.markdown("---")