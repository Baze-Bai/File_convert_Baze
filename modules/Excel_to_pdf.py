import streamlit as st
import pandas as pd
from fpdf import FPDF
import base64
import io
import os
from utils.common import return_to_main
import tempfile
import shutil
import subprocess
from utils.common import cleanup_temp_dirs
import zipfile  # 添加zipfile模块用于创建压缩包

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
    uploaded_files = st.file_uploader("选择一个或多个Excel/CSV文件", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

    def csv_to_excel(csv_file_path, temp_dir):
        """将CSV转换为Excel格式"""
        try:
            # 读取CSV文件
            df = pd.read_csv(csv_file_path)
            # 创建Excel文件路径
            excel_path = os.path.join(temp_dir, os.path.splitext(os.path.basename(csv_file_path))[0] + ".xlsx")
            # 保存为Excel
            df.to_excel(excel_path, index=False)
            return excel_path
        except Exception as e:
            st.error(f"CSV转Excel错误: {str(e)}")
            return None

    def excel_to_pdf_libreoffice(excel_file_path, output_dir):
        """使用LibreOffice转换Excel到PDF"""
        try:
            # 使用LibreOffice命令行转换
            cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, excel_file_path]
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            if process.returncode != 0:
                st.error(f"转换错误: {process.stderr.decode()}")
                return None
                
            # 获取输出文件名
            output_filename = os.path.splitext(os.path.basename(excel_file_path))[0] + ".pdf"
            pdf_path = os.path.join(output_dir, output_filename)
            
            return pdf_path
            
        except Exception as e:
            st.error(f"转换错误: {str(e)}")
            return None

    # 当用户上传文件时
    if uploaded_files:
        # 显示文件预览信息
        for uploaded_file in uploaded_files:
            st.write(f"已上传: {uploaded_file.name} ({uploaded_file.size/1024:.2f} KB)")
        
        # 转换按钮
        if st.button("批量转换为PDF"):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)
            
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
                
                # 处理多个文件的情况
                if len(uploaded_files) > 1:
                    # 创建一个临时目录来存储所有PDF文件
                    pdf_files = []
                    success_count = 0
                    
                    for uploaded_file in uploaded_files:
                        # 创建临时文件保存上传的Excel或CSV
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        
                        # 如果是CSV文件，先转换为Excel
                        if file_path.endswith('.csv'):
                            excel_path = csv_to_excel(file_path, temp_dir)
                            if not excel_path:
                                st.error(f"CSV转Excel失败: {uploaded_file.name}")
                                continue
                        else:
                            excel_path = file_path
                        
                        # 转换为PDF
                        pdf_path = excel_to_pdf_libreoffice(excel_path, output_dir)
                        
                        if pdf_path and os.path.exists(pdf_path):
                            success_count += 1
                            pdf_files.append(pdf_path)
                    
                    if pdf_files:
                        # 创建ZIP文件
                        zip_path = os.path.join(temp_dir, "pdf_files.zip")
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for pdf_file in pdf_files:
                                # 只添加文件名，而不是完整路径
                                zipf.write(pdf_file, os.path.basename(pdf_file))
                        
                        # 读取ZIP文件内容
                        with open(zip_path, "rb") as f:
                            zip_data = f.read()
                        
                        # 创建ZIP文件下载链接
                        b64_zip = base64.b64encode(zip_data).decode()
                        zip_filename = "converted_pdfs.zip"
                        href = f'<a href="data:application/zip;base64,{b64_zip}" download="{zip_filename}" class="download-button">下载所有PDF文件（压缩包）</a>'
                        
                        st.success(f"成功转换 {success_count} 个文件！点击下方链接下载ZIP压缩包。")
                        
                        # 在下载容器中添加下载链接
                        with download_container:
                            st.markdown(href, unsafe_allow_html=True)
                    else:
                        st.error("转换失败，请检查文件格式")
                else:
                    # 单个文件的处理（保持原有逻辑）
                    file_path = os.path.join(temp_dir, uploaded_files[0].name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_files[0].getbuffer())
                    
                    # 如果是CSV文件，先转换为Excel
                    if file_path.endswith('.csv'):
                        excel_path = csv_to_excel(file_path, temp_dir)
                        if not excel_path:
                            st.error("CSV转Excel失败")
                            return
                    else:
                        excel_path = file_path
                    
                    # 转换为PDF
                    pdf_path = excel_to_pdf_libreoffice(excel_path, output_dir)
                    
                    if pdf_path and os.path.exists(pdf_path):
                        # 读取PDF文件内容
                        with open(pdf_path, "rb") as f:
                            pdf_data = f.read()
                        
                        # 创建下载链接
                        b64_pdf = base64.b64encode(pdf_data).decode()
                        pdf_filename = f"{os.path.splitext(uploaded_files[0].name)[0]}.pdf"
                        href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_filename}" class="download-button">下载 {pdf_filename}</a>'
                        
                        st.success(f"文件转换成功！点击下方链接下载 {pdf_filename}。")
                        
                        # 在下载容器中添加下载链接
                        with download_container:
                            st.markdown(href, unsafe_allow_html=True)
                    else:
                        st.error("转换失败，请检查文件格式")

    # 添加页脚
    st.markdown("---")
    st.markdown("""
    ### 使用说明
    1. 此工具使用LibreOffice进行转换，支持XLS、XLSX和CSV格式
    2. 转换后的PDF将保留原始Excel中的表格和数据格式
    """)
