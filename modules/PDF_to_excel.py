import streamlit as st
import PyPDF2
import pandas as pd
import os
from io import BytesIO
from utils.common import return_to_main
import base64
import tempfile
import shutil
from utils.common import cleanup_temp_dirs

def extract_pdf_text(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

def pdf_to_excel():
    return_to_main()
    st.title("PDF➡️Excel")
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
                    # 提取PDF文本
                    text_content = extract_pdf_text(uploaded_file)
                    
                    # 将内容添加到列表中
                    all_pdf_contents.append({
                        "文件名": file_name,
                        "内容": text_content
                    })
                    
                except Exception as e:
                    st.error(f"处理文件 {file_name} 时出错: {str(e)}")

            if all_pdf_contents:
                # 创建DataFrame
                df = pd.DataFrame(all_pdf_contents)
                
                # 转换为Excel
                excel_buffer = BytesIO()
                df.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
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
