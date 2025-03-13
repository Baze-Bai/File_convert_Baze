import streamlit as st
import PyPDF2
import pandas as pd
import os
from io import BytesIO
from utils.common import return_to_main, cleanup_temp_dirs
import base64
import tempfile
import shutil
import tabula
import io

def extract_pdf_text(pdf_file):
    # 保存上传的文件到临时文件，因为tabula需要文件路径
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_path = temp_pdf.name
    
    try:
        # 使用tabula提取表格
        tables = tabula.read_pdf(temp_path, pages='all', multiple_tables=True)
        
        # 如果成功提取到表格
        if tables and len(tables) > 0:
            return tables
        else:
            # 如果tabula无法提取表格，回退到PyPDF2文本提取
            pdf_file.seek(0)  # 重置文件指针
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()
            return pd.DataFrame({"内容": [text]})
    except Exception as e:
        # 发生错误时回退到PyPDF2
        pdf_file.seek(0)  # 重置文件指针
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
