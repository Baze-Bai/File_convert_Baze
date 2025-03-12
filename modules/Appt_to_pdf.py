import streamlit as st
import os
import tempfile
import zipfile
import io
import base64
from utils.common import return_to_main
import subprocess
import shutil
from utils.common import cleanup_temp_dirs

def ppt_to_pdf():
    # 设置页面标题
    return_to_main()
    # 页面标题
    st.title("批量PPT转PDF转换工具")
    st.markdown("上传PPT文件并将其转换为PDF格式")

    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
    
    # 文件上传功能
    uploaded_files = st.file_uploader("选择一个或多个PPT文件", type=["pptx", "ppt"], accept_multiple_files=True)

    def convert_ppt_to_pdf(ppt_path, pdf_dir):
        """使用LibreOffice转换PPT到PDF"""
        try:
            # 确保输出目录存在
            os.makedirs(pdf_dir, exist_ok=True)
            
            # 使用LibreOffice命令行转换
            cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', pdf_dir, ppt_path]
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            if process.returncode != 0:
                st.error(f"LibreOffice转换错误: {process.stderr.decode()}")
                return None
            
            # 获取输入文件的基本名称
            base_name = os.path.basename(ppt_path)
            base_name_without_ext = os.path.splitext(base_name)[0]
            
            # LibreOffice生成的PDF文件路径
            generated_pdf_path = os.path.join(pdf_dir, f"{base_name_without_ext}.pdf")
            
            # 验证文件是否存在
            if not os.path.exists(generated_pdf_path):
                st.error(f"PDF文件未生成: {generated_pdf_path}")
                return None
            
            return generated_pdf_path
            
        except Exception as e:
            st.error(f"转换错误: {str(e)}")
            return None

    # 当用户上传文件时
    if uploaded_files:

        st.write("PPT文件已上传")
        
        # 转换按钮
        if st.button("转换为PDF"):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            with st.spinner("正在转换所有文件..."):
                # 创建一个内存中的ZIP文件
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    conversion_success = False
                    pdf_files_data = []  # 存储所有PDF文件信息
                    
                    for uploaded_file in uploaded_files:
                        # 创建临时文件来保存上传的PPT
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_ppt:
                            tmp_ppt.write(uploaded_file.getvalue())
                            ppt_path = os.path.abspath(tmp_ppt.name)
                        
                        # 创建临时PDF输出目录
                        pdf_dir = os.path.join(temp_dir, "output")
                        
                        try:
                            # 转换为PDF，获取实际生成的PDF路径
                            generated_pdf_path = convert_ppt_to_pdf(ppt_path, pdf_dir)
                            
                            if generated_pdf_path and os.path.exists(generated_pdf_path):
                                conversion_success = True
                                
                                # 读取生成的PDF文件
                                with open(generated_pdf_path, 'rb') as pdf_file:
                                    pdf_data = pdf_file.read()
                                
                                # 添加到ZIP文件
                                output_pdf_filename = f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
                                zip_file.writestr(output_pdf_filename, pdf_data)
                                
                                # 保存PDF数据供后续使用
                                pdf_files_data.append({
                                    "filename": output_pdf_filename,
                                    "data": pdf_data
                                })
                        
                        finally:
                            # 清理临时文件
                            try:
                                os.remove(ppt_path)
                            except:
                                pass
                
                # 添加CSS样式使下载按钮更明显
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
                
                # 只有当至少有一个转换成功时才创建下载链接
                if conversion_success:
                    # 单个文件 - 直接提供PDF下载
                    if len(uploaded_files) == 1 and len(pdf_files_data) == 1:
                        pdf_info = pdf_files_data[0]
                        b64_pdf = base64.b64encode(pdf_info["data"]).decode()
                        href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_info["filename"]}" class="download-button">下载 {pdf_info["filename"]}</a>'
                        st.markdown(href, unsafe_allow_html=True)
                        st.success(f"文件转换成功！您可以下载转换后的PDF文件。")
                    else:
                        # 多个文件 - 提供ZIP包下载
                        zip_buffer.seek(0)
                        zip_data = base64.b64encode(zip_buffer.read()).decode()
                        zip_filename = "所有PDF文件.zip"
                        zip_href = f'<a href="data:application/zip;base64,{zip_data}" download="{zip_filename}" class="download-button">下载所有PDF文件（ZIP压缩包）</a>'
                        st.markdown(zip_href, unsafe_allow_html=True)
                        st.success(f"已成功转换 {len(pdf_files_data)} 个文件！您可以下载包含所有PDF的ZIP压缩包。")
                else:
                    st.error("所有文件转换失败。请确保您的系统安装了LibreOffice。")

