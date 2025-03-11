import streamlit as st
import os
import zipfile
import tempfile
import time
import base64
from utils.common import return_to_main
import shutil
import subprocess
from utils.common import cleanup_temp_dirs

def word_to_pdf():
    return_to_main()

    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
        
    st.title("Word文档批量转PDF工具")
    st.write("上传Word文档(.docx)，将自动转换为PDF格式")

    uploaded_files = st.file_uploader("选择Word文档", type=["docx"], accept_multiple_files=True)

    if uploaded_files:
        if st.button("开始转换"):
            # 先清理之前的临时目录
            cleanup_temp_dirs()
            
            with st.spinner("正在转换中，请稍候..."):
                # 创建新的临时目录
                temp_dir = tempfile.mkdtemp()
                # 记录这个目录以便后续清理
                st.session_state.temp_dirs.append(temp_dir)
                
                output_dir = os.path.join(temp_dir, "output")
                os.makedirs(output_dir, exist_ok=True)
                
                # 保存上传的文件到临时目录
                input_paths = []
                for uploaded_file in uploaded_files:
                    input_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    input_paths.append(input_path)
                
                # 转换文件
                for input_path in input_paths:
                    try:
                        # 在Linux环境中使用LibreOffice转换
                        output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
                        
                        # 使用LibreOffice命令行转换Word到PDF
                        cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path]
                        process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        
                        if process.returncode != 0:
                            st.error(f"转换失败: {process.stderr.decode()}")
                            continue
                        
                        # 添加自定义CSS来美化下载链接
                        st.markdown("""
                        <style>
                            .download-btn {
                                display: inline-block;
                                background-color: #1E88E5;
                                color: white !important;
                                text-align: center;
                                padding: 12px 20px;
                                border-radius: 8px;
                                text-decoration: none;
                                font-weight: bold;
                                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
                                transition: all 0.3s ease;
                                margin: 10px 0;
                                width: auto;
                                font-size: 16px;
                            }
                            .download-btn:hover {
                                background-color: #1565C0;
                                box-shadow: 0 4px 8px rgba(0,0,0,0.3);
                                transform: translateY(-2px);
                            }
                            .download-icon {
                                margin-right: 8px;
                            }
                            .success-box {
                                background-color: #f0f9f4;
                                border-left: 5px solid #4CAF50;
                                padding: 15px;
                                border-radius: 4px;
                                margin: 20px 0;
                            }
                            .file-info {
                                background-color: #f8f9fa;
                                padding: 10px 15px;
                                border-radius: 5px;
                                margin-bottom: 15px;
                                border: 1px solid #e9ecef;
                            }
                        </style>
                        """, unsafe_allow_html=True)

                        # 根据文件数量决定下载方式
                        if len(input_paths) == 1:
                            # 单个文件提供下载链接
                            pdf_filename = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
                            pdf_path = os.path.join(output_dir, pdf_filename)
                            
                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()
                            
                            b64_pdf = base64.b64encode(pdf_bytes).decode()
                            
                            # 显示成功消息和文件信息
                            st.markdown(
                                f"""
                                <div class="success-box">
                                    <h3>✅ 转换成功！</h3>
                                    <div class="file-info">
                                        <strong>文件名：</strong> {pdf_filename}<br>
                                        <strong>文件大小：</strong> {round(len(pdf_bytes)/1024, 2)} KB
                                    </div>
                                    <a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_filename}" class="download-btn">
                                        <span class="download-icon">📥</span> 下载PDF文件
                                    </a>
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                        else:
                            # 多个文件打包为zip
                            zip_filename = "converted_pdfs.zip"
                            zip_path = os.path.join(temp_dir, zip_filename)
                            
                            with zipfile.ZipFile(zip_path, 'w') as zipf:
                                for pdf_file in os.listdir(output_dir):
                                    if pdf_file.endswith('.pdf'):
                                        zipf.write(
                                            os.path.join(output_dir, pdf_file), 
                                            arcname=pdf_file
                                        )
                            
                            with open(zip_path, "rb") as f:
                                zip_bytes = f.read()
                            
                            b64_zip = base64.b64encode(zip_bytes).decode()
                            
                            # 显示成功消息和文件信息
                            st.markdown(
                                f"""
                                <div class="success-box">
                                    <h3>✅ 成功转换 {len(input_paths)} 个文件！</h3>
                                    <div class="file-info">
                                        <strong>压缩包名称：</strong> {zip_filename}<br>
                                        <strong>文件大小：</strong> {round(len(zip_bytes)/1024, 2)} KB<br>
                                        <strong>包含文件数：</strong> {len(input_paths)} 个PDF
                                    </div>
                                    <a href="data:application/zip;base64,{b64_zip}" download="{zip_filename}" class="download-btn">
                                        <span class="download-icon">📥</span> 下载ZIP压缩包
                                    </a>
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                    except Exception as e:
                        st.error(f"转换 {os.path.basename(input_path)} 时出错: {str(e)}")

# 注册退出处理函数（这个在Streamlit中可能不总是有效，但可以尝试）
def cleanup_all_temp_dirs():
    if hasattr(st.session_state, 'temp_dirs'):
        for dir_path in st.session_state.temp_dirs:
            try:
                if os.path.exists(dir_path):
                    shutil.rmtree(dir_path)
            except:
                pass