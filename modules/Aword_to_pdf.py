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
        
    st.title("Word➡️PDF")
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
                conversion_results = []
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
                        
                        # 记录成功转换的文件
                        conversion_results.append(True)
                    except Exception as e:
                        st.error(f"转换 {os.path.basename(input_path)} 时出错: {str(e)}")
                        conversion_results.append(False)
                
                # 添加自定义CSS来美化下载链接
                st.markdown("""
                <style>
                    .download-btn {
                        display: inline-block;
                        background: linear-gradient(135deg, #42A5F5 0%, #1976D2 100%);
                        color: white !important;
                        text-align: center;
                        padding: 14px 24px;
                        border-radius: 50px;
                        text-decoration: none;
                        font-weight: bold;
                        box-shadow: 0 4px 15px rgba(25, 118, 210, 0.3);
                        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                        margin: 15px 0;
                        width: auto;
                        font-size: 16px;
                        border: none;
                        position: relative;
                        overflow: hidden;
                        z-index: 1;
                    }
                    .download-btn:before {
                        content: '';
                        position: absolute;
                        top: 0;
                        left: 0;
                        width: 100%;
                        height: 100%;
                        background: linear-gradient(135deg, #1E88E5 0%, #0D47A1 100%);
                        opacity: 0;
                        z-index: -1;
                        transition: opacity 0.4s ease;
                    }
                    .download-btn:hover {
                        transform: translateY(-3px) scale(1.03);
                        box-shadow: 0 7px 20px rgba(25, 118, 210, 0.4);
                    }
                    .download-btn:hover:before {
                        opacity: 1;
                    }
                    .download-btn:active {
                        transform: translateY(1px) scale(0.98);
                        box-shadow: 0 2px 8px rgba(25, 118, 210, 0.4);
                    }
                    .download-icon {
                        margin-right: 10px;
                        font-size: 18px;
                        display: inline-block;
                        transition: transform 0.3s ease;
                    }
                    .download-btn:hover .download-icon {
                        transform: translateY(2px);
                    }
                    .success-box {
                        background-color: #f0f9f4;
                        border-left: 5px solid #4CAF50;
                        padding: 20px;
                        border-radius: 8px;
                        margin: 25px 0;
                        box-shadow: 0 3px 10px rgba(0,0,0,0.05);
                    }
                    .file-info {
                        background-color: #f8f9fa;
                        padding: 15px 20px;
                        border-radius: 8px;
                        margin-bottom: 20px;
                        border: 1px solid #e9ecef;
                        box-shadow: inset 0 1px 3px rgba(0,0,0,0.05);
                    }
                </style>
                """, unsafe_allow_html=True)

                # 计算成功转换的文件数量
                successful_conversions = sum(conversion_results)
                
                # 只在所有转换完成后显示一次下载选项
                if successful_conversions > 0:
                    # 根据文件数量决定下载方式
                    if successful_conversions == 1 and len(input_paths) == 1:
                        # 单个文件提供下载链接
                        pdf_filename = os.path.splitext(os.path.basename(input_paths[0]))[0] + ".pdf"
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
                                <h3>✅ 成功转换 {successful_conversions} 个文件！</h3>
                                <div class="file-info">
                                    <strong>压缩包名称：</strong> {zip_filename}<br>
                                    <strong>文件大小：</strong> {round(len(zip_bytes)/1024, 2)} KB<br>
                                    <strong>包含文件数：</strong> {successful_conversions} 个PDF
                                </div>
                                <a href="data:application/zip;base64,{b64_zip}" download="{zip_filename}" class="download-btn">
                                    <span class="download-icon">📥</span> 下载ZIP压缩包
                                </a>
                            </div>
                            """, 
                            unsafe_allow_html=True
                        )

# 注册退出处理函数（这个在Streamlit中可能不总是有效，但可以尝试）
def cleanup_all_temp_dirs():
    if hasattr(st.session_state, 'temp_dirs'):
        for dir_path in st.session_state.temp_dirs:
            try:
                if os.path.exists(dir_path):
                    shutil.rmtree(dir_path)
            except:
                pass
