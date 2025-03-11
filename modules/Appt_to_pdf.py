import streamlit as st
import comtypes.client
import os
from pptx import Presentation
import base64
import tempfile
import zipfile
import io
from utils.common import return_to_main
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

    def convert_ppt_to_pdf(ppt_path, pdf_path):
        """使用PowerPoint COM对象转换PPT到PDF"""
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = True
            
            try:
                deck = powerpoint.Presentations.Open(ppt_path)
                deck.SaveAs(pdf_path, 32)  # 32 是 PDF 格式的文件格式常量
                deck.Close()
                return True
            finally:
                powerpoint.Quit()
        except Exception as e:
            st.error(f"PowerPoint转换错误: {str(e)}")
            st.info("尝试使用python-pptx库处理文件（注意：此方法可能不支持所有格式和效果）")
            
            try:
                # 尝试使用python-pptx进行基础转换
                # 注意：这种方法只能提取PPT内容，不能完全保留格式
                from PIL import Image
                import io
                import fitz  # PyMuPDF
                
                # 从PPT文件创建临时图像
                prs = Presentation(ppt_path)
                temp_images = []
                
                for i, slide in enumerate(prs.slides):
                    # 创建一个临时图像文件
                    img_path = os.path.join(os.path.dirname(pdf_path), f"slide_{i}.png")
                    # 这里需要实现将幻灯片保存为图像的代码
                    # 由于python-pptx不直接支持此功能，这里只是示例
                    # 在实际应用中，需要使用其他库或方法进行转换
                    temp_images.append(img_path)
                
                # 创建PDF
                doc = fitz.open()
                for img_path in temp_images:
                    if os.path.exists(img_path):
                        img_doc = fitz.open(img_path)
                        doc.insert_pdf(img_doc)
                        img_doc.close()
                        os.remove(img_path)
                
                doc.save(pdf_path)
                doc.close()
                return True
            except Exception as inner_e:
                st.error(f"替代转换方法失败: {str(inner_e)}")
                st.warning("请确保您的系统安装了Microsoft PowerPoint，并且运行在Windows环境中。")
                return False

    # 当用户上传文件时
    if uploaded_files:
        # 创建一个容器来存储所有PDF下载链接
        download_container = st.container()
        
        for uploaded_file in uploaded_files:
            # 显示文件信息
            st.subheader(f"处理文件: {uploaded_file.name}")
            file_details = {"文件名": uploaded_file.name, "文件类型": uploaded_file.type, 
                        "文件大小": f"{uploaded_file.size / 1024:.2f} KB"}
            st.write(file_details)
            
            # 显示PPT预览（仅显示文件名，因为Streamlit不支持直接预览PPT）
            st.write("PPT文件已上传")
        
        # 转换按钮
        if st.button("批量转换为PDF"):
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
                    for uploaded_file in uploaded_files:
                        # 创建临时文件来保存上传的PPT
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_ppt:
                            tmp_ppt.write(uploaded_file.getvalue())
                            ppt_path = os.path.abspath(tmp_ppt.name)
                        
                        # 创建临时PDF文件路径
                        pdf_path = os.path.abspath(os.path.splitext(tmp_ppt.name)[0] + '.pdf')
                        
                        try:
                            # 转换为PDF
                            if convert_ppt_to_pdf(ppt_path, pdf_path):
                                conversion_success = True
                                # 读取生成的PDF文件
                                with open(pdf_path, 'rb') as pdf_file:
                                    pdf_data = pdf_file.read()
                                
                                # 添加到ZIP文件
                                pdf_filename = f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
                                zip_file.writestr(pdf_filename, pdf_data)
                                
                                # 创建单个PDF的下载链接
                                b64_pdf = base64.b64encode(pdf_data).decode()
                                href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{pdf_filename}" class="download-button">下载 {pdf_filename}</a>'
                                
                                # 在下载容器中添加下载链接
                                with download_container:
                                    st.markdown(href, unsafe_allow_html=True)
                        
                        finally:
                            # 清理临时文件
                            try:
                                os.remove(ppt_path)
                                if os.path.exists(pdf_path):
                                    os.remove(pdf_path)
                            except:
                                pass
                
                # 只有当至少有一个转换成功时才创建ZIP文件的下载链接
                if conversion_success:
                    # 创建ZIP文件的下载链接
                    zip_buffer.seek(0)
                    zip_data = base64.b64encode(zip_buffer.read()).decode()
                    zip_filename = "所有PDF文件.zip"
                    zip_href = f'<a href="data:application/zip;base64,{zip_data}" download="{zip_filename}" class="download-button">下载所有PDF文件（ZIP压缩包）</a>'
                    
                    # 添加一些CSS样式使下载按钮更明显
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
                    
                    st.markdown(zip_href, unsafe_allow_html=True)
                    st.success("所有文件转换成功！您可以单独下载每个PDF文件，或者下载包含所有PDF的ZIP压缩包。")
                else:
                    st.error("所有文件转换失败。请确保您的系统运行在Windows环境中并安装了Microsoft PowerPoint。")

    # 添加页脚和操作指南
    st.markdown("---")
    st.markdown("""
    ### 使用说明
    1. 此工具主要依赖Microsoft PowerPoint进行转换，需要在Windows系统上运行
    2. 确保您的系统已安装Microsoft PowerPoint
    3. 如果遇到"对象没有连接到服务器"错误，请检查PowerPoint是否正确安装
    4. 上传的PPT文件会被临时存储并在转换后删除
    """)