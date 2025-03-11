import streamlit as st
import os
import tempfile
import zipfile
import pdf2image
from PIL import Image
import io
import fitz  # PyMuPDF
from utils.common import return_to_main
import base64
import shutil
from utils.common import cleanup_temp_dirs

def pdf_to_image():
    # 设置页面标题
    return_to_main()
    st.title("PDF转图片转换器")
    st.write("上传PDF文件，将每一页转换为高清晰度图片，并打包下载")

    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
    


    # 上传PDF文件
    uploaded_file = st.file_uploader("选择PDF文件", type="pdf")

    # 图片格式选择
    image_format = st.selectbox(
        "选择输出图片格式",
        options=["PNG", "JPEG", "TIFF", "BMP","JPG"],
        index=0
    )

    # 设置DPI（分辨率）
    dpi = st.slider("选择图片分辨率(DPI)", min_value=100, max_value=1000, value=300, step=50)

    # 转换按钮
    if uploaded_file is not None and st.button("开始转换"):
        # 清理之前的临时目录
        cleanup_temp_dirs()
        
        # 创建新的临时目录
        temp_dir = tempfile.mkdtemp()
        st.session_state.temp_dirs.append(temp_dir)
        
        with st.spinner("正在转换中，请稍候..."):
            # 创建临时文件保存上传的PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                pdf_path = tmp_file.name
            
            try:
                # 创建进度条和状态文本
                pdf_loading_bar = st.progress(0)
                pdf_status_text = st.empty()
                pdf_status_text.text("正在加载PDF文件并准备转换...")
                
                # 打开PDF文件获取页数
                pdf_document = fitz.open(pdf_path)
                total_pages = len(pdf_document)
                pdf_document.close()
                
                # 逐页转换PDF为图片
                images = []
                for i in range(total_pages):
                    # 更新进度条
                    progress = (i + 1) / total_pages
                    pdf_loading_bar.progress(progress)
                    pdf_status_text.text(f"正在转换第 {i+1}/{total_pages} 页...")
                    
                    # 只转换当前页
                    page_images = pdf2image.convert_from_path(
                        pdf_path,
                        dpi=dpi,
                        fmt=image_format.lower(),
                        first_page=i+1,
                        last_page=i+1
                    )
                    
                    if page_images:
                        images.append(page_images[0])
                
                # 清除PDF加载状态
                pdf_status_text.empty()
                
                # 创建临时目录存储图片
                tmp_dir = os.path.join(temp_dir, "images")
                os.makedirs(tmp_dir, exist_ok=True)
                
                # 如果只有一页，直接提供单个图片下载
                if total_pages == 1:
                    # 创建内存中的图片文件
                    img_buffer = io.BytesIO()
                    images[0].save(img_buffer, format=image_format)
                    img_buffer.seek(0)
                    
                    st.success("转换完成！PDF只有一页，将直接下载图片文件")
                    
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
                    
                    # 创建base64编码的图片下载链接
                    b64_img = base64.b64encode(img_buffer.getvalue()).decode()
                    file_ext = image_format.lower()
                    file_name = f"{uploaded_file.name.split('.')[0]}.{file_ext}"
                    mime_type = f"image/{file_ext}"
                    href = f'<a href="data:{mime_type};base64,{b64_img}" download="{file_name}" class="download-button">下载图片文件</a>'
                    st.markdown(href, unsafe_allow_html=True)
                    
                    # 显示预览
                    st.subheader("图片预览")
                    st.image(images[0], caption="第1页", use_container_width=True)
                else:
                    # 处理多页PDF
                    with st.spinner("正在创建压缩包..."):
                        # 创建一个内存中的ZIP文件
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            # 创建进度条
                            saving_bar = st.progress(0)
                            status_text = st.empty()
                            
                            # 保存每一页为图片并添加到ZIP文件
                            total_images = len(images)
                            for i, image in enumerate(images):
                                # 更新进度条和状态文本
                                progress = (i + 1) / total_images
                                saving_bar.progress(progress)
                                status_text.text(f"正在保存第 {i+1}/{total_images} 页...")
                                
                                img_filename = f"page_{i+1}.{image_format.lower()}"
                                img_path = os.path.join(tmp_dir, img_filename)
                                image.save(img_path, format=image_format)
                                zip_file.write(img_path, arcname=img_filename)
                        
                        # 完成后清除状态文本
                        status_text.empty()
                        
                        # 设置ZIP文件供下载
                        zip_buffer.seek(0)
                        st.success(f"转换完成！共转换 {len(images)} 页")
                        
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
                        
                        # 创建base64编码的ZIP下载链接
                        b64_zip = base64.b64encode(zip_buffer.getvalue()).decode()
                        zip_name = f"{uploaded_file.name.split('.')[0]}_images.zip"
                        href = f'<a href="data:application/zip;base64,{b64_zip}" download="{zip_name}" class="download-button">下载图片压缩包</a>'
                        st.markdown(href, unsafe_allow_html=True)
                        
                        # 显示预览
                        st.subheader("图片预览")
                        cols = st.columns(min(3, len(images)))
                        for i, (col, image) in enumerate(zip(cols, images[:3])):
                            with col:
                                st.image(image, caption=f"第 {i+1} 页", use_container_width=True)
                        
                        if len(images) > 3:
                            st.info(f"仅显示前3页预览，压缩包中包含全部 {len(images)} 页图片")
            
            except Exception as e:
                st.error(f"转换过程中出错: {str(e)}")
            
            finally:
                # 删除临时PDF文件
                os.unlink(pdf_path)
