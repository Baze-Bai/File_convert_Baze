import streamlit as st
import os
import tempfile
import shutil
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, A5, A3, legal, TABLOID
import base64
from io import BytesIO
import time
from utils.common import return_to_main


# 定义页面大小选项
PAGE_SIZES = {
    "A3": A3,
    "A4": A4,
    "A5": A5,
    "信纸(Letter)": letter,
    "法律文书(Legal)": legal,
    "小报(Tabloid)": TABLOID
}


def convert_multiple_images_to_pdf(image_paths, output_pdf=None, pagesize=A4, image_quality=95, progress_callback=None):
    """
    将多个图片合并为一个PDF文件，每个图片一页
    
    参数:
        image_paths: 图片路径列表
        output_pdf: 输出PDF路径，如果为None则使用临时文件
        pagesize: PDF页面大小
        image_quality: 图像质量(1-100)
        progress_callback: 进度回调函数，用于更新进度条
    返回:
        生成的PDF文件路径
    """
    if output_pdf is None:
        # 创建临时文件用于存储PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        output_pdf = temp_file.name
    
    c = canvas.Canvas(output_pdf, pagesize=pagesize)
    
    # 创建临时目录存储处理后的图像
    temp_dir = tempfile.mkdtemp()
    processed_images = []
    
    try:
        for i, image_path in enumerate(image_paths):
            # 更新进度
            if progress_callback:
                # 从50%到95%的进度范围内更新
                progress = 0.5 + (i / len(image_paths)) * 0.45
                progress_callback(progress, f"处理图片 {i+1}/{len(image_paths)}...")
                
            # 打开图片并获取尺寸
            img = Image.open(image_path)
            
            # 如果有EXIF方向数据，修正图像方向
            try:
                if hasattr(img, '_getexif') and img._getexif() is not None:
                    exif = img._getexif()
                    orientation_key = 274  # EXIF中的Orientation标签
                    if orientation_key in exif:
                        orientation = exif[orientation_key]
                        if orientation == 2:
                            img = img.transpose(Image.FLIP_LEFT_RIGHT)
                        elif orientation == 3:
                            img = img.transpose(Image.ROTATE_180)
                        elif orientation == 4:
                            img = img.transpose(Image.FLIP_TOP_BOTTOM)
                        elif orientation == 5:
                            img = img.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.ROTATE_90)
                        elif orientation == 6:
                            img = img.transpose(Image.ROTATE_270)
                        elif orientation == 7:
                            img = img.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.ROTATE_270)
                        elif orientation == 8:
                            img = img.transpose(Image.ROTATE_90)
            except (AttributeError, KeyError, IndexError):
                # 忽略缺少EXIF数据的情况
                pass
            
            img_width, img_height = img.size
            
            # 计算图片在PDF页面中的位置和大小，保持宽高比
            pdf_width, pdf_height = pagesize
            ratio = min(pdf_width / img_width, pdf_height / img_height)
            new_width = img_width * ratio
            new_height = img_height * ratio
            x_centered = (pdf_width - new_width) / 2
            y_centered = (pdf_height - new_height) / 2
            
            # 保存处理后的图像到临时文件，应用质量设置
            temp_img_path = os.path.join(temp_dir, f"temp_img_{i}.jpg")
            # 转换为RGB模式以确保兼容性
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(temp_img_path, format='JPEG', quality=image_quality)
            processed_images.append(temp_img_path)
            
            # 添加图片到当前页
            c.drawImage(temp_img_path, x_centered, y_centered, width=new_width, height=new_height)
            
            # 如果不是最后一张图片，添加新页
            if i < len(image_paths) - 1:
                c.showPage()
        
        # 最终保存前更新进度到95%
        if progress_callback:
            progress_callback(0.95, "完成PDF生成...")
            
        c.save()
        
        return output_pdf
    
    finally:
        # 清理临时文件
        for img_path in processed_images:
            try:
                if os.path.exists(img_path):
                    os.remove(img_path)
            except:
                pass
        
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except:
            pass


@st.cache_data
def get_pdf_data(pdf_path):
    """缓存PDF数据以避免重复读取"""
    with open(pdf_path, "rb") as f:
        return f.read()

def image_to_pdf():
    # 页面标题
    return_to_main()
    st.title("图片➡️PDF")
    
    # 创建一个变量来跟踪临时目录，以便后续清理
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = None
    
    # 侧边栏设置
    with st.sidebar:
        st.header("PDF设置")
        
        # 页面大小选择
        page_size_option = st.selectbox(
            "选择PDF页面大小:",
            list(PAGE_SIZES.keys()),
            index=1  # 默认选择A4
        )
        pagesize = PAGE_SIZES[page_size_option]
        
        # 图像质量设置
        st.subheader("图像清晰度设置")
        quality = st.slider("图像质量 (1-100):", 1, 100, 95, 
                           help="较高的质量值会产生更大的文件尺寸但图像更清晰")
        
        # 页面方向
        st.subheader("图片旋转")
        global_rotation = st.radio("顺时针旋转:", ["0°", "90°", "180°", "270°"], index=0)
        
        # 将选择转换为角度值
        global_rotation_angle = int(global_rotation.replace("°", ""))
        
        # 简化会话状态设置
        st.session_state.global_rotation_angle = global_rotation_angle
        
        # 自动调整页面方向选项
        auto_adjust_orientation = st.checkbox("根据旋转角度自动调整页面方向", value=True,
                                             help="当旋转90°或270°时自动切换页面的横向/纵向")
        
        # 文件名输入
        st.subheader("输出设置")
        default_filename = "我的文档.pdf"
        output_filename = st.text_input("PDF文件名:", value=default_filename)
        # 确保文件名以.pdf结尾
        if not output_filename.lower().endswith('.pdf'):
            output_filename += '.pdf'
    
    # 主界面
    # 文件上传
    col1, col2 = st.columns([3, 1])
    
    with col1:
        uploaded_files = st.file_uploader(
            "选择图片文件", 
            type=["jpg", "jpeg", "png", "gif", "bmp", "tiff", "webp"], 
            accept_multiple_files=True,
            label_visibility="collapsed"  # 尝试减少视觉元素
        )
    
    # 限制上传文件数量
    max_files = 50
    if uploaded_files and len(uploaded_files) > max_files:
        st.warning(f"您上传了 {len(uploaded_files)} 个文件，但系统限制最多处理 {max_files} 个文件。将只处理前 {max_files} 个文件。")
        uploaded_files = uploaded_files[:max_files]
    
    # 如果有上传的文件，显示预览和处理选项
    if uploaded_files:
        # 显示成功上传消息
        st.success(f"✅ 已成功上传 {len(uploaded_files)} 个文件")
        
        # 图片预览部分
        st.subheader("图片预览")
        
        # 计算在一行中显示的图片数
        num_cols = 3
        
        # 添加分页功能
        total_pages = (len(uploaded_files) + num_cols - 1) // num_cols
        
        # 初始化当前页码（如果不存在）
        if 'current_preview_page' not in st.session_state:
            st.session_state.current_preview_page = 0
        
        # 计算当前页应显示的图片
        start_idx = st.session_state.current_preview_page * num_cols
        end_idx = min(start_idx + num_cols, len(uploaded_files))
        current_page_files = uploaded_files[start_idx:end_idx]
        
        # 显示当前页的图片
        cols = st.columns(len(current_page_files))
        for i, uploaded_file in enumerate(current_page_files):
            with cols[i]:
                try:
                    # 打开图片并应用全局旋转
                    img = Image.open(uploaded_file)
                    
                    # 获取文件大小
                    file_size = len(uploaded_file.getvalue()) / 1024  # 转换为KB
                    
                    # 获取图片尺寸
                    img_width, img_height = img.size
                    
                    if st.session_state.global_rotation_angle > 0:
                        # 顺时针旋转（PIL中是逆时针，所以用负值）
                        img = img.rotate(-st.session_state.global_rotation_angle, expand=True)
                        # 如果旋转了90°或270°，宽高需要交换
                        if st.session_state.global_rotation_angle in [90, 270]:
                            img_width, img_height = img_height, img_width
                    
                    # 显示旋转后的图片
                    st.image(img, use_container_width=True)
                    
                    # 显示文件名、大小和尺寸
                    if file_size >= 1024:
                        # 如果大于1MB，显示为MB
                        file_size_str = f"{file_size/1024:.1f} MB"
                    else:
                        # 否则显示为KB
                        file_size_str = f"{file_size:.1f} KB"
                        
                    st.caption(f"{uploaded_file.name} | {file_size_str} | {img_width}×{img_height}")
                    
                except Exception as e:
                    st.error(f"无法预览图片: {str(e)}")
        
        # 添加分页控制
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            if st.session_state.current_preview_page > 0:
                if st.button("上一页", key="prev_preview"):
                    st.session_state.current_preview_page -= 1
                    st.rerun()
        
        with col2:
            st.write(f"第 {st.session_state.current_preview_page + 1} 页，共 {total_pages} 页")
        
        with col3:
            if st.session_state.current_preview_page < total_pages - 1:
                if st.button("下一页", key="next_preview"):
                    st.session_state.current_preview_page += 1
                    st.rerun()
        
        # 转换按钮 - 现在占据整个宽度
        convert_button = st.button("转换为PDF", use_container_width=True, type="primary")
        
        # 转换处理逻辑
        if convert_button:
            # 清理之前的临时目录（如果存在）
            if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
                try:
                    shutil.rmtree(st.session_state.temp_dir)
                except Exception as e:
                    st.warning(f"清理临时文件时出错: {str(e)}")
            
            # 显示进度条
            progress_container = st.container()
            with progress_container:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text("处理中...")
                
                # 创建临时目录保存上传的图片
                temp_dir = tempfile.mkdtemp()
                st.session_state.temp_dir = temp_dir  # 保存到会话状态以便后续清理
                temp_image_paths = []
                
                try:
                    # 保存上传的图片到临时文件
                    for i, uploaded_file in enumerate(uploaded_files):
                        progress = (i / len(uploaded_files)) * 0.5  # 前半部分进度用于保存图片
                        progress_bar.progress(progress)
                        
                        try:
                            # 打开图片并应用全局旋转
                            img = Image.open(uploaded_file)
                            if st.session_state.global_rotation_angle > 0:
                                img = img.rotate(-st.session_state.global_rotation_angle, expand=True)
                            
                            # 获取文件扩展名（更安全的方式）
                            _, file_ext = os.path.splitext(uploaded_file.name)
                            if not file_ext:
                                file_ext = ".jpg"  # 默认扩展名
                            
                            # 创建临时文件保存处理后的图片
                            temp_image = os.path.join(temp_dir, f"image_{i}{file_ext}")
                            img.save(temp_image)
                            temp_image_paths.append(temp_image)
                            
                            status_text.text(f"图片加载中... ({i+1}/{len(uploaded_files)})")
                        except Exception as e:
                            st.error(f"处理图片 {uploaded_file.name} 时出错: {str(e)}")
                            continue
                    
                    # 根据旋转角度自动调整页面方向
                    adjusted_pagesize = pagesize
                    if auto_adjust_orientation and st.session_state.global_rotation_angle in [90, 270]:
                        # 如果旋转了90°或270°，且当前是纵向，则切换为横向
                        if pagesize[0] < pagesize[1]:  # 如果是纵向
                            adjusted_pagesize = (pagesize[1], pagesize[0])  # 切换为横向
                    
                    # 检查是否有图片可以处理
                    if not temp_image_paths:
                        st.error("没有可处理的图片，请检查上传的文件。")
                        return
                    
                    # 创建一个进度回调函数
                    def update_progress(progress, status_message):
                        progress_bar.progress(progress)
                        status_text.text(status_message)
                    
                    # 调用转换函数，传入进度回调
                    pdf_path = convert_multiple_images_to_pdf(
                        temp_image_paths, 
                        pagesize=adjusted_pagesize, 
                        image_quality=quality,
                        progress_callback=update_progress
                    )
                    status_text.text("转换完成！")
                    
                    progress_bar.progress(1.0)
                    
                    # 提供下载链接，使用用户指定的文件名
                    success_msg = st.success(f"PDF已生成！点击下方按钮下载 \"{output_filename}\"")

                    # 获取PDF数据
                    pdf_data = get_pdf_data(pdf_path)
                    
                    # 创建CSS样式
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
                    
                    # 创建Base64编码的PDF下载链接
                    b64_pdf = base64.b64encode(pdf_data).decode()
                    href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="{output_filename}" class="download-button">下载PDF文件</a>'
                    
                    # 在列布局中显示下载链接
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col1:
                        st.markdown(href, unsafe_allow_html=True)

                    # 显示文件信息
                    file_info = st.info(f"文件设置: {page_size_option} 页面大小 | 质量: {quality}%")

                    # 显示提示信息
                    st.markdown("""
                    **提示**：
                    1. 大多数浏览器会使用默认下载位置
                    2. 如需选择其他位置保存文件，请在下载时使用浏览器的"另存为"选项（通常可以右键点击下载按钮然后选择"链接另存为..."）
                    """)

                except Exception as e:
                    st.error(f"生成PDF时发生错误: {str(e)}")
                
                finally:
                    # 在会话结束时清理临时文件（这在Streamlit中可能不会立即执行）
                    # 我们已经将temp_dir保存到会话状态，以便在下次运行时清理
                    pass



