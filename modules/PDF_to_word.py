import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import zipfile
from io import BytesIO
from utils.common import return_to_main
import base64
import shutil
from utils.common import cleanup_temp_dirs

def convert_pdf_to_docx(pdf_path, docx_path):
    """将单个PDF文件转换为Word文档"""
    try:
        # 使用更详细的配置选项
        cv = Converter(pdf_path)
        # 设置图片处理参数，针对透明PNG的优化配置
        cv.convert(
            docx_path,
            start=0,  # 从第一页开始
            end=None,  # 转换所有页面
            pages=None,  # 转换所有页面
            zoom=1.5,  # 适中的图片质量
            multi_processing=True,  # 启用多进程处理
            grayscale=False,  # 保持彩色
            use_cropbox=True,  # 使用裁剪框
            ignore_errors=True  # 忽略非致命错误
        )
        cv.close()
    except Exception as e:
        # 检查是否为PNG颜色空间问题
        if "unsupported colorspace for 'png'" in str(e):
            try:
                # 尝试使用灰度模式转换
                cv = Converter(pdf_path)
                cv.convert(
                    docx_path,
                    start=0,
                    end=None,
                    pages=None,
                    zoom=1,
                    multi_processing=False,
                    grayscale=True,  # 使用灰度模式
                    use_cropbox=True,
                    ignore_errors=True
                )
                cv.close()
                return
            except Exception as e2:
                st.warning(f"PNG颜色空间问题：{str(e)}。尝试使用灰度模式转换也失败。")
        
        # 如果转换失败，尝试使用降级方案
        try:
            cv = Converter(pdf_path)
            cv.convert(
                docx_path,
                start=0,
                end=None,
                pages=None,
                zoom=1,
                multi_processing=False,
                grayscale=True,  # 尝试使用灰度模式
                use_cropbox=True,
                ignore_errors=True
            )
            cv.close()
        except Exception as e2:
            # 特别针对PNG颜色空间问题的最终尝试
            if "unsupported colorspace for 'png'" in str(e) or "unsupported colorspace for 'png'" in str(e2):
                try:
                    # 尝试设置最低图像质量，优先完成转换
                    cv = Converter(pdf_path)
                    cv.convert(
                        docx_path,
                        start=0,
                        end=None,
                        pages=None,
                        zoom=0.5,  # 降低图像质量
                        multi_processing=False,
                        grayscale=True,
                        use_cropbox=False,
                        ignore_errors=True  # 移除可能不支持的参数
                    )
                    cv.close()
                    st.warning("由于图像格式问题，部分图像质量可能降低。")
                    return
                except Exception as e3:
                    # 如果所有方法都失败，记录详细错误
                    error_msg = f"PNG颜色空间问题导致转换失败:\n原始错误: {str(e)}\n第二次尝试: {str(e2)}\n第三次尝试: {str(e3)}"
                    st.error(error_msg)
                    raise Exception(f"转换失败，PDF包含不支持的PNG图像格式。请尝试先用其他工具编辑此PDF。")
            
            raise Exception(f"转换失败，请检查PDF文件是否包含不支持的图片格式。错误信息：{str(e2)}")

def create_zip_file(file_paths):
    """创建包含所有转换后文件的zip文件"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_path, original_name in file_paths:
            with open(file_path, 'rb') as f:
                zip_file.writestr(original_name.replace('.pdf', '.docx'), f.read())
    return zip_buffer.getvalue()

def pdf_to_word():
    return_to_main()
    st.title("PDF➡️Word")
    st.write("上传多个PDF文件，将其转换为Word文档")
    
    # 初始化临时目录跟踪
    if 'temp_dirs' not in st.session_state:
        st.session_state.temp_dirs = []
    
    
    # 允许多文件上传
    uploaded_files = st.file_uploader("选择PDF文件（可多选）", 
                                    type=['pdf'], 
                                    accept_multiple_files=True)
    
    if uploaded_files:
        st.write(f"已上传 {len(uploaded_files)} 个文件")
        
        if st.button('开始转换'):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            converted_files = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                for index, uploaded_file in enumerate(uploaded_files):
                    # 更新进度
                    progress = (index) / len(uploaded_files)
                    progress_bar.progress(progress)
                    status_text.text(f"正在转换: {uploaded_file.name}")
                    
                    # 创建临时PDF文件
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                        tmp_pdf.write(uploaded_file.getvalue())
                        pdf_path = tmp_pdf.name
                    
                    # 创建临时Word文件路径
                    docx_path = pdf_path.replace('.pdf', '.docx')
                    
                    try:
                        # 执行转换
                        convert_pdf_to_docx(pdf_path, docx_path)
                        converted_files.append((docx_path, uploaded_file.name))
                    except Exception as e:
                        st.error(f'转换 {uploaded_file.name} 时出错：{str(e)}')
                    finally:
                        # 删除临时PDF文件
                        os.unlink(pdf_path)
                
                # 完成进度条
                progress_bar.progress(1.0)
                status_text.text("转换完成！")
                
                if converted_files:
                    if len(converted_files) == 1:
                        # 单个文件直接提供下载
                        docx_path, original_name = converted_files[0]
                        with open(docx_path, 'rb') as f:
                            docx_data = f.read()
                        
                        st.success('文件转换成功！')
                        
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
                        
                        # 创建base64编码的Word文档下载链接
                        docx_filename = original_name.replace('.pdf', '.docx')
                        b64_docx = base64.b64encode(docx_data).decode()
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        href = f'<a href="data:{mime_type};base64,{b64_docx}" download="{docx_filename}" class="download-button">下载Word文档</a>'
                        st.markdown(href, unsafe_allow_html=True)
                    else:
                        # 多个文件创建zip文件
                        zip_data = create_zip_file(converted_files)
                        
                        # 提供zip文件下载
                        st.success(f'成功转换 {len(converted_files)} 个文件！')
                        
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
                        b64_zip = base64.b64encode(zip_data).decode()
                        href = f'<a href="data:application/zip;base64,{b64_zip}" download="converted_documents.zip" class="download-button">下载所有Word文档</a>'
                        st.markdown(href, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f'处理过程中出现错误：{str(e)}')
            
            finally:
                # 清理所有临时文件
                for docx_path, _ in converted_files:
                    try:
                        os.unlink(docx_path)
                    except:
                        pass
