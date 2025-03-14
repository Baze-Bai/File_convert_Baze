import streamlit as st
import pandas as pd
import os
from io import BytesIO
from utils.common import return_to_main, cleanup_temp_dirs
import base64
import tempfile
import shutil
import io
import tabula
import PyPDF2

def extract_pdf_tables(pdf_file, file_status, page_progress=None, idx=0):
    """使用tabula-py从PDF文件中提取表格，并显示进度"""
    # 保存上传的文件到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_path = temp_pdf.name
    
    try:
        # 获取PDF总页数
        pdf_reader = PyPDF2.PdfReader(temp_path)
        total_pages = len(pdf_reader.pages)
        
        # 不再创建独立进度条，使用传入的进度条或只更新文本状态
        
        # 存储所有页面的表格
        all_tables = []
        
        # 从各页面提取表格
        for page in range(1, total_pages + 1):
            try:
                # 如果提供了进度条，则更新进度
                if page_progress:
                    page_progress.progress(page/total_pages, text=f"页面处理进度: {int(page/total_pages*100)}%")
                else:
                    # 否则只更新文本状态
                    file_status.write(f"正在处理第{idx+1}个文件的第 {page}/{total_pages} 页...")
                
                # 使用tabula提取当前页面的表格
                # lattice=True 适用于有明显边框线的表格
                tables = tabula.read_pdf(
                    temp_path, 
                    pages=page, 
                    multiple_tables=True,
                    lattice=True,
                    guess=False,
                    pandas_options={'header': None}  # 不自动识别表头
                )
                
                # 如果当前页有表格
                if tables and len(tables) > 0:
                    for i, table in enumerate(tables):
                        # 为每个表格添加页码和表格索引信息
                        all_tables.append({
                            'page': page,
                            'table_index': i,
                            'dataframe': table
                        })
                else:
                    file_status.write(f"第 {page} 页未检测到边框表格，尝试无边框模式...")
                    # 如果没有检测到表格，尝试stream模式（适用于无明显边框的表格）
                    tables = tabula.read_pdf(
                        temp_path, 
                        pages=page, 
                        multiple_tables=True,
                        lattice=False,
                        stream=True,
                        guess=False,
                        pandas_options={'header': None}
                    )
                    
                    if tables and len(tables) > 0:
                        file_status.write(f"第 {page} 页无边框模式检测到 {len(tables)} 个表格")
                        for i, table in enumerate(tables):
                            all_tables.append({
                                'page': page,
                                'table_index': i,
                                'dataframe': table
                            })
                    else:
                        file_status.write(f"第 {page} 页未检测到任何表格")
            except Exception as e:
                file_status.warning(f"提取第{page}页表格时出错: {str(e)}")
                continue
            
        # 更新进度条状态为完成（如果提供了进度条）
        if page_progress:
            page_progress.progress(1.0, text="页面处理完成")
        return all_tables
    
    except Exception as e:
        # 检查异常是否是Java RuntimeException
        error_message = str(e)
        if "java.lang.RuntimeException" in error_message:
            file_status.error(f"运行tabula时出错: {error_message}")
            file_status.info("请确保已安装Java运行环境(JRE)，tabula-py依赖Java")
        else:
            file_status.error(f"提取表格时出错: {error_message}")
        return []
    finally:
        # 清理临时文件
        if os.path.exists(temp_path):
            os.unlink(temp_path)

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
        # 检查Java环境
        try:
            import jpype
        except:
            pass
            
        if st.button("开始转换"):
            # 清理之前的临时目录
            cleanup_temp_dirs()
            
            # 创建新的临时目录
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dirs.append(temp_dir)
            
            # 创建一个列表来存储所有PDF的处理结果
            all_pdf_results = []
            
            # 创建一个统一的进度条和状态显示区域
            progress_bar = st.progress(0, text="总体进度: 0%")
            status_area = st.container()
            status_display = status_area.empty()
            
            # 计算总任务数
            total_files = len(uploaded_files)
            
            # 处理每个上传的文件
            for i, uploaded_file in enumerate(uploaded_files):
                file_name = uploaded_file.name
                
                # 更新状态显示区域
                status_display.info(f"正在处理文件 ({i+1}/{total_files}): {file_name}")
                
                try:
                    # 从上传的文件中提取表格
                    tables = extract_pdf_tables(uploaded_file, status_display, idx=i)
                    
                    if tables and len(tables) > 0:
                        status_display.info(f"成功从 {file_name} 提取 {len(tables)} 个表格，正在生成Excel...")
                        
                        # 创建工作表
                        excel_path = os.path.join(temp_dir, file_name.replace('.pdf', '.xlsx'))
                        
                        # 使用一个ExcelWriter来处理所有页面的表格
                        with pd.ExcelWriter(excel_path) as writer:
                            # 按页码组织表格
                            page_tables = {}
                            for table_info in tables:
                                page = table_info['page']
                                if page not in page_tables:
                                    page_tables[page] = []
                                page_tables[page].append(table_info)
                            
                            # 将每页的表格写入对应的工作表
                            total_pages = len(page_tables)
                            for idx, (page, page_table_infos) in enumerate(page_tables.items()):
                                # 更新状态显示
                                status_display.info(f"正在处理 {file_name}: 创建Excel工作表 第{page}页 ({idx+1}/{total_pages})")
                                
                                # 如果页面有多个表格，合并到一个工作表
                                if len(page_table_infos) == 1:
                                    # 只有一个表格时直接使用
                                    df = page_table_infos[0]['dataframe']
                                else:
                                    # 多个表格时，垂直堆叠
                                    dfs = [info['dataframe'] for info in sorted(page_table_infos, key=lambda x: x['table_index'])]
                                    df = pd.concat(dfs, axis=0, ignore_index=True)
                                
                                # 使用唯一的工作表名
                                sheet_name = f"第{page}页"
                                # 写入Excel工作表
                                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        
                        # 读取生成的Excel文件
                        with open(excel_path, 'rb') as f:
                            excel_data = f.read()
                        
                        all_pdf_results.append({
                            "文件名": file_name,
                            "Excel数据": excel_data
                        })
                        
                        status_display.success(f"✅ 文件 {file_name} 处理完成")
                    else:
                        status_display.warning(f"{file_name} 未检测到表格，尝试提取文本内容...")
                        # 回退到文本提取
                        uploaded_file.seek(0)
                        pdf_reader = PyPDF2.PdfReader(uploaded_file)
                        
                        # 创建工作表
                        excel_path = os.path.join(temp_dir, file_name.replace('.pdf', '.xlsx'))
                        with pd.ExcelWriter(excel_path) as writer:
                            total_pages = len(pdf_reader.pages)
                            for page_num, page in enumerate(pdf_reader.pages):
                                # 更新状态
                                if (page_num % 5 == 0) or (page_num == total_pages - 1):  # 每5页更新一次状态，减少更新频率
                                    status_display.info(f"正在处理 {file_name}: 提取文本 第{page_num+1}页 ({page_num+1}/{total_pages})")
                                
                                text = page.extract_text()
                                df = pd.DataFrame({"内容": [text]})
                                sheet_name = f"第{page_num+1}页"
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                        # 读取生成的Excel文件
                        with open(excel_path, 'rb') as f:
                            excel_data = f.read()
                        
                        all_pdf_results.append({
                            "文件名": file_name,
                            "Excel数据": excel_data
                        })
                        
                        status_display.success(f"✅ 文件 {file_name} 处理完成 (仅提取文本)")
                
                except Exception as e:
                    status_display.error(f"处理文件 {file_name} 时出错: {str(e)}")
                
                # 更新总进度条
                progress = (i + 1) / total_files
                progress_bar.progress(progress, text=f"总体进度: {int(progress * 100)}%")
            
            # 完成所有处理，更新最终状态
            status_display.success("所有文件处理完成！")
            progress_bar.progress(1.0, text="转换完成！")
            
            if all_pdf_results:
                # 判断是单个PDF还是多个PDF
                if len(all_pdf_results) == 1:
                    # 单个PDF文件 - 直接提供Excel下载
                    pdf_result = all_pdf_results[0]
                    file_name = os.path.splitext(pdf_result["文件名"])[0] + ".xlsx"
                    
                    # 直接使用已生成的Excel数据
                    excel_data = pdf_result["Excel数据"]
               
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
                    
                    # 创建下载链接
                    b64_excel = base64.b64encode(excel_data).decode()
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    href = f'<a href="data:{mime_type};base64,{b64_excel}" download="{file_name}" class="download-button">下载Excel文件</a>'
                    st.markdown(href, unsafe_allow_html=True)
                    st.success(f"成功将 {pdf_result['文件名']} 转换为Excel！")
                
                else:
                    # 多个PDF文件 - 创建ZIP压缩包
                    with st.spinner("正在创建ZIP压缩包..."):
                        import zipfile
                        from io import BytesIO
                        
                        # 创建内存中的ZIP文件
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            # 循环添加每个Excel文件到ZIP
                            for pdf_result in all_pdf_results:
                                file_name = pdf_result["文件名"]
                                excel_name = os.path.splitext(file_name)[0] + ".xlsx"
                                
                                if "Excel数据" in pdf_result:
                                    # 直接从内存添加Excel数据到ZIP
                                    zip_file.writestr(excel_name, pdf_result["Excel数据"])
                        
                        # 准备下载ZIP文件
                        zip_data = zip_buffer.getvalue()
                        
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
                        
                        # 创建ZIP下载链接
                        b64_zip = base64.b64encode(zip_data).decode()
                        href = f'<a href="data:application/zip;base64,{b64_zip}" download="PDF转换结果.zip" class="download-button">下载所有Excel文件(ZIP压缩包)</a>'
                        st.markdown(href, unsafe_allow_html=True)
                        st.success(f"成功将 {len(all_pdf_results)} 个PDF文件转换为Excel并打包！")
