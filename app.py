import streamlit as st

# 导入功能模块
from modules.Aword_to_pdf import word_to_pdf
from modules.Appt_to_pdf import ppt_to_pdf
from modules.Excel_to_pdf import excel_to_pdf
from modules.Image_to_pdf import image_to_pdf
from modules.PDF_to_ppt import pdf_to_ppt
from modules.PDF_to_excel import pdf_to_excel
from modules.PDF_to_image import pdf_to_image
from modules.PDF_to_word import pdf_to_word

# 页面配置
st.set_page_config(
    page_title="PDF转换工具", 
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="📄"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-title {
        font-size: 3rem !important;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.8rem !important;
        color: #333;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #1E88E5;
    }
    .card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: transform 0.3s, box-shadow 0.3s;
        margin-bottom: 1rem;
    }
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
    }
    .card-title {
        font-size: 1.2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
        color: #1E88E5;
    }
    .card-text {
        font-size: 0.9rem;
        color: #666;
    }
    .stButton>button {
        background-color: #1E88E5;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        width: 100%;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #1565C0;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
    }
    .footer {
        text-align: center;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #ddd;
        color: #666;
        font-size: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# 初始化session state
if 'current_tool' not in st.session_state:
    st.session_state.current_tool = None

# 主页面功能选择
def show_main_menu():
    st.markdown("<h1 class='main-title'>多功能PDF转换工具</h1>", unsafe_allow_html=True)
    
    # 简介
    st.markdown("""
    <div style="text-align: center; margin-bottom: 2rem;">
        <p style="font-size: 1.2rem; color: #555;">
            一站式解决您的所有PDF转换需求，简单高效，完全免费！
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # 其他文件转PDF
    st.markdown("<h2 class='sub-header'>📥 其他文件转PDF</h2>", unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="card">
            <div class="card-title">Word → PDF</div>
            <div class="card-text">将Word文档转换为PDF格式，保持原始格式不变。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Word → PDF", key="word_pdf", use_container_width=True):
            st.session_state.current_tool = "word_to_pdf"
            st.rerun()
    
    with col2:
        st.markdown("""
        <div class="card">
            <div class="card-title">PPT → PDF</div>
            <div class="card-text">将PowerPoint演示文稿转换为PDF文档，便于分享和查看。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("PPT → PDF", key="ppt_pdf", use_container_width=True):
            st.session_state.current_tool = "ppt_to_pdf"
            st.rerun()
    
    with col3:
        st.markdown("""
        <div class="card">
            <div class="card-title">Excel → PDF</div>
            <div class="card-text">将Excel表格转换为PDF格式，方便打印和分发。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Excel → PDF", key="excel_pdf", use_container_width=True):
            st.session_state.current_tool = "excel_to_pdf"
            st.rerun()
    
    with col4:
        st.markdown("""
        <div class="card">
            <div class="card-title">图片 → PDF</div>
            <div class="card-text">将多张图片合并为一个PDF文档，支持多种图片格式。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("图片 → PDF", key="img_pdf", use_container_width=True):
            st.session_state.current_tool = "image_to_pdf"
            st.rerun()
    
    # PDF转其他文件
    st.markdown("<h2 class='sub-header'>📤 PDF转其他文件</h2>", unsafe_allow_html=True)
    
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        st.markdown("""
        <div class="card">
            <div class="card-title">PDF → 图片</div>
            <div class="card-text">将PDF文档转换为高质量图片，支持多种图片格式输出。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("PDF → 图片", key="pdf_img", use_container_width=True):
            st.session_state.current_tool = "pdf_to_image"
            st.rerun()
    
    with col6:
        st.markdown("""
        <div class="card">
            <div class="card-title">PDF → Word</div>
            <div class="card-text">将PDF文档转换为可编辑的Word文档，保留原始排版。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("PDF → Word", key="pdf_word", use_container_width=True):
            st.session_state.current_tool = "pdf_to_word"
            st.rerun()
    
    with col7:
        st.markdown("""
        <div class="card">
            <div class="card-title">PDF → PPT</div>
            <div class="card-text">将PDF文档转换为PowerPoint演示文稿，便于编辑和演示。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("PDF → PPT", key="pdf_ppt", use_container_width=True):
            st.session_state.current_tool = "pdf_to_ppt"
            st.rerun()
    
    with col8:
        st.markdown("""
        <div class="card">
            <div class="card-title">PDF → Excel</div>
            <div class="card-text">将PDF中的表格数据提取并转换为Excel表格，方便分析。</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("PDF → Excel", key="pdf_excel", use_container_width=True):
            st.session_state.current_tool = "pdf_to_excel"
            st.rerun()
    
    # 页脚
    st.markdown("""
    <div class="footer">
        <p>© 2025 PDF转换工具 | 所有转换在本地完成，保障您的数据安全</p>
        <p>如有问题或建议，请联系我们</p>
        <p>作者：@Baze</p>
        <p>邮箱：15583405928@163.com</p>
    </div>
    """, unsafe_allow_html=True)

# 主函数
def main():
    # 根据session_state显示不同的功能页面
    if st.session_state.current_tool is None:
        show_main_menu()
    else:
        if st.session_state.current_tool == "word_to_pdf":
            word_to_pdf()
        elif st.session_state.current_tool == "ppt_to_pdf":
            ppt_to_pdf()
        elif st.session_state.current_tool == "excel_to_pdf":
            excel_to_pdf()
        elif st.session_state.current_tool == "image_to_pdf":
            image_to_pdf()
        elif st.session_state.current_tool == "pdf_to_image":
            pdf_to_image()
        elif st.session_state.current_tool == "pdf_to_word":
            pdf_to_word()
        elif st.session_state.current_tool == "pdf_to_ppt":
            pdf_to_ppt()
        elif st.session_state.current_tool == "pdf_to_excel":
            pdf_to_excel()

if __name__ == "__main__":
    main()