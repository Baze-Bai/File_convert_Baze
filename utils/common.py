import streamlit as st
import shutil
import os

def return_to_main():
    """返回主页的按钮"""
    if st.button("返回主页", key="return_btn"):
        st.session_state.current_tool = None
        st.rerun()

# 清理之前的临时目录
def cleanup_temp_dirs():
    if hasattr(st.session_state, 'temp_dirs'):
        for old_dir in st.session_state.temp_dirs:
            try:
                if os.path.exists(old_dir):
                    shutil.rmtree(old_dir)
            except Exception:
                pass
        st.session_state.temp_dirs = []