import sys
import os
import re
import types

# ==========================================
# 0. POLYFILL (For Python 3.12+ compatibility)
# Must be before any other major imports
# ==========================================
if 'distutils' not in sys.modules:
    class LooseVersion:
        def __init__(self, vstring):
            self.vstring = str(vstring)
            self.version = [int(x) if x.isdigit() else x for x in re.split(r'([0-9]+|\.)', self.vstring) if x and x != '.']
        def __str__(self): return self.vstring
        def __ge__(self, other): return True
        def __lt__(self, other): return False

    d = types.ModuleType('distutils')
    d.version = types.ModuleType('distutils.version')
    d.version.LooseVersion = LooseVersion
    sys.modules['distutils'] = d
    sys.modules['distutils.version'] = d.version

import streamlit as st
import pandas as pd
import time
import threading
from scraper import TaxScraper, format_excel, setup_driver, get_chrome_main_version

# Cấu hình trang
st.set_page_config(
    page_title="VSIC - Công cụ trích xuất Mã Số Thuế",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện Premium
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #f8fafc;
    }
    
    .stButton>button {
        width: 100%;
        border-radius: 12px;
        height: 3em;
        background: linear-gradient(90deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        font-weight: 600;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.4);
        background: linear-gradient(90deg, #2563eb 0%, #1d4ed8 100%);
    }
    
    .status-card {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 20px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin-bottom: 20px;
    }
    
    h1, h2, h3 {
        color: #60a5fa !important;
        font-weight: 600 !important;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #94a3b8;
    }
    
    .metric-value {
        font-size: 1.8rem;
        font-weight: 600;
        color: #ffffff;
    }
    
    .stProgress > div > div > div > div {
        background-color: #3b82f6;
    }
    
    [data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }
</style>
""", unsafe_allow_html=True)

# Khởi tạo Session State
if 'df' not in st.session_state:
    st.session_state.df = None
if 'is_running' not in st.session_state:
    st.session_state.is_running = False
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'output_file' not in st.session_state:
    st.session_state.output_file = ""

def add_log(msg):
    st.session_state.logs.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
    if len(st.session_state.logs) > 50:
        st.session_state.logs.pop(0)

# Sidebar
with st.sidebar:
    st.image("https://www.google.com/images/branding/googlelogo/2x/googlelogo_light_color_92x30dp.png", width=120)
    st.title("Cấu hình")
    
    uploaded_file = st.file_uploader("Tải lên file Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        try:
            if st.session_state.df is None:
                # Đọc toàn bộ file dưới dạng string để tránh lỗi định dạng số (MST dài)
                st.session_state.df = pd.read_excel(uploaded_file, dtype=str)
                st.success(f"Đã tải {len(st.session_state.df)} dòng.")
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

    col_name = st.text_input("Tên cột chứa MST", value="Mã số thuế")
    
    st.divider()
    st.markdown("### Tùy chọn")
    chrome_ver = st.number_input("Phiên bản Chrome (version_main)", value=get_chrome_main_version(), step=1)
    auto_retry = st.checkbox("Tự động thử lại nếu thất bại", value=True)
    headless = st.checkbox("Chạy ẩn (Headless)", value=True, help="Bắt buộc nếu chạy trên Cloud/VPS Linux")

# Main Content
st.title("🚀 VSIC Tax Intelligence")
st.markdown("Trích xuất thông tin doanh nghiệp từ mã số thuế một cách chuyên nghiệp.")

if st.session_state.df is not None:
    # Dashboard Metrics
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown('<div class="status-card"><p class="metric-label">Tổng cộng</p><p class="metric-value">{}</p></div>'.format(len(st.session_state.df)), unsafe_allow_html=True)
    with m2:
        success_count = len(st.session_state.df[st.session_state.df.get('Trạng thái') == 'Thành công']) if 'Trạng thái' in st.session_state.df.columns else 0
        st.markdown('<div class="status-card"><p class="metric-label">Thành công</p><p class="metric-value" style="color: #4ade80;">{}</p></div>'.format(success_count), unsafe_allow_html=True)
    with m3:
        failed_count = len(st.session_state.df[st.session_state.df.get('Trạng thái') == 'Không tìm thấy dữ liệu']) if 'Trạng thái' in st.session_state.df.columns else 0
        st.markdown('<div class="status-card"><p class="metric-label">Thất bại</p><p class="metric-value" style="color: #f87171;">{}</p></div>'.format(failed_count), unsafe_allow_html=True)
    with m4:
        progress_val = (success_count + failed_count) / len(st.session_state.df) * 100 if len(st.session_state.df) > 0 else 0
        st.markdown('<div class="status-card"><p class="metric-label">Tiến độ</p><p class="metric-value" style="color: #60a5fa;">{:.1f}%</p></div>'.format(progress_val), unsafe_allow_html=True)

    if not st.session_state.is_running:
        if st.button("BẮT ĐẦU TRÍCH XUẤT"):
            st.session_state.is_running = True
            st.rerun()
    else:
        if st.button("DỪNG LẠI"):
            st.session_state.is_running = False
            st.rerun()

    if st.session_state.is_running:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        df = st.session_state.df.copy().astype(str)
        df = df.replace('nan', '')
        
        if 'Tên công ty' not in df.columns: df['Tên công ty'] = ""
        if 'Mã Ngành Chính' not in df.columns: df['Mã Ngành Chính'] = ""
        if 'Ngành nghề chính' not in df.columns: df['Ngành nghề chính'] = ""
        if 'Trạng thái' not in df.columns: df['Trạng thái'] = ""

        scraper = None
        try:
            scraper = TaxScraper(setup_driver(headless=headless, version_main=int(chrome_ver)))
            
            # --- LƯỢT 1 ---
            for index, row in df.iterrows():
                if not st.session_state.is_running:
                    break
                    
                code = str(row.get(col_name, "")).strip()
                if not code or code == 'nan' or row.get('Trạng thái') == 'Thành công':
                    continue
                
                status_text.markdown(f"🔍 Lượt 1 - Đang xử lý: **{code}** ({index+1}/{len(df)})")
                progress_bar.progress((index + 1) / len(df))
                
                result = scraper.search_tax_code(code)
                
                if result:
                    for key, value in result.items():
                        if key not in df.columns: df[key] = ""
                        df.at[index, key] = value
                    
                    if result.get('Tên công ty') != "Không tìm thấy":
                        df.at[index, 'Trạng thái'] = "Thành công"
                        add_log(f"✅ {code}: {result.get('Tên công ty')}")
                    else:
                        df.at[index, 'Trạng thái'] = "Không tìm thấy dữ liệu"
                        add_log(f"❌ {code}: Không tìm thấy")
                else:
                    df.at[index, 'Trạng thái'] = "Lỗi kết nối"
                    add_log(f"⚠️ {code}: Lỗi kết nối")
                
                st.session_state.df = df
                if (index + 1) % 5 == 0:
                    temp_file = "ket_qua_tam.xlsx"
                    df.to_excel(temp_file, index=False)
                    format_excel(temp_file)

            # --- LƯỢT 2 (TỰ ĐỘNG THỬ LẠI LỖI) ---
            if auto_retry and st.session_state.is_running:
                failed_mask = df['Trạng thái'].isin(['Không tìm thấy dữ liệu', 'Lỗi kết nối', 'nan', ''])
                failed_indices = df[failed_mask & (df[col_name] != '')].index
                
                if len(failed_indices) > 0:
                    add_log(f"🔄 Thử lại lần 2 cho {len(failed_indices)} mã lỗi...")
                    time.sleep(2) # Nghỉ một chút trước khi thử lại
                    
                    for i, index in enumerate(failed_indices):
                        if not st.session_state.is_running: break
                        
                        code = str(df.at[index, col_name]).strip()
                        status_text.markdown(f"🔄 Lượt 2 - Đang xử lý: **{code}** ({i+1}/{len(failed_indices)})")
                        progress_bar.progress((i + 1) / len(failed_indices))
                        
                        result = scraper.search_tax_code(code)
                        if result and result.get('Tên công ty') != "Không tìm thấy":
                            for key, value in result.items():
                                if key not in df.columns: df[key] = ""
                                df.at[index, key] = value
                            df.at[index, 'Trạng thái'] = "Thành công"
                            add_log(f"✅ [Lượt 2] {code}: {result.get('Tên công ty')}")
                        
                        st.session_state.df = df
            
            st.session_state.is_running = False
            output_file = "ket_qua_cuoi_cung.xlsx"
            df.to_excel(output_file, index=False)
            format_excel(output_file)
            st.session_state.output_file = output_file
            st.success("🎉 Quá trình hoàn tất!")
            st.rerun()

        except Exception as e:
            st.error(f"Lỗi nghiêm trọng: {e}")
            st.session_state.is_running = False
        finally:
            if scraper:
                try: scraper.close()
                except: pass

    st.subheader("Dữ liệu hiện tại")
    st.dataframe(st.session_state.df, use_container_width=True)

    if st.session_state.output_file and os.path.exists(st.session_state.output_file):
        with open(st.session_state.output_file, "rb") as f:
            st.download_button(
                label="📥 TẢI XUỐNG KẾT QUẢ EXCEL",
                data=f,
                file_name="VSIC_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with st.expander("Xem nhật ký hoạt động"):
        for log in reversed(st.session_state.logs):
            st.text(log)
else:
    st.info("👋 Vui lòng tải lên file Excel ở thanh bên trái để bắt đầu.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="status-card"><h3>⚡ Tốc độ cao</h3><p>Tối ưu hóa quy trình tìm kiếm tự động.</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="status-card"><h3>🛡️ Vượt rào cản</h3><p>Sử dụng công nghệ mô phỏng thực tế.</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="status-card"><h3>📊 Định dạng chuẩn</h3><p>Tự động định dạng Excel chuyên nghiệp.</p></div>', unsafe_allow_html=True)
