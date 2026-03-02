import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
import zipfile
import os

st.set_page_config(
    page_title="Excel Formatter Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
        .title-main {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        .success-box {
            background: #d4edda;
            color: #155724;
            padding: 15px;
            border-radius: 5px;
            border-left: 5px solid #28a745;
            margin: 10px 0;
        }
        .error-box {
            background: #f8d7da;
            color: #721c24;
            padding: 15px;
            border-radius: 5px;
            border-left: 5px solid #dc3545;
            margin: 10px 0;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="title-main">
        <h1>📊 Excel Formatter Pro</h1>
        <p>ระบบประมวลผล Master Form พร้อมข้อมูล Color อัตโนมัติ</p>
    </div>
""", unsafe_allow_html=True)

# Functions
def extract_color_data(file_path):
    """ดึงข้อมูล Color จากไฟล์ข้อมูล"""
    wb = load_workbook(file_path)
    ws = wb.active
    
    po_number = ws['H5'].value if ws['H5'].value else 'UNKNOWN'
    colors = []
    
    for row_idx in range(20, ws.max_row + 1):
        cell_a = ws.cell(row=row_idx, column=1)
        cell_j = ws.cell(row=row_idx, column=10)
        
        if cell_a.value and isinstance(cell_a.value, str):
            if '/' in cell_a.value:
                parts = cell_a.value.split('/')
                if len(parts) == 2:
                    code11 = parts[0].strip()
                    code10 = parts[1].strip()
                    qty = cell_j.value if cell_j.value else 0
                    
                    colors.append({
                        'color_code': cell_a.value,
                        'code11': code11,
                        'code10': code10,
                        'qty': int(qty) if qty else 0
                    })
    
    return {
        'po_number': po_number,
        'colors': colors
    }

def process_master_form(master_file_path, data_info):
    """ประมวลผล Master Form"""
    wb = load_workbook(master_file_path)
    ws = wb['Factory code label']
    
    ws['F5'].value = data_info['po_number']
    
    today = datetime.now().strftime('%d/%m/%Y')
    ws['F7'].value = today
    
    ws['B17'].value = 'Tear-Away-Factory-ID-Label'
    
    colors = data_info['colors']
    
    for idx, color_data in enumerate(colors):
        row = 21 + idx
        ws[f'B{row}'].value = 'OPTION 1'
        ws[f'C{row}'].value = color_data['code10']
        ws[f'E{row}'].value = color_data['code11']
        ws[f'F{row}'].value = color_data['qty']
    
    # Return as bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Sidebar
st.sidebar.markdown("### 📋 ขั้นตอน")
st.sidebar.markdown("1. อัพโหลด Master Form")
st.sidebar.markdown("2. อัพโหลดไฟล์ข้อมูล")
st.sidebar.markdown("3. กด 🚀 ประมวลผล")
st.sidebar.markdown("4. ดาวน์โหลดไฟล์")

# Main content
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Master Form")
    master_file = st.file_uploader(
        "อัพโหลด Master Form (.xlsx)",
        type=['xlsx'],
        key='master'
    )

with col2:
    st.subheader("📂 ไฟล์ข้อมูล")
    data_files = st.file_uploader(
        "อัพโหลดไฟล์ข้อมูล (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True,
        key='data'
    )

# Process button
if st.button("🚀 ประมวลผล", key='process_btn', use_container_width=True):
    if not master_file:
        st.markdown("<div class='error-box'>⚠️ โปรดอัพโหลด Master Form</div>", unsafe_allow_html=True)
    elif not data_files:
        st.markdown("<div class='error-box'>⚠️ โปรดอัพโหลดไฟล์ข้อมูล</div>", unsafe_allow_html=True)
    else:
        st.session_state.processed_files = []
        
        progress_bar = st.progress(0)
        total = len(data_files)
        
        for idx, data_file in enumerate(data_files):
            try:
                # Extract data
                with open(f'temp_{data_file.name}', 'wb') as f:
                    f.write(data_file.getbuffer())
                
                data_info = extract_color_data(f'temp_{data_file.name}')
                
                # Process
                output = process_master_form(master_file, data_info)
                
                po_num = data_info['po_number']
                output_name = f"processed_{po_num}_{data_file.name}"
                
                st.session_state.processed_files.append({
                    'name': output_name,
                    'data': output.getvalue(),
                    'po': po_num,
                    'colors': len(data_info['colors'])
                })
                
                # Clean up
                os.remove(f'temp_{data_file.name}')
                
                progress_bar.progress((idx + 1) / total)
                
            except Exception as e:
                st.markdown(f"<div class='error-box'>❌ {data_file.name}: {str(e)}</div>", unsafe_allow_html=True)
        
        if st.session_state.processed_files:
            st.markdown("<div class='success-box'>✅ ประมวลผลสำเร็จ!</div>", unsafe_allow_html=True)

# Display results
if 'processed_files' in st.session_state and st.session_state.processed_files:
    st.subheader("📥 ดาวน์โหลดไฟล์")
    
    col1, col2 = st.columns(2)
    
    # ปุ่มดาวน์โหลดเดี่ยว
    with col1:
        st.markdown("**ดาวน์โหลดแยกไฟล์:**")
        for file_info in st.session_state.processed_files:
            st.download_button(
                label=f"📄 {file_info['name']} ({file_info['colors']} Colors)",
                data=file_info['data'],
                file_name=file_info['name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # ปุ่มดาวน์โหลดทั้งหมด
    with col2:
        st.markdown("**ดาวน์โหลดทั้งหมด:**")
        if st.button("📥 ดาวน์โหลด ZIP ทั้งหมด", use_container_width=True):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for file_info in st.session_state.processed_files:
                    zip_file.writestr(file_info['name'], file_info['data'])
            
            zip_buffer.seek(0)
            st.download_button(
                label="📦 Download ZIP",
                data=zip_buffer.getvalue(),
                file_name="Excel_Formatter_Output.zip",
                mime="application/zip",
                use_container_width=True
            )
    
    # Show summary
    st.subheader("📊 สรุปผลการประมวลผล")
    summary_data = {
        'ไฟล์': [f['name'] for f in st.session_state.processed_files],
        'PO#': [f['po'] for f in st.session_state.processed_files],
        'จำนวน Colors': [f['colors'] for f in st.session_state.processed_files]
    }
    st.table(summary_data)
