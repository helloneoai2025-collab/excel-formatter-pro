import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from io import BytesIO
import zipfile
import os
from copy import copy

st.set_page_config(
    page_title="Excel Formatter F1 Pro",
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
        <h1>📊 Excel Formatter F1 Pro (Streamlit)</h1>
        <p>ระบบประมวลผล Master Form พร้อมข้อมูล Color - Logic F1</p>
    </div>
""", unsafe_allow_html=True)

# ===== F1 LOGIC FUNCTIONS =====

def extract_color_data_f1(file_path):
    """ดึงข้อมูล Color จากไฟล์ข้อมูล (F1 Logic)"""
    wb = load_workbook(file_path)
    ws = wb.active
    
    po_number = ws['H5'].value if ws['H5'].value else 'UNKNOWN'
    colors = []
    
    # หา Blue Zone ก่อน (RGB: FF00B0F0)
    for row_idx in range(20, ws.max_row + 1):
        cell_a = ws.cell(row=row_idx, column=1)
        
        is_blue = False
        if cell_a.fill and cell_a.fill.start_color:
            try:
                if hasattr(cell_a.fill.start_color, 'rgb'):
                    if cell_a.fill.start_color.rgb == 'FF00B0F0':
                        is_blue = True
            except:
                pass
        
        if is_blue and cell_a.value and isinstance(cell_a.value, str) and '/' in cell_a.value:
            cell_j = ws.cell(row=row_idx, column=10)
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
    
    # ถ้าไม่มี Blue Zone ให้ดึงจากข้อมูลธรรมดา
    if not colors:
        for row_idx in range(20, ws.max_row + 1):
            cell_a = ws.cell(row=row_idx, column=1)
            
            if cell_a.value and isinstance(cell_a.value, str) and '/' in cell_a.value:
                cell_j = ws.cell(row=row_idx, column=10)
                parts = cell_a.value.split('/')
                
                if len(parts) == 2:
                    code11 = parts[0].strip()
                    code10 = parts[1].strip()
                    
                    if len(code11) >= 5 and len(code10) >= 2:
                        qty = cell_j.value if cell_j.value else 0
                        colors.append({
                            'color_code': cell_a.value,
                            'code11': code11,
                            'code10': code10,
                            'qty': int(qty) if qty else 0
                        })
    
    return {'po_number': po_number, 'colors': colors}

def copy_cell_style(source_cell, target_cell):
    """Copy Formatting จาก source ไป target"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def process_master_form_f1(master_file_path, data_info):
    """ประมวลผล Master Form (F1 Logic) - Copy Formatting ทั้งหมด"""
    wb = load_workbook(master_file_path)
    ws = wb['Factory code label']
    
    # เติม PO# ใน F5
    ws['F5'].value = data_info['po_number']
    
    # เติม DATE ใน F7
    today = datetime.now().strftime('%d/%m/%Y')
    ws['F7'].value = today
    
    # เติม CUSTOMER ITEM CODE ใน B17
    ws['B17'].value = 'Tear-Away-Factory-ID-Label'
    
    colors = data_info['colors']
    
    # ดึง Template Row (Row 21) สำหรับ Copy Formatting
    template_row = 21
    
    # เติม OPTION 1, CODE, QTY พร้อม Copy Formatting
    for idx, color_data in enumerate(colors):
        row = 21 + idx
        
        # Column B - OPTION 1
        template_b = ws.cell(row=template_row, column=2)
        target_b = ws.cell(row=row, column=2)
        target_b.value = 'OPTION 1'
        copy_cell_style(template_b, target_b)
        
        # Column C - Code10
        template_c = ws.cell(row=template_row, column=3)
        target_c = ws.cell(row=row, column=3)
        target_c.value = color_data['code10']
        copy_cell_style(template_c, target_c)
        
        # Column E - Code11
        template_e = ws.cell(row=template_row, column=5)
        target_e = ws.cell(row=row, column=5)
        target_e.value = color_data['code11']
        copy_cell_style(template_e, target_e)
        
        # Column F - Qty
        template_f = ws.cell(row=template_row, column=6)
        target_f = ws.cell(row=row, column=6)
        target_f.value = color_data['qty']
        copy_cell_style(template_f, target_f)
    
    # Return as bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Sidebar
st.sidebar.markdown("### 📋 ขั้นตอน (F1 Logic)")
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
if st.button("🚀 ประมวลผล (F1 Logic)", key='process_btn', use_container_width=True):
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
                # Save temp file
                with open(f'temp_{data_file.name}', 'wb') as f:
                    f.write(data_file.getbuffer())
                
                # Extract data (F1 Logic)
                data_info = extract_color_data_f1(f'temp_{data_file.name}')
                
                # Process
                output = process_master_form_f1(master_file, data_info)
                
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
            st.markdown("<div class='success-box'>✅ ประมวลผลสำเร็จ! (F1 Logic - Copy Formatting)</div>", unsafe_allow_html=True)

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
                label="📦 ดาวน์โหลด ZIP",
                data=zip_buffer.getvalue(),
                file_name="Excel_Formatter_Output.zip",
                mime="application/zip",
                use_container_width=True
            )
    
    # Show summary
    st.subheader("📊 สรุปผลการประมวลผล (F1 Logic)")
    summary_data = {
        'ไฟล์': [f['name'] for f in st.session_state.processed_files],
        'PO#': [f['po'] for f in st.session_state.processed_files],
        'จำนวน Colors': [f['colors'] for f in st.session_state.processed_files]
    }
    st.table(summary_data)
