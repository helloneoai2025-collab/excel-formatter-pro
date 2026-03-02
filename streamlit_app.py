import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
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
            text-align: center;
            margin-bottom: 30px;
        }
        .success-box {
            background: #d4edda;
            border-left: 5px solid #28a745;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .error-box {
            background: #f8d7da;
            border-left: 5px solid #dc3545;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
    </style>
""", unsafe_allow_html=True)

def extract_color_data(file_bytes, filename):
    try:
        temp_file = f"temp_{filename}"
        with open(temp_file, 'wb') as f:
            f.write(file_bytes)
        
        wb = load_workbook(temp_file)
        ws = wb.active
        
        po_number = ws['H5'].value if ws['H5'].value else 'UNKNOWN'
        colors = []
        
        for row_idx in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=row_idx, column=1)
            cell_j = ws.cell(row=row_idx, column=10)
            
            if cell_a.value and isinstance(cell_a.value, str) and '/' in str(cell_a.value):
                if cell_a.value not in [c['color_full'] for c in colors]:
                    qty = cell_j.value if cell_j.value else 0
                    colors.append({
                        'color_full': str(cell_a.value),
                        'qty': qty
                    })
        
        os.remove(temp_file)
        
        return {
            'success': True,
            'po_number': po_number,
            'colors': colors,
            'color_count': len(colors)
        }
    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }

def process_master_form(master_bytes, color_data_list):
    try:
        results = []
        
        for idx, data_info in enumerate(color_data_list):
            if not data_info['success']:
                results.append({
                    'success': False,
                    'input_file': f"ไฟล์ที่ {idx+1}",
                    'error': data_info['error']
                })
                continue
            
            temp_master = f"temp_master_{idx}.xlsx"
            with open(temp_master, 'wb') as f:
                f.write(master_bytes)
            
            try:
                wb = load_workbook(temp_master)
                ws = wb['Factory code label']
                
                po_number = data_info['po_number']
                colors = data_info['colors']
                
                ws['F5'].value = po_number
                ws['F7'].value = datetime.now().strftime('%d/%m/%Y')
                ws['B17'].value = "Tear-Away-Factory-ID-Label"
                
                start_row = 21
                for col_idx, color_info in enumerate(colors):
                    row = start_row + col_idx
                    
                    color_full = color_info['color_full']
                    parts = color_full.split('/')
                    code11 = parts[0] if len(parts) > 0 else ''
                    code10 = parts[1] if len(parts) > 1 else ''
                    qty = color_info['qty']
                    
                    ws.cell(row=row, column=2).value = "OPTION 1"
                    ws.cell(row=row, column=3).value = code10
                    ws.cell(row=row, column=5).value = code11
                    ws.cell(row=row, column=6).value = qty
                
                output_filename = f"processed_{po_number}_{data_info.get('data_filename', f'file_{idx}')}"
                output_bytes = BytesIO()
                wb.save(output_bytes)
                output_bytes.seek(0)
                
                results.append({
                    'success': True,
                    'input_file': data_info.get('data_filename', f'ไฟล์ที่ {idx+1}'),
                    'po_number': po_number,
                    'color_count': len(colors),
                    'output_filename': output_filename + '.xlsx',
                    'output_bytes': output_bytes.getvalue()
                })
            finally:
                if os.path.exists(temp_master):
                    os.remove(temp_master)
        
        return results
    except Exception as e:
        return [{'success': False, 'error': str(e)}]

st.markdown('<div class="title-main"><h1>📊 Excel Formatter Pro</h1><p>ระบบประมวลผล Master Form พร้อมข้อมูล Color</p></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("ℹ️ ข้อมูล")
    st.markdown("**วิธีใช้:**\n1. อัพโหลด Master Form\n2. อัพโหลดไฟล์ข้อมูล\n3. กด ประมวลผล\n4. ดาวน์โหลดไฟล์")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Master Form")
    master_file = st.file_uploader("เลือก Master Form (.xlsx)", type=['xlsx'], key='master')
    if master_file:
        st.success(f"✅ {master_file.name}")

with col2:
    st.subheader("📂 ไฟล์ข้อมูล")
    data_files = st.file_uploader("เลือกไฟล์ข้อมูล (.xlsx)", type=['xlsx'], accept_multiple_files=True, key='data')
    if data_files:
        st.success(f"✅ {len(data_files)} ไฟล์")

st.markdown("---")

if st.button("🚀 ประมวลผล", use_container_width=True, type="primary"):
    if not master_file:
        st.error("❌ โปรดอัพโหลด Master Form")
    elif not data_files:
        st.error("❌ โปรดอัพโหลดไฟล์ข้อมูล")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        master_bytes = master_file.read()
        
        status_text.text("📊 กำลังดึงข้อมูล...")
        progress_bar.progress(25)
        
        color_data_list = []
        for data_file in data_files:
            data_result = extract_color_data(data_file.read(), data_file.name)
            data_result['data_filename'] = data_file.name
            color_data_list.append(data_result)
        
        status_text.text("⚙️ กำลังประมวลผล...")
        progress_bar.progress(50)
        
        results = process_master_form(master_bytes, color_data_list)
        
        status_text.text("✅ เสร็จสิ้น!")
        progress_bar.progress(100)
        
        st.markdown("---")
        st.subheader("📋 ผลลัพธ์")
        
        success_count = 0
        for result in results:
            if result['success']:
                success_count += 1
                st.success(f"✅ {result['input_file']} | PO#: {result['po_number']} | {result['color_count']} สี")
                st.download_button(
                    label=f"📥 {result['output_filename']}",
                    data=result['output_bytes'],
                    file_name=result['output_filename'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error(f"❌ {result.get('input_file', 'Unknown')}: {result.get('error', 'Unknown error')}")
        
        st.info(f"📈 รวม: {success_count}/{len(results)} ไฟล์สำเร็จ")

st.markdown("---")
st.markdown("<div style='text-align: center; color: #666; font-size: 12px;'><p>Excel Formatter Pro v1.0 | Powered by Streamlit</p></div>", unsafe_allow_html=True)
