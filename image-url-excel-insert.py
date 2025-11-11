import streamlit as st
import pandas as pd
import requests
from PIL import Image as PILImage
from io import BytesIO
import os
import re
import base64
import mimetypes
import time
import xlsxwriter

st.set_page_config(page_title="Excel 图片处理工具", page_icon="📊", layout="wide")
st.title("📊 Excel 图片处理工具 - 图片嵌入单元格版")
st.write("上传包含图片链接的 Excel 文件，自动下载图片并嵌入单元格。支持 webp 转 png。")

# --- 工具函数 ---
def register_webp_mimetype():
    try:
        if '.webp' not in mimetypes.types_map:
            mimetypes.add_type('image/webp', '.webp')
        st.success("✅ 已注册 .webp 格式")
    except Exception as e:
        st.warning(f"⚠️ MIME注册警告: {str(e)}")

def convert_webp_to_png(img, save_path):
    try:
        img = img.convert("RGBA")
        img.save(save_path, 'PNG')
        return save_path
    except:
        return None

def download_image(image_url, save_path):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(image_url, headers=headers, timeout=20)
        response.raise_for_status()
        img = PILImage.open(BytesIO(response.content))
        ext = img.format.lower()
        if ext == "webp":
            return convert_webp_to_png(img, save_path)
        else:
            img.save(save_path)
            return save_path
    except:
        return None

# --- Streamlit 页面 ---
register_webp_mimetype()

uploaded_file = st.file_uploader("📁 上传 Excel 文件 (.xlsx)", type=["xlsx"])
sheet_name = st.text_input("工作表名称（默认Sheet1）", value="Sheet1")

if uploaded_file and st.button("开始处理"):
    st.info("⏳ 开始处理，请稍候...")

    # 读取上传的 Excel
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # 创建输出 Excel
    output_path = f"output_embedded_{uploaded_file.name}"
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet(sheet_name)

    # 设置单元格默认大小
    row_height = 100
    col_width = 20
    for i in range(len(df.columns)):
        worksheet.set_column(i, i, col_width)
    for i in range(len(df)):
        worksheet.set_row(i, row_height)

    # 创建进度条和状态显示
    total = df.size
    progress_bar = st.progress(0)
    status_text = st.empty()
    start_time = time.time()
    success_count = 0
    fail_count = 0
    temp_folder = "temp_images"
    os.makedirs(temp_folder, exist_ok=True)

    # 遍历每个单元格
    for row_idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            progress = (row_idx*len(df.columns)+col_idx+1)/total
            progress_bar.progress(int(progress*100))
            status_text.text(f"处理单元格 {row_idx+1},{col_idx+1}，成功 {success_count} 张，失败 {fail_count} 张")

            if isinstance(cell, str) and cell.startswith("http") and any(ext in cell.lower() for ext in ['jpg','jpeg','png','webp','gif','bmp','svg']):
                safe_name = re.sub(r'[^\w\.]', '_', f"{row_idx}_{col_idx}.png")
                save_path = os.path.join(temp_folder, safe_name)
                img_path = download_image(cell, save_path)
                if img_path:
                    try:
                        # 计算缩放比例，让图片填充单元格
                        img = PILImage.open(img_path)
                        x_scale = col_width*7 / img.width
                        y_scale = row_height*0.75 / img.height
                        scale = min(x_scale, y_scale)
                        worksheet.insert_image(row_idx, col_idx, img_path, {'x_scale': scale, 'y_scale': scale})
                        success_count += 1
                    except:
                        fail_count += 1
                else:
                    fail_count += 1
            else:
                # 普通文字
                worksheet.write(row_idx, col_idx, cell)

    workbook.close()
    elapsed = int(time.time() - start_time)
    st.success(f"✅ 处理完成！成功插入 {success_count} 张图片，失败 {fail_count} 张，耗时 {elapsed} 秒")

    # 下载按钮
    with open(output_path, "rb") as f:
        bytes_data = f.read()
        b64 = base64.b64encode(bytes_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{output_path}">📥 下载处理后的 Excel 文件</a>'
        st.markdown(href, unsafe_allow_html=True)
