import streamlit as st
import pandas as pd
import datetime
import re
import os
from difflib import get_close_matches
from PIL import Image, ImageEnhance, ImageOps
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
import io

# 设置页面配置（手机适配）
st.set_page_config(page_title="食材核对系统 v7.8", layout="centered")


# --- 核心逻辑函数 ---
def init_pdf_font():
    # 尝试多种路径，包括 Linux 下的默认路径和当前目录下的字体文件
    font_paths = [
        "simsun.ttc",  # 建议你把 simsun.ttc 字体文件也上传到 GitHub 仓库
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "C:/Windows/Fonts/simsun.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("SimSun", p))
            return "SimSun"
    return "Helvetica"  # 万不得已返回默认英文字体


def enhance_image(img_file):
    img = Image.open(img_file)
    img = ImageOps.exif_transpose(img)
    img = img.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2.2)
    img = ImageEnhance.Sharpness(img).enhance(1.5)
    return img


def parse_text_input(text_data, db_names, db_records):
    parsed_results = []
    if not text_data: return parsed_results
    lines = text_data.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        match = re.search(r'([^\d\s]+)\s*(\d+\.?\d*)\s*([^\d\s]*)', line)
        if match:
            name_str = match.group(1)
            qty_val = float(match.group(2))
            unit_str = match.group(3)
            if unit_str == "斤" and not line.endswith("公斤"):
                qty_val = qty_val / 2
                unit_str = "kg"
            match_name = get_close_matches(name_str, db_names, n=1, cutoff=0.5)
            if match_name:
                item = next(i for i in db_records if i['名称'] == match_name[0])
                parsed_results.append({
                    'cat': item['分类'], 'name': match_name[0],
                    'brand': '文本录入', 'spec': '-', 'unit': item['单位'], 'qty': qty_val
                })
    return parsed_results


# --- UI 界面 ---
st.title("🍎 食材报价核对终端")
st.caption("v7.8 移动网页版 - 支持拍照与Excel上传")

# 数据库加载
db_file = "数据库02.xlsx"
if not os.path.exists(db_file):
    st.error(f"未找到数据库文件: {db_file}")
    st.stop()

db_df = pd.read_excel(db_file)
db_df['名称'] = db_df['名称'].astype(str).str.strip()
db_names = db_df['名称'].unique().tolist()
db_records = db_df.to_dict('records')

# 1. Excel 上传
uploaded_excels = st.file_uploader("✚ 添加 Excel 订单", type=["xlsx", "xls"], accept_multiple_files=True)

# 2. 图片拍照上传
uploaded_imgs = st.file_uploader("📷 添加手写单图片 (拍照)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# 3. 快捷文本输入
text_input = st.text_area("快捷文本输入 (格式: 名称 数量)", height=150, placeholder="例如：土豆 10\n西红柿 5.5斤")

if st.button("🚀 生成 PDF 核对清单", use_container_width=True):
    all_results = []

    # 处理 Excel
    if uploaded_excels:
        for f in uploaded_excels:
            file_tag = f.name[:2]
            df = pd.read_excel(f, header=1)
            col_qty = next((c for c in df.columns if '5号' in str(c)), None)
            for _, row in df.iterrows():
                name = str(row.get('名称', ''))
                qty = row.get(col_qty)
                if pd.notna(qty) and qty != 0 and name != 'nan':
                    match = get_close_matches(name, db_names, n=1, cutoff=0.6)
                    if match:
                        item = next(i for i in db_records if i['名称'] == match[0])
                        all_results.append({
                            'cat': item['分类'], 'name': match[0],
                            'brand': f"{row.get('品牌', '')}({row.get('备注', '')})".replace('nan', ''),
                            'spec': f"{item['规格']}-({file_tag})", 'unit': item['单位'], 'qty': qty
                        })

    # 处理文本
    text_results = parse_text_input(text_input, db_names, db_records)
    all_results.extend(text_results)

    if not all_results and not uploaded_imgs:
        st.warning("请先输入数据或上传图片")
    else:
        # 生成 PDF
        buffer = io.BytesIO()
        font_name = init_pdf_font()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4

        y = height - 50
        if all_results:
            df_res = pd.DataFrame(all_results)
            categories = ["水菜", "硬菜", "肉类", "冻品", "干货", "豆制品", "豆芽海带丝杂菜", "鸡鸭类", "菌菇类",
                          "水果", "酱菜类"]
            for cat in categories:
                cat_items = df_res[df_res['cat'] == cat]
                if cat_items.empty: continue
                if y < 100: c.showPage(); y = height - 50
                c.setFont(font_name, 12);
                c.setFillColor(colors.blue)
                c.drawString(40, y, f"【{cat}】 ({len(cat_items)}项)");
                y -= 25

                item_index = 1
                c.setFont(font_name, 11);
                c.setFillColor(colors.black)
                for _, r in cat_items.iterrows():
                    c.drawString(60, y, f"{item_index}. {r['name'][:12]}")
                    c.drawString(180, y, f"{r['brand'][:15]}")
                    c.drawString(330, y, f"{r['spec'][:15]}")
                    c.drawString(480, y, f"{round(r['qty'], 2)} {r['unit']}")
                    y -= 20;
                    item_index += 1
                    if y < 50: c.showPage(); y = height - 50; c.setFont(font_name, 11)
                y -= 15

        if uploaded_imgs:
            for img_file in uploaded_imgs:
                c.showPage()
                c.setFont(font_name, 11);
                c.drawString(40, height - 30, f"原单存档：{img_file.name}")
                enhanced = enhance_image(img_file)
                img_w, img_h = enhanced.size
                ratio = min((width - 80) / img_w, (height - 100) / img_h)
                c.drawInlineImage(enhanced, 40, height - 60 - img_h * ratio, width=img_w * ratio, height=img_h * ratio)

        c.save()
        st.success("✅ 清单已生成！")
        st.download_button("📥 点击下载 PDF 清单", data=buffer.getvalue(),
                           file_name=f"核对清单_{datetime.datetime.now().strftime('%m%d_%H%M')}.pdf",
                           mime="application/pdf")