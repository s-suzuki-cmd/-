import streamlit as st
import cv2
import easyocr
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import re
import calendar
from datetime import datetime
import numpy as np
import io
import os
import pypdfium2 as pdfium
from difflib import SequenceMatcher

# PDF生成ライブラリ
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

st.set_page_config(page_title="ピッキングリスト自動解析＆シール指示ツール", layout="wide")

# 対象品番マスター
DEFAULT_MASTER = {
    "99092-77R23-N02": "AD-1957ZS/JP",
    "99092-84UR5-N01": "AD-1957ZS02/JP",
    "99000-79BP3-000": "AN-1327ZS",
    "99000-79Y27-PF2": "AN-MRZ99ZS-71A",
    "99000-79BC1-000": "AN-RZ900ZS-71A",
    "99000-79AR8-PF1": "AN-ZH0777ZS-7A",
    "99000-79Y27-PF1": "AN-ZH09ZS-71A",
    "99000-79BP4-000": "CD-1317ZS",
    "99000-79Y64-000": "CD-7756ZS-E1",
    "99000-79X94-000": "CD-VRM200ZS-E1",
    "9909J-78RM5-N01": "CD-HM022ZSE1",
    "3A108-65T00-000": "CNMV-0159ZS/EU",
    "3A108-65T10-000": "CNMV-0259ZS/AU",
    "99093-55ZR3-N03": "KJ-S103DKZSE1",
    "99000-79W33-000": "ND-ETC3367ZS",
    "99000-79X52-000": "RD-7446ZS",
    "99000-79H98-000": "TS-01142ZS",
    "99000-79H98-001": "TS-01209ZS",
    "9919D-83ST3-R00": "TS-F1640ZSE4",
    "9919D-84SS3-R00": "TS-F1740ZSE3",
    "99000-79BJ0-R00": "TS-G1320FZSE1",
    "9909N-80TY4-N01": "UD-1377ZSE6/WL"
}

st.sidebar.header("⚙️ マスター設定")
if "target_master" not in st.session_state:
    st.session_state.target_master = DEFAULT_MASTER

with st.sidebar.expander("➕ 新規品番の追加", expanded=False):
    new_s = st.text_input("スズキ品番 (例: 99000-79X94-000)")
    new_p = st.text_input("社内部番/パイオニア品番 (例: CD-VRM200ZS-E1)")
    if st.button("追加登録"):
        if new_s and new_p:
            st.session_state.target_master[new_s.strip()] = new_p.strip()
            st.success(f"登録しました: {new_s} -> {new_p}")
        else:
            st.warning("両方の品番を入力してください")

FONT_NAME = "HeiseiKakuGo-W5"
try:
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
except Exception:
    FONT_NAME = "Helvetica"

@st.cache_resource
def load_ocr_reader():
    return easyocr.Reader(['en'], gpu=False)

def clean_code(s):
    s = str(s).upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = s.replace('O', '0').replace('I', '1').replace('Z', '2').replace('S', '5')
    return s

def build_clean_master(master_dict):
    clean_list = []
    for s_code, p_code in master_dict.items():
        clean_list.append({
            's_orig': s_code,
            'p_orig': p_code,
            's_clean': clean_code(s_code),
            'p_clean': clean_code(p_code)
        })
    return clean_list

def match_master(text, clean_master):
    """社内部番重視 ＋ 末尾読み取り漏れ補正 ＋ 対象外仕様(XIJP)完全除外"""
    raw_upper = text.upper()
    c_text = clean_code(text)

    if len(c_text) < 4:
        return None

    # 対象外仕様（XIJP, XI, 84SS3000）を絶対除外
    if "XIJP" in raw_upper or "XI" in raw_upper or "84SS3000" in c_text:
        return None

    for m in clean_master:
        p_cl = m['p_clean']
        s_cl = m['s_clean']

        # 1. 完全・部分一致
        if p_cl in c_text or c_text in p_cl:
            return (m['s_orig'], m['p_orig'])

        # 2. 末尾枝番（E3等）がOCRで切れた場合の補正 (TSF1740ZS 等の前方一致)
        if len(c_text) >= 7 and p_cl.startswith(c_text):
            return (m['s_orig'], m['p_orig'])

        # 3. スズキ品番からの補正
        if s_cl in c_text or (len(c_text) >= 9 and s_cl.startswith(c_text)):
            return (m['s_orig'], m['p_orig'])

    # 類似度判定
    for m in clean_master:
        if SequenceMatcher(None, m['p_clean'], c_text).ratio() > 0.75:
            return (m['s_orig'], m['p_orig'])

    return None

def parse_single_page(ocr_results, clean_master):
    parsed_boxes = []
    full_text = ""
    for item in ocr_results:
        bbox, text, prob = item
        full_text += " " + text
        y_center = (bbox[0][1] + bbox[2][1]) / 2.0
        x_left = bbox[0][0]
        parsed_boxes.append({
            'text': text.strip(),
            'clean': clean_code(text),
            'y': y_center,
            'x': x_left
        })

    date_match = re.search(r'(\d{2,4}/\d{1,2}/\d{1,2})', full_text)
    date_val = date_match.group(1) if date_match else ""

    parsed_boxes.sort(key=lambda b: b['y'])
    rows = []
    for box in parsed_boxes:
        placed = False
        for row in rows:
            if abs(row['y_mean'] - box['y']) < 14:
                row['items'].append(box)
                row['y_mean'] = sum(b['y'] for b in row['items']) / len(row['items'])
                placed = True
                break
        if not placed:
            rows.append({'y_mean': box['y'], 'items': [box]})

    page_items = []

    for row in rows:
        row_items = sorted(row['items'], key=lambda b: b['x'])
        
        hit_master = None
        hit_idx = -1
        hit_x = 0

        for idx, item in enumerate(row_items):
            hit = match_master(item['text'], clean_master)
            if hit:
                hit_master = hit
                hit_idx = idx
                hit_x = item['x']
                break

        if hit_master:
            s_code, p_code = hit_master
            qty = ""
            
            for item in row_items[hit_idx + 1:]:
                if item['x'] - hit_x > 400:
                    continue
                    
                txt = item['text'].replace(',', '').replace(' ', '')
                if txt.isdigit() and 1 <= int(txt) <= 9999:
                    qty = int(txt)
                    break
            
            if qty != "":
                page_items.append({
                    "指示日": date_val,
                    "指示数": qty,
                    "スズキ品番": s_code,
                    "パイオ品番": p_code
                })

    return page_items, date_val

def create_instruction_pdf(items, date_val):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=25
    )
    elements = []
    styles = getSampleStyleSheet()

    title_text = "<b>【作業指示書】 パイオニアラベル貼付作業</b>"
    sub_text = f"指示日: <b>{date_val if date_val else '未特定'}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 発行日時: {datetime.now().strftime('%Y/%m/%d %H:%M')}"
    
    title_style = ParagraphStyle('TitlePDF', parent=styles['Heading1'], fontName=FONT_NAME, fontSize=16, leading=20, alignment=1)
    sub_style = ParagraphStyle('SubPDF', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, leading=14, alignment=1)
    cell_style = ParagraphStyle('CellPDF', parent=styles['Normal'], fontName=FONT_NAME, fontSize=9, leading=12)
    cell_bold = ParagraphStyle('CellBoldPDF', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, leading=13)

    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(sub_text, sub_style))
    elements.append(Spacer(1, 12))

    if not items:
        no_item_msg = "<b>本日、パイオニアラベル貼付の対象品番はありません。</b>"
        elements.append(Paragraph(f"<font color='blue' size=12>{no_item_msg}</font>", sub_style))
    else:
        col_headers = ["No", "スズキ品番", "パイオニア品番", "指示数", "枚数", "作業指示", "完了チェック"]
        table_data = [[Paragraph(f"<b>{h}</b>", cell_bold) for h in col_headers]]

        for idx, item in enumerate(items, 1):
            inst_text = "<font color='green'><b>【シール貼付】</b></font>"
            check_text = "[  ] 完了"
            unit_text = "個"

            table_data.append([
                Paragraph(str(idx), cell_style),
                Paragraph(f"<b><font size=10>{item['スズキ品番']}</font></b>", cell_style),
                Paragraph(str(item['パイオ品番']), cell_style),
                Paragraph(f"<b><font size=11 color='red'>{item['指示数']} {unit_text}</font></b>", cell_style),
                Paragraph("", cell_style),
                Paragraph(inst_text, cell_bold),
                Paragraph(check_text, cell_style)
            ])

        t = Table(table_data, colWidths=[25, 125, 125, 55, 55, 95, 75])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        elements.append(t)

    doc.build(elements)
    buffer.seek(0)
    return buffer

def create_excel(items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "作業時間"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="游ゴシック", size=16, bold=True, underline="single")
    font_header = Font(name="游ゴシック", size=11, bold=True)
    font_body = Font(name="游ゴシック", size=11)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws["A1"] = "パイオニアラベル貼付作業時間"
    ws["A1"].font = font_title

    headers = ["", "作業日", "指示日", "指示数", "スズキ品番", "パイオ品番", "開始時間", "終了時間", "作業時間(分)", "作業人数"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = font_header
        cell.alignment = align_center
        cell.fill = fill_header
        cell.border = border_cell

    max_rows = max(18, len(items))
    for row_idx in range(1, max_rows + 1):
        r = row_idx + 3
        c0 = ws.cell(row=r, column=1, value=row_idx)
        c0.font, c0.alignment, c0.border = font_body, align_center, border_cell
        
        item = items[row_idx - 1] if row_idx - 1 < len(items) else None
        
        c_workdate = ws.cell(row=r, column=2, value="")
        c_date = ws.cell(row=r, column=3, value=item["指示日"] if item else "")
        c_qty = ws.cell(row=r, column=4, value=item["指示数"] if item else "")
        c_suzuki = ws.cell(row=r, column=5, value=item["スズキ品番"] if item else "")
        c_pioneer = ws.cell(row=r, column=6, value=item["パイオ品番"] if item else "")
        c_start = ws.cell(row=r, column=7, value="")
        c_end = ws.cell(row=r, column=8, value="")
        c_calc_time = ws.cell(row=r, column=9, value=f'=IF(AND(G{r}<>"",H{r}<>""),(H{r}-G{r})*24*60,"")')
        c_people = ws.cell(row=r, column=10, value="")

        for c in [c_workdate, c_date, c_start, c_end, c_calc_time, c_people]:
            c.font, c.alignment, c.border = font_body, align_center, border_cell
        c_qty.font, c_qty.alignment, c_qty.border = font_body, align_right, border_cell
        for c in [c_suzuki, c_pioneer]:
            c.font, c.alignment, c.border = font_body, align_left, border_cell

    now = datetime.now()
    year, month = now.year, now.month
    days_in_month = calendar.monthrange(year, month)[1]
    
    for d in range(1, days_in_month + 1):
        ws[f"Z{d}"] = f"{year}/{month:02d}/{d:02d}"

    times = [f"{h:02d}:{m:02d}" for h in range(7, 19) for m in (0, 15, 30, 45)]
    for idx, t in enumerate(times, 1):
        ws[f"AA{idx}"] = t

    for p in range(1, 11):
        ws[f"AB{p}"] = p

    dv_date = DataValidation(type="list", formula1=f"=$Z$1:$Z${days_in_month}", allow_blank=True)
    dv_time = DataValidation(type="list", formula1=f"=$AA$1:$AA${len(times)}", allow_blank=True)
    dv_time.showErrorMessage = False
    dv_people = DataValidation(type="list", formula1="=$AB$1:$AB$10", allow_blank=True)

    ws.add_data_validation(dv_date)
    dv_date.add(f"B4:B{max_rows + 3}")
    ws.add_data_validation(dv_time)
    dv_time.add(f"G4:H{max_rows + 3}")
    ws.add_data_validation(dv_people)
    dv_people.add(f"J4:J{max_rows + 3}")

    column_widths = {"A": 5, "B": 14, "C": 12, "D": 10, "E": 20, "F": 22, "G": 12, "H": 12, "I": 14, "J": 10}
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# --- メインUI ---
st.title("📦 パイオニアラベル貼付 作業指示解析ツール")
st.write("ピッキングリスト（PDF / 画像）を読み込み、**「作業時間記録Excel」** と **「現場用 印刷指示シート(PDF)」** を自動生成します。")

uploaded_files = st.file_uploader("ピッキングリスト（PDF / 画像）をアップロードしてください（複数選択可）", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 解析して指示書を作成", type="primary"):
        with st.spinner("🔍 ページ単位で高精度解析中..."):
            reader = load_ocr_reader()
            clean_master = build_clean_master(st.session_state.target_master)
            all_items = []
            latest_date = ""

            for file in uploaded_files:
                file_bytes = file.read()

                if file.name.lower().endswith(".pdf"):
                    pdf = pdfium.PdfDocument(file_bytes)
                    for page in pdf:
                        image = page.render(scale=2).to_pil()
                        img_np = np.array(image)
                        res = reader.readtext(img_np, detail=1)
                        p_items, p_date = parse_single_page(res, clean_master)
                        all_items.extend(p_items)
                        if p_date:
                            latest_date = p_date
                else:
                    img = cv2.imdecode(np.frombuffer(file_bytes, np.uint8), cv2.IMREAD_COLOR)
                    res = reader.readtext(img, detail=1)
                    p_items, p_date = parse_single_page(res, clean_master)
                    all_items.extend(p_items)
                    if p_date:
                        latest_date = p_date

            st.session_state.parsed_items = all_items
            st.session_state.parsed_date = latest_date

    # 解析結果表示
    if "parsed_items" in st.session_state and st.session_state.parsed_items is not None:
        items = st.session_state.parsed_items
        date_val = st.session_state.parsed_date

        st.divider()
        st.subheader("📋 本日のパイオニアラベル貼付指示（画面確認）")

        if items:
            st.error(f"⚠️ **【シール貼付 作業あり】** 合計 {len(items)} 件の対象品番が検出されました。必ずラベルを貼ってください！")

            cols = st.columns(2)
            for idx, item in enumerate(items):
                col = cols[idx % 2]
                with col:
                    st.markdown(f"""
                    <div style="background-color: #E8F5E9; padding: 15px; border-radius: 10px; border-left: 8px solid #2E7D32; margin-bottom: 12px;">
                        <span style="background-color: #2E7D32; color: white; padding: 3px 8px; border-radius: 5px; font-weight: bold; font-size: 14px;">🏷️ シール貼付あり (No.{idx+1})</span>
                        <h3 style="margin: 8px 0 4px 0; color: #1B5E20;">{item['スズキ品番']}</h3>
                        <p style="margin: 0; font-weight: bold; color: #333;">パイオ品番: {item['パイオ品番']}</p>
                        <p style="margin: 0; font-size: 18px; font-weight: bold; color: #C62828;">指示数量: {item['指示数']} 個</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.success("✅ **【シール貼付 不要】** 対象品番が含まれていません（作業なし）。")

        st.divider()

        excel_buffer = create_excel(items)
        pdf_buffer = create_instruction_pdf(items, date_val)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="🖨️ 【現場用】シール貼付指示書 (PDF) を印刷・ダウンロード",
                data=pdf_buffer,
                file_name=f"作業指示書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                type="primary"
            )
        with col_dl2:
            st.download_button(
                label="📊 【管理用】作業時間記録 Excel をダウンロード",
                data=excel_buffer,
                file_name=f"作業時間_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
