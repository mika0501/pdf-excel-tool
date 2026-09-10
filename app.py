import io
import os
import re
import pandas as pd
import pdfplumber
import pytesseract
import streamlit as st

# Windows 本機 Tesseract 預設路徑備援
if os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

st.set_page_config(page_title="PDF 提單與發票自動整理工具", layout="wide")
st.title("📄 PDF 提單 & 發票自動轉 Excel 工具（升級辨識版）")
st.caption(
    "支援數位 PDF 與紙本掃描提單，自動提取乾淨公司名稱並接續既有 Excel！"
)

# 常見需要過濾的商業組織與物流冗贅詞
LEGAL_AND_BIZ_WORDS = r"\b(Logistics|Consulting|Electronics|Transportation|Transport|Freight|Supply\s+Chain|Worldwide|International|Systems|Services|Solutions|Trading|Inc\.?|LLC\.?|Corp\.?|Corporation|Ltd\.?|Limited|Co\.?|Company|\(USA\)|\(US\))\b"


def clean_company_name(raw_name):
    """通用清洗：去除公司名中的法律後綴與物流字，保留品牌核心簡稱"""
    if not raw_name:
        return ""
    name = re.sub(r"\(.*?\)", "", raw_name)
    name = re.sub(r"[,，\.\-_/]", " ", name)
    name = re.sub(LEGAL_AND_BIZ_WORDS, "", name, flags=re.IGNORECASE)
    words = name.strip().split()
    return " ".join(words) if words else raw_name.strip()


def extract_text_from_pdf(file_bytes):
    """讀取 PDF 文字，若遇掃描圖片頁面則自動呼叫 OCR"""
    pages_text = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            # 如果文字很少（代表是純掃描圖片），自動轉圖片進行 OCR 辨識
            if len(txt.strip()) < 30:
                try:
                    img = page.to_image(resolution=300).original
                    ocr_txt = pytesseract.image_to_string(img, lang="eng")
                    if ocr_txt:
                        txt = ocr_txt
                except Exception:
                    pass
            pages_text.append(txt)
    return "\n--- PAGE BREAK ---\n".join(pages_text)


def parse_pdf(file_bytes):
    full_text = extract_text_from_pdf(file_bytes)

    # 1. INVOICE NUMBER (例如 10161)
    inv_match = re.search(
        r"Invoice\s*no\.?[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_number = inv_match.group(1).strip() if inv_match else ""

    # 2. FREIGHT AMOUNT (Total 後面金額)
    amount_match = re.search(
        r"Total\s*\$?\s*([0-9,]+\.[0-9]{2})", full_text, re.IGNORECASE
    )
    freight_amount = (
        f"${amount_match.group(1).strip()}" if amount_match else ""
    )

    # 3. INVOICE DATE
    date_match = re.search(
        r"Invoice\s*date[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 4. MATERIAL (SKU#)
    sku_match = re.search(
        r"SKU#?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    ) or re.search(r"(CLSAC[A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    material = sku_match.group(1).strip() if sku_match else ""

    # 5. REMARK (修正：只抓 PO#: 後面的純英數號碼，如 PSDELT260902005)
    po_match = re.search(
        r"PO#?[:：]?\s*([A-Za-z0-9]{6,})", full_text, re.IGNORECASE
    )
    po_code = po_match.group(1).strip() if po_match else ""
    remark = f"PO#: {po_code}" if po_code else ""

    # 6. FROM (通用抓取 SHIP FROM 公司名並自動縮減為簡稱)
    from_raw = ""
    sf_match = re.search(
        r"SHIP\s*FROM[\s\S]*?Name[:：]?\s*([^\n\r]+)", full_text, re.IGNORECASE
    )
    if sf_match:
        from_raw = sf_match.group(1).strip()
    from_val = clean_company_name(from_raw)

    # 7. TO (通用抓取 DELIVERY TO 公司名並自動縮減為簡稱)
    to_raw = ""
    dt_match = re.search(
        r"DELIVERY\s*TO[\s\S]*?Name[:：]?\s*([^\n\r]+)",
        full_text,
        re.IGNORECASE,
    )
    if dt_match:
        to_raw = dt_match.group(1).strip()
    to_val = clean_company_name(to_raw)

    return {
        "date": "",
        "INVOICE NUMBER": invoice_number,
        "FREIGHT AMOUNT": freight_amount,
        "INVOICE DATE": invoice_date,
        "MATERIAL": material,
        "REMARK": remark,
        "INBOUND": "",
        "NOTES": "",
        "FROM": from_val,
        "TO": to_val,
    }


# ==================== 操作介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有的 Excel 檔案")
    existing_file = st.file_uploader(
        "如果有先前已經整理好的 Excel 想要累加，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳這次要新增的 PDF 檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 檔案至此（可多選）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    new_records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_pdf(file.read())
        new_records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(new_records)

    # 合併既有 Excel 邏輯
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            if "INVOICE NUMBER" in final_df.columns:
                final_df = final_df.drop_duplicates(
                    subset=["INVOICE NUMBER"], keep="last"
                )
            st.success(
                f"🎉 成功解析 {len(new_records)} 筆 PDF 資料，並已累加至原有的 Excel！目前總計 {len(final_df)} 筆。"
            )
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(new_records)} 筆 PDF 資料！")

    st.subheader("📊 資料預覽")
    st.dataframe(final_df, use_container_width=True)

    # 匯出最新 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Sheet1")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的最新 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="整理結果_最新累加.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
