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
st.title("📄 PDF 提單 & 發票自動轉 Excel 工具（精準雙重解析版）")
st.caption("支援直接從發票明細與提單雙重擷取，解決掃描檔辨識問題！")

LEGAL_AND_BIZ_WORDS = r"\b(Logistics|Consulting|Electronics|Transportation|Transport|Freight|Supply\s+Chain|Worldwide|International|Systems|Services|Solutions|Trading|Inc\.?|LLC\.?|Corp\.?|Corporation|Ltd\.?|Limited|Co\.?|Company|\(USA\)|\(US\))\b"


def clean_company_name(raw_name):
    """去除公司名中的地址數字、法律後綴與物流冗贅詞，只留核心品牌名"""
    if not raw_name:
        return ""
    # 去除前後數字門牌（例如 601 Delta 或 Delta 601）
    name = re.sub(r"^\d+\s*", "", raw_name.strip())
    name = re.sub(r"\s*\d+$", "", name.strip())
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[,，\.\-_/]", " ", name)
    name = re.sub(LEGAL_AND_BIZ_WORDS, "", name, flags=re.IGNORECASE)
    words = name.strip().split()
    return " ".join(words) if words else raw_name.strip()


def extract_all_text(file_bytes):
    """提取純文字 + 掃描頁強制 OCR"""
    full_text_list = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            full_text_list.append(txt)
            # 如果頁面文字很少，或者是提單頁面，進行 OCR
            if len(txt.strip()) < 50 or "LADING" in txt.upper():
                try:
                    img = page.to_image(resolution=300).original
                    ocr_txt = pytesseract.image_to_string(img, lang="eng")
                    if ocr_txt:
                        full_text_list.append(ocr_txt)
                except Exception:
                    pass
    return "\n".join(full_text_list)


def parse_pdf(file_bytes):
    full_text = extract_all_text(file_bytes)

    # 1. INVOICE NUMBER
    inv_match = re.search(
        r"Invoice\s*no\.?[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_number = inv_match.group(1).strip() if inv_match else ""

    # 2. FREIGHT AMOUNT
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

    # 5. REMARK (只抓純代碼 PSDELT260902005)
    po_match = re.search(
        r"PO#?[:：]?\s*([A-Za-z0-9]{6,})", full_text, re.IGNORECASE
    )
    po_code = po_match.group(1).strip() if po_match else ""
    remark = f"PO#: {po_code}" if po_code else ""

    # ================= 6 & 7. FROM 與 TO 雙重解析 =================
    from_val = ""
    to_val = ""

    # 【策略 A】：直接從發票的 Description 抓（最精準！例如：from 601 Delta ... to 1991 Peak Smart）
    route_match = re.search(
        r"from\s+(?:\d+\s+)?([A-Za-z\s]+?)(?:,|\s+Plano|\s+Lewisville|\s+to|\n|$)\s+to\s+(?:\d+\s+)?([A-Za-z\s]+?)(?:\s+on|\s+Dock|\n|$)",
        full_text,
        re.IGNORECASE,
    )
    if route_match:
        from_val = clean_company_name(route_match.group(1))
        to_val = clean_company_name(route_match.group(2))

    # 【策略 B】：若策略 A 未命中，從提單區塊抓取
    if not from_val:
        sf_match = re.search(
            r"SHIP\s*FROM[\s\S]*?Name[:：]?\s*([^\n\r]+)", full_text, re.IGNORECASE
        )
        if sf_match:
            from_val = clean_company_name(sf_match.group(1))

    if not to_val:
        dt_match = re.search(
            r"DELIVERY\s*TO[\s\S]*?Name[:：]?\s*([^\n\r]+)",
            full_text,
            re.IGNORECASE,
        )
        if dt_match:
            to_val = clean_company_name(dt_match.group(1))

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


# ==================== 介面呈現 ====================
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

    # 匯出 Excel
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
