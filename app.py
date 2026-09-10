import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="PDF 提單與發票自動整理工具", page_icon="📄", layout="wide"
)
st.title("📄 PDF 提單 & 發票自動轉 Excel 工具")
st.caption(
    "純文字極速解析版：自動過濾 REMARK 雜訊，FROM/TO 僅取第一個主詞，支援 Excel 累加！"
)


def get_first_word(text):
    """只抓取第一個主要單詞（過濾標點與數字門牌）"""
    if not text:
        return ""
    # 去除前後標點與數字
    cleaned = re.sub(r"^[\d\W]+", "", text.strip())
    # 切割取第一個單詞
    words = cleaned.split()
    if words:
        # 如果第一個字剛好是常見品牌詞，直接回傳
        return words[0].strip()
    return ""


def parse_pdf(file_bytes):
    """解析單一 PDF 全文純文字"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # 1. INVOICE NUMBER
    inv_match = re.search(
        r"Invoice\s*no\.?[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_number = inv_match.group(1).strip() if inv_match else ""

    # 2. FREIGHT AMOUNT (Total 後面的金額)
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

    # 4. MATERIAL (SKU)
    sku_match = re.search(
        r"SKU#?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    ) or re.search(r"(CLSAC[A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    material = sku_match.group(1).strip() if sku_match else ""

    # 5. REMARK (只精準抓 PO 號碼本體，去掉後面的 Inbound&Freight 雜訊)
    po_match = re.search(
        r"PO#?[:：]?\s*([A-Za-z0-9]{6,})", full_text, re.IGNORECASE
    )
    po_code = po_match.group(1).strip() if po_match else ""
    remark = f"PO#: {po_code}" if po_code else ""

    # ================= 6. FROM (只抓第一個字) =================
    from_val = ""
    # 優先從發票明細抓 "from ... to ..."
    route_match = re.search(
        r"from\s+(.*?)\s+to\s+(.*?)(?:\s+on|\s+Dock|\n|$)",
        full_text,
        re.IGNORECASE,
    )
    if route_match:
        from_val = get_first_word(route_match.group(1))
    else:
        # 備援：從 SHIP FROM Name 抓
        sf_match = re.search(
            r"SHIP\s*FROM[\s\S]*?Name[:：]?\s*([^\n\r]+)",
            full_text,
            re.IGNORECASE,
        )
        if sf_match:
            from_val = get_first_word(sf_match.group(1))

    # ================= 7. TO (只抓第一個字) =================
    to_val = ""
    if route_match:
        to_val = get_first_word(route_match.group(2))
    else:
        # 備援：從 DELIVERY TO Name 抓
        dt_match = re.search(
            r"DELIVERY\s*TO[\s\S]*?Name[:：]?\s*([^\n\r]+)",
            full_text,
            re.IGNORECASE,
        )
        if dt_match:
            to_val = get_first_word(dt_match.group(1))

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


# ==================== 畫面介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有的 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將資料接續累加至原有的 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳這次要處理的 PDF 檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 檔案至此（支援多選）",
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

    # 合併既有 Excel
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            if "INVOICE NUMBER" in final_df.columns:
                final_df = final_df.drop_duplicates(
                    subset=["INVOICE NUMBER"], keep="last"
                )
            st.success(
                f"🎉 成功解析 {len(new_records)} 筆 PDF 資料，並已累加至原有的 Excel！目前累積共 {len(final_df)} 筆。"
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
